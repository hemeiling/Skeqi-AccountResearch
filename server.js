/**
 * Account Plan Generator — Local Proxy Server
 * ─────────────────────────────────────────────
 * Keeps the Anthropic API key on the server so the browser never touches it.
 * Supports web search (web_search_20250305) with a proper agentic loop.
 *
 * Run:  node server.js
 * Port: 3000 (change via PORT env var)
 */

'use strict';

const express = require('express');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');
const crypto  = require('crypto');
require('dotenv').config();

const app  = express();
const PORT = process.env.PORT || 3000;

// Where persistent data (logs/, reports/, research-cache/) lives. Defaults to
// the project directory for local dev; on Render etc. point this at the
// mounted persistent disk (e.g. DATA_DIR=/data) so it survives redeploys.
const DATA_DIR = process.env.DATA_DIR || __dirname;
fs.mkdirSync(DATA_DIR, { recursive: true });

// ── File logging ────────────────────────────────────────────────────────────
// Mirrors console output to logs/server.log so failures can be diagnosed
// without needing the terminal that launched the server.
const LOG_DIR  = path.join(DATA_DIR, 'logs');
const LOG_FILE = path.join(LOG_DIR, 'server.log');
fs.mkdirSync(LOG_DIR, { recursive: true });
const logStream = fs.createWriteStream(LOG_FILE, { flags: 'a' });
for (const level of ['log', 'warn', 'error']) {
  const orig = console[level].bind(console);
  console[level] = (...args) => {
    orig(...args);
    const line = args.map(a => typeof a === 'string' ? a : JSON.stringify(a)).join(' ');
    logStream.write(`[${new Date().toISOString()}] [${level}] ${line}\n`);
  };
}

// ── Auth ────────────────────────────────────────────────────────────────────
// Optional shared-credential gate (HTTP Basic Auth) for the whole app. Set
// BASIC_AUTH_USER / BASIC_AUTH_PASS to enable; leave unset to run open (e.g.
// for local dev). /health stays open so uptime checks don't need credentials.
const BASIC_AUTH_USER = process.env.BASIC_AUTH_USER;
const BASIC_AUTH_PASS = process.env.BASIC_AUTH_PASS;

function safeEqual(a, b) {
  const bufA = Buffer.from(String(a));
  const bufB = Buffer.from(String(b));
  if (bufA.length !== bufB.length) {
    crypto.timingSafeEqual(bufA, Buffer.alloc(bufA.length)); // keep timing constant
    return false;
  }
  return crypto.timingSafeEqual(bufA, bufB);
}

if (BASIC_AUTH_USER && BASIC_AUTH_PASS) {
  app.use((req, res, next) => {
    if (req.path === '/health') return next();
    const header = req.headers.authorization || '';
    if (header.startsWith('Basic ')) {
      const decoded = Buffer.from(header.slice(6), 'base64').toString('utf8');
      const sep = decoded.indexOf(':');
      const user = sep === -1 ? decoded : decoded.slice(0, sep);
      const pass = sep === -1 ? '' : decoded.slice(sep + 1);
      if (safeEqual(user, BASIC_AUTH_USER) && safeEqual(pass, BASIC_AUTH_PASS)) {
        return next();
      }
    }
    res.set('WWW-Authenticate', 'Basic realm="Account Plan Generator"');
    res.status(401).send('Authentication required.');
  });
}

// ── Middleware ──────────────────────────────────────────────────────────────
app.use(cors({ origin: '*' }));          // allow the HTML file opened locally
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname))); // serves index.html at /
app.use('/vendor/pptxgenjs', express.static(path.join(__dirname, 'node_modules/pptxgenjs/dist'))); // local PptxGenJS, no CDN needed
app.use('/vendor/docx', express.static(path.join(__dirname, 'node_modules/docx/dist'))); // local docx, no CDN needed
app.use('/vendor/jspdf', express.static(path.join(__dirname, 'node_modules/jspdf/dist'))); // local jsPDF, no CDN needed
app.use('/vendor/html2canvas', express.static(path.join(__dirname, 'node_modules/html2canvas/dist'))); // local html2canvas, no CDN needed

// ── Config ──────────────────────────────────────────────────────────────────
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const DEFAULT_MODEL     = 'claude-sonnet-4-6';
const MAX_TOKENS        = 16000;
const MAX_TOOL_ITER     = 8;   // safety cap for agentic web-search loop

function getApiKey() {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) throw new Error('ANTHROPIC_API_KEY is not set in .env');
  return key;
}

// ── Pricing config (for cost estimation) ────────────────────────────────────
// Rates are USD per 1,000,000 tokens. Override by editing pricing.json or by
// setting the ANTHROPIC_PRICING_JSON env var to a JSON string with the same shape.
const PRICING_FILE = path.join(__dirname, 'pricing.json');
function loadPricing() {
  try {
    if (process.env.ANTHROPIC_PRICING_JSON) return JSON.parse(process.env.ANTHROPIC_PRICING_JSON);
    return JSON.parse(fs.readFileSync(PRICING_FILE, 'utf8'));
  } catch (e) {
    console.warn('[pricing] could not load pricing.json, using built-in defaults:', e.message);
    return { _default: { input: 3.0, output: 15.0, cacheWrite: 3.75, cacheRead: 0.3 } };
  }
}
const PRICING = loadPricing();
function getModelPricing(model) {
  return PRICING[model] || PRICING._default || { input: 3.0, output: 15.0, cacheWrite: 3.75, cacheRead: 0.3 };
}
// usage: { inputTokens, outputTokens, cacheCreationInputTokens, cacheReadInputTokens }
function estimateCost(usage, model) {
  const p = getModelPricing(model);
  return (
    ((usage.inputTokens               || 0) / 1e6) * (p.input ?? 0) +
    ((usage.outputTokens              || 0) / 1e6) * (p.output ?? 0) +
    ((usage.cacheCreationInputTokens  || 0) / 1e6) * (p.cacheWrite ?? p.input ?? 0) +
    ((usage.cacheReadInputTokens      || 0) / 1e6) * (p.cacheRead  ?? p.input ?? 0)
  );
}

// ── Anthropic call with agentic web-search loop ─────────────────────────────
async function callAnthropic({ model, prompt, useSearch }) {
  const key = getApiKey();

  const headers = {
    'Content-Type':      'application/json',
    'x-api-key':         key,
    'anthropic-version': '2023-06-01',
  };
  if (useSearch) headers['anthropic-beta'] = 'web-search-2025-03-05';

  const tools   = useSearch ? [{ type: 'web_search_20250305', name: 'web_search' }] : [];
  let messages  = [{ role: 'user', content: prompt }];
  let finalText = '';
  let iterations = 0;
  const usage = { inputTokens: 0, outputTokens: 0, cacheCreationInputTokens: 0, cacheReadInputTokens: 0 };

  while (iterations < MAX_TOOL_ITER) {
    iterations++;

    const body = { model, max_tokens: MAX_TOKENS, messages };
    if (tools.length) body.tools = tools;

    console.log(`  [Anthropic] iter=${iterations} msgs=${messages.length}`);

    const resp = await fetch(ANTHROPIC_API_URL, {
      method:  'POST',
      headers,
      body:    JSON.stringify(body),
    });

    if (!resp.ok) {
      const errBody = await resp.json().catch(() => ({}));
      throw new Error(errBody?.error?.message || `HTTP ${resp.status}`);
    }

    const data       = await resp.json();
    const stopReason = data.stop_reason;
    const content    = data.content || [];

    if (data.usage) {
      usage.inputTokens              += data.usage.input_tokens || 0;
      usage.outputTokens             += data.usage.output_tokens || 0;
      usage.cacheCreationInputTokens += data.usage.cache_creation_input_tokens || 0;
      usage.cacheReadInputTokens     += data.usage.cache_read_input_tokens || 0;
    }

    console.log(`  [Anthropic] stop_reason=${stopReason} blocks=[${content.map(b => b.type).join(',')}] usage=${JSON.stringify(data.usage||{})}`);

    // Collect text from this turn
    const textBlocks = content.filter(b => b.type === 'text');
    if (textBlocks.length) finalText = textBlocks.map(b => b.text).join('\n').trim();

    if (stopReason === 'end_turn' || stopReason === 'max_tokens') {
      if (!finalText) throw new Error(`No text in response (stop_reason=${stopReason})`);
      console.log(`  [Anthropic] done — text length: ${finalText.length}`);
      return { text: finalText, usage, apiCalls: iterations, model };
    }

    // Handle tool_use — web_search is server-side but we still need to relay
    if (stopReason === 'tool_use') {
      const toolUseBlocks = content.filter(b => b.type === 'tool_use');
      if (!toolUseBlocks.length) throw new Error('stop_reason=tool_use but no tool_use blocks');

      messages.push({ role: 'assistant', content });
      const toolResults = toolUseBlocks.map(tu => ({
        type:        'tool_result',
        tool_use_id: tu.id,
        content:     tu.input ? JSON.stringify(tu.input) : '{}',
      }));
      messages.push({ role: 'user', content: toolResults });
      continue;
    }

    // Unknown stop reason
    if (finalText) return { text: finalText, usage, apiCalls: iterations, model };
    throw new Error(`Unexpected stop_reason: ${stopReason}`);
  }

  throw new Error(`Agentic loop exceeded ${MAX_TOOL_ITER} iterations`);
}

// ── Route: health check ─────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  const hasKey = !!process.env.ANTHROPIC_API_KEY;
  res.json({
    status:  'ok',
    model:   DEFAULT_MODEL,
    hasKey,
    keyHint: hasKey ? process.env.ANTHROPIC_API_KEY.slice(0, 14) + '…' : null,
  });
});

// ── Route: test connection ──────────────────────────────────────────────────
app.post('/api/test', async (req, res) => {
  try {
    const { text } = await callAnthropic({
      model:     DEFAULT_MODEL,
      prompt:    'Reply with exactly: "Connected · 连接成功"',
      useSearch: false,
    });
    res.json({ ok: true, text });
  } catch (err) {
    console.error('[/api/test]', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── Route: single research step ─────────────────────────────────────────────
// Body: { prompt: string, useSearch?: boolean, model?: string }
// Response: { text, usage: {inputTokens,outputTokens,cacheCreationInputTokens,cacheReadInputTokens}, apiCalls, model, costUSD }
app.post('/api/research', async (req, res) => {
  const { prompt, useSearch = true, model = DEFAULT_MODEL } = req.body;

  if (!prompt || typeof prompt !== 'string') {
    return res.status(400).json({ ok: false, error: 'prompt is required' });
  }

  console.log(`\n[/api/research] model=${model} useSearch=${useSearch} promptLen=${prompt.length}`);

  try {
    const { text, usage, apiCalls, model: usedModel } = await callAnthropic({ model, prompt, useSearch });
    const costUSD = estimateCost(usage, usedModel);
    res.json({ ok: true, text, usage, apiCalls, model: usedModel, costUSD });
  } catch (err) {
    console.error('[/api/research] error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── Route: generate full report (all 6 steps, returns structured JSON) ───────
// Body: { target, turl, seller, surl, model?, useSearch? }
// Response: { ok, od, fd, sd, nd, cd, sal, errors }
app.post('/api/generate-report', async (req, res) => {
  const {
    target, turl = '', seller, surl = '',
    model = DEFAULT_MODEL,
    useSearch = true,
  } = req.body;

  if (!target || !seller) {
    return res.status(400).json({ ok: false, error: 'target and seller are required' });
  }

  console.log(`\n[/api/generate-report] target="${target}" seller="${seller}" model=${model} useSearch=${useSearch}`);

  const results  = {};
  const errors   = {};

  // Build prompts inline (same as front-end) so the server is self-contained
  const steps = [
    {
      key: 'od',
      prompt: `You are a bilingual B2B research analyst. Research "${target}" (website: ${turl || 'search for it'}). Seller: "${seller}" (${surl || 'N/A'}).
Use web search. For unverifiable fields use exactly: "Data not available". Never invent figures.
Return ONLY valid JSON, no markdown fences:
{"companyName":"official EN name","companyNameZh":"官方中文名","ticker":"e.g. SZ:300750 or blank",
"profile":{"founded":"year & city","foundedZh":"成立","headquarters":"city, country","headquartersZh":"总部","ceo":"name","ceoZh":"CEO","employees":"~N,000 (YYYY)","employeesZh":"员工","revenue":"$XB (FY20XX)","revenueZh":"营收","countries":"60+ countries","countriesZh":"运营范围","website":"${turl || ''}"},
"segments":[{"pct":"~X%","name":"segment","nameZh":"分部","desc":"one sentence","descZh":"一句话"}],
"quote":"CEO quote or tagline","quoteZh":"语录",
"stats":[{"value":"figure","label":"metric EN","labelZh":"指标"}],
"sources":[{"name":"source","url":"URL or N/A","type":"primary|secondary|ai-inferred","status":"found|partial|unavailable","excerpt":"what came from here","excerptZh":"来源说明"}],
"dataReliability":"high|medium|low","reliabilityNote":"note EN","reliabilityNoteZh":"说明"}
3-4 segments, 4 stats. PURE JSON ONLY.`,
    },
    {
      key: 'fd',
      prompt: `Research financials for "${target}" (${turl || ''}). Sources: company IR → Yahoo Finance → SEC EDGAR 10-K → press releases.
Use "Data not available" for unverifiable. Return ONLY valid JSON, no markdown:
{"kpis":[{"value":"figure","label":"EN","labelZh":"中文","sub":"FY period","subZh":"财年","color":"#hex","sourceName":"source","sourceUrl":"URL or N/A"}],
"context":"2-3 sentence financial narrative (EN)","contextZh":"财务分析（中文）",
"industry":"industry (EN)","industryZh":"行业","market":"size+growth with source","marketZh":"市场规模",
"trends":"3-4 trends plain text EN","trendsZh":"趋势中文",
"sources":[{"name":"source","url":"URL or N/A","type":"primary|secondary|ai-inferred","status":"found|partial|unavailable","excerpt":"what came from here","excerptZh":"来源说明"}],
"dataReliability":"high|medium|low","reliabilityNote":"note","reliabilityNoteZh":"说明"}
Colors: #0D1B3E #0369A1 #1E2F5C #D4860A. 4 KPIs. PURE JSON ONLY.`,
    },
    {
      key: 'sd',
      prompt: `Research "${target}" (${turl || ''}) strategic priorities and technology investments.
Sources: website → annual report/10-K → earnings transcripts → press releases → industry news.
Mark inferred [AI-inferred]. Return ONLY valid JSON:
{"priorities":[{"title":"EN","titleZh":"中文","body":"2-3 sentences (EN)","bodyZh":"中文","sourceName":"source","sourceUrl":"URL or N/A","confidence":"high|medium|low"}],
"techNumbers":[{"value":"figure","desc":"EN","descZh":"中文","sourceName":"source","sourceUrl":"URL or N/A"}],
"techInvestments":[{"category":"EN","categoryZh":"中文","items":["EN"],"itemsZh":["中文"]}],
"signals":"buying signals (EN)","signalsZh":"购买信号",
"sources":[{"name":"source","url":"URL or N/A","type":"primary|secondary|ai-inferred","status":"found|partial|unavailable","excerpt":"what","excerptZh":"说明"}],
"dataReliability":"high|medium|low","reliabilityNote":"note","reliabilityNoteZh":"说明"}
4 priorities, 5 tech numbers, 3 tech categories. PURE JSON ONLY.`,
    },
    {
      key: 'nd',
      prompt: `Find recent significant news about "${target}" (${turl || ''}) from the past 6 months.
Sources: company press page → Reuters/Bloomberg/WSJ → Google News → SEC 8-K.
NEVER fabricate. Return ONLY valid JSON:
{"news":[{"date":"Month YYYY","headline":"EN","headlineZh":"中文","summary":"1-2 sentences EN","summaryZh":"中文","source":"publication","sourceUrl":"URL or N/A","category":"earnings|strategy|technology|leadership|partnership|regulatory|other","verified":true}],
"sources":[{"name":"source","url":"URL or N/A","type":"primary|secondary","status":"found|partial|unavailable","excerpt":"found","excerptZh":"说明"}]}
PURE JSON ONLY. Do not fabricate.`,
    },
    {
      key: 'cd',
      prompt: `Analyze "${target}" (${turl || ''}) from "${seller}" account team perspective.
Research: website → 10-K risk section → LinkedIn → industry news. Mark inferred [AI-inferred].
Return ONLY valid JSON:
{"competitors":[{"name":"competitor","strengths":"EN","strengthsZh":"中文","advantage":"where ${seller} wins EN","advantageZh":"中文","threat":"HIGH|MEDIUM|LOW","sourceNote":"basis"}],
"differentiation":"2-3 sentences EN","differentiationZh":"中文",
"pains":[{"title":"EN","titleZh":"中文","body":"2-3 sentences EN","bodyZh":"中文","urgency":"HIGH|MED-HIGH|MEDIUM","sourceNote":"basis","confidence":"high|medium|low"}],
"sources":[{"name":"source","url":"URL or N/A","type":"primary|secondary|ai-inferred","status":"found|partial|unavailable","excerpt":"what","excerptZh":"说明"}],
"dataReliability":"high|medium|low","reliabilityNote":"note","reliabilityNoteZh":"说明"}
4 competitors, 6 pain points. PURE JSON ONLY.`,
    },
    {
      key: 'sal',
      prompt: `Sales playbook for "${seller}" (${surl || ''}) selling into "${target}" (${turl || ''}).
Stakeholders: company leadership page → LinkedIn → press releases → 10-K officers.
Unverifiable names: use role + "[Name not publicly listed]". Return ONLY valid JSON:
{"stakeholders":[{"name":"Name or Role [Name not publicly listed]","title":"EN","titleZh":"中文","priority":"Primary|Champion|Influencer|Awareness|Expansion","msg":"engagement note EN","msgZh":"中文","color":"#hex","sourceName":"source","sourceUrl":"URL or N/A","verified":true}],
"valueProps":[{"title":"${seller} solution","titleZh":"方案","match":"need EN","matchZh":"需求","body":"2-3 sentences EN","bodyZh":"中文","segment":"biz segment"}],
"phases":[{"phase":"PHASE 1","phaseZh":"第一阶段","period":"Month 1-2","periodZh":"第1-2个月","name":"phase name EN","nameZh":"阶段名","actions":["action EN"],"actionsZh":["行动中文"]}],
"goal":"goal + revenue target EN","goalZh":"目标",
"questions":"5-6 questions plain text EN","questionsZh":"5-6问题",
"sources":[{"name":"source","url":"URL or N/A","type":"primary|secondary|ai-inferred","status":"found|partial|unavailable","excerpt":"what","excerptZh":"说明"}],
"dataReliability":"high|medium|low","reliabilityNote":"note","reliabilityNoteZh":"说明"}
Colors: #0D1B3E #0369A1 #D4860A #1E2F5C #2D4278. 5-6 stakeholders, 4 value props, 3 phases. PURE JSON ONLY.`,
    },
  ];

  const usage = { model, apiCalls: 0, inputTokens: 0, outputTokens: 0, cacheCreationInputTokens: 0, cacheReadInputTokens: 0, costUSD: 0, bySection: {} };

  for (const step of steps) {
    try {
      console.log(`  [step:${step.key}] running…`);
      const { text: raw, usage: stepUsage, apiCalls } = await callAnthropic({ model, prompt: step.prompt, useSearch });
      const stepCost = estimateCost(stepUsage, model);
      usage.apiCalls += apiCalls;
      usage.inputTokens += stepUsage.inputTokens;
      usage.outputTokens += stepUsage.outputTokens;
      usage.cacheCreationInputTokens += stepUsage.cacheCreationInputTokens;
      usage.cacheReadInputTokens += stepUsage.cacheReadInputTokens;
      usage.costUSD += stepCost;
      usage.bySection[step.key] = { apiCalls, ...stepUsage, costUSD: stepCost };

      const json = await parseAnthropicJson(raw, model, step.key);
      results[step.key] = json;
      console.log(`  [step:${step.key}] ✅ ok`);
    } catch (err) {
      console.error(`  [step:${step.key}] ❌`, err.message);
      errors[step.key] = err.message;
      results[step.key] = null;  // front-end will use fallback
    }
  }

  const successCount = Object.values(results).filter(v => v !== null).length;
  console.log(`[/api/generate-report] done — ${successCount}/${steps.length} succeeded — cost≈$${usage.costUSD.toFixed(4)}`);

  res.json({ ok: true, results, errors, usage });
});

// ── JSON parse helper ────────────────────────────────────────────────────────
// Escape stray unescaped " inside JSON string values (e.g. Chinese text quoting a
// term with straight double-quotes: 为"特斯拉"). Tracks object/array nesting so it
// knows whether a string is a key (ends only at ":") or a value (ends at , } ]).
function repairJSON(s) {
  let out = '', inStr = false, esc = false, isKeyStr = false, awaitingKey = false;
  const stack = [];
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inStr) {
      if (esc) { out += c; esc = false; continue; }
      if (c === '\\') { out += c; esc = true; continue; }
      if (c === '"') {
        let j = i + 1; while (j < s.length && /\s/.test(s[j])) j++;
        const next = s[j];
        const isEnd = isKeyStr
          ? (next === ':' || next === undefined)
          : (next === undefined || ',}]'.includes(next));
        if (isEnd) { inStr = false; out += c; if (isKeyStr) awaitingKey = false; }
        else out += '\\"';
        continue;
      }
      out += c;
      continue;
    }
    out += c;
    switch (c) {
      case '{': stack.push('obj'); awaitingKey = true; break;
      case '[': stack.push('arr'); break;
      case '}': case ']': stack.pop(); break;
      case ',': if (stack[stack.length - 1] === 'obj') awaitingKey = true; break;
      case ':': awaitingKey = false; break;
      case '"': inStr = true; isKeyStr = (stack[stack.length - 1] === 'obj' && awaitingKey); break;
    }
  }
  return out;
}

async function parseAnthropicJson(raw, model, stepKey) {
  try {
    return parseJSON(raw);
  } catch (firstErr) {
    console.warn(`  [step:${stepKey}] parse failed, retrying as clean JSON — ${firstErr.message}`);
    const repairPrompt = `The previous response was intended to be pure JSON, but it was malformed. Please return only valid JSON and nothing else, using the same content where possible.\n\n${raw}`;
    const { text: repairedRaw } = await callAnthropic({ model, prompt: repairPrompt, useSearch: false });
    try {
      return parseJSON(repairedRaw);
    } catch (secondErr) {
      throw new Error(`JSON parse failed: ${firstErr.message}; retry also failed: ${secondErr.message}`);
    }
  }
}

function parseJSON(txt) {
  let clean = txt
    .replace(/```json[\s\S]*?```/g, m => m.slice(m.indexOf('\n') + 1, m.lastIndexOf('```')))
    .replace(/```json|```/g, '')
    .trim();
  const first = clean.indexOf('{');
  const last  = clean.lastIndexOf('}');
  if (first === -1 || last === -1) throw new Error('No JSON object in response');
  clean = clean.slice(first, last + 1);
  try {
    return JSON.parse(clean);
  } catch (e) {
    return JSON.parse(repairJSON(clean));
  }
}

// ── Report repository ───────────────────────────────────────────────────────
// File-based store so generated reports can be saved, searched, versioned, and
// reloaded later instead of being re-researched from scratch.
const REPORTS_DIR = path.join(DATA_DIR, 'reports');
const REPORTS_DATA_DIR = path.join(REPORTS_DIR, 'data');
const REPORTS_INDEX_FILE = path.join(REPORTS_DIR, 'index.json');
fs.mkdirSync(REPORTS_DATA_DIR, { recursive: true });
if (!fs.existsSync(REPORTS_INDEX_FILE)) fs.writeFileSync(REPORTS_INDEX_FILE, '[]');

function slugify(s) {
  return String(s || '').trim().toLowerCase()
    .replace(/[^a-z0-9一-鿿]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'na';
}

function readReportIndex() {
  try { return JSON.parse(fs.readFileSync(REPORTS_INDEX_FILE, 'utf8')); }
  catch { return []; }
}

function writeReportIndex(list) {
  fs.writeFileSync(REPORTS_INDEX_FILE, JSON.stringify(list, null, 2));
}

// Body: full report object { target, turl, seller, surl, od, fd, sd, nd, cd, sal, results? }
// Response: { ok, id, key, version }
app.post('/api/reports', (req, res) => {
  const D = req.body || {};
  if (!D.target || !D.seller) {
    return res.status(400).json({ ok: false, error: 'target and seller are required' });
  }

  const key = `${slugify(D.target)}__${slugify(D.seller)}`;
  const index = readReportIndex();
  const version = index.filter(r => r.key === key).reduce((max, r) => Math.max(max, r.version), 0) + 1;
  const id = `${key}__v${version}__${Date.now()}`;
  const createdAt = new Date().toISOString();

  fs.writeFileSync(path.join(REPORTS_DATA_DIR, `${id}.json`), JSON.stringify({ ...D, id, key, version, createdAt }, null, 2));

  index.unshift({
    id, key, version, createdAt,
    target: D.target, targetZh: D.od?.companyNameZh || '', turl: D.turl || '',
    companyName: D.od?.companyName || '', companyNameZh: D.od?.companyNameZh || '',
    reportType: 'Executive Account Plan · 高管账户计划',
    seller: D.seller, surl: D.surl || '',
    sections: ['od', 'fd', 'sd', 'nd', 'cd', 'sal'].reduce((acc, k) => {
      const status = D._sectionStatus && D._sectionStatus[k];
      acc[k] = status === 'success' ? 'live' : (status === 'fallback' ? 'fallback' : (D[k] ? 'live' : 'fallback'));
      return acc;
    }, {}),
    usage: D._usage ? {
      model:        D._usage.model || '',
      apiCalls:     D._usage.apiCalls || 0,
      inputTokens:  D._usage.inputTokens || 0,
      outputTokens: D._usage.outputTokens || 0,
      totalTokens:  (D._usage.inputTokens || 0) + (D._usage.outputTokens || 0),
      costUSD:      D._usage.costUSD || 0,
    } : null,
  });
  writeReportIndex(index);

  console.log(`[/api/reports] saved id=${id} key=${key} version=${version}`);
  res.json({ ok: true, id, key, version });
});

// Query: ?q=search term (matches target/targetZh/seller/turl/surl, case-insensitive)
// Response: { ok, reports: [...metadata, latest-first grouped by key with version history] }
app.get('/api/reports', (req, res) => {
  const q = String(req.query.q || '').trim().toLowerCase();
  let index = readReportIndex();
  if (q) {
    index = index.filter(r =>
      [r.target, r.targetZh, r.seller, r.turl, r.surl].some(v => (v || '').toLowerCase().includes(q)));
  }
  res.json({ ok: true, reports: index });
});

// Response: { ok, report: <full saved report> }
app.get('/api/reports/:id', (req, res) => {
  const file = path.join(REPORTS_DATA_DIR, `${req.params.id}.json`);
  if (!fs.existsSync(file)) return res.status(404).json({ ok: false, error: 'Report not found' });
  res.json({ ok: true, report: JSON.parse(fs.readFileSync(file, 'utf8')) });
});

app.delete('/api/reports/:id', (req, res) => {
  const file = path.join(REPORTS_DATA_DIR, `${req.params.id}.json`);
  if (!fs.existsSync(file)) return res.status(404).json({ ok: false, error: 'Report not found' });
  fs.unlinkSync(file);
  writeReportIndex(readReportIndex().filter(r => r.id !== req.params.id));
  console.log(`[/api/reports] deleted id=${req.params.id}`);
  res.json({ ok: true });
});

// ── Research cache ───────────────────────────────────────────────────────────
// Per-section research results (company overview, financials, competitive
// intelligence, etc.) are cached on disk keyed by section type + target/seller
// so future report generations can reuse them instead of calling Claude again.
const CACHE_DIR = path.join(DATA_DIR, 'research-cache');
fs.mkdirSync(CACHE_DIR, { recursive: true });

// Time-to-live per section type — fast-moving content (news) expires sooner
// than slow-moving content (company overview, strategy).
const CACHE_TTL_MS = {
  s1: 30 * 24 * 60 * 60 * 1000, // company overview
  s2: 7  * 24 * 60 * 60 * 1000, // financials
  s3: 14 * 24 * 60 * 60 * 1000, // strategy & technology
  s4: 1  * 24 * 60 * 60 * 1000, // recent news
  s5: 7  * 24 * 60 * 60 * 1000, // competitive & pain points
  s6: 14 * 24 * 60 * 60 * 1000, // stakeholders & sales
  default: 7 * 24 * 60 * 60 * 1000,
};

function cacheFile(type, key) {
  return path.join(CACHE_DIR, `${slugify(type)}__${slugify(key)}.json`);
}

// Response: { ok, found, data?, costUSD?, cachedAt?, ageMs?, expired? }
app.get('/api/research-cache/:type/:key', (req, res) => {
  const { type, key } = req.params;
  const file = cacheFile(type, key);
  if (!fs.existsSync(file)) return res.json({ ok: true, found: false });
  try {
    const entry = JSON.parse(fs.readFileSync(file, 'utf8'));
    const ttl = CACHE_TTL_MS[type] || CACHE_TTL_MS.default;
    const ageMs = Date.now() - (entry.cachedAt || 0);
    if (ageMs > ttl) return res.json({ ok: true, found: false, expired: true });
    res.json({ ok: true, found: true, data: entry.data, costUSD: entry.costUSD || 0, cachedAt: entry.cachedAt, ageMs });
  } catch (err) {
    console.error('[/api/research-cache] read error:', err.message);
    res.json({ ok: true, found: false });
  }
});

// Body: { data, costUSD? }
app.post('/api/research-cache/:type/:key', (req, res) => {
  const { type, key } = req.params;
  const { data, costUSD } = req.body || {};
  if (data === undefined) return res.status(400).json({ ok: false, error: 'data is required' });
  fs.writeFileSync(cacheFile(type, key), JSON.stringify({ data, costUSD: costUSD || 0, cachedAt: Date.now() }, null, 2));
  res.json({ ok: true });
});

app.delete('/api/research-cache/:type/:key', (req, res) => {
  const { type, key } = req.params;
  const file = cacheFile(type, key);
  if (fs.existsSync(file)) fs.unlinkSync(file);
  res.json({ ok: true });
});

// ── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════════════════════════╗
║   Account Plan Generator — Proxy Server              ║
║   http://localhost:${PORT}                              ║
╠══════════════════════════════════════════════════════╣
║  API key loaded: ${process.env.ANTHROPIC_API_KEY ? '✅ Yes (' + process.env.ANTHROPIC_API_KEY.slice(0,14) + '…)' : '❌ NOT SET — add to .env'}
║  Open browser:   http://localhost:${PORT}               ║
╚══════════════════════════════════════════════════════╝
`);
});
