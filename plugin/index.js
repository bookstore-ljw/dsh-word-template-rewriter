/**
 * Host half of the Word pipeline panel.
 *
 * It owns three HTTP endpoints under `/word-pipeline` and nothing else:
 * `state`, `upload`, `op`.  The panel runs the project's own CLI
 * (`python -m dsh_word ...`) so the UI and the agent share one implementation;
 * the browser never executes anything itself.
 *
 * A random token is minted at load and injected into the page as a global, so
 * the endpoints reject anything that did not come from this page.
 */

import { spawn, spawnSync } from 'node:child_process';
import { createReadStream, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createHash, randomBytes } from 'node:crypto';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { skillProvider } from './lib/skill-provider.js';

export const name = 'word-pipeline';
export const inject = ['webServer'];

const TOKEN = randomBytes(18).toString('hex');
const HERE = path.dirname(fileURLToPath(import.meta.url));
function resolveWordScripts() {
  const home = process.env.USERPROFILE || process.env.HOME || '';
  const candidates = [
    path.join(home, '.dsh', 'profiles', 'word', 'node_modules', 'dsh-word-master', 'skills', 'word-master', 'scripts'),
    path.join(home, '.dsh', 'profiles', 'desktop', 'node_modules', 'dsh-word-master', 'skills', 'word-master', 'scripts'),
  ];
  for (const dir of candidates) {
    if (existsSync(dir)) return dir;
  }
  return '';
}

const WORD_SCRIPTS = {
  convert: 'convert.py', template: 'template.py', gate: 'gate.py', render: 'render.py',
  export: 'export.py', spec: 'spec.py', project: 'project.py', fonts: 'fonts.py',
  draft: 'draft_model.py', citations: 'citations.py', omml: 'omml.py', visual: 'visual_check.py',
};

const ROOT = process.env.DSH_WORD_ROOT || 'D:\\desktop\\project\\DSH-word';

function resolvePython() {
  const candidates = [
    process.env.DSH_WORD_PYTHON,
    'python',
    'py',
    path.join(process.env.LOCALAPPDATA || '', 'Python', 'pythoncore-3.12-64', 'python.exe'),
  ].filter((item) => item && (item === 'python' || item === 'py' || existsSync(item)));
  for (const candidate of candidates) {
    try {
      const probe = spawnSync(candidate, ['-c', 'print(1)'], { encoding: 'utf8', timeout: 8000 });
      if (probe.status === 0) return candidate;
    } catch {
      /* try the next candidate */
    }
  }
  return 'python';
}

const PYTHON = resolvePython();
const BASE = '/word-pipeline';
const OP_TIMEOUT_MS = 180_000;

const DEFAULTS = {
  template: path.join(ROOT, 'fixtures', 'build', 'template.docx'),
  source: path.join(ROOT, 'fixtures', 'build', 'source.docx'),
  plan: path.join(ROOT, 'fixtures', 'build', 'plan.yaml'),
  corrections: path.join(ROOT, 'fixtures', 'corrections.example.yaml'),
  out: path.join(ROOT, 'fixtures', 'build', 'out.docx'),
};

const KINDS = {
  template: ['.docx'],
  source: ['.docx'],
  plan: ['.yaml', '.yml'],
  image: ['.png', '.jpg', '.jpeg'],
};

const PREVIEW_DIR = path.join(ROOT, 'fixtures', 'build', 'preview-cache');
const progress = { running: false, op: '', steps: [], error: '', startedAt: 0, ok: null };
const previewManifest = new Map();
let renderChain = Promise.resolve();
let livePreview = { path: '', step: 0 };
let liveTarget = '';
let liveChain = Promise.resolve();
let lastLiveSave = 0;
const mappingDebug = new Map();
const MAPPING_DIR = path.join(ROOT, 'logs', 'mapping');

function liveSync(target, source, save, focus) {
  const job = liveChain.then(async () => {
    const args = ['live-sync', '--target', target, '--source', source];
    if (!save) args.push('--no-save');
    if (focus !== undefined && focus !== null) args.push('--focus', String(focus));
    const result = await runCli(args);
    if (result.code !== 0) throw new Error(result.stderr.trim() || 'Word 同步失败');
    if (save) lastLiveSave = Date.now();
  });
  liveChain = job.catch(() => {});
  return job;
}

function liveSyncBlock(target, step, index, save) {
  const job = liveChain.then(async () => {
    const args = ['live-sync-block', '--target', target, '--step', step, '--index', String(index)];
    if (!save) args.push('--no-save');
    const result = await runCli(args);
    if (result.code !== 0) throw new Error(result.stderr.trim() || 'Word 段落同步失败');
    if (save) lastLiveSave = Date.now();
  });
  liveChain = job.catch(() => {});
  return job;
}

function syncStepToWord(item) {
  if (!liveTarget || !item.step) return;
  const index = item.location && item.location.paragraph;
  const kind = String(item.kind || '');
  const save = Date.now() - lastLiveSave >= 30000;
  if (Number.isInteger(index) && ['paragraph', 'text', 'empty', 'heading'].includes(kind)) {
    liveSyncBlock(liveTarget, item.step, index, save).catch(() => {});
    return;
  }
  const highlight = path.join(path.dirname(item.step), 'highlight', path.basename(item.step));
  const source = existsSync(highlight) ? highlight : item.step;
  liveSync(liveTarget, source, save, Number.isInteger(index) ? index : undefined).catch(() => {});
}

const RUN_DIR = path.join(ROOT, 'logs', 'runs');
const runs = new Map();
const MAX_RUNS = 60;

function startRun(kind, meta = {}) {
  const id = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
  const run = { id, kind, meta, startedAt: Date.now(), endedAt: 0, status: 'running', events: [] };
  runs.set(id, run);
  if (runs.size > MAX_RUNS) runs.delete(runs.keys().next().value);
  fs.mkdir(RUN_DIR, { recursive: true }).catch(() => {});
  logEvent(run, 'info', 'start', `${kind} 开始`, meta);
  return run;
}

function logEvent(run, level, phase, message, extra = {}) {
  if (!run) return null;
  const entry = { t: Date.now(), level, phase, message, ...extra };
  run.events.push(entry);
  fs.appendFile(path.join(RUN_DIR, `${run.id}.jsonl`), `${JSON.stringify(entry)}\n`).catch(() => {});
  return entry;
}

function endRun(run, status, extra = {}) {
  if (!run) return;
  run.status = status;
  run.endedAt = Date.now();
  logEvent(run, status === 'ok' ? 'info' : 'error', 'done', `${run.kind} ${status}`, {
    ms: run.endedAt - run.startedAt,
    ...extra,
  });
}

function runSummary(run) {
  return {
    id: run.id,
    kind: run.kind,
    status: run.status,
    startedAt: run.startedAt,
    durationMs: (run.endedAt || Date.now()) - run.startedAt,
    events: run.events.length,
    meta: run.meta,
  };
}
let lastImageSlot = 'image.fig1';
let hostContext = null;

function ensurePreview(docx, png) {
  const job = renderChain.then(async () => {
    if (existsSync(png)) return;
    const pdf = png.replace(/\.png$/i, '.pdf');
    const run = await runCli(['render', '--docx', docx, '--pdf', pdf, '--png', png, '--dpi', '140']);
    if (run.code !== 0) throw new Error(run.stderr.trim() || `渲染失败（退出码 ${run.code}）`);
  });
  renderChain = job.catch(() => {});
  return job;
}

function json(res, status, body) {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(payload);
}

async function readJson(req, limit = 64 * 1024 * 1024) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) throw new Error('请求体过大');
    chunks.push(chunk);
  }
  const text = Buffer.concat(chunks).toString('utf8').trim();
  return text ? JSON.parse(text) : {};
}

function within(candidate, fallback) {
  const raw = typeof candidate === 'string' && candidate.length > 0 ? candidate : fallback;
  const resolved = path.resolve(ROOT, raw);
  const relative = path.relative(ROOT, resolved);
  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error('只允许操作工作目录内的文件');
  }
  return resolved;
}

function runCli(args, onLine) {
  return new Promise((resolve) => {
    let stdout = '';
    let stderr = '';
    let buffer = '';
    let settled = false;
    const feed = (chunk) => {
      if (!onLine) return;
      buffer += chunk;
      let index;
      while ((index = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, index).trim();
        buffer = buffer.slice(index + 1);
        if (line) onLine(line);
      }
    };
    const finish = (code, extra = '') => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code, stdout, stderr: stderr + extra });
    };

    const child = spawn(PYTHON, ['-m', 'dsh_word', ...args], {
      cwd: ROOT,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
      windowsHide: true,
    });
    const timer = setTimeout(() => {
      child.kill();
      finish(-1, `\n命令超时（${OP_TIMEOUT_MS / 1000}s）`);
    }, OP_TIMEOUT_MS);

    child.stdout.on('data', (chunk) => {
      stdout += chunk;
      feed(chunk.toString('utf8'));
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('error', (error) => finish(-1, `\n无法启动 ${PYTHON}: ${error.message}`));
    child.on('close', (code) => finish(code ?? -1));
  });
}

function buildArgs(body) {
  const template = within(body.template, DEFAULTS.template);
  const out = within(body.out, DEFAULTS.out);
  const op = String(body.op ?? '');

  if (op === 'fill') {
    const args = ['fill', '--template', template];
    if (body.source) {
      args.push('--source', within(body.source, DEFAULTS.source), '--auto-plan');
    } else {
      args.push('--plan', within(body.plan, DEFAULTS.plan));
    }
    // Unique per run: Word may still hold the previous run's step files open.
    const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
    args.push('--out', out, '--steps-dir', stepsDir, '--progress');
    return { args, artifacts: [out], stepsDir, out };
  }
  if (op === 'check') {
    const report = path.join(path.dirname(out), 'panel-check.json');
    return {
      args: ['check', '--docx', out, '--template', template, '--json', report],
      artifacts: [out, report],
    };
  }
  if (op === 'apply') {
    const batches = path.join(path.dirname(out), 'batches');
    const report = path.join(path.dirname(out), 'panel-router.json');
    return {
      args: [
        'apply',
        '--template', template,
        '--base-plan', within(body.plan, DEFAULTS.plan),
        '--corrections', within(body.corrections, DEFAULTS.corrections),
        '--out-dir', batches,
        '--json', report,
      ],
      artifacts: [batches, report],
    };
  }
  if (op === 'render') {
    const pdf = path.join(path.dirname(out), 'panel-preview.pdf');
    const png = path.join(path.dirname(out), 'panel-preview.png');
    return {
      args: ['render', '--docx', out, '--pdf', pdf, '--png', png, '--dpi', '160'],
      artifacts: [pdf, png],
    };
  }
  throw new Error(`未知操作: ${op}`);
}

async function executeOp(body, hooks = {}) {
  const { args, artifacts, stepsDir } = buildArgs(body);
  const ledger = hooks.run || null;
  const opStarted = Date.now();
  if (ledger) {
    hooks.emit?.(logEvent(ledger, 'info', String(body.op || 'op'), `执行 ${body.op}`, {
      command: `python -m dsh_word ${args.join(' ')}`,
    }));
  }
  progress.running = true;
  progress.op = String(body.op ?? '');
  progress.steps = [];
  progress.error = '';
  progress.ok = null;
  progress.preview = '';
  progress.startedAt = Date.now();
  if (body.op === 'fill' || body.op === 'convert') livePreview = { path: '', step: 0 };

  const run = await runCli(args, (line) => {
    if (!line.startsWith('PROGRESS ')) return;
    try {
      const item = JSON.parse(line.slice('PROGRESS '.length));
      progress.steps.push(item);
      if (item.kind === 'image' && item.slot) lastImageSlot = item.slot;
      if (hooks.onStep) hooks.onStep(item);
      if (ledger) {
        hooks.emit?.(logEvent(ledger, 'info', 'fill.op', `步骤 ${item.op}：${item.slot || item.kind}`, {
          op: item.op,
          slot: item.slot || '',
          kind: item.kind,
        }));
      }
      if (liveTarget && item.step) syncStepToWord(item);
    } catch {
      /* ignore a malformed progress line */
    }
  });

  if (run.code === 0 && stepsDir && progress.steps.length) {
    const previewDir = path.join(stepsDir, 'previews');
    const manifestPath = path.join(previewDir, 'manifest.json');
    await fs.rm(previewDir, { recursive: true, force: true });
    const highlightDir = path.join(stepsDir, 'highlight');
    const renderDir = existsSync(highlightDir) ? highlightDir : stepsDir;
    const render = await runCli(
      [
        'render-steps',
        '--dir', renderDir,
        '--out-dir', previewDir,
        '--dpi', '110',
        '--json', manifestPath,
        '--stream',
      ],
      (line) => {
        if (!line.startsWith('STEP ')) return;
        const path = line.slice('STEP '.length).trim();
        livePreview = { path, step: livePreview.step + 1 };
        if (hooks.onPreview) hooks.onPreview(path);
        if (ledger) hooks.emit?.(logEvent(ledger, 'info', 'preview', '步骤预览已渲染', { artifact: path }));
      },
    );
    if (render.code === 0 && existsSync(manifestPath)) {
      const manifest = JSON.parse(await fs.readFile(manifestPath, 'utf8'));
      for (const [key, value] of Object.entries(manifest)) {
        previewManifest.set(path.resolve(key), value);
        const original = path.join(stepsDir, path.basename(key));
        if (original !== key) previewManifest.set(path.resolve(original), value);
      }
    } else {
      progress.error = render.stderr.trim() || progress.error;
    }
  }

  progress.running = false;
  progress.ok = run.code === 0;
  if (run.code !== 0) progress.error = run.stderr.trim() || `退出码 ${run.code}`;
  if (ledger) {
    hooks.emit?.(
      logEvent(
        ledger,
        run.code === 0 ? 'info' : 'error',
        String(body.op || 'op'),
        `${body.op} ${run.code === 0 ? '完成' : '失败'}`,
        { ms: Date.now() - opStarted, code: run.code },
      ),
    );
  }

  if (body.op === 'fill' && run.code === 0 && progress.steps.length === 0) {
    const message =
      '模板中没有找到 {{ }} 槽位标记，未做任何填充。请确认资料/模板没有传反，并在模板里放好占位符，例如 {{ text.title }}。';
    progress.ok = false;
    progress.error = message;
    return {
      ok: false,
      op: 'fill',
      code: 0,
      stdout: run.stdout,
      stderr: message,
      artifacts,
      steps: [],
      command: `python -m dsh_word ${args.join(' ')}`,
    };
  }

  if (body.op === 'apply' && run.code === 0) {
    try {
      const names = (await fs.readdir(artifacts[0])).filter((name) =>
        name.toLowerCase().endsWith('.docx'),
      );
      for (const name of names) artifacts.push(path.join(artifacts[0], name));
      if (names.length) {
        const previewPng = path.join(path.dirname(artifacts[0]), 'panel-live-apply.png');
        await ensurePreview(path.join(artifacts[0], names[0]), previewPng);
        artifacts.push(previewPng);
        livePreview = { path: previewPng, step: livePreview.step + 1 };
      }
    } catch (error) {
      progress.error = error.message || String(error);
    }
  }

  return {
    ok: run.code === 0,
    op: body.op,
    code: run.code,
    stdout: run.stdout,
    stderr: run.stderr,
    artifacts,
    steps: progress.steps,
    command: `python -m dsh_word ${args.join(' ')}`,
  };
}

async function handleOp(res, body) {
  json(res, 200, await executeOp(body));
}

async function handleInstruct(res, body) {
  const text = String(body.text ?? '').trim();
  const template = within(body.template, DEFAULTS.template);
  const source = body.source ? within(body.source, DEFAULTS.source) : null;
  const out = within(body.out, DEFAULTS.out);
  const lower = text.toLowerCase();

  if (/填充|fill/.test(lower)) {
    const result = await executeOp({ op: 'fill', template, source, out });
    return json(res, 200, {
      ok: result.ok,
      reply: result.ok
        ? `填充完成：${result.steps.length} 个原子操作，已生成 ${path.basename(result.artifacts[0])}。`
        : `填充失败：${(result.stderr || '').trim() || `退出码 ${result.code}`}`,
      result,
    });
  }
  if (/检查|check/.test(lower)) {
    const result = await executeOp({ op: 'check', template, out });
    return json(res, 200, {
      ok: result.ok,
      reply: (result.stdout || '').trim() || (result.ok ? '检查通过。' : '检查未通过。'),
      result,
    });
  }
  if (/合批|apply|修正/.test(lower)) {
    const result = await executeOp({ op: 'apply', template, out });
    return json(res, 200, {
      ok: result.ok,
      reply: (result.stdout || '').trim() || (result.ok ? '合批完成。' : '合批失败。'),
      result,
    });
  }
  if (/渲染|render|预览/.test(lower)) {
    const result = await executeOp({ op: 'render', template, out });
    return json(res, 200, { ok: result.ok, reply: '已用 Word 重新渲染预览。', result });
  }

  const widthMatch = text.match(/(\d+(?:\.\d+)?)\s*(cm|厘米|公分)/i);
  if (widthMatch) {
    const width = Number(widthMatch[1]);
    const corrections = path.join(ROOT, 'fixtures', 'build', 'agent-corrections.yaml');
    await fs.writeFile(
      corrections,
      `corrections:\n  - slot: ${lastImageSlot}\n    width: ${width}cm\n    note: 来自应用面板对话\n`,
      'utf8',
    );
    const result = await executeOp({ op: 'apply', template, out, corrections });
    return json(res, 200, {
      ok: result.ok,
      reply: `已把 ${lastImageSlot} 调整为 ${width}cm，并重新合批。`,
      result,
    });
  }

  return json(res, 200, {
    ok: true,
    reply: '可以这样指挥我：填充 / 检查 / 合批 / 渲染，或「把图改成 10cm」。',
    result: null,
  });
}

async function analyzeDocument(docx) {
  const run = await runCli(['analyze', '--docx', docx]);
  if (run.code !== 0) throw new Error(run.stderr.trim() || '结构分析失败');
  const line = run.stdout.trim().split('\n').filter(Boolean).pop() || '{}';
  return JSON.parse(line);
}

function deterministicMapping(template, source) {
  const body = template.blocks.filter((block) => block.kind === 'paragraph' || block.kind === 'empty');
  const mapping = [];
  let index = 0;
  let anchor = null;
  for (const block of source.blocks) {
    if (['image', 'formula', 'table'].includes(block.kind)) {
      mapping.push({ template: anchor || 'append', source: block.id, action: anchor ? 'insert-after' : 'append' });
      continue;
    }
    if (block.kind !== 'paragraph' && block.kind !== 'heading') continue;
    const target = body[index];
    if (target) {
      mapping.push({ template: target.id, source: block.id, action: 'replace' });
      anchor = target.id;
      index += 1;
    } else {
      mapping.push({ template: anchor || 'append', source: block.id, action: anchor ? 'insert-after' : 'append' });
    }
  }
  while (index < body.length) {
    mapping.push({ template: body[index].id, action: 'skip' });
    index += 1;
  }
  return mapping;
}

async function llmMapping(template, source, issues) {
  const llm = hostContext && hostContext.get('llm');
  const selector = hostContext && hostContext.get('agentDefaultModel');
  if (!llm || !selector) return { mapping: null, raw: '', reason: 'no-llm' };
  let provider;
  let model;
  try {
    const selection = selector.currentSelection();
    provider = selection.provider;
    model = selection.model;
  } catch (error) {
    return { mapping: null, raw: '', reason: `selection:${error.message || error}` };
  }
  const system = [
    '你是文档结构映射器。输入包含 template.sections / source.sections（heading → 子块，按文档顺序）与完整文本。',
    '目标：把资料内容映射到模板正文，保持资料的原始顺序与章节结构。',
    '只输出 JSON：{"mapping":[{"template":"paragraph:1","source":"paragraph:0","action":"replace"}],"notes":"一句话"}',
    '规则：',
    '- 模板 heading 不映射；正文段落按**章节语义**优先对齐（同名/近义标题下的内容优先），其次才按顺序；',
    '- 文本用 replace；',
    '- 图片/公式/表格必须保持资料中的原始顺序：放在它前面最近的文本块所映射的模板块之后，action 用 insert-after，template 填那个模板块 id；',
    '- 没有可用锚点时用 {"template":"append","source":"<id>","action":"append"}；',
    '- 每个资料块最多用一次；不要重复使用。',
  ].join('\n');
  const payload = {
    template: {
      sections: template.sections || [],
      blocks: template.blocks.map(({ id, kind, style, chars, text, section }) => ({ id, kind, style, chars, section, text })),
    },
    source: {
      sections: source.sections || [],
      blocks: source.blocks.map(({ id, kind, chars, text, section }) => ({ id, kind, chars, section, text })),
    },
    issues: issues || [],
  };
  const stream = llm.stream({
    provider,
    model,
    system,
    messages: [{ role: 'user', content: [{ type: 'text', text: JSON.stringify(payload) }] }],
    maxTokens: 8000,
    temperature: 0,
  });
  let raw = '';
  let chunks = 0;
  const sample = [];
  try {
    for await (const chunk of stream) {
      chunks += 1;
      if (sample.length < 4) sample.push(chunk && chunk.type ? chunk.type : typeof chunk);
      if (typeof chunk === 'string') raw += chunk;
      else if (chunk && chunk.type === 'text-delta' && typeof chunk.text === 'string') raw += chunk.text;
      else if (chunk && chunk.delta && typeof chunk.delta.text === 'string') raw += chunk.delta.text;
    }
  } catch (error) {
    return { mapping: null, raw, chunks, system, payload, reason: `stream:${error.message || error}` };
  }
  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) return { mapping: null, raw, chunks, system, payload, reason: `no-json:sample=${sample.join(',')}` };
  const parsed = JSON.parse(match[0]);
  if (!Array.isArray(parsed.mapping)) return { mapping: null, raw, system, payload };
  return { ...parsed, raw, system, payload };
}

async function semanticReview(template, source, mapping) {
  const llm = hostContext && hostContext.get('llm');
  const selector = hostContext && hostContext.get('agentDefaultModel');
  if (!llm || !selector) return null;
  let provider;
  let model;
  try {
    const selection = selector.currentSelection();
    provider = selection.provider;
    model = selection.model;
  } catch {
    return null;
  }
  const system = [
    '你是映射审校员。给定模板/资料结构、完整文本和映射列表，逐条判断：',
    '1) 资料内容是否放在了语义合适的模板段落（章节是否对应，而不是只看顺序）；',
    '2) 图片/公式/表格是否保持资料中的原始顺序、并紧跟其前置文本；',
    '3) 是否有内容遗漏或重复。',
    '只输出 JSON：{"ok":true|false,"issues":[{"template":"","source":"","reason":""}],"suggestions":"一句话"}',
  ].join('\n');
  const payload = {
    template: { sections: template.sections || [], blocks: template.blocks.map(({ id, kind, section, text }) => ({ id, kind, section, text: String(text || '').slice(0, 240) })) },
    source: { sections: source.sections || [], blocks: source.blocks.map(({ id, kind, section, text }) => ({ id, kind, section, text: String(text || '').slice(0, 240) })) },
    mapping,
  };
  const stream = llm.stream({
    provider,
    model,
    system,
    messages: [{ role: 'user', content: [{ type: 'text', text: JSON.stringify(payload) }] }],
    maxTokens: 4000,
    temperature: 0,
  });
  let raw = '';
  for await (const chunk of stream) {
    if (chunk && chunk.type === 'text-delta') raw += chunk.text;
  }
  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return { ...JSON.parse(match[0]), raw };
  } catch {
    return null;
  }
}

async function runAgenticFill(template, source, out, hooks, run) {
  const tpl = await analyzeDocument(template);
  const src = await analyzeDocument(source);
  let mappingPayload = null;
  const attempts = [];
  try {
    mappingPayload = await llmMapping(tpl, src, null);
    const valid = mappingPayload && Array.isArray(mappingPayload.mapping) && mappingPayload.mapping.length;
    if (valid) {
      attempts.push({ stage: 'initial', raw: mappingPayload.raw, mapping: mappingPayload.mapping, notes: mappingPayload.notes });
      if (run) hooks.emit?.(logEvent(run, 'info', 'map', '模型生成结构映射', {
        blocks: mappingPayload.mapping.length,
        notes: mappingPayload.notes || '',
      }));
    } else if (run) {
      hooks.emit?.(logEvent(run, 'warn', 'map', '模型输出无法解析，回退确定性映射', {
        raw: String((mappingPayload && mappingPayload.raw) || '').slice(0, 300),
        reason: (mappingPayload && mappingPayload.reason) || '',
        chunks: (mappingPayload && mappingPayload.chunks) || 0,
      }));
      mappingPayload = null;
    }
  } catch (error) {
    if (run) hooks.emit?.(logEvent(run, 'warn', 'map', `模型映射失败：${error.message || error}`));
  }
  if (!mappingPayload || !mappingPayload.mapping || !mappingPayload.mapping.length) {
    mappingPayload = { mapping: deterministicMapping(tpl, src), notes: '按文档顺序确定性映射' };
    if (run) hooks.emit?.(logEvent(run, 'info', 'map', '使用确定性映射', { blocks: mappingPayload.mapping.length }));
  }

  const mappingPath = path.join(path.dirname(out), 'mapping.json');
  await fs.writeFile(
    mappingPath,
    JSON.stringify({ mapping: mappingPayload.mapping, notes: mappingPayload.notes || '' }, null, 2),
    'utf8',
  );

  let verify = await runCli(['map-verify', '--template', template, '--source', source, '--mapping', mappingPath]);
  let verifyPayload = null;
  try {
    verifyPayload = JSON.parse(verify.stdout.trim().split('\n').pop() || '{}');
  } catch {
    verifyPayload = null;
  }
  if (run) {
    hooks.emit?.(logEvent(run, verify.code === 0 ? 'info' : 'warn', 'map-check', '映射检查', {
      ok: verify.code === 0,
      issues: (verifyPayload && verifyPayload.issues) || [],
    }));
  }
  if (verify.code !== 0) {
    const fixed = await llmMapping(tpl, src, (verifyPayload && verifyPayload.issues) || []).catch(() => null);
    if (fixed && Array.isArray(fixed.mapping) && fixed.mapping.length) {
      attempts.push({ stage: 'fix', raw: fixed.raw, mapping: fixed.mapping, notes: fixed.notes });
      mappingPayload = fixed;
      await fs.writeFile(
    mappingPath,
    JSON.stringify({ mapping: mappingPayload.mapping, notes: mappingPayload.notes || '' }, null, 2),
    'utf8',
  );
      verify = await runCli(['map-verify', '--template', template, '--source', source, '--mapping', mappingPath]);
      if (run) hooks.emit?.(logEvent(run, 'info', 'map-check', '按检查意见修正映射', { ok: verify.code === 0 }));
    }
  }

  try {
    const review = await semanticReview(tpl, src, mappingPayload.mapping);
    if (review) {
      attempts.push({ stage: 'review', raw: review.raw, issues: review.issues || [], suggestions: review.suggestions || '' });
      if (run) {
        hooks.emit?.(
          logEvent(
            run,
            review.ok ? 'info' : 'warn',
            'map-review',
            review.ok ? '语义审校通过' : `语义审校发现问题：${(review.issues || []).length} 条`,
            { suggestions: review.suggestions || '' },
          ),
        );
      }
      if (!review.ok && Array.isArray(review.issues) && review.issues.length) {
        const fixed2 = await llmMapping(tpl, src, review.issues).catch(() => null);
        if (fixed2 && Array.isArray(fixed2.mapping) && fixed2.mapping.length) {
          attempts.push({ stage: 'remap', raw: fixed2.raw, mapping: fixed2.mapping, notes: fixed2.notes });
          mappingPayload = fixed2;
          await fs.writeFile(
            mappingPath,
            JSON.stringify({ mapping: mappingPayload.mapping, notes: mappingPayload.notes || '' }, null, 2),
            'utf8',
          );
          verify = await runCli(['map-verify', '--template', template, '--source', source, '--mapping', mappingPath]);
          if (run) hooks.emit?.(logEvent(run, 'info', 'map-review', '按语义审校重映射', { ok: verify.code === 0 }));
        }
      }
    }
  } catch (error) {
    if (run) hooks.emit?.(logEvent(run, 'warn', 'map-review', `语义审校失败：${error.message || error}`));
  }

  const debug = {
    runId: run ? run.id : '',
    at: Date.now(),
    template: tpl,
    source: src,
    attempts,
    final: mappingPayload.mapping,
    notes: mappingPayload.notes || '',
    verify: verifyPayload,
  };
  try {
    await fs.mkdir(MAPPING_DIR, { recursive: true });
    const debugPath = path.join(MAPPING_DIR, `${run ? run.id : Date.now()}.json`);
    await fs.writeFile(debugPath, JSON.stringify(debug, null, 2), 'utf8');
    mappingDebug.set(run ? run.id : String(Date.now()), debug);
    if (run) hooks.emit?.(logEvent(run, 'info', 'map', '映射详情已保存', { path: debugPath }));
  } catch {
    /* debug dump must never break the run */
  }

  try {
    const review = await semanticReview(tpl, src, mappingPayload.mapping);
    if (review) {
      attempts.push({ stage: 'review', raw: review.raw, issues: review.issues || [], suggestions: review.suggestions || '' });
      if (run) {
        hooks.emit?.(
          logEvent(
            run,
            review.ok ? 'info' : 'warn',
            'map-review',
            review.ok ? '语义审校通过' : `语义审校发现问题：${(review.issues || []).length} 条`,
            { suggestions: review.suggestions || '' },
          ),
        );
      }
      if (!review.ok && Array.isArray(review.issues) && review.issues.length) {
        const fixed2 = await llmMapping(tpl, src, review.issues).catch(() => null);
        if (fixed2 && Array.isArray(fixed2.mapping) && fixed2.mapping.length) {
          attempts.push({ stage: 'remap', raw: fixed2.raw, mapping: fixed2.mapping, notes: fixed2.notes });
          mappingPayload = fixed2;
          await fs.writeFile(
            mappingPath,
            JSON.stringify({ mapping: mappingPayload.mapping, notes: mappingPayload.notes || '' }, null, 2),
            'utf8',
          );
          verify = await runCli(['map-verify', '--template', template, '--source', source, '--mapping', mappingPath]);
          if (run) hooks.emit?.(logEvent(run, 'info', 'map-review', '按语义审校重映射', { ok: verify.code === 0 }));
        }
      }
    }
  } catch (error) {
    if (run) hooks.emit?.(logEvent(run, 'warn', 'map-review', `语义审校失败：${error.message || error}`));
  }

  const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
  const steps = [];
  const checks = [];
  const apply = await runCli(
    ['map-apply', '--template', template, '--source', source, '--mapping', mappingPath,
      '--out', out, '--steps-dir', stepsDir, '--progress'],
    (line) => {
      if (line.startsWith('CHECK ')) {
        let check = null;
        try {
          check = JSON.parse(line.slice('CHECK '.length));
        } catch {
          return;
        }
        checks.push(check);
        if (run) {
          hooks.emit?.(
            logEvent(
              run,
              check.ok ? 'info' : 'warn',
              '块检查',
              `块 ${check.index} ${check.ok ? '通过' : `问题：${check.issues.join('、')}`}`,
              { block: check.index, kind: check.kind },
            ),
          );
        }
        return;
      }
      if (!line.startsWith('PROGRESS ')) return;
      let item = null;
      try {
        item = JSON.parse(line.slice('PROGRESS '.length));
      } catch {
        return;
      }
      steps.push(item);
      if (hooks.onStep) hooks.onStep(item);
      if (liveTarget && item.step) syncStepToWord(item);
    },
  );
  return {
    ok: apply.code === 0,
    code: apply.code,
    steps,
    checks,
    stdout: apply.stdout,
    stderr: apply.stderr,
  };
}

async function runShell(command) {
  return new Promise((resolve) => {
    let out = '';
    let err = '';
    const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], {
      cwd: ROOT,
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    });
    const timer = setTimeout(() => {
      child.kill();
      resolve({ code: -1, stdout: out, stderr: '命令超时（120s）' });
    }, 120000);
    child.stdout.on('data', (chunk) => { out += chunk.toString('utf8'); });
    child.stderr.on('data', (chunk) => { err += chunk.toString('utf8'); });
    child.on('error', (error) => { clearTimeout(timer); resolve({ code: -1, stdout: out, stderr: error.message }); });
    child.on('close', (code) => { clearTimeout(timer); resolve({ code: code ?? -1, stdout: out, stderr: err }); });
  });
}

async function agentRunLoop(run, body, send) {
  const template = within(body.template, DEFAULTS.template);
  const source = body.source ? within(body.source, DEFAULTS.source) : null;
  const out = within(body.out, DEFAULTS.out);
  const system = [
    '你是 Word 流水线 Agent（DSH Agent 的大脑拷贝），运行在 Word 流水线助手这个外壳里。',
    '你可以使用工具读写文件、运行命令、分析/填充 Word 文档。必须用工具完成，不要编造内容。',
    '可用工具（每步只调用一个）：',
    'list_dir({"path":"."})',
    'read_file({"path":"docs/architecture.md"})',
    'write_file({"path":"tmp/x.py","content":"..."})',
    'run_command({"command":"python -m dsh_word analyze --docx fixtures/build/template.docx"})',
    'word({"cmd":"gate","args":["--project","...","--input","..."]})：调用 dsh-word-master 的 Word 工具。',
    '  cmd 可选：convert / template / spec / gate / render / export / project / fonts / draft / citations / omml / visual。',
    '  例：word({"cmd":"convert","args":["--input","C:/x.pdf","--out-dir","fixtures/build/materials"]})',
    'analyze({"docx":"绝对路径"})：结构块与章节树。',
    'safe_fill({})：非破坏填充（只改已有段落文字，保留版式，单双栏安全）。',
    'sectional({})：按章节逐段填充。',
    'map_apply({"mapping":[{"template":"paragraph:1","source":"paragraph:0","action":"replace"}]})',
    'check({}) / heal({}) / render({})',
    '工作方式：先读项目与模板结构（analyze），理解样式/章节/题注/公式规则；',
    '再按章节逐块处理，每改一部分就用 check 对比模板；不合适就调整，不要一次性乱改。',
    '缺内容时记录并继续，最后用 final 汇报：产物路径、检查问题、缺失章节。',
    '输出只能是 JSON：{"thought":"一句想法","tool":"工具名","args":{...}} 或 {"final":"给用户的中文结论"}。',
  ].join('\n');
  const messages = [
    {
      role: 'user',
      content: [{
        type: 'text',
        text: `任务：把资料按模板格式填成成稿，模板只提供格式（可能单栏或双栏），不要破坏模板版式。\n模板：${template}\n资料：${source || '（未提供，使用示例）'}\n输出：${out}\n先 analyze 模板与资料，再决定如何逐节填充。`,
      }],
    },
  ];

  const executeTool = async (tool, args) => {
    const call = String(tool || '');
    try {
      if (call === 'list_dir') {
        const target = within(args.path || '.', '.');
        const names = await fs.readdir(target);
        return names.slice(0, 200).join('\n');
      }
      if (call === 'read_file') {
        const target = within(args.path, '');
        const text = await fs.readFile(target, 'utf8');
        return text.slice(0, 20000);
      }
      if (call === 'write_file') {
        const target = within(args.path, '');
        await fs.mkdir(path.dirname(target), { recursive: true });
        await fs.writeFile(target, String(args.content || ''), 'utf8');
        return `写入 ${target}`;
      }
      if (call === 'run_command') {
        const result = await runShell(String(args.command || ''));
        return `exit=${result.code}\n${result.stdout.slice(-4000)}\n${result.stderr.slice(-2000)}`;
      }
      if (call === 'word') {
        const scriptsDir = resolveWordScripts();
        if (!scriptsDir) return '没有找到 dsh-word-master 脚本目录（请先安装该 bundle）';
        const script = WORD_SCRIPTS[String(args.cmd || '')];
        if (!script) return `未知 word 子命令：${args.cmd}（可选：${Object.keys(WORD_SCRIPTS).join('/')}）`;
        const scriptArgs = Array.isArray(args.args) ? args.args.map(String) : [];
        const result = await new Promise((resolve) => {
          let out = '';
          let err = '';
          const child = spawn('python', [path.join(scriptsDir, script), ...scriptArgs], {
            cwd: scriptsDir, windowsHide: true,
            env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
          });
          const timer = setTimeout(() => { child.kill(); resolve({ code: -1, stdout: out, stderr: '超时（180s）' }); }, 180000);
          child.stdout.on('data', (chunk) => { out += chunk.toString('utf8'); });
          child.stderr.on('data', (chunk) => { err += chunk.toString('utf8'); });
          child.on('error', (error) => { clearTimeout(timer); resolve({ code: -1, stdout: out, stderr: error.message }); });
          child.on('close', (code) => { clearTimeout(timer); resolve({ code: code ?? -1, stdout: out, stderr: err }); });
        });
        return `exit=${result.code}\n${result.stdout.slice(-4000)}\n${result.stderr.slice(-2000)}`;
      }
      if (call === 'analyze') {
        const target = within(args.docx, template);
        const result = await runCli(['analyze', '--docx', target]);
        return result.code === 0 ? result.stdout.trim().slice(0, 3000) : result.stderr.trim();
      }
      if (call === 'safe_fill' || call === 'sectional') {
        const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
        const command = call === 'sectional' ? 'sectional' : 'safe-fill';
        const result = await runCli(
          [command, '--template', template, '--source', source || DEFAULTS.source, '--out', out, '--steps-dir', stepsDir, '--progress'],
          (line) => {
            if (line.startsWith('SEGMENT ')) {
              try { send({ type: 'section', item: JSON.parse(line.slice('SEGMENT '.length)) }); } catch { /* ignore */ }
            }
          },
        );
        return result.code === 0 ? `已填充：${out}` : result.stderr.trim();
      }
      if (call === 'map_apply') {
        const mappingPath = path.join(path.dirname(out), 'agent-mapping.json');
        await fs.writeFile(mappingPath, JSON.stringify({ mapping: args.mapping || [] }, null, 2), 'utf8');
        const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
        const result = await runCli(['map-apply', '--template', template, '--source', source || DEFAULTS.source,
          '--mapping', mappingPath, '--out', out, '--steps-dir', stepsDir, '--progress'], (line) => {
          if (line.startsWith('PROGRESS ')) {
            try {
              const item = JSON.parse(line.slice('PROGRESS '.length));
              send({ type: 'step', item });
              if (liveTarget && item.step) syncStepToWord(item);
            } catch { /* ignore */ }
          }
        });
        return result.code === 0 ? `已应用映射：${out}` : result.stderr.trim();
      }
      if (call === 'check') {
        const report = path.join(path.dirname(out), 'panel-check.json');
        const result = await runCli(['check', '--docx', out, '--template', template, '--json', report]);
        const issues = await readCheckIssues(out);
        return JSON.stringify({ ok: result.code === 0, issues });
      }
      if (call === 'heal') {
        const result = await runCli(['heal', '--template', template, '--docx', out]);
        return result.code === 0 ? `已恢复模板格式：${result.stdout.trim()}` : result.stderr.trim();
      }
      if (call === 'render') {
        const pdf = path.join(path.dirname(out), 'panel-preview.pdf');
        const png = path.join(path.dirname(out), 'panel-preview.png');
        const result = await runCli(['render', '--docx', out, '--pdf', pdf, '--png', png, '--dpi', '160']);
        return result.code === 0 ? `已渲染 ${png}` : result.stderr.trim();
      }
      return `未知工具：${call}`;
    } catch (error) {
      return `工具执行失败：${error.message || error}`;
    }
  };

  let parseFailures = 0;
  const ran = new Set();
  for (let step = 1; step <= 18; step += 1) {
    let raw = '';
    try {
      const completed = await completeText(system, messages, 6000);
      raw = completed.raw || '';
    } catch (error) {
      send({ type: 'final', text: `模型调用失败：${error.message || error}` });
      return;
    }
    const match = raw.match(/\{[\s\S]*\}/);
    let decision = null;
    if (match) {
      try { decision = JSON.parse(match[0]); } catch { decision = null; }
    }
    if (!decision) {
      parseFailures += 1;
      if (parseFailures >= 3) {
        send({ type: 'final', text: '模型连续没有返回可执行动作。' });
        return;
      }
      messages.push(
        { role: 'assistant', content: [{ type: 'text', text: raw.slice(0, 800) }] },
        { role: 'user', content: [{ type: 'text', text: '只输出一个 JSON：{"thought":..,"tool":..,"args":..} 或 {"final":..}。' }] },
      );
      continue;
    }
    if (decision.final || !decision.tool) {
      const missing = [];
      const usedWord = ran.has('word');
      if (!usedWord) {
        if (!ran.has('analyze')) missing.push('用 analyze 读模板/资料');
        if (!ran.has('safe_fill') && !ran.has('sectional') && !ran.has('map_apply')) missing.push('填充 out.docx');
        if (!ran.has('check')) missing.push('用 check 检查');
        if (!ran.has('render')) missing.push('用 render 渲染');
      }
      if (missing.length) {
        messages.push(
          { role: 'assistant', content: [{ type: 'text', text: String(decision.final || '') }] },
          { role: 'user', content: [{ type: 'text', text: `还不能结束，请先完成：${missing.join('；')}。继续用工具，完成后再 final。` }] },
        );
        continue;
      }
      send({ type: 'final', text: String(decision.final || '完成。') });
      return;
    }
    send({ type: 'thought', text: String(decision.thought || '') });
    send({ type: 'tool', name: String(decision.tool), args: decision.args || {} });
    ran.add(String(decision.tool));
    const summary = await executeTool(decision.tool, decision.args || {});
    send({ type: 'result', name: String(decision.tool), summary: String(summary).slice(0, 3000) });
    messages.push(
      { role: 'assistant', content: [{ type: 'text', text: JSON.stringify({ tool: decision.tool }) }] },
      { role: 'user', content: [{ type: 'text', text: `工具 ${decision.tool} 结果：${String(summary).slice(0, 4000)}` }] },
    );
  }
  send({ type: 'final', text: '达到步数上限，先停在这里。' });
}

async function handleAgentRunStream(req, res, body) {
  res.writeHead(200, { 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-store' });
  const send = (event) => {
    try {
      res.write(`${JSON.stringify(event)}\n`);
    } catch {
      /* client gone */
    }
  };
  const run = startRun('agent', { template: body.template, source: body.source });
  try {
    await agentRunLoop(run, body, send);
    endRun(run, 'ok');
    send({ type: 'done' });
  } catch (error) {
    send({ type: 'error', error: String(error.message || error) });
    endRun(run, 'fail');
  }
  res.end();
}

async function runSafeFill(template, source, out, hooks, run) {
  const segments = [];
  let summary = '';
  const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
  const result = await runCli(
    ['safe-fill', '--template', template, '--source', source, '--out', out, '--steps-dir', stepsDir, '--progress'],
    (line) => {
      if (line.startsWith('SEGMENT ')) {
        try {
          const item = JSON.parse(line.slice('SEGMENT '.length));
          segments.push(item);
          if (run) hooks.emit?.(logEvent(run, 'info', 'section', `${item.label}：替换 ${item.replaced}/${item.slots} 段`, item));
        } catch {
          /* ignore */
        }
      } else if (line.startsWith('{')) {
        summary = line;
      }
    },
  );
  return { ok: result.code === 0, code: result.code, steps: [], checks: [], stdout: result.stdout, stderr: result.stderr, segments, summary };
}

async function runCompose(template, source, out, hooks, run) {
  const steps = [];
  const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
  const result = await runCli(
    ['compose', '--template', template, '--source', source, '--out', out, '--steps-dir', stepsDir, '--progress'],
    (line) => {
      if (!line.startsWith('PROGRESS ')) return;
      let item = null;
      try {
        item = JSON.parse(line.slice('PROGRESS '.length));
      } catch {
        return;
      }
      steps.push(item);
      if (hooks.onStep) hooks.onStep(item);
      if (liveTarget && item.step) syncStepToWord(item);
    },
  );
  return { ok: result.code === 0, code: result.code, steps, checks: [], stdout: result.stdout, stderr: result.stderr };
}

async function runSectionedFill(template, source, out, hooks, run) {
  const steps = [];
  const checks = [];
  const missing = [];
  const stepsDir = path.join(path.dirname(out), 'panel-steps', String(Date.now()));
  const apply = await runCli(
    ['section-fill', '--template', template, '--source', source, '--out', out, '--steps-dir', stepsDir, '--progress'],
    (line) => {
      const cut = line.indexOf(' ');
      if (cut < 0) return;
      const kind = line.slice(0, cut);
      let payload = null;
      try {
        payload = JSON.parse(line.slice(cut + 1));
      } catch {
        return;
      }
      if (kind === 'SEGMENT') {
        if (run) hooks.emit?.(logEvent(run, 'info', 'segment', `章节 ${payload.segment}${payload.title ? `（${payload.title}）` : ''}：${payload.mapped} 块${payload.missing ? '，资料缺失' : ''}`, payload));
      } else if (kind === 'MISSING') {
        missing.push(payload);
        if (run) hooks.emit?.(logEvent(run, 'warn', 'missing', `资料缺少「${payload.segment}」内容：${payload.reason}`, payload));
        if (hooks.onMissing) hooks.onMissing(payload);
      } else if (kind === 'PROGRESS') {
        steps.push(payload);
        if (hooks.onStep) hooks.onStep(payload);
        if (liveTarget && payload.step) syncStepToWord(payload);
      } else if (kind === 'CHECK') {
        checks.push(payload);
        if (run) hooks.emit?.(logEvent(run, payload.ok ? 'info' : 'warn', '块检查', `块 ${payload.index} ${payload.ok ? '通过' : `问题：${payload.issues.join('、')}`}`, payload));
      }
    },
  );
  return { ok: apply.code === 0, code: apply.code, steps, checks, missing, stdout: apply.stdout, stderr: apply.stderr };
}

async function readCheckIssues(out) {
  try {
    const raw = await fs.readFile(path.join(path.dirname(out), 'panel-check.json'), 'utf8');
    const list = JSON.parse(raw);
    if (!Array.isArray(list)) return [];
    return list
      .filter((item) => item.severity !== 'info' || item.rule === 'template-conformance')
      .slice(0, 10)
      .map((item) => ({
        rule: item.rule,
        severity: item.severity || 'error',
        message: item.message,
        detail: item.detail || '',
      }));
  } catch {
    return [];
  }
}

async function runWordMasterScript(script, args, timeoutMs = 180000) {
  const scriptsDir = resolveWordScripts();
  if (!scriptsDir) return { code: -1, stdout: '', stderr: '没有找到 dsh-word-master 脚本' };
  return new Promise((resolve) => {
    let out = '';
    let err = '';
    const child = spawn('python', [path.join(scriptsDir, script), ...args], {
      cwd: scriptsDir, windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    });
    const timer = setTimeout(() => { child.kill(); resolve({ code: -1, stdout: out, stderr: '超时' }); }, timeoutMs);
    child.stdout.on('data', (chunk) => { out += chunk.toString('utf8'); });
    child.stderr.on('data', (chunk) => { err += chunk.toString('utf8'); });
    child.on('error', (error) => { clearTimeout(timer); resolve({ code: -1, stdout: out, stderr: error.message }); });
    child.on('close', (code) => { clearTimeout(timer); resolve({ code: code ?? -1, stdout: out, stderr: err }); });
  });
}

async function wordMasterRender(docx) {
  try {
    const projectDir = path.join(ROOT, 'fixtures', 'build', 'wm-project');
    if (!existsSync(path.join(projectDir, 'project.json'))) {
      await runWordMasterScript('project.py', ['init', projectDir, '--title', 'word-pipeline', '--force']);
    }
    const result = await runWordMasterScript('render.py', ['--project', projectDir, '--input', docx, '--dpi', '110']);
    const last = result.stdout.trim().split('\n').pop() || '{}';
    return JSON.parse(last);
  } catch {
    return null;
  }
}

async function wordMasterGate(docx) {
  try {
    const projectDir = path.join(ROOT, 'fixtures', 'build', 'wm-project');
    if (!existsSync(path.join(projectDir, 'project.json'))) {
      await runWordMasterScript('project.py', ['init', projectDir, '--title', 'word-pipeline', '--force']);
    }
    const result = await runWordMasterScript('gate.py', ['--project', projectDir, '--input', docx, '--json']);
    const last = result.stdout.trim().split('\n').pop() || '{}';
    const report = JSON.parse(last);
    return { ...report, issues: [...(report.p0 || []), ...(report.p1 || [])] };
  } catch {
    return null;
  }
}

async function runConvert(body, hooks = {}) {
  const template = within(body.template, DEFAULTS.template);
  const source = body.source ? within(body.source, DEFAULTS.source) : null;
  const out = within(body.out, DEFAULTS.out);
  const stages = [];
  const run = hooks.run || startRun('convert', { template, source });
  const emit = (entry) => hooks.onLog?.(entry);
  emit(logEvent(run, 'info', 'start', '开始转换', { template, source, out }));
  if (hooks.onStage) hooks.onStage('fill');

  const fill = await (async () => {
    let markerless = false;
    let tpl = null;
    try {
      tpl = await analyzeDocument(template);
      markerless = !tpl.blocks.some((block) => String(block.text || '').includes('{{'));
    } catch {
      markerless = false;
    }
    if (markerless && source) {
      emit(logEvent(run, 'info', 'section', '非破坏填充：只改已有段落文字，保留模板版式'));
      return runSafeFill(template, source, out, { ...hooks, emit }, run);
    }
    return executeOp({ op: 'fill', template, source, out }, { ...hooks, run, emit });
  })();
  stages.push({ name: 'fill', ok: fill.ok, steps: fill.steps.length || 0 });
  if (!fill.ok) {
    emit(logEvent(run, 'error', 'fill', '填充失败', { error: (fill.stderr || '').trim() }));
    endRun(run, 'fail', { stage: 'fill' });
    return {
      ok: false,
      stage: 'fill',
      stages,
      error: (fill.stderr || '').trim() || `填充退出码 ${fill.code}`,
    };
  }

  if (fill.checks && fill.checks.length) {
    const failed = fill.checks.filter((check) => !check.ok);
    emit(
      logEvent(
        run,
        failed.length ? 'warn' : 'info',
        '块检查',
        `逐块检查：${fill.checks.length} 块，${failed.length} 块有问题`,
        { failed: failed.map((check) => ({ block: check.index, issues: check.issues })) },
      ),
    );
  }

  if (hooks.onStage) hooks.onStage('check');
  emit(logEvent(run, 'info', 'fill', `填充完成 ${fill.steps.length} 步`, { steps: fill.steps.length }));
  const check = await executeOp({ op: 'check', template, out }, { ...hooks, run, emit });
  stages.push({ name: 'check', ok: check.ok, summary: (check.stdout || '').trim() });
  let checkIssues = await readCheckIssues(out);
  let checkOk = check.ok;
  const needsHeal = checkIssues.some(
    (item) =>
      item.rule === 'template-conformance' && /(样式|字符格式|页面设置)/.test(item.message),
  );
  if (needsHeal) {
    const heal = await runCli(['heal', '--template', template, '--docx', out]);
    let cost = null;
    try {
      cost = JSON.parse(heal.stdout.trim().split('\n').pop() || '{}');
    } catch {
      cost = null;
    }
    if (heal.code === 0) {
      emit(logEvent(run, 'info', 'heal', '格式自愈：已恢复模板样式', cost || {}));
      const recheck = await executeOp({ op: 'check', template, out }, { ...hooks, run, emit });
      checkOk = recheck.ok;
      checkIssues = await readCheckIssues(out);
    } else {
      emit(logEvent(run, 'warn', 'heal', `格式自愈失败：${(heal.stderr || '').trim()}`));
    }
  }
  let templateHasJump = false;
  try {
    const tplDoc = await analyzeDocument(template);
    let prev = 0;
    for (const block of tplDoc.blocks) {
      const level = Number((String(block.style || '').match(/(\d+)/) || [])[1] || 0);
      if (!level) continue;
      if (level > prev + 1) { templateHasJump = true; break; }
      prev = level;
    }
  } catch { templateHasJump = false; }
  const gate = await wordMasterGate(out);
  if (gate) {
    const p0 = gate.p0 || [];
    const p1 = gate.p1 || [];
    for (const item of [...p0, ...p1].slice(0, 10)) {
      const message = item.message || item.title || JSON.stringify(item);
      if (templateHasJump && /标题.*跳到|heading.*jump/i.test(message)) {
        emit(logEvent(run, 'info', 'gate', '模板本身存在该标题层级，忽略此提示', { message }));
        continue;
      }
      checkIssues.push({
        severity: p0.includes(item) ? 'error' : 'info',
        rule: 'word-master-gate',
        message,
      });
    }
    emit(logEvent(run, p0.length ? 'warn' : 'info', 'gate',
      `Word Master 质检：${gate.verdict || 'UNKNOWN'}（P0 ${p0.length}，P1 ${p1.length}）`,
      { verdict: gate.verdict, p0, p1 }));
    if (p0.length) checkOk = false;
  }
  const headingJump = [...(gate?.p0 || []), ...(gate?.p1 || [])].some((item) =>
    /标题.*跳到|heading.*jump/i.test(item.message || ''),
  );
  if (headingJump) {
    const fixed = await runCli(['heal', '--template', template, '--docx', out, '--headings']);
    if (fixed.code === 0) {
      emit(logEvent(run, 'info', 'heal', '标题跳级：已按模板层级自动修正'));
      const recheck = await executeOp({ op: 'check', template, out }, { ...hooks, run, emit });
      checkOk = recheck.ok;
      checkIssues = await readCheckIssues(out);
      const gate2 = await wordMasterGate(out);
      if (gate2) {
        for (const item of [...(gate2.p0 || []), ...(gate2.p1 || [])].slice(0, 10)) {
          checkIssues.push({
            severity: (gate2.p0 || []).includes(item) ? 'error' : 'info',
            rule: 'word-master-gate',
            message: item.message || item.title || JSON.stringify(item),
          });
        }
        if ((gate2.p0 || []).length) checkOk = false;
      }
    }
  }
  const hasCheckError = !checkOk || checkIssues.some((item) => item.severity === 'error');
  const checkFailed = hasCheckError;
  if (hasCheckError) {
    emit(
      logEvent(run, 'warn', 'check', `检查未通过：${checkIssues.length || 1} 个问题，继续渲染以便查看`, {
        issues: checkIssues.slice(0, 5),
      }),
    );
    stages[stages.length - 1] = {
      name: 'check',
      ok: false,
      summary: (check.stdout || '').trim(),
      issues: checkIssues,
    };
  } else if (checkIssues.length) {
    emit(
      logEvent(run, 'info', 'check', `检查通过；与模板对比发现 ${checkIssues.length} 处差异`, {
        issues: checkIssues.slice(0, 5),
      }),
    );
    stages[stages.length - 1] = {
      name: 'check',
      ok: true,
      summary: (check.stdout || '').trim(),
      issues: checkIssues,
    };
  } else {
    emit(logEvent(run, 'info', 'check', '检查通过', { summary: (check.stdout || '').trim() }));
  }
  if (liveTarget) await liveSync(liveTarget, out, true).catch(() => {});
  if (hooks.onStage) hooks.onStage('render');
  const render = await executeOp({ op: 'render', template, out }, { ...hooks, run, emit });
  stages.push({ name: 'render', ok: render.ok });
  const wmRender = await wordMasterRender(out);
  if (wmRender && wmRender.ok) {
    emit(logEvent(run, 'info', 'render', `Word Master 渲染：${wmRender.pages || 0} 页（LibreOffice）`, {
      pdf: wmRender.pdf, pages: wmRender.images,
    }));
    if (Array.isArray(wmRender.images) && wmRender.images[0]) {
      await fs.copyFile(wmRender.images[0], path.join(path.dirname(out), 'panel-preview.png')).catch(() => {});
    }
  } else if (wmRender) {
    emit(logEvent(run, 'warn', 'render', `Word Master 渲染跳过：${wmRender.reason || wmRender.message || '未知'}`, {}));
  }
  if (!render.ok) {
    emit(logEvent(run, 'error', 'render', '渲染失败', { error: (render.stderr || '').trim() }));
    endRun(run, 'fail', { stage: 'render' });
    return {
      ok: false,
      stage: 'render',
      stages,
      error: (render.stderr || '').trim() || '渲染失败',
    };
  }

  const docxOut = out;
  const pdfOut = path.join(path.dirname(out), 'panel-preview.pdf');
  const pngOut = path.join(path.dirname(out), 'panel-preview.png');
  if (hooks.onPreview) hooks.onPreview(pngOut);
  if (hooks.onStage) hooks.onStage('done');
  emit(logEvent(run, 'info', 'done', '转换完成', { out, steps: fill.steps.length }));
  endRun(run, 'ok', { out });

  return {
    ok: true,
    runId: run.id,
    checkFailed,
    checkIssues,
    stages,
    docx: docxOut,
    pdf: pdfOut,
    png: pngOut,
    steps: fill.steps,
    check: (check.stdout || '').trim(),
  };
}

async function handleConvertStream(req, res, body) {
  res.writeHead(200, {
    'Content-Type': 'application/x-ndjson; charset=utf-8',
    'Cache-Control': 'no-store',
    'X-Accel-Buffering': 'no',
  });
  const send = (event) => {
    try {
      res.write(`${JSON.stringify(event)}\n`);
    } catch {
      /* client went away */
    }
  };
  try {
    const run = startRun('convert', { template: body.template, source: body.source });
    send({ type: 'run', runId: run.id });
    const result = await runConvert(body, {
      run,
      onStage: (name) => send({ type: 'stage', name }),
      onStep: (item) => send({ type: 'step', item }),
      onPreview: (path) => send({ type: 'preview', path }),
      onMissing: (payload) => send({ type: 'missing', ...payload }),
      onLog: (entry) => send({ type: 'log', entry }),
    });
    send({ type: 'done', result });
  } catch (error) {
    send({ type: 'error', error: String(error.message || error) });
  }
  res.end();
}

async function handleConvert(res, body) {
  try {
    json(res, 200, await runConvert(body));
  } catch (error) {
    json(res, 500, { ok: false, error: String(error.message || error) });
  }
}

const MODEL_USAGE_PATH = path.join(ROOT, 'data', 'model-usage.json');

function loadModelUsage() {
  try {
    return JSON.parse(readFileSync(MODEL_USAGE_PATH, 'utf8'));
  } catch {
    return { calls: 0, inputTokens: 0, outputTokens: 0, cost: 0, failures: 0, byProvider: {} };
  }
}

function saveModelUsage(usage) {
  mkdirSync(path.dirname(MODEL_USAGE_PATH), { recursive: true });
  writeFileSync(MODEL_USAGE_PATH, JSON.stringify(usage, null, 2), 'utf8');
}

function recordUsage(provider, usage, ok, error) {
  const totals = loadModelUsage();
  const key = provider.id;
  const entry = totals.byProvider[key] || { name: provider.name, calls: 0, inputTokens: 0, outputTokens: 0, cost: 0, failures: 0 };
  entry.name = provider.name || entry.name;
  if (ok && usage) {
    const input = Number(usage.input || 0);
    const output = Number(usage.output || 0);
    const priceIn = Number(provider.priceIn || 0);
    const priceOut = Number(provider.priceOut || 0);
    const cost = (input / 1e6) * priceIn + (output / 1e6) * priceOut;
    totals.calls += 1;
    totals.inputTokens += input;
    totals.outputTokens += output;
    totals.cost += cost;
    entry.calls += 1;
    entry.inputTokens += input;
    entry.outputTokens += output;
    entry.cost += cost;
  } else if (!ok) {
    totals.failures = (totals.failures || 0) + 1;
    entry.failures = (entry.failures || 0) + 1;
    entry.lastError = String(error || '').slice(0, 200);
  }
  totals.byProvider[key] = entry;
  saveModelUsage(totals);
  return totals;
}

const MODEL_CONFIG_PATH = path.join(ROOT, 'data', 'model-config.json');

function loadModelConfig() {
  try {
    return JSON.parse(readFileSync(MODEL_CONFIG_PATH, 'utf8'));
  } catch {
    return { active: '', providers: [] };
  }
}

function saveModelConfig(config) {
  mkdirSync(path.dirname(MODEL_CONFIG_PATH), { recursive: true });
  writeFileSync(MODEL_CONFIG_PATH, JSON.stringify(config, null, 2), 'utf8');
}

function maskKey(key) {
  const value = String(key || '');
  if (value.length <= 8) return value ? '****' : '';
  return `${value.slice(0, 4)}…${value.slice(-4)}`;
}

function messagesToText(messages) {
  return messages
    .map((m) => {
      const content = Array.isArray(m.content) ? m.content : [{ type: 'text', text: String(m.content || '') }];
      const text = content.map((c) => (c.type === 'text' ? c.text : '')).join('');
      return `${m.role === 'assistant' ? 'Assistant' : 'User'}: ${text}`;
    })
    .join('\n\n');
}

async function callCustomProvider(provider, system, messages, maxTokens = 3000) {
  const base = String(provider.baseUrl || '').replace(/\/$/, '');
  if (!base || !provider.model) throw new Error('模型配置不完整（baseUrl / model）');
  if (provider.type === 'anthropic') {
    const response = await fetch(`${base}/v1/messages`, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': provider.apiKey || '',
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: provider.model,
        max_tokens: maxTokens,
        system,
        messages: [{ role: 'user', content: messagesToText(messages) }],
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error?.message || `HTTP ${response.status}`);
    const usage = data.usage || {};
    return {
      text: (data.content || []).map((part) => part.text || '').join(''),
      usage: { input: usage.input_tokens || 0, output: usage.output_tokens || 0 },
    };
  }
  const response = await fetch(`${base}/chat/completions`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', authorization: `Bearer ${provider.apiKey || ''}` },
    body: JSON.stringify({
      model: provider.model,
      max_tokens: maxTokens,
      temperature: 0,
      messages: [{ role: 'system', content: system }, { role: 'user', content: messagesToText(messages) }],
    }),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error?.message || `HTTP ${response.status}`);
  const usage = data.usage || {};
  return {
    text: data.choices?.[0]?.message?.content || '',
    usage: { input: usage.prompt_tokens || 0, output: usage.completion_tokens || 0 },
  };
}

async function completeText(system, messages, maxTokens = 3000) {
  const config = loadModelConfig();
  const providers = (config.providers || []).slice();
  const activeId = config.active;
  providers.sort((a, b) => (a.id === activeId ? -1 : b.id === activeId ? 1 : 0));
  let lastError = null;
  for (const provider of providers) {
    try {
      const result = await callCustomProvider(provider, system, messages, maxTokens);
      recordUsage(provider, result.usage, true);
      return { raw: result.text, source: provider.name || provider.model };
    } catch (error) {
      lastError = error;
      recordUsage(provider, null, false, error.message || error);
    }
  }
  if (providers.length && lastError) throw lastError;
  const llm = hostContext && hostContext.get('llm');
  const selector = hostContext && hostContext.get('agentDefaultModel');
  if (!llm || !selector) throw new Error('没有可用模型：请在「模型」里添加 API');
  const selection = selector.currentSelection();
  const stream = llm.stream({ provider: selection.provider, model: selection.model, system, messages, maxTokens, temperature: 0 });
  let raw = '';
  for await (const chunk of stream) {
    if (chunk && chunk.type === 'text-delta' && typeof chunk.text === 'string') raw += chunk.text;
    else if (chunk && chunk.delta && typeof chunk.delta.text === 'string') raw += chunk.delta.text;
  }
  return { raw, source: `${selection.provider}/${selection.model}` };
}

async function handleChat(res, body) {
  const text = String(body.text ?? '').trim();
  const run = startRun('chat', { text });
  const started = Date.now();
  const template = within(body.template, DEFAULTS.template);
  const source = body.source ? within(body.source, DEFAULTS.source) : null;
  const out = within(body.out, DEFAULTS.out);
  if (!text) {
    return json(res, 200, { ok: true, action: 'none', reply: '请说点什么。' });
  }

  const llm = hostContext && hostContext.get('llm');
  const selector = hostContext && hostContext.get('agentDefaultModel');
  let provider = null;
  let model = null;
  try {
    const selection = selector && selector.currentSelection();
    provider = selection && selection.provider;
    model = selection && selection.model;
  } catch {
    provider = null;
    model = null;
  }
  if (!llm || !provider || !model) {
    return json(res, 200, {
      ok: true,
      fallback: true,
      action: 'none',
      reply: 'Host 上没有可用的模型服务，已改用本地指令解析。',
    });
  }

  const system = [
    '你是 Word 流水线助手的调度器，负责把用户的话翻译成一个动作。',
    '可用动作：',
    '- convert：一键执行 填充→检查→渲染',
    '- fill / check / render：单独执行某一步',
    '- set_image_width：调整图片宽度，width 是厘米数字',
    '- none：只回复文字，不执行动作',
    '只输出一个 JSON 对象，不要代码块、不要多余文字：',
    '{"action":"convert|fill|check|render|set_image_width|none","width":数字可选,"reply":"给用户的中文一句话"}',
    '用户只是提问或闲聊时用 none，并在 reply 里直接回答。',
  ].join('\n');

  const history = Array.isArray(body.history) ? body.history.slice(-6) : [];
  const messages = history
    .filter((item) => item && item.role === 'user' && item.text)
    .map((item) => ({ role: 'user', content: [{ type: 'text', text: String(item.text) }] }));
  messages.push({ role: 'user', content: [{ type: 'text', text }] });

  let raw = '';
  try {
    const completed = await completeText(system, messages, 4000);
    raw = completed.raw || '';
  } catch (error) {
    return json(res, 200, {
      ok: false,
      fallback: true,
      action: 'none',
      reply: `模型调用失败：${error.message || error}`,
    });
  }
  const decisionMatch = raw.match(/\{[\s\S]*\}/);
  let decision = null;
  if (decisionMatch) {
    try { decision = JSON.parse(decisionMatch[0]); } catch { decision = null; }
  }
  if (!decision) {
    return json(res, 200, { ok: true, action: 'none', reply: raw.trim().slice(0, 800) || '模型没有返回可执行动作。' });
  }
  const action = String(decision.action || 'none');
  const width = Number(decision.width);
  logEvent(run, 'info', 'chat', `模型判定：${action}`, {
    provider,
    model,
    action,
    ms: Date.now() - started,
  });
  let result = null;
  try {
    if (action === 'convert') {
      result = await runConvert({ template, source, out });
    } else if (action === 'fill' || action === 'check' || action === 'render') {
      result = await executeOp({ op: action, template, source, out });
    } else if (action === 'set_image_width' && Number.isFinite(width)) {
      const corrections = path.join(ROOT, 'fixtures', 'build', 'agent-corrections.yaml');
      await fs.writeFile(
        corrections,
        `corrections:\n  - slot: ${lastImageSlot}\n    width: ${width}cm\n    note: 来自面板对话\n`,
        'utf8',
      );
      result = await executeOp({ op: 'apply', template, out, corrections });
    }
  } catch (error) {
    logEvent(run, 'error', 'chat', `执行失败：${error.message || error}`);
    endRun(run, 'fail', { action });
    return json(res, 200, {
      ok: false,
      action,
      reply: `执行失败：${error.message || error}`,
      runId: run.id,
      logs: run.events,
    });
  }

  endRun(run, 'ok', { action });
  return json(res, 200, {
    ok: true,
    action,
    reply: String(decision.reply || ''),
    result,
    runId: run.id,
    logs: run.events,
  });
}

async function handleDownload(res, url) {
  const target = within(url.searchParams.get('path'), '');
  const stat = await fs.stat(target);
  if (!stat.isFile()) return json(res, 400, { ok: false, error: '只能下载文件' });
  const extension = path.extname(target).toLowerCase();
  const types = {
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.pdf': 'application/pdf',
    '.png': 'image/png',
  };
  const name = path.basename(target);
  res.writeHead(200, {
    'Content-Type': types[extension] || 'application/octet-stream',
    'Content-Length': stat.size,
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(name)}`,
  });
  createReadStream(target).pipe(res);
}

async function handlePreview(res, url) {
  const target = within(url.searchParams.get('path'), '');
  const lowerPath = target.toLowerCase();
  if (lowerPath.endsWith('.png')) {
    const bytes = await fs.readFile(target);
    res.writeHead(200, { 'Content-Type': 'image/png', 'Cache-Control': 'no-store' });
    return res.end(bytes);
  }
  if (!lowerPath.endsWith('.docx')) {
    return json(res, 400, { ok: false, error: '只能预览 .docx 或 .png' });
  }
  const cached = previewManifest.get(path.resolve(target));
  if (cached && existsSync(cached)) {
    const [pngStat, docStat] = await Promise.all([fs.stat(cached), fs.stat(target)]);
    if (pngStat.mtimeMs >= docStat.mtimeMs) {
      const bytes = await fs.readFile(cached);
      res.writeHead(200, { 'Content-Type': 'image/png', 'Cache-Control': 'no-store' });
      return res.end(bytes);
    }
    previewManifest.delete(path.resolve(target));
  }
  const stat = await fs.stat(target);
  const key = createHash('sha1')
    .update(`${target}:${stat.mtimeMs}:${stat.size}`)
    .digest('hex')
    .slice(0, 20);
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  const png = path.join(PREVIEW_DIR, `${key}.png`);
  await ensurePreview(target, png);
  const bytes = await fs.readFile(png);
  res.writeHead(200, { 'Content-Type': 'image/png', 'Cache-Control': 'no-store' });
  res.end(bytes);
}

async function handleUpload(res, body) {
  const kind = String(body.kind ?? '');
  const allowed = KINDS[kind];
  if (!allowed) return json(res, 400, { ok: false, error: `未知上传类型: ${kind}` });

  const provided = String(body.name ?? 'file');
  const safe = provided.replace(/[^\w.\-\u4e00-\u9fff]+/g, '_').slice(-90) || 'file';
  const extension = path.extname(safe).toLowerCase();
  if (!allowed.includes(extension)) {
    return json(res, 400, { ok: false, error: `${kind} 只接受 ${allowed.join(' / ')}` });
  }
  const bytes = Buffer.from(String(body.data ?? ''), 'base64');
  if (bytes.length === 0) return json(res, 400, { ok: false, error: '文件为空' });

  const directory = path.join(ROOT, 'uploads', kind);
  await fs.mkdir(directory, { recursive: true });
  const target = path.join(directory, `${Date.now()}-${safe}`);
  await fs.writeFile(target, bytes);
  return json(res, 200, { ok: true, path: target, name: safe, bytes: bytes.length });
}

function loopback(req) {
  const address = req.socket && req.socket.remoteAddress;
  return address === '127.0.0.1' || address === '::1' || address === '::ffff:127.0.0.1';
}

function originAllowed(req) {
  const origin = req.headers.origin;
  if (!origin) return true; // same-origin GET, image loads, non-browser clients
  if (origin === 'null') return true; // file:// / dsh-app:// renderers
  try {
    const parsed = new URL(origin);
    return parsed.hostname === '127.0.0.1' || parsed.hostname === 'localhost';
  } catch {
    return false;
  }
}

function trusted(req, url) {
  const header = req.headers['x-word-token'];
  const query = url.searchParams.get('t');
  if (header === TOKEN || query === TOKEN) return true;
  // The desktop shell renders from its own scheme; loops back without a token.
  return loopback(req) && originAllowed(req);
}

async function handle(ctx, req, res) {
  const url = new URL(req.url ?? '/', 'http://localhost');
  const route = url.pathname.slice(BASE.length) || '/';
  res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Word-Token');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Vary', 'Origin');
  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }
  if (req.method === 'GET' && route === '/app') {
    const html = await fs.readFile(path.join(HERE, 'app.html'));
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
    return res.end(html);
  }
  if (!trusted(req, url)) {
    return json(res, 403, { ok: false, error: '请求来源不被允许' });
  }
  if (req.method === 'GET' && route === '/state') {
    return json(res, 200, { ok: true, root: ROOT, defaults: DEFAULTS, python: PYTHON, autoPanel: process.env.DSH_WORD_AUTO_PANEL === '1' });
  }
  if (req.method === 'GET' && route === '/progress') {
    return json(res, 200, { ok: true, ...progress, preview: livePreview.path });
  }
  if (req.method === 'GET' && route === '/runs') {
    const limit = Math.min(Math.max(Number(url.searchParams.get('limit')) || 20, 1), MAX_RUNS);
    const list = [...runs.values()]
      .sort((a, b) => b.startedAt - a.startedAt)
      .slice(0, limit)
      .map(runSummary);
    return json(res, 200, { ok: true, runs: list });
  }
  if (req.method === 'GET' && route === '/run') {
    const run = runs.get(String(url.searchParams.get('id') || ''));
    if (!run) return json(res, 404, { ok: false, error: '没有这个运行' });
    return json(res, 200, { ok: true, run });
  }
  if (req.method === 'GET' && route === '/model-usage') {
    return json(res, 200, { ok: true, usage: loadModelUsage() });
  }
  if (req.method === 'GET' && route === '/model-config') {
    const config = loadModelConfig();
    return json(res, 200, {
      ok: true,
      active: config.active || '',
      providers: (config.providers || []).map((item) => ({
        id: item.id, name: item.name, type: item.type, baseUrl: item.baseUrl,
        model: item.model, key: maskKey(item.apiKey), priceIn: item.priceIn || 0, priceOut: item.priceOut || 0,
      })),
    });
  }
  if (req.method === 'POST' && route === '/model-config') {
    const body = await readJson(req);
    const config = loadModelConfig();
    const providers = config.providers || [];
    if (body.action === 'delete') {
      config.providers = providers.filter((item) => item.id !== body.id);
      if (config.active === body.id) config.active = config.providers[0]?.id || '';
    } else {
      const id = body.id || `p${Date.now().toString(36)}`;
      const existing = providers.find((item) => item.id === id);
      const entry = {
        id,
        name: String(body.name || (existing && existing.name) || '自定义模型'),
        type: body.type === 'anthropic' ? 'anthropic' : 'openai',
        baseUrl: String(body.baseUrl || (existing && existing.baseUrl) || ''),
        model: String(body.model || (existing && existing.model) || ''),
        apiKey: String(body.apiKey || (existing && existing.apiKey) || ''),
        priceIn: Number(body.priceIn ?? (existing && existing.priceIn) ?? 0),
        priceOut: Number(body.priceOut ?? (existing && existing.priceOut) ?? 0),
      };
      if (existing) Object.assign(existing, entry);
      else providers.push(entry);
      config.providers = providers;
      if (body.activate !== false) config.active = id;
    }
    saveModelConfig(config);
    return json(res, 200, { ok: true, active: config.active });
  }
  if (req.method === 'POST' && route === '/model-test') {
    const body = await readJson(req);
    const config = loadModelConfig();
    const provider = (config.providers || []).find((item) => item.id === (body.id || config.active));
    if (!provider) return json(res, 200, { ok: false, error: '没有选择模型' });
    try {
      const reply = await callCustomProvider(provider, 'You are a test.', [{ role: 'user', content: [{ type: 'text', text: 'ping' }] }], 20);
      return json(res, 200, { ok: true, reply: String(reply).slice(0, 80) });
    } catch (error) {
      return json(res, 200, { ok: false, error: String(error.message || error) });
    }
  }
  if (req.method === 'GET' && route === '/models') {
    const llm = hostContext && hostContext.get('llm');
    let providers = [];
    try {
      providers = llm && typeof llm.listProviders === 'function' ? llm.listProviders() : [];
    } catch {
      providers = [];
    }
    let selection = null;
    try {
      const selector = hostContext && hostContext.get('agentDefaultModel');
      selection = selector ? selector.currentSelection() : null;
    } catch {
      selection = null;
    }
    return json(res, 200, { ok: true, providers, selection });
  }
  if (req.method === 'GET' && route === '/result') {
    const out = DEFAULTS.out;
    const issues = await readCheckIssues(out);
    return json(res, 200, {
      ok: true,
      docx: out,
      pdf: path.join(path.dirname(out), 'panel-preview.pdf'),
      png: path.join(path.dirname(out), 'panel-preview.png'),
      issues,
    });
  }
  if (req.method === 'GET' && route === '/mapping') {
    const id = String(url.searchParams.get('id') || '');
    let debug = id ? mappingDebug.get(id) : null;
    if (!debug && mappingDebug.size) debug = [...mappingDebug.values()].pop();
    return json(res, 200, { ok: Boolean(debug), debug: debug || null });
  }
  if (req.method === 'POST' && route === '/upload') {
    return handleUpload(res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/op') {
    return handleOp(res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/agent/instruct') {
    return handleInstruct(res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/agent/chat') {
    return handleChat(res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/convert') {
    return handleConvert(res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/agent/run-stream') {
    return handleAgentRunStream(req, res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/convert-stream') {
    return handleConvertStream(req, res, await readJson(req));
  }
  if (req.method === 'POST' && route === '/live-start') {
    const body = await readJson(req);
    const template = within(body.template, DEFAULTS.template);
    const target = path.join(path.dirname(DEFAULTS.out), 'live-preview.docx');
    try {
      await fs.copyFile(template, target);
    } catch (error) {
      return json(res, 400, { ok: false, error: `复制模板失败：${error.message || error}` });
    }
    const opened = await runCli(['live-open', '--file', target]);
    liveTarget = opened.code === 0 ? target : '';
    return json(res, 200, {
      ok: opened.code === 0,
      path: target,
      error: opened.code === 0 ? '' : opened.stderr.trim(),
    });
  }
  if (req.method === 'POST' && route === '/open-word') {
    const body = await readJson(req);
    const requested = within(body.path, DEFAULTS.out);
    if (!existsSync(requested)) {
      return json(res, 400, { ok: false, error: '文件还不存在，请先转换' });
    }
    const opened = await runCli(['live-open', '--file', requested]);
    liveTarget = opened.code === 0 ? requested : '';
    return json(res, 200, {
      ok: opened.code === 0,
      path: requested,
      live: Boolean(liveTarget),
      error: opened.code === 0 ? '' : opened.stderr.trim(),
    });
  }
  if (req.method === 'GET' && route === '/preview') {
    return handlePreview(res, url);
  }
  if (req.method === 'GET' && route === '/download') {
    return handleDownload(res, url);
  }
  return json(res, 404, { ok: false, error: '未知接口' });
}

/**
 * 把 plugin/skills 下的技能注册进 DSH 技能目录。
 *
 * DSH 不会自动扫描插件包的 skills 目录，必须通过 skills 缝注册 provider；
 * 宿主没有该缝（老版本）时静默降级，插件其余功能不受影响。
 */
function registerSkillProvider(ctx, enabled = true) {
  if (!enabled) return;
  if (typeof ctx?.inject !== 'function') return;
  try {
    ctx.inject(['skills'], (skillCtx) => {
      try {
        skillCtx.skills.registerProvider(() => skillProvider);
        ctx.logger?.info?.('[word-pipeline] 已注册技能：docx-retypeset、word-pipeline');
      } catch (error) {
        ctx.logger?.warn?.(`[word-pipeline] 技能注册失败：${error?.message ?? error}`);
      }
    });
  } catch (error) {
    ctx.logger?.debug?.(`[word-pipeline] skills 缝不可用，跳过技能注册：${error?.message ?? error}`);
  }
}

export function apply(ctx, config = {}) {
  hostContext = ctx;
  registerSkillProvider(ctx, config?.skill !== false);
  // Debug aid: persist the advertised (authenticated) GUI URL so local tooling
  // such as the UI debugger can open the running page.
  try {
    const webApp = ctx.get('webApp');
    if (webApp && typeof webApp.getAdvertisedUrl === 'function') {
      fs.writeFile(
        path.join(ROOT, 'fixtures', 'build', 'web-url.txt'),
        String(webApp.getAdvertisedUrl()),
      ).catch(() => {});
    }
  } catch {
    /* the web app service is optional */
  }

  ctx.effect(() =>
    ctx.webServer.register({
      kind: 'prefix',
      path: BASE,
      handler: (req, res) =>
        handle(ctx, req, res).catch((error) =>
          json(res, 500, { ok: false, error: String(error?.message ?? error) }),
        ),
    }),
  );

  ctx.on('webserver/index-inject', (table) => {
    table.push({
      kind: 'global',
      name: '__DSH_WORD_PIPELINE__',
      value: { token: TOKEN },
    });
  });
}
