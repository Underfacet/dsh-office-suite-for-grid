/**
 * 统一操作层（v0.1 既有 + P1 增量）：所有操作（对话工具、远程命令、面板）返回同一结果形状。
 * 成功：{ ok: true, result: ... }
 * 失败：{ ok: false, error: { code, message, hint? } }
 *
 * P1 增量（不改变既有导出与行为）：
 * - OpsError / toErrorPayload / guarded：结构化错误（错误码见 docs/P1-设计.md §3.8）；
 * - 运行配置：DEFAULT_CONFIG（§3.7）、runtime\ui-config.json 覆盖层、白名单键校验；
 * - 运行时目录常量：runtime / runtime\tmp / runtime\logs / runtime\backups（§5.3）。
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { cpus } from 'node:os';
import { resolve, isAbsolute, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { FRAME_SIZES, DEFAULT_DRAWING_NO_PATTERNS, DEFAULT_TITLE_BLOCK_ZONE, DEFAULT_FRAME_BLOCK_NAMES } from './ops/drawingNo.js';

const HOST_DIR = dirname(fileURLToPath(import.meta.url));

/** 插件包目录（suite/）。 */
export const PACKAGE_DIR = resolve(HOST_DIR, '..');
/** 项目根目录（含 docs/ tests/ runtime/ output/）。 */
export const PROJECT_DIR = resolve(PACKAGE_DIR, '..');
/** 运行时目录（git 忽略）。 */
export const RUNTIME_DIR = join(PROJECT_DIR, 'runtime');
export const TMP_DIR = join(RUNTIME_DIR, 'tmp');
export const LOG_DIR = join(RUNTIME_DIR, 'logs');
export const BACKUP_DIR = join(RUNTIME_DIR, 'backups');
/** 面板可写配置覆盖层。 */
export const UI_CONFIG_PATH = join(RUNTIME_DIR, 'ui-config.json');

/** 从 PowerShell 输出中提取首个 JSON 对象（容忍前导警告行）。 */
export function parseJson(text) {
  if (!text) return undefined;
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) return undefined;
  try {
    return JSON.parse(text.slice(start, end + 1));
  } catch {
    return undefined;
  }
}

/** 当前会话的工作目录（无会话时退回进程目录）。 */
export async function sessionCwd(ctx, exec) {
  if (exec?.agent) {
    try {
      return await ctx.workingDirectory.ensure(exec.agent, exec.signal);
    } catch { /* 回退 */ }
  }
  return process.cwd();
}

/** 绝对路径原样返回；相对路径按会话工作目录解析。 */
export async function resolveInputPath(ctx, exec, p) {
  if (isAbsolute(p)) return p;
  return resolve(await sessionCwd(ctx, exec), p);
}

export const dirnameOf = (p) => dirname(p);

export const ok = (result) => ({ ok: true, result });
export const fail = (code, message, hint) => ({ ok: false, error: { code, message, ...(hint ? { hint } : {}) } });

/**
 * 操作层结构化错误：throw 后由 guarded()/命令层转成 { ok:false, error }。
 * code 采用 §3.8 错误码（如 E1021 PATH_OUTSIDE_ROOT）。
 */
export class OpsError extends Error {
  constructor(code, message, hint) {
    super(message);
    this.name = 'OpsError';
    this.code = code;
    if (hint) this.hint = hint;
  }
}

/** 判断是否取消（AbortSignal / DOMException AbortError）。 */
export function isAbortError(err) {
  return err?.name === 'AbortError' || err?.code === 'ABORT_ERR' || err?.code === 'E1098';
}

/** 把任意异常规整为 { code, message, hint? }。 */
export function toErrorPayload(err, fallbackCode = 'E1099') {
  if (isAbortError(err)) return { code: 'E1098', message: '任务已取消' };
  if (err && typeof err === 'object' && typeof err.code === 'string' && err.code) {
    return { code: err.code, message: String(err.message ?? ''), ...(err.hint ? { hint: String(err.hint) } : {}) };
  }
  return { code: fallbackCode, message: String(err?.message ?? err ?? '未知错误') };
}

/** 执行操作并捕获异常 → 统一结果形状。取消错误原样抛出，由任务层处理。 */
export async function guarded(fn, fallbackCode = 'E1099') {
  try {
    return await fn();
  } catch (err) {
    if (isAbortError(err)) throw err;
    return { ok: false, error: toErrorPayload(err, fallbackCode) };
  }
}

/** 取消检查：任务流水线每一步之前调用。 */
export function throwIfAborted(signal) {
  if (signal?.aborted) {
    const err = new Error('任务已取消');
    err.name = 'AbortError';
    throw err;
  }
}

/** 中止态错误（供 AbortController.abort(reason) 使用）。 */
export function abortError(message = '任务已取消') {
  const err = new Error(message);
  err.name = 'AbortError';
  return err;
}

// ───────────────────────────── 配置（§3.7） ─────────────────────────────

/** 默认配置：与 cordis.patch.yml 的 config 以及 §3.7 表格一致。 */
export const DEFAULT_CONFIG = Object.freeze({
  projectRoot: '',
  outputRoot: '',
  backupCopies: 3,
  parallel: 'auto',
  pythonPath: '',
  /** P2 §7：Office 编辑/排版引擎（auto = COM 优先 → Python venv）。 */
  officeEngine: 'auto',
  /** P2 §7：显式指定建 venv 用的 python.exe；空 = 自动探测。 */
  venvPython: '',
  /** P2 §7：officeToPdf.convert 渲染优先级。 */
  previewPriority: 'foreground',
  /** P2 §7：用户模板/样式集目录；空 = runtime\templates。 */
  officeTemplatesDir: '',
  visioVisible: false,
  allowExternalPaths: false,
  pdfPaper: 'A4-landscape',
  sortMode: 'drawingNo',
  frameBlockNames: [...DEFAULT_FRAME_BLOCK_NAMES],
  frameSizes: FRAME_SIZES,
  frameToleranceMm: 2,
  drawingNoPatterns: [...DEFAULT_DRAWING_NO_PATTERNS],
  titleBlockZone: [...DEFAULT_TITLE_BLOCK_ZONE],
  maxConcurrentPdf: 4,
  cadTimeoutMs: 300_000,
  odaPath: '',
  odaVersion: 'ACAD2018',
  odaRecurse: true,
  keepTempFiles: false,
  logLevel: 'info',
  // ── P3 §5.4 / §2.3 / §2.1 / §3.1 / §4.4⑤ 增量（默认值与各 ops 模块一致） ──
  /** P3 §5.4：客户端预览 blob 缓存条数（LRU，1–10；可写）。 */
  previewCacheSize: 5,
  /** P3 §2.3：图纸目录表图名区相对坐标 [x0,y0,x1,y1]（原点左下；不可写，随包默认）。 */
  catalogTitleNameZone: [0.3, 0.25, 1.0, 0.55],
  /**
   * P3 §2.3：比例提取正则（不可写，随包默认）。
   * 第一则「比例: x:x」，第二则全框裸写比号（排除图号连字符干扰）；与
   * suite/host/ops/cad-catalog.js 的 DEFAULT_CATALOG_SCALE_PATTERNS 保持逐字一致。
   */
  catalogScalePatterns: [
    '比例\\s*[:：]?\\s*(\\d{1,3}\\s*[:：]\\s*\\d{1,4})',
    '(?<![\\dA-Za-z])(\\d{1,3}\\s*[:：]\\s*\\d{1,4})(?![\\dA-Za-z])',
  ],
  /** P3 §2.1：目录表单次图框上限（同步命令防长阻塞；不可写）。 */
  catalogMaxFrames: 500,
  /** P3 §3.1：CAD 批量队列条目上限（1–100；不可写）。 */
  queueMaxItems: 100,
  /** P3 §4.4⑤：端子排每页行数（1–40；模型 meta.terminalsPerPage 优先；不可写）。 */
  terminalPerPage: 20,
});

/** 面板可写配置键（suite.config.set 白名单，§2）。 */
export const WRITABLE_CONFIG_KEYS = Object.freeze([
  'outputRoot',
  'parallel',
  'pdfPaper',
  'sortMode',
  'allowExternalPaths',
  'odaPath',
  // ── P2 §7 增量 ──
  'officeEngine',
  'venvPython',
  'previewPriority',
  'officeTemplatesDir',
  // ── P3 §5.4 增量（P3 §6 标注可写者仅此一项） ──
  'previewCacheSize',
]);

/** 同步读取 runtime\ui-config.json 覆盖层；缺失/损坏时返回 {}。 */
export function loadUiConfig() {
  try {
    const raw = readFileSync(UI_CONFIG_PATH, 'utf8');
    const value = JSON.parse(raw);
    return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
  } catch {
    return {};
  }
}

/** 同步写入 runtime\ui-config.json（目录不存在则创建）。 */
export function saveUiConfig(overrides) {
  mkdirSync(RUNTIME_DIR, { recursive: true });
  writeFileSync(UI_CONFIG_PATH, JSON.stringify(overrides ?? {}, null, 2), 'utf8');
  return UI_CONFIG_PATH;
}

/** 有效配置 = 默认值 ← 插件 Config ← UI 覆盖层（后写覆盖前写）。 */
export function effectiveConfig(pluginConfig = {}) {
  const overrides = loadUiConfig();
  const merged = { ...DEFAULT_CONFIG };
  for (const [key, value] of Object.entries(pluginConfig ?? {})) {
    if (value !== undefined) merged[key] = value;
  }
  for (const [key, value] of Object.entries(overrides)) {
    if (value !== undefined) merged[key] = value;
  }
  return merged;
}

/** P3 §6 标注「不可写（随包默认，真机校准后改默认值）」的只读配置键（取值仍走同一校验口径）。 */
export const READONLY_CONFIG_KEYS = Object.freeze([
  'catalogTitleNameZone',
  'catalogScalePatterns',
  'catalogMaxFrames',
  'queueMaxItems',
  'terminalPerPage',
]);

/**
 * 校验单个配置项；仅做语法/取值范围校验（路径是否越界由命令层经 boundary 校验）。
 * 只读键（P3 §6）同样做取值校验，但合法值最终仍以 CONFIG_INVALID_KEY 拒绝写入——
 * 这样「值非法」（CONFIG_INVALID_VALUE）与「键只读」（CONFIG_INVALID_KEY）两类错误各自明确。
 * @returns {{ok:true,value:any}|{ok:false,code:string,message:string}}
 */
export function validateConfigEntry(key, value) {
  if (!WRITABLE_CONFIG_KEYS.includes(key) && !READONLY_CONFIG_KEYS.includes(key)) {
    return {
      ok: false,
      code: 'CONFIG_INVALID_KEY',
      message: `不支持的配置键: ${String(key)}（可写键：${WRITABLE_CONFIG_KEYS.join(', ')}）`,
    };
  }
  const invalid = (expect) => ({ ok: false, code: 'CONFIG_INVALID_VALUE', message: `配置项 ${key} 期望 ${expect}，收到: ${JSON.stringify(value)}` });
  /** 取值校验通过后：可写键原样返回；只读键拒绝写入（值本身仍是合法的）。 */
  const accept = (normalized) => {
    if (WRITABLE_CONFIG_KEYS.includes(key)) return { ok: true, value: normalized };
    return {
      ok: false,
      code: 'CONFIG_INVALID_KEY',
      message: `配置键 ${key} 为只读（P3 §6：随包默认，真机校准后改默认值）；可写键：${WRITABLE_CONFIG_KEYS.join(', ')}`,
    };
  };
  switch (key) {
    case 'outputRoot': {
      if (typeof value !== 'string') return invalid('字符串');
      const trimmed = value.trim();
      if (trimmed && !isAbsolute(trimmed)) return invalid('绝对路径（或空字符串 = 会话工作区 output\\pdf）');
      return accept(trimmed);
    }
    case 'parallel': {
      if (value === 'auto') return accept('auto');
      if (Number.isInteger(value) && value >= 1 && value <= 8) return accept(value);
      return invalid("'auto' 或 1–8 的整数");
    }
    case 'pdfPaper': {
      if (value === 'A4-landscape' || value === 'keep') return accept(value);
      return invalid("'A4-landscape' | 'keep'");
    }
    case 'sortMode': {
      if (value === 'drawingNo' || value === 'filename') return accept(value);
      return invalid("'drawingNo' | 'filename'");
    }
    case 'allowExternalPaths': {
      if (typeof value === 'boolean') return accept(value);
      return invalid('布尔值');
    }
    case 'odaPath': {
      if (typeof value !== 'string') return invalid('字符串（ODAFileConverter.exe 路径）');
      const trimmed = value.trim();
      if (trimmed && !isAbsolute(trimmed)) return invalid('绝对路径（或空字符串 = 自动探测）');
      return accept(trimmed);
    }
    case 'officeEngine': {
      if (value === 'auto' || value === 'com' || value === 'python') return accept(value);
      return invalid("'auto' | 'com' | 'python'");
    }
    case 'venvPython': {
      if (typeof value !== 'string') return invalid('字符串（python.exe 路径）');
      const trimmed = value.trim();
      if (trimmed && !isAbsolute(trimmed)) return invalid('绝对路径（或空字符串 = 自动探测本机 Python）');
      return accept(trimmed);
    }
    case 'previewPriority': {
      if (value === 'foreground' || value === 'background') return accept(value);
      return invalid("'foreground' | 'background'");
    }
    case 'officeTemplatesDir': {
      if (typeof value !== 'string') return invalid('字符串（自定义模板/样式集目录）');
      const trimmed = value.trim();
      if (trimmed && !isAbsolute(trimmed)) return invalid('绝对路径（或空字符串 = runtime\\templates）');
      return accept(trimmed);
    }
    // ── P3 §5.4 / §2 / §3 增量 ──
    case 'previewCacheSize': {
      if (Number.isInteger(value) && value >= 1 && value <= 10) return accept(value);
      return invalid('1–10 的整数（预览 blob 缓存条数）');
    }
    // 以下五个键 P3 §6 标注为「不可写（随包默认）」：取值照常校验，写入仍被 accept 拒绝。
    case 'catalogMaxFrames': {
      if (Number.isInteger(value) && value >= 1 && value <= 5000) return accept(value);
      return invalid('1–5000 的整数（目录表单次图框上限）');
    }
    case 'queueMaxItems': {
      if (Number.isInteger(value) && value >= 1 && value <= 100) return accept(value);
      return invalid('1–100 的整数（CAD 批量队列条目上限）');
    }
    case 'terminalPerPage': {
      if (Number.isInteger(value) && value >= 1 && value <= 40) return accept(value);
      return invalid('1–40 的整数（端子排每页行数）');
    }
    case 'catalogTitleNameZone': {
      if (!Array.isArray(value) || value.length !== 4) return invalid('[x0,y0,x1,y1] 四个 0–1 数字');
      const zone = value.map((item) => Number(item));
      if (zone.some((item) => !Number.isFinite(item) || item < 0 || item > 1)) return invalid('[x0,y0,x1,y1] 四个 0–1 数字');
      if (!(zone[0] < zone[2]) || !(zone[1] < zone[3])) return invalid('x0 < x1 且 y0 < y1 的 0–1 数字四元组');
      return accept(zone);
    }
    case 'catalogScalePatterns': {
      if (!Array.isArray(value) || !value.length) return invalid('非空正则字符串数组');
      const patterns = [];
      for (const item of value) {
        if (typeof item !== 'string' || !item.trim()) return invalid('非空正则字符串数组');
        try {
          new RegExp(item, 'i');
        } catch {
          return invalid(`可编译的正则（无法编译: ${item}）`);
        }
        patterns.push(item);
      }
      return accept(patterns);
    }
    default:
      return { ok: false, code: 'CONFIG_INVALID_KEY', message: `不支持的配置键: ${key}` };
  }
}

/** 解析 parallel 配置 → 实际并发度（auto = min(cpus-1, maxConcurrentPdf)，上限 8，≥1）。 */
export function resolveParallel(config = {}) {
  const value = config.parallel;
  if (Number.isInteger(value)) return Math.min(Math.max(value, 1), 8);
  const logical = (() => {
    try {
      const list = cpus();
      return list.length || 4;
    } catch {
      return 4;
    }
  })();
  const cap = Number.isFinite(config.maxConcurrentPdf) ? config.maxConcurrentPdf : 4;
  return Math.min(Math.max(Math.min(logical - 1, cap), 1), 8);
}
