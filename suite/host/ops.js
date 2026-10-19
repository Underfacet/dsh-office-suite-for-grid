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
});

/** 面板可写配置键（suite.config.set 白名单，§2）。 */
export const WRITABLE_CONFIG_KEYS = Object.freeze([
  'outputRoot',
  'parallel',
  'pdfPaper',
  'sortMode',
  'allowExternalPaths',
  'odaPath',
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

/**
 * 校验单个可写配置项；仅做语法/取值范围校验（路径是否越界由命令层经 boundary 校验）。
 * @returns {{ok:true,value:any}|{ok:false,code:string,message:string}}
 */
export function validateConfigEntry(key, value) {
  if (!WRITABLE_CONFIG_KEYS.includes(key)) {
    return {
      ok: false,
      code: 'CONFIG_INVALID_KEY',
      message: `不支持的配置键: ${String(key)}（可写键：${WRITABLE_CONFIG_KEYS.join(', ')}）`,
    };
  }
  const invalid = (expect) => ({ ok: false, code: 'CONFIG_INVALID_VALUE', message: `配置项 ${key} 期望 ${expect}，收到: ${JSON.stringify(value)}` });
  switch (key) {
    case 'outputRoot': {
      if (typeof value !== 'string') return invalid('字符串');
      const trimmed = value.trim();
      if (trimmed && !isAbsolute(trimmed)) return invalid('绝对路径（或空字符串 = 会话工作区 output\\pdf）');
      return { ok: true, value: trimmed };
    }
    case 'parallel': {
      if (value === 'auto') return { ok: true, value: 'auto' };
      if (Number.isInteger(value) && value >= 1 && value <= 8) return { ok: true, value };
      return invalid("'auto' 或 1–8 的整数");
    }
    case 'pdfPaper': {
      if (value === 'A4-landscape' || value === 'keep') return { ok: true, value };
      return invalid("'A4-landscape' | 'keep'");
    }
    case 'sortMode': {
      if (value === 'drawingNo' || value === 'filename') return { ok: true, value };
      return invalid("'drawingNo' | 'filename'");
    }
    case 'allowExternalPaths': {
      if (typeof value === 'boolean') return { ok: true, value };
      return invalid('布尔值');
    }
    case 'odaPath': {
      if (typeof value !== 'string') return invalid('字符串（ODAFileConverter.exe 路径）');
      const trimmed = value.trim();
      if (trimmed && !isAbsolute(trimmed)) return invalid('绝对路径（或空字符串 = 自动探测）');
      return { ok: true, value: trimmed };
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
