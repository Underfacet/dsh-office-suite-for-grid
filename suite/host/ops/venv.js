/**
 * Python venv 生命周期（docs/P2-设计.md §3.5）
 *
 * 目录：
 *   runtime\venv\             Python 环境（python -m venv 创建）
 *   runtime\venv\SUITE_VERSION 当前套件版本标记（升级后自动重建依赖）
 *   runtime\venv.lock\        并发锁（目录存在=持锁；mkdir 原子；>120s 视为残留可抢占）
 *   runtime\wheels\           离线 wheel 缓存（存在则 pip --find-links）
 *
 * 依赖锁定：suite\requirements-office.txt（python-docx==1.2.0 / openpyxl==3.1.5）。
 * 失败降级链（提示文案）：便携 Python 3.12（config.venvPython 指向）→ runtime\wheels 离线 wheel → E2002/E2003。
 *
 * 导出：ensureVenv(ctx, config, exec) → {ok:true,result:{python,reason,...}} | {ok:false,error}
 *       venvStatus() → 同步状态（probe/诊断页用）
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, utimesSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { PACKAGE_DIR, RUNTIME_DIR, fail, ok, parseJson, toErrorPayload } from '../ops.js';
import { runPs1 } from '../engines.js';

/** venv 目录（runtime\venv）。 */
export const VENV_DIR = join(RUNTIME_DIR, 'venv');
/** 并发锁目录（runtime\venv.lock）。 */
export const VENV_LOCK_DIR = join(RUNTIME_DIR, 'venv.lock');
/** 离线 wheel 目录（runtime\wheels）。 */
export const WHEELS_DIR = join(RUNTIME_DIR, 'wheels');
/** 顶层依赖锁定文件（随包）。 */
export const REQUIREMENTS_PATH = join(PACKAGE_DIR, 'requirements-office.txt');

/** 锁超时（120s 后视为残留）。 */
const LOCK_STALE_MS = 120_000;
/** 等待锁的总超时。 */
const LOCK_WAIT_MS = 120_000;
/** venv 创建 + 装包超时。 */
const VENV_SETUP_TIMEOUT_MS = 600_000;

const VENV_PYTHON = join(VENV_DIR, 'Scripts', 'python.exe');
const VERSION_MARKER = join(VENV_DIR, 'SUITE_VERSION');

function readSuiteVersion() {
  try {
    const pkg = JSON.parse(readFileSync(join(PACKAGE_DIR, 'package.json'), 'utf8'));
    return typeof pkg?.version === 'string' ? pkg.version : '';
  } catch {
    return '';
  }
}

function readText(path) {
  try {
    return readFileSync(path, 'utf8').replace(/^\uFEFF/, '');
  } catch {
    return '';
  }
}

/** 从 pyvenv.cfg 读取 Python 版本。 */
function readPythonVersion(venvDir) {
  const cfg = readText(join(venvDir, 'pyvenv.cfg'));
  const match = cfg.match(/^\s*version(?:_info)?\s*=\s*([^\r\n]+)/im);
  return match ? match[1].trim() : '';
}

/** 从 site-packages 的 dist-info 目录读取已安装版本。 */
function readPackageVersion(venvDir, distribution) {
  const site = join(venvDir, 'Lib', 'site-packages');
  let names;
  try {
    names = readdirSync(site);
  } catch {
    return null;
  }
  const prefix = `${distribution.replace(/-/g, '_')}-`;
  for (const name of names) {
    if (!name.toLowerCase().endsWith('.dist-info')) continue;
    if (!name.toLowerCase().startsWith(prefix.toLowerCase())) continue;
    const rest = name.slice(prefix.length, name.length - '.dist-info'.length);
    return rest.split('-')[0] || null;
  }
  return null;
}

/**
 * 同步读取 venv 状态（不启动子进程；healthy = 版本标记匹配 + 两个依赖的 dist-info 存在）。
 * @param {{venvDir?:string}} [options]
 */
export function venvStatus(options = {}) {
  const dir = options.venvDir ?? VENV_DIR;
  const pythonExe = join(dir, 'Scripts', 'python.exe');
  const present = existsSync(pythonExe);
  const expectedVersion = readSuiteVersion();
  const marker = readText(join(dir, 'SUITE_VERSION')).trim();
  const docx = readPackageVersion(dir, 'python-docx');
  const openpyxl = readPackageVersion(dir, 'openpyxl');
  const healthy = present && !!marker && marker === expectedVersion && !!docx && !!openpyxl;
  const stat = present ? (() => { try { return statSync(pythonExe); } catch { return null; } })() : null;
  return {
    present,
    dir,
    healthy,
    python: present ? pythonExe : null,
    pythonVersion: present ? readPythonVersion(dir) : '',
    packages: { docx: docx ?? null, openpyxl: openpyxl ?? null },
    suiteVersion: expectedVersion,
    versionMarker: marker || null,
    markerMatches: !!expectedVersion && marker === expectedVersion,
    updatedAt: stat ? stat.mtime.toISOString() : null,
  };
}

function isReady(expectedVersion) {
  const status = venvStatus();
  return status.present && status.markerMatches && status.pythonVersion && status.packages.docx && status.packages.openpyxl
    && status.versionMarker === expectedVersion;
}

/** 尝试创建锁目录；返回 true=已持锁。 */
function tryAcquireLock() {
  try {
    mkdirSync(VENV_LOCK_DIR, { recursive: false });
    return true;
  } catch (err) {
    if (err?.code === 'EEXIST') return false;
    // 目录不存在就创建父目录后重试一次
    try {
      mkdirSync(RUNTIME_DIR, { recursive: true });
      mkdirSync(VENV_LOCK_DIR, { recursive: false });
      return true;
    } catch {
      return false;
    }
  }
}

function lockAgeMs() {
  try {
    return Date.now() - statSync(VENV_LOCK_DIR).mtimeMs;
  } catch {
    return 0;
  }
}

function releaseLock() {
  try {
    rmSync(VENV_LOCK_DIR, { recursive: true, force: true });
  } catch { /* 释放失败忽略（残留锁 120s 后可抢占） */ }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * 获取 venv 锁（mkdir 原子）。残留锁（>120s）自动抢占；总等待超 120s 抛 VENV_BUSY。
 */
async function acquireLock(signal) {
  const started = Date.now();
  for (;;) {
    if (signal?.aborted) {
      const err = new Error('任务已取消');
      err.name = 'AbortError';
      throw err;
    }
    if (tryAcquireLock()) return;
    if (lockAgeMs() > LOCK_STALE_MS) {
      // 残留锁：抢占
      try { rmSync(VENV_LOCK_DIR, { recursive: true, force: true }); } catch { /* 忽略 */ }
      if (tryAcquireLock()) return;
    }
    if (Date.now() - started > LOCK_WAIT_MS) {
      const err = new Error('venv 正被其它任务创建/安装，等待超时');
      err.code = 'VENV_BUSY';
      err.hint = `如确认无其它任务在运行，可删除残留锁目录后重试：${VENV_LOCK_DIR}`;
      throw err;
    }
    await sleep(500);
  }
}

/** 解析建 venv 用的基础 Python：config.venvPython 优先，其次 pythonPath，最后 PATH 中的 python。 */
function resolveBasePython(config = {}) {
  const explicit = typeof config.venvPython === 'string' ? config.venvPython.trim() : '';
  if (explicit) return explicit;
  const fallback = typeof config.pythonPath === 'string' ? config.pythonPath.trim() : '';
  if (fallback) return fallback;
  return 'python';
}

function setupHint(extra) {
  const parts = [
    extra,
    '可安装便携 Python 3.12 后用配置 venvPython 指向其 python.exe',
    `离线环境可把 python-docx/openpyxl/lxml 的 wheel 放入 ${WHEELS_DIR} 后重试`,
  ].filter(Boolean);
  return parts.join('；');
}

/**
 * 确保 venv 可用（创建/装包/版本升级重建）。
 * @param {object} ctx DSH 插件上下文（需 subprocess）
 * @param {object} config 有效配置
 * @param {object} exec 工具/命令执行上下文（agent/signal）
 * @returns {Promise<{ok:true,result:{python:string,reason:string,[key:string]:any}}|{ok:false,error:{code:string,message:string,hint?:string}}>}
 */
export async function ensureVenv(ctx, config = {}, exec = {}) {
  const expectedVersion = readSuiteVersion();
  const signal = exec?.signal;
  try {
    const force = config.venvForceRebuild === true;
    if (!force && isReady(expectedVersion)) {
      const status = venvStatus();
      return ok({ python: status.python, reason: 'cached', ...status });
    }

    await acquireLock(signal);
    // 心跳：长时间 pip 安装期间刷新锁目录 mtime，避免被其它进程当作残留锁抢占
    const heartbeat = setInterval(() => {
      try {
        const now = new Date();
        utimesSync(VENV_LOCK_DIR, now, now);
      } catch { /* 锁已被释放 */ }
    }, 30_000);
    if (typeof heartbeat.unref === 'function') heartbeat.unref();
    try {
      // 二次检查（防并发重复创建）
      if (!force && isReady(expectedVersion)) {
        const status = venvStatus();
        return ok({ python: status.python, reason: 'cached-after-lock', ...status });
      }

      const python = resolveBasePython(config);
      const pyenvResult = await runPs1(ctx, 'pyenv.ps1', {
        venvPath: VENV_DIR,
        requirements: REQUIREMENTS_PATH,
        python,
        wheelsDir: WHEELS_DIR,
        suiteVersion: expectedVersion,
      }, {
        signal,
        timeoutMs: Number.isFinite(config.venvSetupTimeoutMs) ? config.venvSetupTimeoutMs : VENV_SETUP_TIMEOUT_MS,
        graceMs: 180_000,
        cwd: dirname(VENV_DIR),
      });

      const parsed = parseJson(pyenvResult.stdout);
      if (!parsed) {
        return fail(
          'E2002',
          `venv 创建/装包失败（pyenv.ps1 未返回可解析 JSON）: ${(pyenvResult.stderr || pyenvResult.stdout || '').slice(0, 500)}`,
          setupHint(`所用 Python：${python}`),
        );
      }
      if (!parsed.ok) {
        const code = parsed.error?.code === 'E2003' ? 'E2003' : (parsed.error?.code ?? 'E2002');
        return fail(
          code,
          String(parsed.error?.message ?? 'venv 创建失败'),
          setupHint(parsed.error?.hint ?? `所用 Python：${python}`),
        );
      }

      const status = venvStatus();
      if (!status.present) {
        return fail('E2002', 'venv 创建后未找到 python.exe', setupHint(`目录：${VENV_DIR}`));
      }
      return ok({
        python: status.python,
        reason: force ? 'rebuilt' : (parsed.result?.created === true ? 'created' : 'repaired'),
        pythonVersion: parsed.result?.pythonVersion ?? status.pythonVersion,
        packages: parsed.result?.packages ?? status.packages,
        installMode: parsed.result?.installMode ?? null,
        suiteVersion: expectedVersion,
        dir: VENV_DIR,
        warnings: Array.isArray(parsed.result?.warnings) ? parsed.result.warnings : [],
      });
    } finally {
      clearInterval(heartbeat);
      releaseLock();
    }
  } catch (err) {
    if (err?.name === 'AbortError') throw err;
    const payload = toErrorPayload(err, 'E2002');
    const code = payload.code === 'VENV_BUSY' ? 'VENV_BUSY' : (payload.code || 'E2002');
    return fail(code, payload.message, payload.hint ?? setupHint(''));
  }
}

/** 删除 venv（诊断页「重建」可先清理；P2 不默认自动调用）。 */
export async function removeVenv() {
  await rm(VENV_DIR, { recursive: true, force: true });
  return ok({ removed: VENV_DIR });
}
