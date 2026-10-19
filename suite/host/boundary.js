/**
 * 2026 边界校验（docs/P1-设计.md §5.3）
 *
 * 规则：
 * - 边界根 resolveBoundaryRoot() = config.projectRoot（非空时）或插件包目录上溯两级的 `E:\工作\2026`；
 * - 所有 ops 的输入读取与输出写入路径必须先过 assertInsideRoot(abs, { write })；
 * - 越界抛 OpsError('E1021', …)（= §3.8 PATH_OUTSIDE_ROOT）；
 * - allowExternalPaths:true 只放宽**读取**（输出仍限边界内），越界读取通过 warn 记录告警。
 */
import { dirname, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { OpsError } from './ops.js';

const HERE = dirname(fileURLToPath(import.meta.url));

/** 插件包目录（suite/）。 */
export const PACKAGE_DIR = resolve(HERE, '..');

/** 项目根目录（含 docs/ tests/ runtime/ output/ 的仓库根，= suite 的上一级）。 */
export const PROJECT_DIR = resolve(PACKAGE_DIR, '..');

/**
 * 默认边界根：项目根目录上溯两级（= E:\工作\2026，本项目所在年份目录）。
 * 与设计 §5.3 一致：写盘仅限 E:\工作\2026 之内；config.projectRoot 可覆盖。
 */
export const DEFAULT_BOUNDARY_ROOT = resolve(PROJECT_DIR, '..', '..');

/** 大小写不敏感比较用的归一化路径（Windows）。 */
function normalizeForCompare(path) {
  let value = resolve(path);
  while (value.length > 3 && (value.endsWith('\\') || value.endsWith('/'))) value = value.slice(0, -1);
  return value.toLowerCase();
}

/**
 * 计算边界根：config.projectRoot 非空时使用，否则默认 `dirname(dirname(suite))`。
 * @param {{projectRoot?:string}} [config]
 */
export function resolveBoundaryRoot(config = {}) {
  const configured = typeof config?.projectRoot === 'string' ? config.projectRoot.trim() : '';
  const base = configured || DEFAULT_BOUNDARY_ROOT;
  return resolve(base);
}

/** 路径是否位于边界根内（大小写不敏感，含根本身）。 */
export function isInsideRoot(abs, root = DEFAULT_BOUNDARY_ROOT) {
  if (!abs) return false;
  const target = normalizeForCompare(abs);
  const base = normalizeForCompare(root);
  if (target === base) return true;
  const prefix = base.endsWith(sep) ? base : `${base}${sep.toLowerCase()}`;
  return target.startsWith(prefix);
}

/**
 * 断言路径在边界根内，返回 resolve 后的绝对路径（便于链式使用）。
 * @param {string} abs 待校验路径
 * @param {object} [options]
 * @param {boolean} [options.write=false] true = 写操作（allowExternalPaths 不放宽）
 * @param {boolean} [options.allowExternalPaths=false] 是否允许越界读取（仅读取）
 * @param {string} [options.root] 边界根（默认 DEFAULT_BOUNDARY_ROOT）
 * @param {(message:string, detail:object)=>void} [options.warn] 越界读取告警回调
 * @returns {string}
 * @throws {OpsError} E1021 PATH_OUTSIDE_ROOT
 */
export function assertInsideRoot(abs, options = {}) {
  if (typeof abs !== 'string' || !abs.trim()) {
    throw new OpsError('E1021', '路径为空，无法通过边界校验', '请提供绝对路径或相对会话工作区的路径');
  }
  const target = resolve(abs);
  const root = options.root ? resolve(options.root) : DEFAULT_BOUNDARY_ROOT;
  const write = options.write === true;
  if (isInsideRoot(target, root)) return target;
  if (!write && options.allowExternalPaths === true) {
    try {
      options.warn?.(`越界读取（allowExternalPaths=true）: ${target}`, { path: target, root, write });
    } catch { /* 告警失败不影响读取 */ }
    return target;
  }
  const hint = write
    ? `输出只能写入 ${root} 之内（allowExternalPaths 不放宽写操作）`
    : `输入需位于 ${root} 之内；如需读取外部路径请在设置中开启 allowExternalPaths`;
  throw new OpsError('E1021', `路径超出允许边界（${root}）: ${target}`, hint);
}

/** 边界描述（供诊断页展示）。 */
export function boundaryInfo(config = {}) {
  const root = resolveBoundaryRoot(config);
  return {
    root,
    defaultValue: DEFAULT_BOUNDARY_ROOT,
    custom: root !== resolve(DEFAULT_BOUNDARY_ROOT),
    allowExternalPaths: config?.allowExternalPaths === true,
  };
}
