/**
 * 工作区目录列举（docs/P1-设计.md §2 suite.files）
 *
 * 直接使用 node:fs（readdir + stat），路径先过 boundary 校验：
 *   assertInsideRoot(abs, { write:false, allowExternalPaths, root })
 * 返回 { path, entries:[{name,type:'file'|'directory'|'other',size?}], truncated }。
 */
import { readdir, stat } from 'node:fs/promises';
import { basename, resolve } from 'node:path';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import { OpsError, fail, ok, resolveInputPath, toErrorPayload } from '../ops.js';
import { compareNatural } from './drawingNo.js';

/** 默认返回条目上限。 */
export const DEFAULT_LIST_LIMIT = 2000;

function normalizeExtensions(input) {
  if (!Array.isArray(input)) return null;
  const list = input
    .filter((value) => typeof value === 'string' && value.trim())
    .map((value) => {
      const trimmed = value.trim().toLowerCase();
      return trimmed.startsWith('.') ? trimmed : `.${trimmed}`;
    });
  return list.length ? new Set(list) : null;
}

function extensionOf(name) {
  const index = name.lastIndexOf('.');
  return index >= 0 ? name.slice(index).toLowerCase() : '';
}

/**
 * 列举目录（面板文件树 / suite.files 命令共用）。
 * @param {object} ctx
 * @param {{path?:string, extensions?:string[], limit?:number}} args
 * @param {{agent?:any, signal?:AbortSignal}} exec
 * @param {object} config 运行配置（allowExternalPaths / projectRoot）
 * @returns {Promise<{ok:true,result:{path:string,entries:Array,truncated:boolean}}|{ok:false,error:object}>}
 */
export async function listDirectory(ctx, args = {}, exec = {}, config = {}) {
  try {
    const requested = typeof args.path === 'string' && args.path.trim() ? args.path.trim() : '.';
    const resolved = await resolveInputPath(ctx, exec, requested);
    const root = resolveBoundaryRoot(config);
    const abs = assertInsideRoot(resolved, {
      write: false,
      allowExternalPaths: config.allowExternalPaths === true,
      root,
      warn: (message) => console.warn(`[office-suite] ${message}`),
    });

    let dirents;
    try {
      dirents = await readdir(abs, { withFileTypes: true });
    } catch (err) {
      if (err?.code === 'ENOENT') {
        return fail('E1023', `目录不存在: ${abs}`, '请确认路径正确，或先在工作区打开会话');
      }
      if (err?.code === 'EACCES' || err?.code === 'EPERM') {
        return fail('E1022', `无权限读取目录: ${abs}`);
      }
      return fail('E1099', `列举目录失败: ${String(err?.message ?? err)}`);
    }

    const extensions = normalizeExtensions(args.extensions);
    const limit = Number.isInteger(args.limit) && args.limit > 0 ? args.limit : DEFAULT_LIST_LIMIT;
    const entries = [];
    for (const dirent of dirents) {
      const isDirectory = dirent.isDirectory();
      const isFile = dirent.isFile();
      if (!isDirectory && !isFile) {
        entries.push({ name: dirent.name, type: 'other' });
        continue;
      }
      if (isFile && extensions && !extensions.has(extensionOf(dirent.name))) continue;
      const entry = { name: dirent.name, type: isDirectory ? 'directory' : 'file' };
      if (isFile) {
        try {
          const info = await stat(resolve(abs, dirent.name));
          entry.size = info.size;
        } catch { /* 文件在列举过程中消失则忽略大小 */ }
      }
      entries.push(entry);
    }
    entries.sort((a, b) => {
      const aDir = a.type === 'directory' ? 0 : 1;
      const bDir = b.type === 'directory' ? 0 : 1;
      if (aDir !== bDir) return aDir - bDir;
      return compareNatural(a.name, b.name);
    });
    const truncated = entries.length > limit;
    return ok({
      path: abs,
      name: basename(abs),
      entries: truncated ? entries.slice(0, limit) : entries,
      truncated,
    });
  } catch (err) {
    if (err instanceof OpsError) return fail(err.code, err.message, err.hint);
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}
