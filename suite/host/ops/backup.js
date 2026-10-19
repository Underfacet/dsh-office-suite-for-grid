/**
 * 修改前自动备份（docs/P2-设计.md §3.6）
 *
 * 布局（默认）：
 *   runtime\backups\index.json                  { "<absPath>": [ {id,time,file,name,size}, ... 最新在前 ] }
 *   runtime\backups\docs\<md5(absPath 小写)前12位>\<文件名>.<ISO时间戳>.bak
 *
 * 规则：
 * - 每个文档保留最近 backupCopies（默认 3）份，写新备份后删除更旧的；
 * - 备份失败返回 E2008 并中止编辑（宁可不动，不可无保底）；
 * - index.json 损坏时按「空索引 + 磁盘扫描」自愈并告警；
 * - 恢复前先备份当前版本（preRestoreBackup），备份失败即中止恢复。
 *
 * 导出：
 *   backupFile(path, options)       → {ok:true,result:{backupPath,id,time,...},backupPath} | {ok:false,error}
 *   listBackups(path, options)      → {ok:true,result:{path,backups:[...]}} | {ok:false,error}
 *   restoreBackup(backupPath, targetPath, options) → {ok:true,result:{restoredTo,preRestoreBackup,...}} | ...
 *   backupOp / backupCreateOp / backupListOp / backupRestoreOp（命令层统一形状包装）
 */
import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, readdir, rename, stat, unlink, writeFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { BACKUP_DIR, OpsError, fail, ok, resolveInputPath, toErrorPayload } from '../ops.js';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';

/** 默认保留份数（config.backupCopies）。 */
export const DEFAULT_BACKUP_COPIES = 3;

/** 单个文档最多保留份数上限（防御异常配置）。 */
const MAX_BACKUP_COPIES = 50;

/** 索引写串行化（进程内），避免并发批量排版互相覆盖 index.json。 */
let indexQueue = Promise.resolve();

function withIndexLock(fn) {
  const run = indexQueue.then(fn, fn);
  indexQueue = run.then(() => undefined, () => undefined);
  return run;
}

/** 边界根（config.projectRoot 可覆盖，默认 E:\工作\2026）。 */
function boundaryRoot(options = {}) {
  if (options.root) return resolve(options.root);
  return resolveBoundaryRoot(options.config ?? {});
}

/** 备份根目录（默认 runtime\backups）。 */
function backupRoot(options = {}) {
  return resolve(options.backupDir ?? BACKUP_DIR);
}

function normalizeCopies(value) {
  const n = Number(value);
  if (Number.isInteger(n) && n >= 1 && n <= MAX_BACKUP_COPIES) return n;
  return DEFAULT_BACKUP_COPIES;
}

/** 路径摘要：md5(小写绝对路径) 前 12 位。 */
function digestOf(abs) {
  return createHash('md5').update(abs.toLowerCase()).digest('hex').slice(0, 12);
}

/** yyyyMMdd-HHmmss-SSS（本地时区无关，统一 UTC ISO 去分隔符）。 */
function timestamp(date = new Date()) {
  const iso = date.toISOString(); // 2026-10-20T01:35:14.123Z
  return `${iso.slice(0, 10).replace(/-/g, '')}-${iso.slice(11, 19).replace(/:/g, '')}-${iso.slice(20, 23)}`;
}

async function exists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

async function isFile(path) {
  try {
    return (await stat(path)).isFile();
  } catch {
    return false;
  }
}

async function readIndex(dir) {
  const file = join(dir, 'index.json');
  let text;
  try {
    text = await readFile(file, 'utf8');
  } catch (err) {
    if (err?.code === 'ENOENT') return {};
    console.warn(`[office-suite] 备份索引读取失败，按空索引处理：${String(err?.message ?? err)}`);
    return {};
  }
  try {
    const value = JSON.parse(text.replace(/^\uFEFF/, ''));
    if (value && typeof value === 'object' && !Array.isArray(value)) return value;
    throw new Error('索引不是对象');
  } catch (err) {
    console.warn(`[office-suite] 备份索引损坏，重建为空索引：${String(err?.message ?? err)}`);
    return {};
  }
}

async function writeIndex(dir, index) {
  await mkdir(dir, { recursive: true });
  const file = join(dir, 'index.json');
  const tmp = `${file}.tmp-${process.pid}-${Date.now()}`;
  await writeFile(tmp, JSON.stringify(index, null, 2), 'utf8');
  try {
    await rename(tmp, file);
  } catch {
    // Windows 上目标存在时 rename 可能失败：退化为直接覆盖
    await writeFile(file, JSON.stringify(index, null, 2), 'utf8');
    try { await unlink(tmp); } catch { /* 忽略 */ }
  }
}

/** 备份目录内某个文档的子目录。 */
function folderFor(backupDir, abs) {
  return join(backupDir, 'docs', digestOf(abs));
}

/** 生成不冲突的备份文件名：<文件名>.<ISO 时间戳>.bak（同毫秒冲突时追加 -N）。 */
async function uniqueBackupName(folder, abs) {
  const base = `${basename(abs)}.${timestamp()}`;
  let candidate = join(folder, `${base}.bak`);
  let seq = 0;
  while (await exists(candidate)) {
    seq += 1;
    candidate = join(folder, `${base}-${seq}.bak`);
  }
  return candidate;
}

/** 备份文件名排序键（同一毫秒时按 -N 序）。 */
function entryTime(entry) {
  return String(entry?.time ?? '');
}

/**
 * 备份一个文件。
 * @param {string} path 待备份文件（绝对路径）
 * @param {{backupDir?:string, copies?:number, root?:string, config?:object}} [options]
 * @returns {Promise<{ok:true,result:object,backupPath:string}|{ok:false,error:object}>}
 */
export async function backupFile(path, options = {}) {
  const abs = resolve(String(path ?? ''));
  const root = boundaryRoot(options);
  try {
    assertInsideRoot(abs, { write: true, root });
  } catch (err) {
    return fail(err.code ?? 'E1021', err.message, err.hint);
  }
  if (!(await isFile(abs))) {
    return fail('E1023', `待备份文件不存在或不是文件: ${abs}`);
  }

  const backupDir = backupRoot(options);
  const copies = normalizeCopies(options.copies);
  try {
    assertInsideRoot(backupDir, { write: true, root });
  } catch (err) {
    return fail(err.code ?? 'E1021', err.message, err.hint);
  }

  const folder = folderFor(backupDir, abs);
  try {
    await mkdir(folder, { recursive: true });
  } catch (err) {
    return fail('E2008', `创建备份目录失败: ${String(err?.message ?? err)}`, `目录：${folder}`);
  }

  let dest;
  try {
    dest = await uniqueBackupName(folder, abs);
    await copyFile(abs, dest);
  } catch (err) {
    return fail('E2008', `备份文件失败: ${String(err?.message ?? err)}`, `目标：${dest ?? folder}`);
  }

  const info = await stat(dest).catch(() => null);
  const now = new Date();
  const entry = {
    id: basename(dest),
    name: basename(dest),
    file: dest,
    time: now.toISOString(),
    size: info?.size ?? 0,
  };

  try {
    await withIndexLock(async () => {
      const index = await readIndex(backupDir);
      const existing = Array.isArray(index[abs]) ? index[abs].filter((item) => item && item.file !== dest) : [];
      const merged = [entry, ...existing].sort((a, b) => entryTime(b).localeCompare(entryTime(a)));
      const kept = merged.slice(0, copies);
      const removed = merged.slice(copies);
      for (const item of removed) {
        if (!item?.file) continue;
        // 只删除备份目录内、文件名以 .bak 结尾的文件，避免误删
        try {
          const target = resolve(item.file);
          if (target.startsWith(resolve(backupDir)) && target.toLowerCase().endsWith('.bak')) {
            await unlink(target);
          }
        } catch { /* 已被清理或占用，忽略 */ }
      }
      index[abs] = kept;
      await writeIndex(backupDir, index);
    });
  } catch (err) {
    console.warn(`[office-suite] 备份索引更新失败（备份文件已写入）：${String(err?.message ?? err)}`);
    return fail('E2008', `备份索引更新失败: ${String(err?.message ?? err)}`, `备份文件：${dest}`);
  }

  return {
    ok: true,
    result: { backupPath: dest, id: entry.id, name: entry.name, time: entry.time, size: entry.size, copies },
    backupPath: dest,
  };
}

/** 扫描备份目录，重建某文档的备份清单（索引损坏/手工拷贝时用）。 */
async function scanBackups(folder, abs) {
  let names;
  try {
    names = await readdir(folder);
  } catch {
    return [];
  }
  const prefix = `${basename(abs)}.`;
  const entries = [];
  for (const name of names) {
    if (!name.startsWith(prefix) || !name.toLowerCase().endsWith('.bak')) continue;
    const file = join(folder, name);
    try {
      const info = await stat(file);
      if (!info.isFile()) continue;
      entries.push({ id: name, name, file, time: info.mtime.toISOString(), size: info.size });
    } catch { /* 读取失败忽略 */ }
  }
  entries.sort((a, b) => entryTime(b).localeCompare(entryTime(a)));
  return entries;
}

/**
 * 列举某文档的备份（最新在前）。索引缺失/损坏时自动扫描磁盘。
 * @param {string} path 文档绝对路径（可以已不存在）
 * @param {{backupDir?:string, root?:string, config?:object, copies?:number}} [options]
 */
export async function listBackups(path, options = {}) {
  const abs = resolve(String(path ?? ''));
  const backupDir = backupRoot(options);
  const folder = folderFor(backupDir, abs);
  try {
    assertInsideRoot(backupDir, { write: false, root: boundaryRoot(options) });
  } catch (err) {
    return fail(err.code ?? 'E1021', err.message, err.hint);
  }

  let entries = [];
  await withIndexLock(async () => {
    const index = await readIndex(backupDir);
    const raw = Array.isArray(index[abs]) ? index[abs] : [];
    const byFile = new Map();
    // 索引条目优先（保留原始备份时间）；文件丢失的条目丢弃
    for (const item of raw) {
      if (!item?.file) continue;
      const file = resolve(item.file);
      try {
        const info = await stat(file);
        if (!info.isFile()) continue;
        byFile.set(file, {
          id: String(item.id ?? basename(file)),
          name: String(item.name ?? basename(file)),
          file,
          time: String(item.time ?? info.mtime.toISOString()),
          size: info.size,
        });
      } catch { /* 文件已丢失：跳过 */ }
    }
    // 磁盘扫描补齐索引缺失/损坏的条目（自愈）
    for (const scanned of await scanBackups(folder, abs)) {
      const file = resolve(scanned.file);
      if (!byFile.has(file)) byFile.set(file, scanned);
    }
    entries = [...byFile.values()];
    if (entries.length !== raw.length || raw.some((item) => item?.file && !byFile.has(resolve(item.file)))) {
      const index2 = await readIndex(backupDir);
      if (entries.length) index2[abs] = entries;
      else delete index2[abs];
      try { await writeIndex(backupDir, index2); } catch { /* 索引不可写不阻断列举 */ }
    }
  });

  entries.sort((a, b) => entryTime(b).localeCompare(entryTime(a)));
  return ok({ path: abs, backups: entries.map((entry) => ({ ...entry })) });
}

/**
 * 恢复备份：先把当前版本再备份一次（失败即中止），再复制备份回目标路径。
 * @param {string} backupPath 备份文件绝对路径
 * @param {string} targetPath 目标文档绝对路径
 * @param {{backupDir?:string, copies?:number, root?:string, config?:object}} [options]
 */
export async function restoreBackup(backupPath, targetPath, options = {}) {
  const source = resolve(String(backupPath ?? ''));
  const target = resolve(String(targetPath ?? ''));
  const root = boundaryRoot(options);
  try {
    assertInsideRoot(source, { write: false, root });
    assertInsideRoot(target, { write: true, root });
  } catch (err) {
    return fail(err.code ?? 'E1021', err.message, err.hint);
  }

  if (!(await isFile(source))) {
    return fail('BACKUP_NOT_FOUND', `备份文件不存在: ${source}`);
  }

  let preRestoreBackup = null;
  if (await exists(target)) {
    const pre = await backupFile(target, { ...options, root });
    if (!pre.ok) {
      return fail('E2008', `恢复前备份当前版本失败，已中止恢复: ${pre.error.message}`, pre.error.hint ?? '请检查备份目录是否可写');
    }
    preRestoreBackup = pre.result.backupPath;
  }

  try {
    await mkdir(dirname(target), { recursive: true });
    await copyFile(source, target);
  } catch (err) {
    return fail('E2008', `恢复备份失败: ${String(err?.message ?? err)}`, `目标：${target}`);
  }

  return ok({
    restoredTo: target,
    restoredFrom: source,
    preRestoreBackup,
  });
}

// ───────────────────────── 命令层包装（统一 {ok,result}|{ok,error}） ─────────────────────────

/** 解析文档路径并过边界（写）。 */
async function resolveDocPath(ctx, config, args, exec) {
  const raw = typeof args?.path === 'string' ? args.path.trim() : '';
  if (!raw) throw new OpsError('E1041', '缺少 path');
  const abs = assertInsideRoot(await resolveInputPath(ctx, exec, raw), {
    write: true,
    root: resolveBoundaryRoot(config),
  });
  return abs;
}

/** 备份（命令：suite.office.backup 的 create 动作，供面板手动备份）。 */
export async function backupCreateOp(ctx, config = {}, args = {}, exec = {}) {
  try {
    const abs = await resolveDocPath(ctx, config, args, exec);
    const result = await backupFile(abs, { copies: config.backupCopies, root: resolveBoundaryRoot(config) });
    return result;
  } catch (err) {
    return { ok: false, error: toErrorPayload(err) };
  }
}

/** suite.office.backup.list：{path} → {path, backups[]}；无备份返回 BACKUP_NONE。 */
export async function backupListOp(ctx, config = {}, args = {}, exec = {}) {
  try {
    const abs = await resolveDocPath(ctx, config, args, exec);
    const listed = await listBackups(abs, { copies: config.backupCopies, root: resolveBoundaryRoot(config) });
    if (!listed.ok) return listed;
    if (!listed.result.backups.length) {
      return fail('BACKUP_NONE', `该文档没有备份记录: ${abs}`, '编辑/排版时会自动生成备份');
    }
    return listed;
  } catch (err) {
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}

/** suite.office.backup.restore：{path, id} → {restoredTo, preRestoreBackup}。 */
export async function backupRestoreOp(ctx, config = {}, args = {}, exec = {}) {
  try {
    const abs = await resolveDocPath(ctx, config, args, exec);
    const id = typeof args?.id === 'string' ? args.id.trim() : '';
    if (!id) return fail('E1041', '缺少 id（suite.office.backup.list 返回的备份 id）');
    const listed = await listBackups(abs, { copies: config.backupCopies, root: resolveBoundaryRoot(config) });
    if (!listed.ok) return listed;
    const wanted = id.toLowerCase();
    const entry = listed.result.backups.find((item) => (
      String(item.id).toLowerCase() === wanted
      || String(item.name).toLowerCase() === wanted
      || String(item.id).toLowerCase() === `${wanted}.bak`
    ));
    if (!entry) {
      return fail('BACKUP_NOT_FOUND', `未找到备份: ${id}`, `可用备份：${listed.result.backups.map((item) => item.id).join('、') || '(无)'}`);
    }
    const restored = await restoreBackup(entry.file, abs, {
      copies: config.backupCopies,
      root: resolveBoundaryRoot(config),
    });
    if (!restored.ok) return restored;
    return ok({
      ...restored.result,
      backup: { id: entry.id, time: entry.time },
    });
  } catch (err) {
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}

/**
 * 备份统一入口（命令层可只接一个函数）：
 *   args.action = 'list' | 'restore' | 'create'（默认 list）
 *   list:   {path}
 *   restore:{path, id}
 *   create: {path}
 */
export async function backupOp(ctx, config = {}, args = {}, exec = {}) {
  const action = typeof args?.action === 'string' ? args.action.trim().toLowerCase() : 'list';
  switch (action) {
    case 'list':
      return backupListOp(ctx, config, args, exec);
    case 'restore':
      return backupRestoreOp(ctx, config, args, exec);
    case 'create':
      return backupCreateOp(ctx, config, args, exec);
    default:
      return fail('E1041', `未知备份动作: ${action}`, "action 必须为 'list' | 'restore' | 'create'");
  }
}
