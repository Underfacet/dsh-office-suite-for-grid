/**
 * D1：备份轮换 / 列举 / 恢复（docs/P2-设计.md §3.6 / §9）
 *
 * 全部使用 runtime\tmp\tests 下的临时目录，不触碰 tests-local 真实样例。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { backupFile, listBackups, restoreBackup } from '../suite/host/ops/backup.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `backup-${process.pid}`);

before(async () => {
  await mkdir(TMP, { recursive: true });
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

async function writeDoc(name, content) {
  const path = join(TMP, name);
  await writeFile(path, content, 'utf8');
  return path;
}

test('备份轮换：每文档默认保留最近 3 份，且最新在前', async () => {
  const doc = await writeDoc('rotate.txt', 'v1');
  const backupDir = join(TMP, 'backups-rotation');

  for (const [index, content] of ['v1', 'v2', 'v3', 'v4'].entries()) {
    if (index > 0) await writeFile(doc, content, 'utf8');
    const result = await backupFile(doc, { backupDir, copies: 3 });
    assert.equal(result.ok, true, JSON.stringify(result));
    assert.match(result.result.backupPath, /rotate\.txt\..*\.bak$/);
  }

  const listed = await listBackups(doc, { backupDir });
  assert.equal(listed.ok, true, JSON.stringify(listed));
  assert.equal(listed.result.backups.length, 3, '超出 3 份的旧备份应被删除');
  assert.equal(await readFile(listed.result.backups[0].file, 'utf8'), 'v4');
  assert.equal(await readFile(listed.result.backups[1].file, 'utf8'), 'v3');
  assert.equal(await readFile(listed.result.backups[2].file, 'utf8'), 'v2');

  const files = [...new Set(listed.result.backups.map((entry) => entry.file))];
  assert.equal(files.length, 3);
});

test('备份轮换：copies 可配置（2 份）', async () => {
  const doc = await writeDoc('rotate2.txt', 'a');
  const backupDir = join(TMP, 'backups-copies2');
  for (const content of ['a', 'b', 'c']) {
    await writeFile(doc, content, 'utf8');
    const result = await backupFile(doc, { backupDir, copies: 2 });
    assert.equal(result.ok, true);
  }
  const listed = await listBackups(doc, { backupDir });
  assert.equal(listed.result.backups.length, 2);
  assert.equal(await readFile(listed.result.backups[0].file, 'utf8'), 'c');
});

test('恢复：先把当前版本再备份，再写回目标文件', async () => {
  const doc = await writeDoc('restore.txt', 'old');
  const backupDir = join(TMP, 'backups-restore');
  const first = await backupFile(doc, { backupDir });
  assert.equal(first.ok, true);

  await writeFile(doc, 'current', 'utf8');
  const second = await backupFile(doc, { backupDir });
  assert.equal(second.ok, true);

  const restored = await restoreBackup(first.result.backupPath, doc, { backupDir });
  assert.equal(restored.ok, true, JSON.stringify(restored));
  assert.equal(restored.result.restoredTo, doc);
  assert.ok(restored.result.preRestoreBackup, '恢复前必须备份当前版本');
  assert.equal(await readFile(doc, 'utf8'), 'old');
  assert.equal(await readFile(restored.result.preRestoreBackup, 'utf8'), 'current');

  const listed = await listBackups(doc, { backupDir });
  assert.ok(listed.result.backups.length >= 2);
  assert.equal(listed.result.backups[0].file, restored.result.preRestoreBackup);
});

test('恢复：备份文件不存在 → BACKUP_NOT_FOUND', async () => {
  const doc = await writeDoc('missing-restore.txt', 'x');
  const result = await restoreBackup(join(TMP, 'no-such.bak'), doc, { backupDir: join(TMP, 'backups-missing') });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'BACKUP_NOT_FOUND');
});

test('备份：源文件不存在 → E1023；越界路径 → E1021', async () => {
  const missing = await backupFile(join(TMP, 'does-not-exist.docx'), { backupDir: join(TMP, 'backups-x') });
  assert.equal(missing.ok, false);
  assert.equal(missing.error.code, 'E1023');

  const outside = await backupFile('C:\\Windows\\notepad.exe', { backupDir: join(TMP, 'backups-x') });
  assert.equal(outside.ok, false);
  assert.equal(outside.error.code, 'E1021');
});

test('索引损坏：可自愈重建并继续备份/列举', async () => {
  const doc = await writeDoc('corrupt.txt', 'v1');
  const backupDir = join(TMP, 'backups-corrupt');
  const first = await backupFile(doc, { backupDir });
  assert.equal(first.ok, true);

  // 故意写坏 index.json
  await writeFile(join(backupDir, 'index.json'), '{ this is not json', 'utf8');

  const second = await backupFile(doc, { backupDir });
  assert.equal(second.ok, true, JSON.stringify(second));

  const listed = await listBackups(doc, { backupDir });
  assert.equal(listed.ok, true);
  assert.ok(listed.result.backups.length >= 2);
  const newest = await readFile(listed.result.backups[0].file, 'utf8');
  assert.equal(newest, 'v1');
});

test('列举：无备份记录时返回空数组（命令层再映射 BACKUP_NONE）', async () => {
  const doc = await writeDoc('none.txt', 'x');
  const listed = await listBackups(doc, { backupDir: join(TMP, 'backups-none') });
  assert.equal(listed.ok, true);
  assert.deepEqual(listed.result.backups, []);
});
