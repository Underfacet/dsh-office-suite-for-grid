/**
 * D1：Office 编辑/排版指令解析与校验 + 引擎选择（docs/P2-设计.md §3.2/§3.4/§3.3）
 *
 * COM 冒烟测试默认跳过；设置 SUITE_COM_SMOKE=1 时对 tests-local 样例的副本
 * （runtime\tmp 下）执行一次段落替换/单元格写入并读回验证，不修改真实样例。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync } from 'node:fs';
import { mkdir, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  appForFamily,
  chooseOfficeEngine,
  comUsableFromProbe,
  familyOfExtension,
  validateEditOps,
} from '../suite/host/ops/office-edit.js';
import { validateFormatSpec, validateScope } from '../suite/host/ops/office-format.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `office-edit-${process.pid}`);
const ENGINE_DIR = join(REPO_ROOT, 'suite', 'engines');

before(async () => {
  await mkdir(TMP, { recursive: true });
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

const isOpsError = (code) => (err) => err?.name === 'OpsError' && err.code === code;

// ───────────────────────── Word 编辑指令 ─────────────────────────

test('Word 指令：合法操作全量通过并归一化', () => {
  const ops = validateEditOps('demo.docx', '.docx', [
    { action: 'paragraph', op: 'setText', at: 1, text: '正文' },
    { action: 'paragraph', op: 'append', text: '追加段' },
    { action: 'heading', level: 2, op: 'setText', at: 2, text: '标题' },
    { action: 'heading', level: 1, op: 'insert', at: 1, text: '新标题' },
    { action: 'table', table: 1, op: 'setCell', row: 1, col: 2, text: 'x' },
    { action: 'table', table: 1, op: 'addRow', at: 'end' },
    { action: 'table', table: 1, op: 'addColumn', at: 3 },
    { action: 'table', table: 1, op: 'deleteColumn', col: 1 },
    { action: 'table', table: 1, op: 'deleteRow', row: 2 },
    { action: 'saveAs', path: 'out.docx' },
  ]);
  assert.equal(ops.length, 10);
  assert.equal(ops[0].action, 'paragraph');
});

test('Word 指令：create 只能作为首条且 kind=docx', () => {
  const ops = validateEditOps('new.docx', '.docx', [
    { action: 'create', kind: 'docx' },
    { action: 'paragraph', op: 'append', text: 'hello' },
  ]);
  assert.equal(ops[0].action, 'create');

  assert.throws(
    () => validateEditOps('new.docx', '.docx', [
      { action: 'paragraph', op: 'append', text: 'x' },
      { action: 'create', kind: 'docx' },
    ]),
    isOpsError('E2005'),
  );
  assert.throws(
    () => validateEditOps('new.docx', '.docx', [{ action: 'create', kind: 'xlsx' }]),
    isOpsError('E2005'),
  );
});

test('Word 指令：必填字段/索引/样式名/扩展名校验', () => {
  const bad = [
    [{ action: 'paragraph', op: 'setText', at: 0, text: 'x' }],
    [{ action: 'paragraph', op: 'setText', at: 1 }],
    [{ action: 'paragraph', op: 'setStyle', at: 1 }],
    [{ action: 'paragraph', op: 'insert', at: 2 }],
    [{ action: 'heading', level: 4, op: 'setText', at: 1, text: 'x' }],
    [{ action: 'heading', level: 1, op: 'append', at: 1, text: 'x' }],
    [{ action: 'table', table: 0, op: 'addRow' }],
    [{ action: 'table', table: 1, op: 'setCell', row: 1, text: 'x' }],
    [{ action: 'saveAs', path: 'out.xlsx' }],
    [{ action: 'unknown' }],
    [],
  ];
  for (const ops of bad) {
    assert.throws(() => validateEditOps('demo.docx', '.docx', ops), isOpsError('E2005'), JSON.stringify(ops));
  }
});

test('指令集按扩展名隔离：docx 不接受 Excel 指令，xlsx 不接受 Word 指令', () => {
  assert.throws(
    () => validateEditOps('demo.docx', '.docx', [{ action: 'cell', at: 'A1', value: 1 }]),
    isOpsError('E2005'),
  );
  assert.throws(
    () => validateEditOps('demo.xlsx', '.xlsx', [{ action: 'paragraph', op: 'append', text: 'x' }]),
    isOpsError('E2005'),
  );
});

test('不支持的文件类型 → E1042', () => {
  assert.throws(() => validateEditOps('a.pdf', '.pdf', [{ action: 'cell', at: 'A1', value: 1 }]), isOpsError('E1042'));
});

// ───────────────────────── Excel 编辑指令 ─────────────────────────

test('Excel 指令：单元格/公式/行列/工作表合法', () => {
  const ops = validateEditOps('demo.xlsx', '.xlsx', [
    { action: 'cell', sheet: 'Sheet1', at: 'b3', value: 1.5 },
    { action: 'cell', sheet: 'Sheet1', at: 'C3', formula: '=SUM(B1:B2)' },
    { action: 'row', sheet: 'Sheet1', op: 'insert', at: 2 },
    { action: 'row', sheet: 'Sheet1', op: 'delete', at: 3 },
    { action: 'column', sheet: 'Sheet1', op: 'insert', at: 'B' },
    { action: 'column', sheet: 'Sheet1', op: 'delete', at: 2 },
    { action: 'sheet', op: 'rename', name: 'Sheet1', newName: '数据' },
    { action: 'sheet', op: 'add', name: '汇总' },
    { action: 'sheet', op: 'delete', name: '临时' },
    { action: 'saveAs', path: '副本.xlsx' },
  ]);
  assert.equal(ops[0].at, 'B3', '单元格地址应归一化为大写');
  assert.equal(ops[4].at, 'B');
});

test('Excel 指令：value/formula 互斥、公式前缀、A1 地址与列字母校验', () => {
  const bad = [
    [{ action: 'cell', at: 'A1', value: 1, formula: '=B1' }],
    [{ action: 'cell', at: 'A1' }],
    [{ action: 'cell', at: 'A1', formula: 'SUM(A2:A3)' }],
    [{ action: 'cell', at: '1A', value: 1 }],
    [{ action: 'cell', at: 'XFE1', value: 1 }],
    [{ action: 'cell', at: 'A1', value: null }],
    [{ action: 'row', op: 'insert', at: 0 }],
    [{ action: 'row', op: 'move', at: 1 }],
    [{ action: 'column', op: 'insert', at: 'AAAA' }],
    [{ action: 'sheet', op: 'rename', name: 'S1' }],
    [{ action: 'sheet', op: 'unknown', name: 'S1' }],
    [{ action: 'saveAs', path: '副本.docx' }],
  ];
  for (const ops of bad) {
    assert.throws(() => validateEditOps('demo.xlsx', '.xlsx', ops), isOpsError('E2005'), JSON.stringify(ops));
  }
});

test('Excel 指令：create kind=xlsx 且仅首条', () => {
  const ops = validateEditOps('new.xlsx', '.xlsx', [
    { action: 'create', kind: 'xlsx' },
    { action: 'cell', at: 'A1', value: 'x' },
  ]);
  assert.equal(ops[0].action, 'create');
  assert.throws(
    () => validateEditOps('new.xlsx', '.xlsx', [{ action: 'create', kind: 'docx' }]),
    isOpsError('E2005'),
  );
});

// ───────────────────────── 扩展名 / 引擎选择 ─────────────────────────

test('扩展名与 app 映射', () => {
  assert.equal(familyOfExtension('.docx'), 'word');
  assert.equal(familyOfExtension('.DOC'), 'word');
  assert.equal(familyOfExtension('.xls'), 'excel');
  assert.equal(familyOfExtension('.pdf'), null);
  assert.equal(appForFamily('word'), 'word');
  assert.equal(appForFamily('excel'), 'excel');
});

const wordComProbe = (comUsable, present = true) => ({
  officeCom: { word: { present, comUsable }, excel: { present, comUsable } },
});

test('引擎选择：auto 优先 COM', () => {
  const decision = chooseOfficeEngine(wordComProbe(true), { officeEngine: 'auto' }, undefined, { app: 'word', venvHealthy: true });
  assert.equal(decision.engine, 'com');
  assert.equal(decision.fallback, false);
});

test('引擎选择：auto 在 COM 不可用时降级 python（标记 fallback）', () => {
  const decision = chooseOfficeEngine(wordComProbe(false), { officeEngine: 'auto' }, undefined, { app: 'word', venvHealthy: true });
  assert.equal(decision.engine, 'python');
  assert.equal(decision.fallback, true);
  assert.equal(decision.needsVenv, false);
});

test('引擎选择：auto 且无 COM/venv → E2001', () => {
  assert.throws(
    () => chooseOfficeEngine(wordComProbe(false), { officeEngine: 'auto' }, undefined, { app: 'word', venvHealthy: false }),
    isOpsError('E2001'),
  );
});

test('引擎选择：显式 engine=com 不可用 → E2001；显式 python 可用 → python', () => {
  assert.throws(
    () => chooseOfficeEngine(wordComProbe(false), { officeEngine: 'auto' }, 'com', { app: 'word' }),
    isOpsError('E2001'),
  );
  const decision = chooseOfficeEngine(wordComProbe(true), { officeEngine: 'auto' }, 'python', { app: 'word', venvHealthy: true });
  assert.equal(decision.engine, 'python');
  assert.equal(decision.fallback, false);
});

test('引擎选择：非法 engine → PARAM_INVALID；v0.1 probe 兼容 Present', () => {
  assert.throws(
    () => chooseOfficeEngine(wordComProbe(true), {}, 'native', { app: 'word' }),
    isOpsError('PARAM_INVALID'),
  );
  assert.equal(comUsableFromProbe({ 'Word.Application': { present: true } }, 'word'), true);
  assert.equal(comUsableFromProbe({ 'Word.Application': { present: false } }, 'word'), false);
});

// ───────────────────────── 排版规格 ─────────────────────────

test('排版规格：公文样式集形态通过校验', () => {
  const spec = validateFormatSpec({
    rules: [
      {
        target: 'heading',
        level: 1,
        font: { name: '方正小标宋简体', sizePt: 22 },
        paragraph: { align: 'center', lineSpacing: 1.5 },
        numbering: 'chinese',
      },
      { target: 'body', font: { name: '仿宋_GB2312', sizePt: 16 }, paragraph: { lineSpacingPt: 28 } },
      { target: 'page', page: { marginsMm: { top: 37, bottom: 35, left: 28, right: 26 } } },
      { target: 'footer', footer: { text: '第 X 页 共 Y 页', font: { sizePt: 12 } } },
    ],
    scope: 'all',
  });
  assert.equal(spec.rules.length, 4);
  assert.deepEqual(spec.scope, { type: 'all' });
});

test('排版规格：非法 target/level/numbering/字段类型 → E2006', () => {
  const bad = [
    [{ target: 'title' }],
    [{ target: 'heading', level: 4, font: { sizePt: 12 } }],
    [{ target: 'heading', level: 1 }],
    [{ target: 'body', font: { name: '' } }],
    [{ target: 'body', font: { sizePt: 'big' } }],
    [{ target: 'body', font: { sizePt: 12 }, paragraph: { lineSpacing: 1.5, lineSpacingPt: 20 } }],
    [{ target: 'body', font: { sizePt: 12 }, numbering: 'roman' }],
    [{ target: 'body', font: { sizePt: 12 }, paragraph: { align: 'middle' } }],
    [{ target: 'page', page: { marginsMm: { top: 0 } } }],
    [{ target: 'page' }],
    [{ target: 'header' }],
    [{ target: 'footer', footer: {} }],
  ];
  for (const rules of bad) {
    assert.throws(() => validateFormatSpec({ rules }), isOpsError('E2006'), JSON.stringify(rules));
  }
});

test('排版规格：scope 合法/非法', () => {
  assert.deepEqual(validateScope(undefined), { type: 'all' });
  assert.deepEqual(validateScope({ headings: [3, 1, 1] }), { type: 'headings', headings: [1, 3] });
  assert.deepEqual(validateScope({ paragraphs: [2, 1] }), { type: 'paragraphs', paragraphs: [1, 2] });
  for (const scope of [{ headings: [] }, { headings: [0] }, { paragraphs: [0] }, { paragraphs: '1' }, 3]) {
    assert.throws(() => validateScope(scope), isOpsError('E2006'), JSON.stringify(scope));
  }
});

test('排版规格：styleSet/template/placeholders 类型校验', () => {
  assert.throws(() => validateFormatSpec({ styleSet: 42 }), isOpsError('E2006'));
  assert.throws(() => validateFormatSpec({ template: '' }), isOpsError('E2006'));
  assert.throws(() => validateFormatSpec({ placeholders: [] }), isOpsError('E2006'));
  assert.throws(() => validateFormatSpec({ placeholders: { 标题: { bad: true } } }), isOpsError('E2006'));
  const spec = validateFormatSpec({ template: 'gongwen', placeholders: { 标题: '关于 X 的通知', 日期: '2026-10-20' } });
  assert.deepEqual(spec.placeholders, { 标题: '关于 X 的通知', 日期: '2026-10-20' });
});

// ───────────────────────── COM 冒烟（可选） ─────────────────────────

function runPowerShell(args) {
  return new Promise((resolvePromise) => {
    const child = spawn('powershell', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', ...args], {
      windowsHide: true,
      cwd: REPO_ROOT,
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => { stdout += String(chunk); });
    child.stderr.on('data', (chunk) => { stderr += String(chunk); });
    child.on('error', (err) => resolvePromise({ code: -1, stdout, stderr: `${stderr}${String(err)}` }));
    child.on('close', (code) => resolvePromise({ code, stdout, stderr }));
  });
}

function parseEngineJson(text) {
  if (!text) return null;
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) return null;
  try { return JSON.parse(text.slice(start, end + 1)); } catch { return null; }
}

const COM_SMOKE = process.env.SUITE_COM_SMOKE === '1';

test('COM 冒烟：Word 段落替换并读回', { skip: !COM_SMOKE ? '设置 SUITE_COM_SMOKE=1 启用本机 COM 冒烟' : false }, async () => {
  const source = join(REPO_ROOT, 'tests-local', 'sample.docx');
  if (!existsSync(source)) return;
  mkdirSync(TMP, { recursive: true });
  const docPath = join(TMP, 'sample-com.docx');
  copyFileSync(source, docPath);

  const payload = {
    mode: 'edit',
    path: docPath,
    ext: '.docx',
    app: 'word',
    ops: [{ action: 'paragraph', op: 'setText', at: 1, text: 'P2 COM 冒烟测试段落' }],
    outputPath: docPath,
    create: false,
  };
  const edited = await runPowerShell(['-File', join(ENGINE_DIR, 'office-com.ps1'), '-Json', JSON.stringify(payload)]);
  const editJson = parseEngineJson(edited.stdout);
  assert.ok(editJson, `引擎输出不可解析：${edited.stdout} ${edited.stderr}`);
  assert.equal(editJson.ok, true, JSON.stringify(editJson));

  const read = await runPowerShell(['-File', join(ENGINE_DIR, 'office-read.ps1'), '-Json', JSON.stringify({ path: docPath })]);
  const readJson = parseEngineJson(read.stdout);
  assert.ok(readJson?.ok, `读回失败：${read.stdout} ${read.stderr}`);
  const preview = readJson.result?.info?.preview ?? [];
  assert.ok(preview.includes('P2 COM 冒烟测试段落'), `读回内容未包含写入文本：${JSON.stringify(preview)}`);
});

test('COM 冒烟：Excel 单元格写入并读回', { skip: !COM_SMOKE ? '设置 SUITE_COM_SMOKE=1 启用本机 COM 冒烟' : false }, async () => {
  const source = join(REPO_ROOT, 'tests-local', 'sample.xlsx');
  if (!existsSync(source)) return;
  mkdirSync(TMP, { recursive: true });
  const xlsxPath = join(TMP, 'sample-com.xlsx');
  copyFileSync(source, xlsxPath);

  const payload = {
    mode: 'edit',
    path: xlsxPath,
    ext: '.xlsx',
    app: 'excel',
    ops: [{ action: 'cell', at: 'A1', value: 'P2COM' }],
    outputPath: xlsxPath,
    create: false,
  };
  const edited = await runPowerShell(['-File', join(ENGINE_DIR, 'office-com.ps1'), '-Json', JSON.stringify(payload)]);
  const editJson = parseEngineJson(edited.stdout);
  assert.ok(editJson, `引擎输出不可解析：${edited.stdout} ${edited.stderr}`);
  assert.equal(editJson.ok, true, JSON.stringify(editJson));

  const read = await runPowerShell(['-File', join(ENGINE_DIR, 'office-read.ps1'), '-Json', JSON.stringify({ path: xlsxPath })]);
  const readJson = parseEngineJson(read.stdout);
  assert.ok(readJson?.ok, `读回失败：${read.stdout} ${read.stderr}`);
  const firstSheet = readJson.result?.info?.sheets?.[0];
  assert.ok(firstSheet, '未读取到工作表');
  assert.ok((firstSheet.sample ?? []).some((line) => line.includes('P2COM')), JSON.stringify(firstSheet.sample));
});
