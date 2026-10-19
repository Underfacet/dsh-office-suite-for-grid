/**
 * D1：图纸目录表字段抽取 / 排序 / 张数 / 工作表 / CSV（docs/P3-设计.md §2.3 / §2.4）
 *
 * 夹具：tests/fixtures/make-multi-frame-pdf.js（A2 页内 4 个 A4 图框 + 图签图号 + SCALE 1:100）
 * 断言：PDF 级逐框抽取、图名/比例纯函数规则、排序/张数/多来源工作表命名、CSV 降级输出。
 * 不真写 xlsx（写表引擎由 engines/catalog-write.py 的 D2 实测覆盖）。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const requireFromSuite = createRequire(new URL('../suite/package.json', import.meta.url));
const { PDFDocument } = requireFromSuite('pdf-lib');
import { writeMultiFramePdf } from './fixtures/make-multi-frame-pdf.js';
import {
  CATALOG_HEADERS,
  annotateCatalogCounts,
  buildCatalogSheets,
  extractCatalogFields,
  extractCatalogScale,
  extractCatalogTitle,
  formatFrameSizeLabel,
  normalizeScaleText,
  readCatalogFromPdf,
  rowsToCsv,
  sortCatalogRows,
} from '../suite/host/ops/cad-catalog.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `catalog-${process.pid}`);

let fixture;

before(async () => {
  await mkdir(TMP, { recursive: true });
  fixture = await writeMultiFramePdf(join(TMP, 'multi-frame-a2.pdf'));
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

const isOpsError = (code) => (err) => err?.name === 'OpsError' && err.code === code;

/** 合成相对坐标文本项（0..1，原点左下；与 PDF/COM 两引擎同口径）。 */
function relativeItems(entries) {
  return entries.map(([text, x, y]) => ({ text, x, y, height: 0.02 }));
}

// ───────────────────────── 图名 / 比例 ─────────────────────────

test('图名：取图名区内中文最多的行，去「图名：」前缀，排除图号/比例/日期行', () => {
  const items = relativeItems([
    ['图号：D0101-01', 0.65, 0.08], // 图签区（y < 0.25），不在图名区
    ['白云变电站', 0.35, 0.42],
    ['110kV 主接线图', 0.55, 0.42], // 同 y → 合并为一行（中文 7 字）
    ['比例 1:100', 0.6, 0.3],
    ['2026-10-20', 0.7, 0.36],
    ['12345', 0.4, 0.5],
    ['图名：金亭变电站二次回路接线图', 0.4, 0.48], // 中文 12 字，最长
    ['设计：张三', 0.5, 0.33],
  ]);
  assert.equal(extractCatalogTitle(items, {}), '金亭变电站二次回路接线图');
});

test('图名：中文同分时优先非中文字符更少者（避免把电压等级行当图名）', () => {
  const items = relativeItems([
    ['白云变电站 110kV 主接线图', 0.4, 0.42], // 中文 9 + 非中文 6
    ['白云变电站主接线图', 0.4, 0.48], // 中文 9 + 非中文 0
  ]);
  assert.equal(extractCatalogTitle(items, {}), '白云变电站主接线图');
});

test('图名：图名区外不参与；无中文候选返回 null；「图名：」前缀剥离', () => {
  assert.equal(extractCatalogTitle(relativeItems([['主接线图', 0.1, 0.1]]), {}), null);
  assert.equal(extractCatalogTitle(relativeItems([['SCALE 1:100', 0.5, 0.4]]), {}), null);
  assert.equal(extractCatalogTitle(relativeItems([['图名：白云变总平面布置图', 0.5, 0.45]]), {}), '白云变总平面布置图');
  // 自定义图名区
  assert.equal(
    extractCatalogTitle(relativeItems([['白云变主接线图', 0.35, 0.7]]), { zone: [0.3, 0.6, 1.0, 0.9] }),
    '白云变主接线图',
  );
});

test('比例：前缀优先 / 裸写比号 / 全角冒号归一 / 排除日期时间行', () => {
  assert.equal(extractCatalogScale(relativeItems([['比例：1：100', 0.6, 0.12]]), {}), '1:100');
  assert.equal(extractCatalogScale(relativeItems([['SCALE 1:500', 0.6, 0.12]]), {}), '1:500');
  assert.equal(extractCatalogScale(relativeItems([['日期 2026/10/20', 0.6, 0.12]]), {}), null);
  assert.equal(extractCatalogScale(relativeItems([['时间 10:30', 0.6, 0.12]]), {}), null);
  assert.equal(extractCatalogScale(relativeItems([['D0101-01', 0.7, 0.06]]), {}), null);
  assert.equal(normalizeScaleText(' 1 ： 200 '), '1:200');
  assert.equal(normalizeScaleText('abc'), null);
});

// ───────────────────────── 统一字段抽取 ─────────────────────────

test('extractCatalogFields(pdf)：图号/图名/图幅/比例一次抽取', () => {
  const frame = { x: 0, y: 0, width: 297 * (72 / 25.4), height: 210 * (72 / 25.4), widthMm: 297, heightMm: 210 };
  const pageItems = [
    { text: 'D0101-01', x: 0.66 * frame.width, y: 0.06 * frame.height, height: 12 },
    { text: '比例 1:100', x: 0.66 * frame.width, y: 0.11 * frame.height, height: 8 },
    { text: '白云变 110kV 主接线图', x: 0.4 * frame.width, y: 0.42 * frame.height, height: 10 },
  ];
  const fields = extractCatalogFields('pdf', { frame, pageItems });
  assert.equal(fields.drawingNo, 'D0101-01');
  assert.equal(fields.title, '白云变 110kV 主接线图');
  assert.equal(fields.titleSource, 'text');
  assert.equal(fields.frameSize, 'A4'); // 由 frame 尺寸回退匹配
  assert.equal(fields.orientation, 'landscape');
  assert.equal(fields.scale, '1:100');
  assert.equal(fields.kind, 'pdf');
  assert.ok(fields.texts >= 3);
});

test('extractCatalogFields(pdf)：图名未命中回退文件名；COM 预留接口标记 enginePending', () => {
  const frame = { x: 0, y: 0, width: 841.89, height: 595.28 };
  const fields = extractCatalogFields('pdf', { frame, pageItems: [], fileName: join(TMP, '白云变-二次图纸.pdf') });
  assert.equal(fields.title, '白云变-二次图纸');
  assert.equal(fields.titleSource, 'filename');

  const com = extractCatalogFields('com', {
    items: relativeItems([
      ['D0101-02', 0.66, 0.06],
      ['比例 1:200', 0.66, 0.11],
      ['金亭变主接线图', 0.5, 0.42],
    ]),
    frameSize: 'A3',
    orientation: 'landscape',
    fileName: 'jinting.dwg',
  });
  assert.equal(com.kind, 'com');
  assert.equal(com.enginePending, true, 'COM 级目录表接口为 P3 预留（待真机）');
  assert.equal(com.drawingNo, 'D0101-02');
  assert.equal(com.title, '金亭变主接线图');
  assert.equal(com.frameSize, 'A3');
  assert.equal(com.scale, '1:200');
});

test('extractCatalogFields：非法 kind → E1041', () => {
  assert.throws(() => extractCatalogFields('oda', {}), isOpsError('E1041'));
  assert.throws(() => extractCatalogFields('pdf', {}), isOpsError('E1041'), 'pdf 级缺少 frame 应报 E1041');
});

// ───────────────────────── PDF 级逐框抽取（合成夹具） ─────────────────────────

test('readCatalogFromPdf：4 个 A4 图框逐框抽取图号/图幅/比例，图名回退文件名', async () => {
  const read = await readCatalogFromPdf(fixture.path, {});
  assert.equal(read.framesTotal, 4);
  assert.equal(read.pageCount, 1);
  assert.equal(read.rows.length, 4);
  assert.equal(read.warnings.length, 0);
  const nos = read.rows.map((row) => row.drawingNo).sort();
  assert.deepEqual(nos, ['D0101-01', 'D0101-02', 'D0101-03', 'D0101-04']);
  for (const row of read.rows) {
    assert.equal(row.frameSize, 'A4');
    assert.equal(row.orientation, 'landscape');
    assert.equal(row.scale, '1:100');
    assert.equal(row.titleSource, 'filename');
    assert.equal(row.title, 'multi-frame-a2');
    assert.equal(row.source, fixture.path);
    assert.equal(row.pageIndex, 0);
  }
  // 页内按 Z 形平面序（左上 → 左下 → 右上 → 右下）
  assert.deepEqual(read.rows.map((row) => row.drawingNo), ['D0101-02', 'D0101-03', 'D0101-04', 'D0101-01']);
});

test('readCatalogFromPdf：无图框 PDF → E1012', async () => {
  const doc = await PDFDocument.create();
  doc.addPage([200, 200]).drawRectangle({ x: 10, y: 10, width: 50, height: 30 });
  const inputPath = join(TMP, 'no-frame.pdf');
  await writeFile(inputPath, await doc.save());
  await assert.rejects(() => readCatalogFromPdf(inputPath, {}), isOpsError('E1012'));
});

// ───────────────────────── 排序 / 张数 / 工作表 ─────────────────────────

test('排序：drawingNo 图号优先（缺失排后）；filename 按来源文件+页序', () => {
  const rows = [
    { drawingNo: null, source: 'b.pdf', pageIndex: 0, frameIndex: 0, title: '乙图', frameSize: 'A4' },
    { drawingNo: 'D0101-10', source: 'a.pdf', pageIndex: 1, frameIndex: 0 },
    { drawingNo: 'D0101-2', source: 'a.pdf', pageIndex: 0, frameIndex: 0 },
    { drawingNo: null, source: 'a.pdf', pageIndex: 1, frameIndex: 1, title: '甲图', frameSize: 'A3' },
  ];
  assert.deepEqual(
    sortCatalogRows(rows, 'drawingNo').map((row) => row.drawingNo ?? row.title),
    ['D0101-2', 'D0101-10', '甲图', '乙图'],
  );
  assert.deepEqual(
    sortCatalogRows(rows, 'filename').map((row) => `${row.source}#${row.pageIndex}.${row.frameIndex}`),
    ['a.pdf#0.0', 'a.pdf#1.0', 'a.pdf#1.1', 'b.pdf#0.0'],
  );
});

test('张数：同来源按图号计数；图号缺失按「图名+图幅」组合计数', () => {
  const rows = [
    { drawingNo: 'D0101-01', source: 'a.pdf', title: '主接线', frameSize: 'A4' },
    { drawingNo: 'D0101-01', source: 'a.pdf', title: '主接线', frameSize: 'A4' },
    { drawingNo: null, source: 'a.pdf', title: '平面布置', frameSize: 'A3' },
    { drawingNo: null, source: 'a.pdf', title: '平面布置', frameSize: 'A3' },
    { drawingNo: null, source: 'a.pdf', title: '平面布置', frameSize: 'A1' },
    { drawingNo: 'D0101-01', source: 'b.pdf', title: '主接线', frameSize: 'A4' },
  ];
  const annotated = annotateCatalogCounts(rows);
  assert.deepEqual(annotated.map((row) => row.count), [2, 2, 2, 2, 1, 1]);
});

test('工作表：单来源只有「汇总」；多来源追加同名截断去重的工作表', () => {
  const rows = [
    { drawingNo: 'D0101-01', source: join(TMP, '甲站.pdf'), title: '主接线', frameSize: 'A4', orientation: 'landscape', count: 1, seq: 1 },
    { drawingNo: 'D0101-02', source: join(TMP, '乙站.pdf'), title: '平面布置', frameSize: 'A3', orientation: 'portrait', count: 1, seq: 2 },
  ];
  assert.deepEqual(buildCatalogSheets(rows, [{ file: rows[0].source }]).map((sheet) => sheet.name), ['汇总']);

  const longName = join(TMP, `${'超长文件名'.repeat(8)}.pdf`);
  const many = [
    { ...rows[0], source: join(TMP, 'sub', '同名.pdf') },
    { ...rows[1], source: join(TMP, '其它', '同名.pdf') },
    { ...rows[1], source: longName },
  ];
  const sheets = buildCatalogSheets(many, many.map((row) => ({ file: row.source })));
  assert.equal(sheets[0].name, '汇总');
  assert.equal(sheets.length, 4);
  const names = sheets.slice(1).map((sheet) => sheet.name);
  assert.equal(new Set(names).size, names.length, '工作表名去重');
  assert.ok(names.every((name) => name.length <= 31), '工作表名 ≤31 字符');
  assert.equal(sheets[0].rows.length, 3, '汇总含全部行');
  assert.equal(sheets[0].rows[1].source, '同名.pdf');
});

// ───────────────────────── CSV 降级 / 展示格式 ─────────────────────────

test('CSV 降级：UTF-8 BOM + 表头 + CRLF + 引号转义', () => {
  const rows = [{
    seq: 1,
    drawingNo: 'D0101-01',
    title: '主接线,（含逗号）',
    frameSize: 'A4',
    orientation: 'landscape',
    scale: '1:100',
    count: 2,
    source: '甲站"引号".pdf',
  }];
  const csv = rowsToCsv(rows);
  assert.ok(csv.startsWith('\uFEFF'), 'UTF-8 BOM 前缀');
  assert.ok(csv.includes('\r\n'), 'CRLF 行尾');
  const lines = csv.replace(/^\uFEFF/, '').split('\r\n');
  assert.equal(lines[0], CATALOG_HEADERS.join(','));
  assert.equal(lines[1], '1,D0101-01,"主接线,（含逗号）",A4（横）,1:100,2,"甲站""引号"".pdf"');
});

test('图幅展示：横向/纵向标签', () => {
  assert.equal(formatFrameSizeLabel('A4', 'landscape'), 'A4（横）');
  assert.equal(formatFrameSizeLabel('A3', 'portrait'), 'A3（竖）');
  assert.equal(formatFrameSizeLabel('A0', undefined), 'A0（横）');
  assert.equal(formatFrameSizeLabel(null, 'landscape'), '');
});
