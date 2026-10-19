/**
 * D1：渲染导出公共层 office-pdf.js（docs/P3-设计.md §5.1 / §5.3）
 *
 * 全部使用 stub officeToPdf 服务验证 convert 契约与错误分支：
 *   - 服务缺失/扩展名不支持 → E2012；写盘失败/越界 → E2013/E1021；渲染超限 → E2011；
 *   - 契约字段：{ extension, priority, source:{ key, version, bytes, read(signal,maxBytes) } }；
 *   - 返回：pdfBase64 / missingFonts / generation / cacheKey（= version|generation）；
 *   - saveAs：写入显式临时路径（**不写 output\pdf**），重名追加时间戳。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { convertOfficeToPdf, previewPdfOp } from '../suite/host/ops/office-pdf.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `office-pdf-${process.pid}`);

/** 合成 PDF 字节（stub 服务返回；内容不必可解析）。 */
const PDF_BYTES = Buffer.from('%PDF-1.7\n% stub officeToPdf output\n%%EOF\n', 'utf8');

let docxPath;
let docxSize;

before(async () => {
  await mkdir(TMP, { recursive: true });
  docxPath = join(TMP, '白云变-主接线.docx');
  await writeFile(docxPath, Buffer.alloc(2048, 0x41));
  docxSize = (await stat(docxPath)).size;
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

/** stub Host 服务：记录请求，按 options 返回/抛错。 */
function makeService(options = {}) {
  const calls = [];
  const service = {
    calls,
    async convert(request, signal) {
      calls.push({ request, signal });
      if (options.onConvert) return options.onConvert(request, signal);
      if (options.fail) throw new Error(options.fail);
      if (options.noPdf) return { missingFonts: [] };
      // 模拟真实服务读取 source：验证 read(signal, maxBytes) 契约
      if (options.readBytes) await request.source.read(undefined, options.maxRead ?? 10 * 1024 * 1024);
      return {
        pdf: options.pdf ?? new Uint8Array(PDF_BYTES),
        missingFonts: options.missingFonts ?? [],
        generation: options.generation ?? 7,
      };
    },
  };
  return service;
}

function makeCtx(service) {
  return {
    get(name) {
      return name === 'officeToPdf' ? service : undefined;
    },
  };
}

function decodeBase64(value) {
  return Buffer.from(String(value ?? ''), 'base64');
}

// ───────────────────────── 服务缺失 / 扩展名 ─────────────────────────

test('E2012：officeToPdf 服务缺失（ctx.get 返回 undefined）', async () => {
  const result = await previewPdfOp(makeCtx(undefined), {}, { path: docxPath }, {});
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2012');
  assert.match(result.error.message, /officeToPdf/);
});

test('E2012：扩展名不在 6 类（.pdf/.txt）', async () => {
  const service = makeService();
  const ctx = makeCtx(service);
  const pdfPath = join(TMP, 'x.pdf');
  await writeFile(pdfPath, '%PDF-1.7\n');
  for (const path of [pdfPath, join(TMP, 'x.txt')]) {
    const result = await previewPdfOp(ctx, {}, { path }, {});
    assert.equal(result.ok, false, path);
    assert.equal(result.error.code, 'E2012');
  }
  assert.equal(service.calls.length, 0, '扩展名不支持时不调用服务');
});

test('E1041：缺少 path / saveAs 非法类型 / engine 非法', async () => {
  const ctx = makeCtx(makeService());
  assert.equal((await previewPdfOp(ctx, {}, {}, {})).error.code, 'E1041');
  assert.equal((await previewPdfOp(ctx, {}, { path: docxPath, saveAs: 42 }, {})).error.code, 'E1041');
  assert.equal((await previewPdfOp(ctx, {}, { path: docxPath, engine: 'com' }, {})).error.code, 'E1041');
});

// ───────────────────────── 成功路径（base64 预览） ─────────────────────────

test('预览成功：pdfBase64 / missingFonts / generation / cacheKey，convert 契约字段齐全', async () => {
  const service = makeService({ missingFonts: ['SimSun'], generation: 11 });
  const result = await previewPdfOp(makeCtx(service), {}, { path: docxPath }, {});
  assert.equal(result.ok, true);
  const info = await stat(docxPath);
  assert.deepEqual(decodeBase64(result.result.pdfBase64), PDF_BYTES);
  assert.deepEqual(result.result.missingFonts, ['SimSun']);
  assert.equal(result.result.generation, 11);
  assert.equal(result.result.cacheKey, `${info.mtimeMs}|11`);
  assert.equal(result.result.extension, 'docx');
  assert.equal(result.result.engine, 'officeToPdf');
  assert.equal(result.result.size, PDF_BYTES.length);
  assert.equal(result.result.outputPath, undefined, '预览不落盘');

  const { request } = service.calls[0];
  assert.equal(request.extension, 'docx');
  assert.equal(request.priority, 'foreground');
  assert.equal(request.source.key, 'suite-preview');
  assert.equal(request.source.version, String(info.mtimeMs));
  assert.equal(request.source.bytes, docxSize);
  const read = await request.source.read(undefined, 64 * 1024 * 1024);
  assert.equal(read.version, String(info.mtimeMs));
  assert.equal(Buffer.compare(Buffer.from(read.bytes), await readFile(docxPath)), 0);
});

test('config.previewPriority=background → request.priority=background', async () => {
  const service = makeService();
  const result = await previewPdfOp(makeCtx(service), { previewPriority: 'background' }, { path: docxPath }, {});
  assert.equal(result.ok, true);
  assert.equal(service.calls[0].request.priority, 'background');
});

// ───────────────────────── 渲染异常 ─────────────────────────

test('E2011：服务返回缺少 pdf 字节 / 服务抛错', async () => {
  const noPdf = await previewPdfOp(makeCtx(makeService({ noPdf: true })), {}, { path: docxPath }, {});
  assert.equal(noPdf.error.code, 'E2011');
  assert.match(noPdf.error.message, /pdf/);

  const boom = await previewPdfOp(makeCtx(makeService({ fail: '渲染进程崩溃' })), {}, { path: docxPath }, {});
  assert.equal(boom.error.code, 'E2011');
  assert.match(boom.error.message, /渲染进程崩溃/);
});

test('E2011：source.read 的 maxBytes 上限（服务按小上限读取）', async () => {
  const service = makeService({ readBytes: true, maxRead: 16 });
  const result = await previewPdfOp(makeCtx(service), {}, { path: docxPath }, {});
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2011');
  assert.match(result.error.message, /渲染大小上限/);
});

test('扩展名覆盖 6 类：pptx 请求 extension=pptx（P3 §5.4 预览扩展）', async () => {
  const pptxPath = join(TMP, '汇报.pptx');
  await writeFile(pptxPath, Buffer.alloc(512, 0x42));
  const service = makeService();
  const result = await previewPdfOp(makeCtx(service), {}, { path: pptxPath }, {});
  assert.equal(result.ok, true);
  assert.equal(service.calls[0].request.extension, 'pptx');
  assert.equal(result.result.extension, 'pptx');
});

// ───────────────────────── saveAs 落盘 ─────────────────────────

test('saveAs：写显式临时路径（原子写），返回 outputPath/size/saved', async () => {
  const service = makeService({ missingFonts: [], generation: 3 });
  const target = join(TMP, '导出', '主接线.pdf');
  const result = await previewPdfOp(makeCtx(service), {}, { path: docxPath, saveAs: target }, {});
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.result.outputPath, target);
  assert.equal(result.result.saved, true);
  assert.equal(result.result.size, PDF_BYTES.length);
  assert.equal(result.result.pdfBase64, undefined, '落盘模式不回传 base64');
  assert.deepEqual(await readFile(target), PDF_BYTES);
});

test('saveAs 目录 → 同名 .pdf；重名追加 -<时间戳> 且不覆盖既有文件', async () => {
  const service = makeService();
  const existing = join(TMP, '白云变-主接线.pdf');
  await writeFile(existing, 'old');
  const first = await previewPdfOp(makeCtx(service), {}, { path: docxPath, saveAs: TMP }, {});
  assert.equal(first.ok, true);
  assert.notEqual(first.result.outputPath, existing, '重名不覆盖既有文件');
  assert.match(first.result.outputPath, /白云变-主接线-\d{14}\.pdf$/);
  assert.equal(await readFile(existing, 'utf8'), 'old');
  assert.deepEqual(await readFile(first.result.outputPath), PDF_BYTES);
});

test('saveAs=true：落到 config.outputRoot（这里指向临时目录，避免写 output\\pdf）', async () => {
  const service = makeService();
  const outputRoot = join(TMP, 'output-root');
  const result = await previewPdfOp(makeCtx(service), { outputRoot }, { path: docxPath, saveAs: true }, {});
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.result.outputPath, join(outputRoot, '白云变-主接线.pdf'));
  assert.deepEqual(await readFile(result.result.outputPath), PDF_BYTES);
});

test('E1021：saveAs 越界（C:\\Windows）被 boundary 拒绝（写盘不受 allowExternalPaths 放宽）', async () => {
  const service = makeService();
  const result = await previewPdfOp(
    makeCtx(service),
    { allowExternalPaths: true },
    { path: docxPath, saveAs: 'C:\\Windows\\office-suite-export.pdf' },
    {},
  );
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E1021');
  assert.equal(service.calls.length, 0, '越界时不应调用渲染服务');
});

test('E2013：目标目录不可用（父路径是文件）→ 写盘失败且清理半成品', async () => {
  const service = makeService();
  const parentAsFile = join(TMP, 'parent-is-file.pdf');
  await writeFile(parentAsFile, 'not a directory');
  const result = await previewPdfOp(makeCtx(service), {}, { path: docxPath, saveAs: join(parentAsFile, 'out.pdf') }, {});
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2013');
  assert.equal(await readFile(parentAsFile, 'utf8'), 'not a directory', '既有文件不被破坏');
});

// ───────────────────────── 公共函数直调 ─────────────────────────

test('convertOfficeToPdf：服务缺失 E2012；扩展名 E2012；文件缺失 E1023', async () => {
  await assert.rejects(
    () => convertOfficeToPdf(makeCtx(undefined), {}, docxPath),
    (err) => err?.name === 'OpsError' && err.code === 'E2012',
  );
  await assert.rejects(
    () => convertOfficeToPdf(makeCtx(makeService()), {}, join(TMP, 'x.txt')),
    (err) => err?.name === 'OpsError' && err.code === 'E2012',
  );
  await assert.rejects(
    () => convertOfficeToPdf(makeCtx(makeService()), {}, join(TMP, 'missing.docx')),
    (err) => err?.name === 'OpsError' && err.code === 'E1023',
  );
});
