/**
 * 渲染导出公共层（docs/P3-设计.md §5.1 / §5.3）
 *
 * 把 suite.preview 中的 officeToPdf.convert 调用逻辑抽为共用函数，供：
 *   - suite.preview（面板预览，iframe blob）→ previewPdfOp（saveAs 省略 → pdfBase64）；
 *   - suite.export / office_export（导出 PDF 文件）→ previewPdfOp({saveAs}) 或 convertOfficeToPdf。
 *
 * 契约（P2 §3.7 已核实，P3 §5.1 沿用）：
 *   convert({ extension, priority, source:{ key, version, bytes, read(signal,maxBytes) } }, signal)
 *     → { pdf: Uint8Array, missingFonts?: string[], generation?: number|string, cacheKey?: string }
 *   缓存键 = `${version}|${generation}`（version = 源文件 mtimeMs）。
 *
 * 错误码：E2012 EXPORT_UNSUPPORTED（服务缺失/扩展名不支持）、E2013 EXPORT_WRITE_FAILED（写盘失败）、
 *         E2011（渲染失败/输出超上限）、E1021/E1023/E1041（路径参数）。
 *
 * 导出：
 *   EXPORT_EXTENSIONS / MAX_EXPORT_BYTES
 *   getOfficeToPdfService(ctx)
 *   convertOfficeToPdf(ctx, config, absPath, signal, options?)
 *   resolvePdfOutputPath(ctx, config, exec, sourceAbs, saveAs)
 *   previewPdfOp(ctx, config, args{path,saveAs?,engine?}, exec)
 */
import { mkdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import {
  OpsError,
  fail,
  isAbortError,
  ok,
  resolveInputPath,
  sessionCwd,
  throwIfAborted,
  toErrorPayload,
} from '../ops.js';

/** 支持的导出/预览扩展名（§5.1：officeToPdf OfficeExtension 的 6 类）。 */
export const EXPORT_EXTENSIONS = Object.freeze(['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']);

/** 单文件渲染上限（与 P2 预览的 64MB 一致，R-24）。 */
export const MAX_EXPORT_BYTES = 64 * 1024 * 1024;

/** 默认优先级（config.previewPriority 可覆盖）。 */
const DEFAULT_PRIORITY = 'foreground';

function timestamp() {
  return new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
}

function randomSuffix() {
  return Math.random().toString(36).slice(2, 8);
}

function validateEngineHint(value) {
  if (value === undefined || value === null || value === '') return 'auto';
  if (value === 'auto' || value === 'officeToPdf') return value;
  throw new OpsError(
    'E1041',
    "engine 必须为 'auto' | 'officeToPdf'",
    'PDF 渲染固定由 Host officeToPdf.convert 完成（无 COM/Python 引擎可选）',
  );
}

function isAbort(err) {
  return isAbortError(err) || err?.name === 'AbortError';
}

/**
 * 读取 Host officeToPdf 服务（access.optional：未挂载时 ctx.get 返回 undefined）。
 * @returns {{convert:Function}|null}
 */
export function getOfficeToPdfService(ctx) {
  try {
    const service = typeof ctx?.get === 'function' ? ctx.get('officeToPdf') : undefined;
    return service && typeof service.convert === 'function' ? service : null;
  } catch {
    return null;
  }
}

function serviceMissingError() {
  return new OpsError(
    'E2012',
    'Host officeToPdf 服务不可用（未挂载或缺少 convert）',
    '预览/导出需要宿主提供 officeToPdf.convert；可改用系统应用打开文件，或在 Client 走 remote.officeToPdf 直连（P3 §5.2 方案 A）',
  );
}

function unsupportedExtensionError(ext) {
  return new OpsError(
    'E2012',
    `该类型不支持导出/预览: ${ext || '(无扩展名)'}`,
    `仅支持 ${EXPORT_EXTENSIONS.join('/')}（officeToPdf 的 OfficeExtension 不含 vsdx/dwg/pdf）`,
  );
}

/** PDF 字节归一化（Uint8Array / ArrayBuffer / Buffer）。 */
function toPdfBytes(value) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  return null;
}

/**
 * 公共转换：文件 → PDF 字节（officeToPdf.convert）。
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 有效配置（previewPriority / 上限）
 * @param {string} absPath 已过 boundary 的源文件绝对路径
 * @param {AbortSignal} [signal]
 * @param {{key?:string, service?:object, maxBytes?:number}} [options]
 * @returns {Promise<{pdf:Uint8Array, missingFonts:string[], generation:any, cacheKey:string, version:string, size:number, extension:string}>}
 * @throws {OpsError} E2012 / E2011 / E1023
 */
export async function convertOfficeToPdf(ctx, config = {}, absPath, signal, options = {}) {
  const service = options.service ?? getOfficeToPdfService(ctx);
  if (!service) throw serviceMissingError();
  const ext = extname(absPath).toLowerCase();
  if (!EXPORT_EXTENSIONS.includes(ext)) throw unsupportedExtensionError(ext);

  let info;
  try {
    info = await stat(absPath);
  } catch {
    throw new OpsError('E1023', `文件不存在: ${absPath}`);
  }
  if (!info.isFile()) throw new OpsError('E1023', `不是文件: ${absPath}`);
  const limit = Number.isFinite(options.maxBytes) && options.maxBytes > 0
    ? Math.min(Number(options.maxBytes), MAX_EXPORT_BYTES)
    : MAX_EXPORT_BYTES;
  if (info.size > limit) {
    throw new OpsError('E2011', `文件超过渲染大小上限（${info.size} > ${limit} 字节）`, '可改用系统应用打开；或拆分/压缩后重试');
  }

  const version = String(info.mtimeMs ?? 0);
  const priority = config?.previewPriority === 'background' ? 'background' : DEFAULT_PRIORITY;
  const request = {
    extension: ext.slice(1), // 'doc' | 'docx' | 'xls' | 'xlsx' | 'ppt' | 'pptx'
    priority,
    source: {
      key: typeof options.key === 'string' && options.key.trim() ? options.key.trim() : 'suite-preview',
      version,
      bytes: info.size,
      read: async (readSignal, maxBytes) => {
        const requested = Number(maxBytes);
        const readLimit = Number.isFinite(requested) && requested > 0 ? Math.min(requested, limit) : limit;
        if (info.size > readLimit) {
          throw new OpsError('E2011', `文件超过渲染大小上限（${info.size} > ${readLimit} 字节）`);
        }
        const bytes = await readFile(absPath, { signal: readSignal ?? signal });
        return { bytes, version };
      },
    },
  };

  let output;
  try {
    output = await service.convert(request, signal);
  } catch (err) {
    if (isAbort(err)) throw err;
    if (err instanceof OpsError) throw err;
    throw new OpsError('E2011', `PDF 渲染失败: ${String(err?.message ?? err)}`, '请确认文件未被占用且未超过渲染大小上限；也可改用系统应用打开');
  }
  const pdf = toPdfBytes(output?.pdf);
  if (!pdf || !pdf.byteLength) {
    throw new OpsError('E2011', 'officeToPdf.convert 返回结果缺少 pdf 字节');
  }
  const generation = output?.generation ?? null;
  return {
    pdf,
    missingFonts: Array.isArray(output?.missingFonts) ? output.missingFonts.map((item) => String(item)) : [],
    generation,
    cacheKey: `${version}|${generation ?? ''}`,
    version,
    size: pdf.byteLength,
    extension: ext.slice(1),
  };
}

/** 默认 PDF 输出根：config.outputRoot 或 会话工作区 output\pdf（§5.3）。 */
async function defaultPdfOutputRoot(ctx, exec, config) {
  if (typeof config?.outputRoot === 'string' && config.outputRoot.trim()) {
    const configured = config.outputRoot.trim();
    return isAbsolute(configured) ? resolve(configured) : resolve(await sessionCwd(ctx, exec), configured);
  }
  return join(await sessionCwd(ctx, exec), 'output', 'pdf');
}

async function statOrNull(path) {
  try {
    return await stat(path);
  } catch {
    return null;
  }
}

/**
 * 解析导出 PDF 目标路径（boundary 写校验；重名加 -<时间戳>）。
 * @param {string|boolean} saveAs true=默认目录同名 .pdf；字符串=目录或 .pdf 路径
 * @returns {Promise<string>} 绝对路径
 */
export async function resolvePdfOutputPath(ctx, config, exec, sourceAbs, saveAs) {
  const root = resolveBoundaryRoot(config);
  const defaultName = `${basename(String(sourceAbs), extname(String(sourceAbs)))}.pdf`;
  let target;
  if (saveAs === true) {
    target = join(await defaultPdfOutputRoot(ctx, exec, config), defaultName);
  } else if (typeof saveAs === 'string' && saveAs.trim()) {
    const abs = assertInsideRoot(await resolveInputPath(ctx, exec, saveAs.trim()), { write: true, root });
    const info = await statOrNull(abs);
    if (info?.isDirectory() || (!info && !/\.pdf$/i.test(abs))) {
      target = join(abs, defaultName);
    } else if (extname(abs).toLowerCase() === '.pdf') {
      target = abs;
    } else {
      throw new OpsError('E1041', `导出目标必须是 .pdf 文件或目录: ${abs}`);
    }
  } else {
    throw new OpsError('E1041', 'saveAs 必须为 true 或 .pdf 路径/目录');
  }
  target = assertInsideRoot(target, { write: true, root });
  if (extname(target).toLowerCase() !== '.pdf') {
    throw new OpsError('E1041', `导出目标必须是 .pdf 文件: ${target}`);
  }
  if (await statOrNull(target)) {
    // 重名：<原名>-<时间戳>.pdf（§5.3）
    target = target.replace(/\.pdf$/i, `-${timestamp()}.pdf`);
    target = assertInsideRoot(target, { write: true, root });
  }
  return target;
}

/**
 * 原子写 PDF：先写临时文件再 rename；失败清理半成品（§5.3）。
 * @throws {OpsError} E2013
 */
async function writePdfAtomic(target, bytes) {
  const tmp = `${target}.${Date.now().toString(36)}-${randomSuffix()}.tmp`;
  try {
    await mkdir(dirname(target), { recursive: true });
    await writeFile(tmp, bytes);
    await rename(tmp, target);
  } catch (err) {
    try { await rm(tmp, { force: true }); } catch { /* 清理失败忽略 */ }
    throw new OpsError(
      'E2013',
      `导出 PDF 写盘失败: ${target}`,
      String(err?.message ?? err) || '请检查输出目录是否可写、磁盘空间，以及目标文件是否被占用',
    );
  }
}

/**
 * 预览/导出 PDF（suite.preview / suite.export 共用）。
 *
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 有效配置
 * @param {{path:string, saveAs?:string|boolean, engine?:'auto'|'officeToPdf'}} args
 *   - saveAs 省略 → 返回 { pdfBase64, missingFonts, generation, cacheKey }（面板 iframe）；
 *   - saveAs 给值 → 落盘（默认 output\pdf，重名加 -<时间戳>）并返回 { outputPath, size, ... }。
 * @param {object} exec 工具/命令执行上下文（agent/signal）
 * @returns {Promise<{ok:true,result:object}|{ok:false,error:object}>}
 */
export async function previewPdfOp(ctx, config = {}, args = {}, exec = {}) {
  try {
    const engine = validateEngineHint(args?.engine);
    const rawPath = typeof args?.path === 'string' ? args.path.trim() : '';
    if (!rawPath) throw new OpsError('E1041', '缺少 path');
    // 先判空 Host 服务（§5.1：access.optional → 缺失即 E2012）
    const service = getOfficeToPdfService(ctx);
    if (!service) throw serviceMissingError();

    const abs = assertInsideRoot(await resolveInputPath(ctx, exec, rawPath), {
      write: false,
      allowExternalPaths: config?.allowExternalPaths === true,
      root: resolveBoundaryRoot(config),
      warn: (message) => console.warn(`[office-suite] ${message}`),
    });
    const ext = extname(abs).toLowerCase();
    if (!EXPORT_EXTENSIONS.includes(ext)) throw unsupportedExtensionError(ext);
    throwIfAborted(exec?.signal);

    // 先解析落盘目标（boundary 校验 + 重名处理），避免渲染完才发现路径非法
    const wantsFile = args?.saveAs !== undefined && args?.saveAs !== null;
    const target = wantsFile ? await resolvePdfOutputPath(ctx, config, exec, abs, args.saveAs) : null;

    const converted = await convertOfficeToPdf(ctx, config, abs, exec?.signal, { service });
    const base = {
      missingFonts: converted.missingFonts,
      generation: converted.generation,
      cacheKey: converted.cacheKey,
      extension: converted.extension,
      engine: 'officeToPdf',
      engineHint: engine,
      size: converted.size,
    };

    if (!wantsFile) {
      return ok({
        ...base,
        pdfBase64: Buffer.from(converted.pdf).toString('base64'),
      });
    }
    await writePdfAtomic(target, converted.pdf);
    return ok({
      ...base,
      outputPath: target,
      saved: true,
    });
  } catch (err) {
    if (isAbortError(err)) throw err;
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}
