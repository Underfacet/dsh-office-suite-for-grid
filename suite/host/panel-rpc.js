/**
 * 面板本机 HTTP RPC 路由（替代 `remote.commands.execute` 网关调用）
 *
 * 背景：面板经 `ctx.remote.commands.execute` 调宿主命令时，typert 网关按 descriptor
 * 校验参数并报 `missing "submittedAttachments"`（框架层 descriptor 不一致）。本文件由
 * 插件自注册同源 HTTP 路由，把请求直接分发给 commands-panel.js 的 `runSuiteCommand`；
 * 命令注册（commands-panel.js 里的 `ctx.commands.register`）保持不变，命令栏/智能体
 * 仍走原通道。
 *
 * 契约：
 *   POST /office-suite/rpc
 *   body: { command: string, args?: object }
 *   - 命令层结果（成功或业务错误）：200 + `{kind:'success'|'error', text}`（沿用既有信封）；
 *   - body 非法：400 `{ok:false,error:{code:'BAD_REQUEST',message}}`；
 *   - body 超限（>1MB）：413 `{ok:false,error:{code:'PAYLOAD_TOO_LARGE',message}}`；
 *   - 非 POST：405（Allow: POST）；
 *   - 鉴权失败：`connection.requestRejection(req)` 返回的 401/403（connection 缺省放行，
 *     仅限本机 loopback 部署）；
 *   - 未捕获异常：500 `{kind:'error', text: JSON.stringify({ok:false,error:{code:'RPC_FAILED',...}})}`。
 *
 * 日志只记录 command 与耗时：args 可能含文件路径/正文，按涉密内容处理，不打印。
 */
import { runSuiteCommand } from './commands-panel.js';

/** 面板 RPC 路由路径（与 client.js 的同源相对地址一致）。 */
export const PANEL_RPC_PATH = '/office-suite/rpc';

/** 单请求 body 上限（1MB）。 */
export const MAX_RPC_BODY_BYTES = 1024 * 1024;

/** 写 JSON 响应（连接已结束/销毁时静默跳过，避免二次写触发 unhandled error）。 */
function sendJson(res, status, payload) {
  if (!res || res.writableEnded || res.destroyed) return;
  const body = JSON.stringify(payload);
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.end(body);
}

/** 400/413 等路由层错误（不是命令层信封）。 */
function sendRouteError(res, status, code, message, headers) {
  if (!res || res.writableEnded || res.destroyed) return;
  if (headers && typeof headers === 'object') {
    for (const [key, value] of Object.entries(headers)) res.setHeader(key, value);
  }
  sendJson(res, status, { ok: false, error: { code, message } });
}

function payloadTooLargeError() {
  const error = new Error('payload too large');
  error.code = 'PAYLOAD_TOO_LARGE';
  return error;
}

/**
 * 读取请求体（上限 1MB）。超限时立即 reject（code=PAYLOAD_TOO_LARGE），
 * 后续数据块继续被忽略以避免破坏 socket。
 * @param {import('node:http').IncomingMessage} req
 * @param {number} [limit]
 * @returns {Promise<string>}
 */
function readRequestBody(req, limit = MAX_RPC_BODY_BYTES) {
  return new Promise((resolve, reject) => {
    let size = 0;
    let tooLarge = false;
    const chunks = [];
    req.on('data', (chunk) => {
      if (tooLarge) return;
      const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
      size += buffer.length;
      if (size > limit) {
        tooLarge = true;
        chunks.length = 0;
        reject(payloadTooLargeError());
        return;
      }
      chunks.push(buffer);
    });
    req.on('end', () => {
      if (tooLarge) return;
      resolve(Buffer.concat(chunks).toString('utf8'));
    });
    req.once('aborted', () => {
      if (tooLarge) return;
      reject(new Error('request aborted'));
    });
    req.on('error', (error) => {
      if (tooLarge) return;
      reject(error);
    });
  });
}

/** 解析并校验 `{command,args}` 信封；失败返回 {error:{code,message}}。 */
function parseEnvelope(raw) {
  let value;
  try {
    value = JSON.parse(raw);
  } catch {
    return { error: { code: 'BAD_REQUEST', message: 'invalid JSON body' } };
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { error: { code: 'BAD_REQUEST', message: 'body must be a JSON object' } };
  }
  const command = value.command;
  if (typeof command !== 'string' || !command.trim()) {
    return { error: { code: 'BAD_REQUEST', message: 'command must be a non-empty string' } };
  }
  const args = value.args === undefined || value.args === null ? {} : value.args;
  if (typeof args !== 'object' || Array.isArray(args)) {
    return { error: { code: 'BAD_REQUEST', message: 'args must be a JSON object' } };
  }
  return { command: command.trim(), args };
}

/**
 * 构造 HTTP handler。
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 插件 Config（与命令注册同一份）
 */
function createHandler(ctx, config) {
  return async function panelRpcHandler(req, res) {
    // 1) 仅 POST
    if (req.method !== 'POST') {
      res.setHeader('Allow', 'POST');
      res.statusCode = 405;
      res.end();
      return;
    }

    // 2) 读取 body（≤1MB）；Content-Length 已超限时直接 413，不读流。
    const declared = Number(req.headers?.['content-length']);
    if (Number.isFinite(declared) && declared > MAX_RPC_BODY_BYTES) {
      try { req.resume?.(); } catch { /* 忽略 */ }
      sendRouteError(res, 413, 'PAYLOAD_TOO_LARGE', `请求体超过 ${MAX_RPC_BODY_BYTES} 字节`);
      return;
    }
    let raw;
    try {
      raw = await readRequestBody(req);
    } catch (error) {
      if (error?.code === 'PAYLOAD_TOO_LARGE') {
        sendRouteError(res, 413, 'PAYLOAD_TOO_LARGE', `请求体超过 ${MAX_RPC_BODY_BYTES} 字节`);
        return;
      }
      sendRouteError(res, 400, 'BAD_REQUEST', `读取请求体失败: ${String(error?.message ?? error)}`);
      return;
    }

    // 3) 鉴权：connection 缺省时放行（仅限本机 loopback 部署）。
    try {
      const connection = typeof ctx?.get === 'function' ? ctx.get('connection') : undefined;
      const rejection = connection?.requestRejection?.(req);
      if (rejection) {
        res.statusCode = rejection;
        res.end();
        return;
      }
    } catch (error) {
      console.error('[office-suite] RPC 鉴权检查失败:', error);
      res.statusCode = 403;
      res.end();
      return;
    }

    // 4) 解析信封
    const envelope = parseEnvelope(raw);
    if (envelope.error) {
      sendRouteError(res, 400, envelope.error.code, envelope.error.message);
      return;
    }
    const { command, args } = envelope;

    // 5) 分发到命令层（与命令注册共用同一实现）
    const controller = new AbortController();
    const abort = () => {
      if (!res.writableEnded && !res.destroyed) controller.abort();
    };
    req.once?.('aborted', abort);
    res.once?.('close', abort);

    const startedAt = Date.now();
    try {
      const result = await runSuiteCommand(ctx, config, command, args, controller.signal);
      const elapsed = Date.now() - startedAt;
      // 只记 command 与耗时，不打印 args（涉密）。
      console.info(`[office-suite] rpc ${command} ${elapsed}ms`);
      sendJson(res, 200, result);
    } catch (error) {
      const elapsed = Date.now() - startedAt;
      console.error(`[office-suite] rpc ${command} failed ${elapsed}ms:`, error);
      sendJson(res, 500, {
        kind: 'error',
        text: JSON.stringify({
          ok: false,
          error: { code: 'RPC_FAILED', message: String(error?.message ?? error) },
        }),
      });
    } finally {
      req.off?.('aborted', abort);
      res.off?.('close', abort);
    }
  };
}

/**
 * 注册面板 RPC 路由。
 * @param {object} ctx DSH 插件上下文
 * @param {object} [config] 插件 Config
 * @returns {() => void} disposer（webServer 缺失/注册失败时返回空函数）
 */
export function registerPanelRpc(ctx, config = {}) {
  let web;
  try {
    web = typeof ctx?.get === 'function' ? ctx.get('webServer') : undefined;
  } catch {
    web = undefined;
  }
  if (!web || typeof web.register !== 'function') {
    console.warn('[office-suite] webServer 不可用：面板 RPC 路由未注册，面板将无法调用宿主命令');
    return () => {};
  }
  try {
    const dispose = web.register({
      kind: 'exact',
      path: PANEL_RPC_PATH,
      handler: createHandler(ctx, config),
    });
    return typeof dispose === 'function' ? dispose : () => {};
  } catch (error) {
    console.error('[office-suite] 面板 RPC 路由注册失败:', error);
    return () => {};
  }
}
