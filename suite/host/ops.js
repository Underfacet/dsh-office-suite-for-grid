/**
 * 统一操作层：所有操作（工具、远程命令、面板）返回同一结果形状。
 * 成功：{ ok: true, result: ... }
 * 失败：{ ok: false, error: { code, message, hint? } }
 */
import { resolve, isAbsolute, dirname } from 'node:path';

/** 从 PowerShell 输出中提取首个 JSON 对象（容忍前导警告行）。 */
export function parseJson(text) {
  if (!text) return undefined;
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) return undefined;
  try {
    return JSON.parse(text.slice(start, end + 1));
  } catch {
    return undefined;
  }
}

/** 当前会话的工作目录（无会话时退回进程目录）。 */
export async function sessionCwd(ctx, exec) {
  if (exec?.agent) {
    try {
      return await ctx.workingDirectory.ensure(exec.agent, exec.signal);
    } catch { /* 回退 */ }
  }
  return process.cwd();
}

/** 绝对路径原样返回；相对路径按会话工作目录解析。 */
export async function resolveInputPath(ctx, exec, p) {
  if (isAbsolute(p)) return p;
  return resolve(await sessionCwd(ctx, exec), p);
}

export const dirnameOf = (p) => dirname(p);

export const ok = (result) => ({ ok: true, result });
export const fail = (code, message, hint) => ({ ok: false, error: { code, message, ...(hint ? { hint } : {}) } });
