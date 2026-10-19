/**
 * 引擎适配层（v0.1）：通过 ctx.subprocess 运行随包 PowerShell 引擎脚本。
 * 后续版本在此扩展：引擎探测缓存、操作级降级链、并行调度、Python worker 等。
 */
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));
const scriptPath = (name) => join(HERE, '..', 'engines', name);

/**
 * 解析本机可用的 PowerShell：优先 pwsh（7+），回退 powershell（5.1）。
 * 两个版本的脚本保持 5.1 兼容语法（不使用 7+ 独有运算符）。
 */
async function resolveShell(ctx, signal) {
  for (const name of ['pwsh', 'powershell']) {
    try {
      return await ctx.subprocess.resolveExecutable(name, {}, signal);
    } catch { /* 尝试下一个 */ }
  }
  throw new Error('本机未找到 PowerShell（pwsh / powershell），套件引擎无法运行');
}

/**
 * 运行随包 PowerShell 脚本，传入单个 JSON 参数（脚本以 -Json 接收）。
 * @returns {{ exitCode: number|null, signal: string|null, stdout: string, stderr: string }}
 */
export async function runPs1(ctx, name, arg, options = {}) {
  const signal = options.signal;
  const exe = await resolveShell(ctx, signal);
  const handle = ctx.subprocess.spawn({
    argv: [
      exe,
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', scriptPath(name),
      '-Json', JSON.stringify(arg ?? {}),
    ],
    cwd: options.cwd ?? process.cwd(),
    stdio: {
      stdin: 'ignore',
      stdout: { maxBytes: 8 * 1024 * 1024, spill: { maxBytes: 64 * 1024 * 1024 } },
      stderr: { maxBytes: 2 * 1024 * 1024 },
    },
    graceMs: options.graceMs ?? 60_000,
    signal,
    env: { ...process.env, ...(options.env ?? {}) },
  });

  const timer = options.timeoutMs
    ? setTimeout(() => { try { handle.terminate(); } catch { /* 已结束 */ } }, options.timeoutMs)
    : null;
  let outcome;
  try {
    outcome = await handle.done;
  } finally {
    if (timer) clearTimeout(timer);
  }

  return {
    exitCode: outcome.exitCode,
    signal: outcome.signal,
    stdout: handle.collected.stdout?.readFrom(0).text ?? '',
    stderr: handle.collected.stderr?.readFrom(0).text ?? '',
  };
}
