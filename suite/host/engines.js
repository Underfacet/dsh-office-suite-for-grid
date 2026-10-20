/**
 * 引擎适配层（v0.1）：通过 ctx.subprocess 运行随包 PowerShell 引擎脚本。
 * 后续版本在此扩展：引擎探测缓存、操作级降级链、并行调度、Python worker 等。
 */
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { existsSync } from 'node:fs';

const HERE = dirname(fileURLToPath(import.meta.url));
const scriptPath = (name) => join(HERE, '..', 'engines', name);

/**
 * 候选 PowerShell 绝对路径（宿主子进程服务的 PATH 可能被净化，裸命令解析不到时兜底）。
 * 可用环境变量 SUITE_SHELL_PATH 强制指定。
 */
function shellCandidates() {
  const candidates = [];
  const forced = typeof process.env.SUITE_SHELL_PATH === 'string' ? process.env.SUITE_SHELL_PATH.trim() : '';
  if (forced) candidates.push(forced);
  const sysRoot = process.env.SystemRoot || 'C:\\Windows';
  const programFiles = process.env.ProgramFiles || 'C:\\Program Files';
  const localApp = process.env.LOCALAPPDATA || '';
  candidates.push(
    join(sysRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'),
    join(programFiles, 'PowerShell', '7', 'pwsh.exe'),
    join(programFiles, 'PowerShell', '7-preview', 'pwsh.exe'),
  );
  if (localApp) candidates.push(join(localApp, 'Microsoft', 'WindowsApps', 'pwsh.exe'));
  return candidates;
}

/**
 * 解析本机可用的 PowerShell：优先 pwsh（7+），回退 powershell（5.1）；
 * 裸命令解析失败时按绝对路径候选逐一校验存在性。
 * 两个版本的脚本保持 5.1 兼容语法（不使用 7+ 独有运算符）。
 */
async function resolveShell(ctx, signal) {
  for (const name of ['pwsh', 'powershell']) {
    try {
      return await ctx.subprocess.resolveExecutable(name, {}, signal);
    } catch { /* 尝试下一个 */ }
  }
  for (const candidate of shellCandidates()) {
    try {
      if (candidate && existsSync(candidate)) return candidate;
    } catch { /* 继续 */ }
  }
  throw new Error(
    '本机未找到 PowerShell：裸命令（pwsh/powershell）解析失败，绝对路径候选也不存在。'
      + '可通过环境变量 SUITE_SHELL_PATH 指定 powershell.exe 的绝对路径。',
  );
}

/**
 * 运行任意可执行文件（P1 增量，供 ODAFileConverter 等外部引擎使用）。
 * 与 runPs1 相同的 stdout/stderr 上限、graceMs 与超时处理。
 * @returns {{ exitCode: number|null, signal: string|null, stdout: string, stderr: string }}
 */
export async function runExe(ctx, exe, args = [], options = {}) {
  const handle = ctx.subprocess.spawn({
    argv: [exe, ...args.map((arg) => String(arg))],
    cwd: options.cwd ?? process.cwd(),
    stdio: {
      stdin: 'ignore',
      stdout: { maxBytes: 8 * 1024 * 1024, spill: { maxBytes: 64 * 1024 * 1024 } },
      stderr: { maxBytes: 2 * 1024 * 1024 },
    },
    graceMs: options.graceMs ?? 60_000,
    signal: options.signal,
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
