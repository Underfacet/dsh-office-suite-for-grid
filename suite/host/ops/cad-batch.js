/**
 * CAD 批量队列（docs/P3-设计.md §3）
 *
 * 入口：batchCadOp(ctx, config, args, exec, store?)
 *   args: { inputs: [文件/目录…], mode:'merge'|'split', paper?, sort?, parallel?, output? }
 *   - inputs 每项 = 一个独立子任务（复用 cad.js 的 convertCad，不重复实现转换流水线）；
 *   - merge 条目必须是目录、split 条目必须是单文件（.dwg/.pdf）；
 *   - 目录输出时子任务输出为 `<mode>-<时间戳>-<两位序号>.pdf`（避免同秒同名覆盖）；
 *   - 与既有 `input` 单任务语义互斥（同时给 → E1041）。
 *
 * 聚合（tasks.js 零侵入 + §3.3 最小扩展）：
 *   - 父任务 kind='cad.queue'（子任务 parentId=父任务 id，经 TaskStore.setParent 回填）；
 *   - 父任务 progress.itemsDone/itemsTotal 聚合子任务终态数；
 *   - 子任务沿用既有状态机与 taskId（面板/命令可分别轮询）；父任务只记摘要日志；
 *   - 全部成功/部分失败 → 父任务 done（failures 记录失败项，语义同 suite.format.batch）；
 *     全部失败/取消 → 父任务 failed（E1033 QUEUE_ALL_FAILED）；
 *   - 取消父任务 → 递归 cancel 未终态子任务；取消子任务只影响自身（父任务继续）。
 *
 * 返回（§3.1，立即返回）：{ ok:true, result:{ taskId, kind:'cad.queue', mode, engine, items, itemsTotal } }
 * 错误码：E1041（参数/条目类型）、E1021/E1023（路径）、E1001（无引擎）、E1033（全部失败）。
 */
import { readdir, stat } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import {
  OpsError,
  fail,
  isAbortError,
  ok,
  resolveInputPath,
  sessionCwd,
  toErrorPayload,
} from '../ops.js';
import { analyzeEngines, chooseEngine, convertCad, engineChain, probeEngines } from './cad.js';
import { compareNatural } from './drawingNo.js';
import { TERMINAL_STATES, tasks as defaultTasks } from '../tasks.js';

/** 队列父任务 kind。 */
export const QUEUE_KIND = 'cad.queue';
/** 队列条目上限（config.queueMaxItems 可调，硬上限 100）。 */
export const QUEUE_MAX_ITEMS = 100;

const EXCLUDED_FILE = /^~\$|\.(bak|dwl|dwl2|tmp)$/i;

function timestamp() {
  return new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
}

function isTerminal(state) {
  return TERMINAL_STATES.includes(state);
}

function validateQueueMode(value) {
  if (value !== 'merge' && value !== 'split') {
    throw new OpsError('E1041', "mode 必须为 'merge' | 'split'", `收到: ${JSON.stringify(value)}`);
  }
  return value;
}

function validateQueuePaper(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'A4-landscape' || value === 'keep') return value;
  throw new OpsError('E1041', "paper 必须为 'A4-landscape' | 'keep'", `收到: ${JSON.stringify(value)}`);
}

function validateQueueSort(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'drawingNo' || value === 'filename') return value;
  throw new OpsError('E1041', "sort 必须为 'drawingNo' | 'filename'", `收到: ${JSON.stringify(value)}`);
}

function validateQueueParallel(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'auto') return 'auto';
  if (Number.isInteger(value) && value >= 1 && value <= 8) return value;
  throw new OpsError('E1041', "parallel 必须为 'auto' 或 1–8 的整数", `收到: ${JSON.stringify(value)}`);
}

function queueMaxItems(config) {
  const configured = Number(config?.queueMaxItems);
  if (Number.isFinite(configured) && configured >= 1) return Math.min(Math.floor(configured), QUEUE_MAX_ITEMS);
  return QUEUE_MAX_ITEMS;
}

async function statOrNull(path) {
  try {
    return await stat(path);
  } catch {
    return null;
  }
}

/** 默认输出根：config.outputRoot 或 会话工作区 output\pdf（语义同 cad.convert）。 */
async function defaultQueueOutputRoot(ctx, exec, config) {
  if (typeof config?.outputRoot === 'string' && config.outputRoot.trim()) {
    const configured = config.outputRoot.trim();
    return isAbsolute(configured) ? resolve(configured) : resolve(await sessionCwd(ctx, exec), configured);
  }
  return join(await sessionCwd(ctx, exec), 'output', 'pdf');
}

/**
 * 解析队列条目输出路径（目录输出时 `<mode>-<时间戳>-<两位序号>.pdf`）。
 * @returns {Promise<string[]>} 与 inputs 等长的绝对路径
 */
async function resolveQueueOutputs(ctx, exec, config, mode, inputs, outputArg) {
  const stamp = timestamp();
  const nameFor = (index) => `${mode}-${stamp}-${String(index + 1).padStart(2, '0')}.pdf`;
  const raw = typeof outputArg === 'string' ? outputArg.trim() : '';
  if (!raw) {
    const root = await defaultQueueOutputRoot(ctx, exec, config);
    return inputs.map((input, index) => join(root, nameFor(index)));
  }
  const abs = assertInsideRoot(await resolveInputPath(ctx, exec, raw), {
    write: true,
    root: resolveBoundaryRoot(config),
  });
  const info = await statOrNull(abs);
  const isDirectory = info ? info.isDirectory() : !/\.pdf$/i.test(abs);
  if (!isDirectory) {
    if (inputs.length > 1) {
      throw new OpsError('E1041', '多条目时 output 必须是目录', `收到单个 PDF 文件: ${abs}（单文件请直接用 suite.cad.convert）`);
    }
    return [abs];
  }
  return inputs.map((input, index) => join(abs, nameFor(index)));
}

/** 校验单个队列条目：merge=目录，split=单文件（.dwg/.pdf）。 */
async function validateQueueItem(inputAbs, mode) {
  const info = await statOrNull(inputAbs);
  if (!info) throw new OpsError('E1023', `输入路径不存在: ${inputAbs}`);
  if (mode === 'merge') {
    if (!info.isDirectory()) throw new OpsError('E1041', `合并队列的条目必须是目录: ${inputAbs}`, '单文件请用 mode=split');
    const names = await readdir(inputAbs).catch(() => []);
    const usable = names.filter((name) => !EXCLUDED_FILE.test(name) && /\.(dwg|pdf)$/i.test(name));
    if (!usable.length) {
      throw new OpsError('E1023', `目录内没有可转换的 .dwg / .pdf 文件: ${inputAbs}`);
    }
    return;
  }
  if (!info.isFile()) throw new OpsError('E1041', `分割队列的条目必须是单个 .dwg/.pdf 文件: ${inputAbs}`, '文件夹请用 mode=merge');
  const ext = extname(inputAbs).toLowerCase();
  if (ext !== '.dwg' && ext !== '.pdf') {
    throw new OpsError('E1041', '分割队列仅支持 .dwg / .pdf 输入', `收到: ${ext || '(无扩展名)'}（${basename(inputAbs)}）`);
  }
}

/**
 * 子任务聚合器：订阅 TaskStore（不新增轮询），推进父任务 itemsDone/itemsTotal 与终态。
 * items 在提交阶段会逐步补上 taskId，因此按 taskId 动态查找（最多 100 项，开销可忽略）。
 * @returns {{start:() => void, sync:() => void}}
 */
function createQueueAggregator({ parent, items, store }) {
  let stopped = false;
  let unsubChild = null;
  let unsubParent = null;

  const log = (level, message) => {
    try { store.appendLog(parent.id, level, message); } catch { /* 父任务可能已被清理 */ }
  };
  const terminalCount = () => items.filter((item) => isTerminal(item.state)).length;
  const stop = () => {
    if (unsubChild) { try { unsubChild(); } catch { /* ignore */ } unsubChild = null; }
    if (unsubParent) { try { unsubParent(); } catch { /* ignore */ } unsubParent = null; }
  };
  const cancelChildren = () => {
    for (const item of items) {
      if (!item.taskId || isTerminal(item.state)) continue;
      try { store.cancel(item.taskId); } catch { /* 已终态 */ }
    }
  };

  const finalize = () => {
    if (stopped) return;
    if (!items.every((item) => isTerminal(item.state))) return;
    stopped = true;
    stop();
    const succeeded = items.filter((item) => item.state === 'done');
    const failures = items
      .filter((item) => item.state !== 'done')
      .map((item) => ({
        index: item.index,
        input: item.input,
        code: item.code ?? (item.state === 'cancelled' ? 'E1098' : null),
        message: item.message ?? (item.state === 'cancelled' ? '任务已取消' : null),
      }));
    if (!succeeded.length) {
      const first = failures[0];
      try {
        store.fail(parent.id, {
          code: 'E1033',
          message: `批量队列全部 ${items.length} 个条目失败`,
          hint: first ? [first.code, first.message].filter(Boolean).join(' ') : undefined,
        });
      } catch { /* 父任务可能已被清理 */ }
      log('error', `批量队列全部 ${items.length} 个条目失败（E1033）`);
      return;
    }
    try {
      store.finish(parent.id, {
        itemsTotal: items.length,
        itemsDone: items.length,
        succeeded: succeeded.length,
        failed: failures.length,
        failures: failures.slice(0, 100),
        outputs: succeeded.slice(0, 100).map((item) => ({
          index: item.index,
          input: item.input,
          output: store.get(item.taskId)?.output ?? item.output ?? null,
        })),
      });
    } catch { /* 父任务可能已被清理 */ }
    log('info', `批量队列完成：成功 ${succeeded.length}，失败 ${failures.length}（共 ${items.length}）`);
  };

  const onSnapshot = (snapshot) => {
    if (stopped || !snapshot) return;
    if (snapshot.taskId === parent.id) {
      if (snapshot.state === 'cancelled') cancelChildren();
      return;
    }
    const item = items.find((candidate) => candidate.taskId === snapshot.taskId);
    if (!item) return;
    const previous = item.state;
    item.state = snapshot.state;
    item.error = snapshot.error ?? item.error ?? null;
    if (isTerminal(snapshot.state)) {
      if (previous === snapshot.state) return; // finish 后的 appendLog 会再次 emit，避免重复记账
      item.code = snapshot.error?.code ?? null;
      item.message = snapshot.error?.message ?? null;
      if (snapshot.state === 'done' && snapshot.output) item.output = snapshot.output;
      try { store.tick(parent.id, { itemsDone: terminalCount(), currentFile: basename(item.input) }); } catch { /* 父任务可能已被清理 */ }
      if (snapshot.state === 'done') {
        log('info', `[${item.index}/${items.length}] ${basename(item.input)} ok → ${snapshot.output ?? item.output ?? ''}`);
      } else {
        log('error', `[${item.index}/${items.length}] ${basename(item.input)} ${snapshot.state}${item.code ? ` ${item.code}` : ''}${item.message ? `: ${item.message}` : ''}`);
      }
      finalize();
    }
  };

  /** 主动同步一遍子任务快照（订阅前/提交后调用；不依赖轮询）。 */
  const sync = () => {
    if (stopped) return;
    for (const item of items) {
      if (!item.taskId || isTerminal(item.state)) continue;
      const snapshot = store.snapshot(item.taskId);
      if (snapshot) onSnapshot(snapshot);
      else {
        // 任务记录可能被 MAX_TASKS 清理；若已由事件记为终态则保留（上面已跳过），否则按失败处理
        item.state = 'failed';
        item.code = 'TASK_NOT_FOUND';
        item.message = '子任务记录已过期';
      }
    }
    try { store.tick(parent.id, { itemsTotal: items.length, itemsDone: terminalCount() }); } catch { /* 父任务可能已被清理 */ }
    finalize();
  };

  return {
    start() {
      unsubChild = store.subscribe(onSnapshot);
      unsubParent = store.subscribe((snapshot) => {
        if (snapshot?.taskId === parent.id && snapshot.state === 'cancelled') cancelChildren();
      });
      // 父任务可能在提交期间被取消：启动时补一次取消传递
      const parentSnapshot = store.snapshot(parent.id);
      if (parentSnapshot?.state === 'cancelled') cancelChildren();
      sync();
    },
    sync,
  };
}

/**
 * suite.cad.convert 批量队列：inputs 每项一个独立子任务，父任务聚合进度与终态。
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 有效配置
 * @param {{inputs:string[],mode:'merge'|'split',paper?:string,sort?:string,parallel?:string|number,output?:string}} args
 * @param {object} exec 工具/命令执行上下文（agent/signal）
 * @param {import('../tasks.js').TaskStore} [store] 任务注册表（默认进程内共享实例）
 * @returns {Promise<{ok:true,result:object}|{ok:false,error:object}>}
 */
export async function batchCadOp(ctx, config = {}, args = {}, exec = {}, store = defaultTasks) {
  const warnings = [];
  try {
    const mode = validateQueueMode(args?.mode);
    const paper = validateQueuePaper(args?.paper);
    const sort = validateQueueSort(args?.sort);
    const parallel = validateQueueParallel(args?.parallel);
    if (typeof args?.input === 'string' && args.input.trim()) {
      throw new OpsError(
        'E1041',
        'input 与 inputs 互斥（单任务用 suite.cad.convert，批量用 inputs）',
        `收到 input=${JSON.stringify(args.input)} 与 inputs 数组`,
      );
    }
    if (!Array.isArray(args?.inputs) || !args.inputs.length) {
      throw new OpsError('E1041', '缺少 inputs（文件/目录路径数组，1–100 项）');
    }
    const limit = queueMaxItems(config);
    if (args.inputs.length > limit) {
      throw new OpsError('E1041', `队列条目数超过上限（${limit}）`, `收到 ${args.inputs.length} 项；请分批提交`);
    }

    const root = resolveBoundaryRoot(config);
    const allowExternalPaths = config.allowExternalPaths === true;
    const resolved = [];
    const seen = new Set();
    for (const raw of args.inputs) {
      if (typeof raw !== 'string' || !raw.trim()) throw new OpsError('E1041', 'inputs 中存在空路径');
      const abs = assertInsideRoot(await resolveInputPath(ctx, exec, raw.trim()), {
        write: true,
        allowExternalPaths,
        root,
        warn: (message) => console.warn(`[office-suite] ${message}`),
      });
      const key = abs.toLowerCase();
      if (seen.has(key)) {
        warnings.push(`忽略重复条目：${abs}`);
        continue;
      }
      seen.add(key);
      resolved.push(abs);
    }
    if (!resolved.length) throw new OpsError('E1041', 'inputs 去重后为空');

    for (const inputAbs of resolved) {
      await validateQueueItem(inputAbs, mode);
    }

    const outputs = await resolveQueueOutputs(ctx, exec, config, mode, resolved, args?.output);

    // 引擎预判（避免把全部必然失败的条目塞进队列）
    const isPdfInputs = resolved.map((file) => extname(file).toLowerCase() === '.pdf');
    const needsDwg = isPdfInputs.some((isPdf) => !isPdf);
    let engines = null;
    if (needsDwg) {
      const probe = await probeEngines(ctx, config);
      if (!probe.ok) {
        throw new OpsError(probe.error.code ?? 'E1051', probe.error.message ?? '环境探测失败', probe.error.hint);
      }
      engines = analyzeEngines(probe.result);
      const primary = chooseEngine({ isPdf: false, engines });
      if (primary === 'none') {
        throw new OpsError(
          'E1001',
          '本机未检测到 AutoCAD/中望/浩辰 COM，也未检测到 ODAFileConverter，且队列含 .dwg 输入',
          '可安装 CAD 或 ODAFileConverter，或先把图纸导出为 PDF 后只提交 .pdf',
        );
      }
      if (primary !== engineChain(engines, false)[0]) {
        warnings.push(`引擎降级：${engineChain(engines, false).join(' → ')}（选定 ${primary}）`);
      }
    }
    const enginePerItem = resolved.map((file, index) => (isPdfInputs[index] ? 'pdf' : chooseEngine({ isPdf: false, engines })));
    const uniqueEngines = [...new Set(enginePerItem)];
    const engineSummary = uniqueEngines.length === 1 ? uniqueEngines[0] : 'mixed';

    const parent = store.create({
      kind: QUEUE_KIND,
      mode,
      engine: engineSummary,
      input: `${resolved.length} 个条目`,
      output: outputs.length === 1 ? outputs[0] : dirname(outputs[0]),
    });
    store.appendLog(parent.id, 'info', `批量队列创建：${resolved.length} 个条目（mode=${mode}，engine=${engineSummary}）`);
    store.tick(parent.id, { itemsTotal: resolved.length, itemsDone: 0 });
    store.setState(parent.id, 'running');

    const items = resolved.map((inputAbs, index) => ({
      index: index + 1,
      input: inputAbs,
      output: outputs[index],
      taskId: null,
      state: 'queued',
      code: null,
      message: null,
      error: null,
    }));

    const submissionFailures = [];
    // 先订阅再提交：子任务可能在提交后续条目期间就进入终态（100 条时 MAX_TASKS 会清理旧终态任务）
    const aggregator = createQueueAggregator({ parent, items, store });
    aggregator.start();
    for (const item of items) {
      try {
        const submitted = await convertCad(ctx, {
          mode,
          input: item.input,
          output: item.output,
          paper,
          sort,
          parallel,
        }, exec, config, store);
        if (submitted.ok) {
          item.taskId = submitted.result.taskId;
          item.state = 'queued';
          item.output = submitted.result.output ?? item.output;
          try { store.setParent(item.taskId, parent.id); } catch { /* 子任务可能已被清理 */ }
        } else {
          item.state = 'failed';
          item.code = submitted.error?.code ?? null;
          item.message = submitted.error?.message ?? null;
          submissionFailures.push({ index: item.index, input: item.input, code: item.code, message: item.message });
          store.appendLog(parent.id, 'error', `[${item.index}/${items.length}] ${basename(item.input)} 提交失败 ${item.code}: ${item.message}`);
        }
      } catch (err) {
        if (isAbortError(err)) throw err;
        const payload = toErrorPayload(err);
        item.state = 'failed';
        item.code = payload.code;
        item.message = payload.message;
        submissionFailures.push({ index: item.index, input: item.input, code: payload.code, message: payload.message });
        store.appendLog(parent.id, 'error', `[${item.index}/${items.length}] ${basename(item.input)} 提交失败 ${payload.code}: ${payload.message}`);
      }
    }
    aggregator.sync();

    const submittedCount = items.filter((item) => item.taskId).length;
    if (!submittedCount) {
      // 全部条目在提交阶段失败：父任务直接进入 E1033，调用方拿到明确错误
      const first = submissionFailures[0];
      try {
        store.fail(parent.id, {
          code: 'E1033',
          message: `批量队列全部 ${items.length} 个条目提交失败`,
          hint: first ? [first.code, first.message].filter(Boolean).join(' ') : undefined,
        });
      } catch { /* 父任务可能已被清理 */ }
      return fail(
        'E1033',
        `批量队列全部 ${items.length} 个条目提交失败`,
        first ? [first.code, first.message].filter(Boolean).join(' ') : undefined,
      );
    }

    return ok({
      taskId: parent.id,
      kind: QUEUE_KIND,
      mode,
      engine: engineSummary,
      items: items.map((item) => ({
        index: item.index,
        input: item.input,
        output: item.output,
        taskId: item.taskId,
        state: item.state,
        ...(item.code ? { code: item.code, message: item.message } : {}),
      })),
      itemsTotal: items.length,
      itemsSubmitted: submittedCount,
      state: store.get(parent.id)?.state ?? 'running',
      ...(warnings.length ? { warnings } : {}),
    });
  } catch (err) {
    if (isAbortError(err)) throw err;
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}

/** 队列条目按文件名自然序（面板清单编辑器排序辅助）。 */
export function sortQueueInputs(inputs) {
  return (Array.isArray(inputs) ? inputs : []).slice().sort((a, b) => compareNatural(a, b));
}
