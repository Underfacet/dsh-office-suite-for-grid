/**
 * CAD 转换任务注册表（docs/P1-设计.md §3.6）
 *
 * 内存态（进程重启即失效）；每个任务：
 * - 状态机：queued → running → converting → splitting → merging → done | failed | cancelled
 * - 进度：progress { filesDone, filesTotal, framesDone, framesTotal, currentFile }
 * - 日志：环形缓冲 200 行/任务，元素 { t, level, msg }
 * - 取消：每个任务一个 AbortController；cancel() 后任务状态置 cancelled 并中止流水线
 * - 订阅：subscribe(fn) 供（可选的）Host→Client 事件推送使用；P1 面板以轮询 suite.cad.status 为主。
 */
import { OpsError } from './ops.js';

/** 任务状态（顺序即合法推进方向）。 */
export const TASK_STATES = Object.freeze([
  'queued',
  'running',
  'converting',
  'splitting',
  'merging',
  'done',
  'failed',
  'cancelled',
]);

/** 终态：不可再推进、不可取消。 */
export const TERMINAL_STATES = Object.freeze(['done', 'failed', 'cancelled']);

/** 每个任务保留的日志行数。 */
export const MAX_LOG_LINES = 200;

/** 注册表内最多保留的任务数（超出时清理最旧的终态任务）。 */
const MAX_TASKS = 64;

function nowIso() {
  return new Date().toISOString();
}

function isTerminal(state) {
  return TERMINAL_STATES.includes(state);
}

/** 进度字段默认值。 */
function initialProgress() {
  return {
    filesDone: 0,
    filesTotal: 0,
    framesDone: 0,
    framesTotal: 0,
    currentFile: null,
  };
}

/**
 * 任务注册表。命令层与 ops 层共用一个实例（进程内单例见 tasks）。
 */
export class TaskStore {
  constructor(options = {}) {
    this.tasks = new Map();
    this.listeners = new Set();
    this.seq = 0;
    this.maxLogLines = Number.isInteger(options.maxLogLines) ? options.maxLogLines : MAX_LOG_LINES;
  }

  /** 创建任务（初始态 queued），返回内部任务对象（含 abort controller）。 */
  create(meta = {}) {
    const id = `cad-${Date.now().toString(36)}-${(++this.seq).toString(36)}`;
    const controller = new AbortController();
    const timestamp = nowIso();
    const task = {
      id,
      taskId: id,
      kind: meta.kind ?? 'cad.convert',
      mode: meta.mode ?? null,
      engine: meta.engine ?? null,
      input: meta.input ?? null,
      output: meta.output ?? null,
      state: 'queued',
      createdAt: timestamp,
      updatedAt: timestamp,
      progress: initialProgress(),
      log: [],
      error: null,
      abort: controller,
      controller,
      settled: false,
    };
    this.tasks.set(id, task);
    this.prune();
    this.emit(task);
    return task;
  }

  get(taskId) {
    return this.tasks.get(String(taskId ?? '')) ?? null;
  }

  /** 任务摘要（列表用；不含完整日志）。 */
  list() {
    return [...this.tasks.values()]
      .sort((a, b) => (a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : 0))
      .map((task) => this.summary(task.id));
  }

  /** 任务摘要：状态/进度/输出/错误，日志只给条数与最后一条。 */
  summary(taskId) {
    const task = this.get(taskId);
    if (!task) return null;
    return {
      taskId: task.id,
      state: task.state,
      mode: task.mode,
      engine: task.engine,
      input: task.input,
      output: task.output,
      progress: { ...task.progress },
      error: task.error ? { ...task.error } : null,
      logCount: task.log.length,
      lastLog: task.log.length ? { ...task.log[task.log.length - 1] } : null,
      createdAt: task.createdAt,
      updatedAt: task.updatedAt,
    };
  }

  /** 任务完整快照（可直接作为 suite.cad.status 的 result）。 */
  snapshot(taskId) {
    const task = this.get(taskId);
    if (!task) return null;
    return {
      taskId: task.id,
      state: task.state,
      mode: task.mode,
      engine: task.engine,
      input: task.input,
      output: task.output,
      progress: { ...task.progress },
      log: task.log.map((entry) => ({ ...entry })),
      error: task.error ? { ...task.error } : null,
      createdAt: task.createdAt,
      updatedAt: task.updatedAt,
    };
  }

  /** 推进状态（终态不再变化）。 */
  setState(taskId, state) {
    const task = this.requireTask(taskId);
    if (!TASK_STATES.includes(state)) {
      throw new OpsError('E1099', `未知任务状态: ${state}`);
    }
    if (isTerminal(task.state)) return task;
    if (task.state !== state) {
      task.state = state;
      task.updatedAt = nowIso();
      this.emit(task);
    }
    return task;
  }

  /** 合并进度补丁（只接受 progress 字段），返回最新进度。 */
  tick(taskId, patch = {}) {
    const task = this.requireTask(taskId);
    if (isTerminal(task.state)) return { ...task.progress };
    const progress = task.progress;
    const numericKeys = ['filesDone', 'filesTotal', 'framesDone', 'framesTotal'];
    for (const key of numericKeys) {
      if (patch[key] !== undefined && Number.isFinite(Number(patch[key]))) {
        progress[key] = Math.max(0, Math.floor(Number(patch[key])));
      }
    }
    if (patch.currentFile !== undefined) progress.currentFile = patch.currentFile == null ? null : String(patch.currentFile);
    task.updatedAt = nowIso();
    this.emit(task);
    return { ...progress };
  }

  /** 追加日志（环形缓冲）。level: info | warn | error。 */
  appendLog(taskId, level, msg) {
    const task = this.requireTask(taskId);
    const normalizedLevel = ['info', 'warn', 'error', 'debug'].includes(level) ? level : 'info';
    task.log.push({ t: nowIso(), level: normalizedLevel, msg: String(msg ?? '') });
    if (task.log.length > this.maxLogLines) task.log.splice(0, task.log.length - this.maxLogLines);
    task.updatedAt = nowIso();
    this.emit(task);
    return task;
  }

  /** appendLog 的简写：log.run(taskId, 'info', '...')。 */
  get runner() {
    return {
      info: (taskId, msg) => this.appendLog(taskId, 'info', msg),
      warn: (taskId, msg) => this.appendLog(taskId, 'warn', msg),
      error: (taskId, msg) => this.appendLog(taskId, 'error', msg),
    };
  }

  /** 任务成功结束。 */
  finish(taskId, output) {
    const task = this.requireTask(taskId);
    if (isTerminal(task.state)) return task;
    if (output !== undefined && output !== null) task.output = output;
    task.state = 'done';
    task.updatedAt = nowIso();
    task.settled = true;
    this.emit(task);
    return task;
  }

  /** 任务失败结束。error 为 { code, message, hint? }。 */
  fail(taskId, error) {
    const task = this.requireTask(taskId);
    if (isTerminal(task.state)) return task;
    task.state = 'failed';
    task.error = error && typeof error === 'object'
      ? { code: String(error.code ?? 'E1099'), message: String(error.message ?? '未知错误'), ...(error.hint ? { hint: String(error.hint) } : {}) }
      : { code: 'E1099', message: String(error ?? '未知错误') };
    task.updatedAt = nowIso();
    task.settled = true;
    this.emit(task);
    return task;
  }

  /**
   * 取消任务：终态任务抛 E1032；否则 abort + 状态置 cancelled。
   * @returns {{taskId:string,state:'cancelled'}}
   */
  cancel(taskId) {
    const task = this.get(taskId);
    if (!task) throw new OpsError('TASK_NOT_FOUND', `任务不存在或已过期: ${String(taskId)}`);
    if (isTerminal(task.state)) {
      throw new OpsError('TASK_NOT_CANCELLABLE', `任务已结束（${task.state}），无法取消`, '仅排队/运行中的任务可取消');
    }
    task.state = 'cancelled';
    task.updatedAt = nowIso();
    task.settled = true;
    try {
      task.controller.abort(new Error('用户取消'));
    } catch { /* 已中止 */ }
    this.appendLog(task.id, 'warn', '任务已由用户取消（已产出的中间 PDF 保留在临时目录）');
    this.emit(task);
    return { taskId: task.id, state: 'cancelled' };
  }

  /** 订阅任务变化（可选事件推送用）；返回取消订阅函数。 */
  subscribe(listener) {
    if (typeof listener !== 'function') return () => {};
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  emit(task) {
    if (!this.listeners.size) return;
    const snapshot = this.snapshot(task.id);
    for (const listener of [...this.listeners]) {
      try {
        listener(snapshot);
      } catch { /* 订阅者异常不影响任务 */ }
    }
  }

  /** 任务对象不存在时抛 E1031。 */
  requireTask(taskId) {
    const task = this.get(taskId);
    if (!task) throw new OpsError('TASK_NOT_FOUND', `任务不存在或已过期: ${String(taskId)}`);
    return task;
  }

  /** 清理：超过上限时移除最旧的终态任务。 */
  prune() {
    if (this.tasks.size <= MAX_TASKS) return;
    const settled = [...this.tasks.values()]
      .filter((task) => isTerminal(task.state))
      .sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1));
    while (this.tasks.size > MAX_TASKS && settled.length) {
      const oldest = settled.shift();
      this.tasks.delete(oldest.id);
    }
  }
}

/** 进程内共享注册表。 */
export const tasks = new TaskStore();

export const createTaskStore = (options) => new TaskStore(options);
