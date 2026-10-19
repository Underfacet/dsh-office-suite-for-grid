/**
 * DSH 办公·电网绘图套件 —— Host 半（v0.1）
 *
 * v0.1 能力：引擎环境探测（suite_env）、Office 文档/表格读取（office_read，
 * 本机 Word/Excel COM）、Visio 流程图绘制（visio_draw，本机 Visio COM）。
 * 面板与远程命令共用同一操作层，为 P1 的 Client 面板预留同一套 ops。
 */
import { mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { runPs1 } from './host/engines.js';
import { parseJson, sessionCwd, resolveInputPath, dirnameOf } from './host/ops.js';

void fileURLToPath;

export const inject = ['tools', 'commands', 'subprocess', 'workingDirectory', 'systemPrompt'];

/** 工具结果统一为 JSON 文本呈现。 */
const jsonOutput = {
  schema: { type: 'object', properties: {}, additionalProperties: true },
  render: (_args, value) => [{ type: 'text', text: typeof value === 'string' ? value : JSON.stringify(value, null, 2) }],
};

export function apply(ctx, config = {}) {
  const disposers = [];

  // ── 工具：环境探测 / 诊断 ──
  disposers.push(ctx.tools.register({
    name: 'suite_env',
    description:
      '检测本机办公与绘图软件环境（Word/Excel/PowerPoint/Visio、AutoCAD/中望/浩辰、LibreOffice、Python、ODA 转换工具）及其版本，返回各引擎可用能力，供诊断与降级判断。',
    parameters: { type: 'object', properties: {}, additionalProperties: false },
    output: jsonOutput,
    async execute(_args, exec) {
      const r = await runPs1(ctx, 'probe.ps1', {}, { signal: exec.signal, timeoutMs: 60_000 });
      return parseJson(r.stdout) ?? { ok: false, error: { code: 'ENGINE_FAIL', message: r.stderr || r.stdout || '环境探测失败' } };
    },
  }));

  // ── 工具：读取 Word / Excel ──
  disposers.push(ctx.tools.register({
    name: 'office_read',
    description:
      '读取 Word(.doc/.docx) 或 Excel(.xls/.xlsx)：Word 返回段落数、表格数、标题大纲与正文预览；Excel 返回工作表清单、行列规模与数据样例。需要本机装有 Microsoft Office。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '文件路径（绝对路径，或相对当前会话工作目录）' },
      },
      required: ['path'],
    },
    output: jsonOutput,
    async execute(args, exec) {
      const abs = await resolveInputPath(ctx, exec, args.path);
      const r = await runPs1(ctx, 'office-read.ps1', { path: abs }, { signal: exec.signal, timeoutMs: 180_000 });
      return parseJson(r.stdout) ?? { ok: false, error: { code: 'ENGINE_FAIL', message: r.stderr || r.stdout || '读取失败' } };
    },
  }));

  // ── 工具：Visio 绘图 ──
  disposers.push(ctx.tools.register({
    name: 'visio_draw',
    description:
      '调用本机 Visio 绘制图表并保存为 .vsdx。v0.1 支持 flowchart（开始/结束为椭圆、判断为矩形、过程为矩形，箭头连线，按有向边自动分列布局）。需要本机装有 Visio。',
    parameters: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['flowchart'] },
        title: { type: 'string', description: '图表标题' },
        nodes: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              id: { type: 'string' },
              label: { type: 'string' },
              kind: { type: 'string', enum: ['start', 'process', 'decision', 'end'] },
            },
            required: ['id', 'label'],
          },
        },
        edges: {
          type: 'array',
          items: {
            type: 'object',
            properties: { from: { type: 'string' }, to: { type: 'string' } },
            required: ['from', 'to'],
          },
        },
        outputPath: { type: 'string', description: '输出 .vsdx 路径；缺省保存到会话工作目录 output/diagrams/ 下并自动命名' },
      },
      required: ['kind', 'nodes', 'edges'],
    },
    output: jsonOutput,
    async execute(args, exec) {
      const cwd = await sessionCwd(ctx, exec);
      const outputPath = args.outputPath
        ? await resolveInputPath(ctx, exec, args.outputPath)
        : join(cwd, 'output', 'diagrams', `flowchart-${Date.now()}.vsdx`);
      mkdirSync(dirnameOf(outputPath), { recursive: true });
      const payload = { kind: args.kind, title: args.title ?? '', nodes: args.nodes, edges: args.edges, outputPath };
      const r = await runPs1(ctx, 'visio-flowchart.ps1', payload, { signal: exec.signal, timeoutMs: 300_000, cwd });
      return parseJson(r.stdout) ?? { ok: false, error: { code: 'ENGINE_FAIL', message: r.stderr || r.stdout || '绘图失败' } };
    },
  }));

  // ── 命令（界面/输入框可用，与工具同一操作层） ──
  disposers.push(ctx.commands.register({
    name: 'office-env',
    description: '检测本机办公与绘图软件环境（Word/Excel/Visio/AutoCAD 等）',
    async handler() {
      const r = await runPs1(ctx, 'probe.ps1', {}, { timeoutMs: 60_000 });
      return { kind: 'success', text: r.stdout || '(no output)' };
    },
  }));

  // ── 系统提示：指引模型使用本套件工具 ──
  disposers.push(ctx.systemPrompt.section({
    name: 'tool:grid-office-suite',
    order: 900,
    text: () =>
      'Office 与绘图任务优先使用本套件工具：suite_env 检测本机软件环境与版本；office_read 读取 doc/docx/xls/xlsx（大纲/概况/样例）；visio_draw 调用本机 Visio 绘制流程图并保存为 .vsdx。',
  }));

  void config;
  return () => {
    for (const dispose of [...disposers].reverse()) {
      try { dispose(); } catch { /* 卸载时忽略个别清理失败 */ }
    }
  };
}
