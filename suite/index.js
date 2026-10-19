/**
 * DSH 办公·电网绘图套件 —— Host 半（v0.1 + P1/P2/P3 接线）
 *
 * 工具：suite_env（引擎环境探测）、office_read（Word/Excel 读取）、office_edit（结构化编辑）、
 * office_format（Word 排版）、office_export（导出 PDF，P3 §5.3）、visio_draw（Visio 绘图：
 * flowchart + P2 电网三图 + P3 五类图）。
 * 面板与远程命令共用同一操作层（host/ops/*），命令见 host/commands-panel.js。
 */
import { mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { runPs1 } from './host/engines.js';
import { effectiveConfig, parseJson, sessionCwd, resolveInputPath, dirnameOf } from './host/ops.js';
import { editOp } from './host/ops/office-edit.js';
import { formatOp } from './host/ops/office-format.js';
import { previewPdfOp } from './host/ops/office-pdf.js';
import {
  GRID_KINDS,
  P3_GRID_KINDS,
  expandGridLinks,
  loadSymbols,
  validateGridModel,
} from './host/ops/grid-model.js';
import { registerPanelCommands } from './host/commands-panel.js';

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

  // ── 工具：Office 结构化编辑（Word / Excel） ──
  disposers.push(ctx.tools.register({
    name: 'office_edit',
    description:
      '结构化编辑 Word(.doc/.docx) 或 Excel(.xls/.xlsx)：Word 段落/标题/样式、表格单元格与行列；Excel 单元格/公式、行列与工作表。按 ops 数组顺序执行，修改前自动备份。需要本机 Word/Excel（COM）或已就绪的 Python venv。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '文件路径（绝对路径，或相对当前会话工作目录）；不存在时可用 ops[0] 的 create 指令新建' },
        ops: {
          type: 'array',
          items: { type: 'object', additionalProperties: true },
          description:
            '编辑指令数组（非空，顺序执行）。每条为双字段 {action, op, ...}：action ∈ create|paragraph|heading|table|saveAs（Word）或 create|cell|row|column|sheet|saveAs（Excel）；op 为子操作（如 paragraph 的 setText/setStyle/insert/delete/append，table 的 setCell/addRow/deleteRow/addColumn/deleteColumn，cell 的 value/formula）。完整字段见 docs/P2-设计.md §3.2',
        },
        engine: {
          type: 'string',
          enum: ['auto', 'com', 'python'],
          description: "引擎选择；缺省用配置 officeEngine（默认 'auto'：COM 优先，不可用时降级 Python venv）",
        },
      },
      required: ['path', 'ops'],
    },
    output: jsonOutput,
    async execute(args, exec) {
      return editOp(ctx, effectiveConfig(config), args, exec);
    },
  }));

  // ── 工具：Office 排版（Word） ──
  disposers.push(ctx.tools.register({
    name: 'office_format',
    description:
      'Word 排版：按 rules 应用标题/正文字体字号、行距段距、编号、页边距、页眉页脚；可按 styleSet 套用样式集，或按 template 套用 docx 模板并填充占位符。修改前自动备份；批量排版请用面板命令 suite.format.batch。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Word(.doc/.docx) 文件路径（绝对路径，或相对当前会话工作目录）；套用模板时目标可以不存在' },
        rules: {
          type: 'array',
          items: { type: 'object', additionalProperties: true },
          description:
            '排版规则数组。每条 {target: heading|body|page|header|footer, level?, font?{name,sizePt,bold}, paragraph?{lineSpacing,lineSpacingPt,spaceBeforePt,spaceAfterPt,align}, numbering?, page?{marginsMm}, header?/footer?{text,font}}；见 docs/P2-设计.md §3.3',
        },
        styleSet: { type: 'string', description: '样式集 id（suite\\templates\\styles 或 runtime\\templates\\styles 下的 <id>.json；同名用户样式集优先）' },
        template: { type: 'string', description: '模板名或 docx 路径（suite\\templates\\docs / runtime\\templates\\docs）；套用模板=从模板新建并按 placeholders 填充' },
        placeholders: {
          type: 'object',
          additionalProperties: true,
          description: '模板占位符键值，如 {"标题":"关于 X 的通知","日期":"2026-10-20"}',
        },
        scope: { type: 'string', description: "作用范围，缺省 'all'（全部）；'all' 之外的标题/段落范围请通过命令 suite.format 传对象" },
        engine: {
          type: 'string',
          enum: ['auto', 'com', 'python'],
          description: "引擎选择；缺省用配置 officeEngine（默认 'auto'：COM 优先，不可用时降级 Python venv）",
        },
      },
      required: ['path'],
    },
    output: jsonOutput,
    async execute(args, exec) {
      return formatOp(ctx, effectiveConfig(config), args, exec);
    },
  }));

  // ── 工具：Visio 绘图（flowchart + 电网三图 + P3 五类） ──
  disposers.push(ctx.tools.register({
    name: 'visio_draw',
    description:
      "调用本机 Visio 绘制图表并保存为 .vsdx。kind='flowchart'：流程图（开始/结束椭圆、过程/判断矩形、箭头连线，按有向边自动分列）；kind='singleLine'（电气主接线图）/ 'distribution'（配电系统图）/ 'networkTopology'（通信网络拓扑图）需提供 model 数据模型（见 docs/P2-设计.md §4.1）；kind='protection'（二次回路/继保原理图）/ 'substationPlan'（变电站总平面布置示意图）/ 'geoNetwork'（地理接线图）/ 'powerFlow'（潮流图/系统接线图）/ 'terminalStrip'（端子排图，可多页）见 docs/P3-设计.md §4.4。Host 先校验符号与接线再绘制。需要本机装有 Visio。",
    parameters: {
      type: 'object',
      properties: {
        kind: {
          type: 'string',
          enum: [
            'flowchart',
            'singleLine',
            'distribution',
            'networkTopology',
            'protection',
            'substationPlan',
            'geoNetwork',
            'powerFlow',
            'terminalStrip',
          ],
        },
        title: { type: 'string', description: '图表标题' },
        model: {
          type: 'object',
          additionalProperties: true,
          description: "电网图数据模型 JSON（kind 为 P2/P3 八类电网图之一时必填）：P2 三类见 docs/P2-设计.md §4.1，P3 五类见 docs/P3-设计.md §4.4",
        },
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
      required: ['kind'],
    },
    output: jsonOutput,
    async execute(args, exec) {
      const cwd = await sessionCwd(ctx, exec);

      // P1 行为不变：流程图走 visio-flowchart.ps1
      if (args.kind === 'flowchart') {
        const outputPath = args.outputPath
          ? await resolveInputPath(ctx, exec, args.outputPath)
          : join(cwd, 'output', 'diagrams', `flowchart-${Date.now()}.vsdx`);
        mkdirSync(dirnameOf(outputPath), { recursive: true });
        const payload = { kind: args.kind, title: args.title ?? '', nodes: args.nodes, edges: args.edges, outputPath };
        const r = await runPs1(ctx, 'visio-flowchart.ps1', payload, { signal: exec.signal, timeoutMs: 300_000, cwd });
        return parseJson(r.stdout) ?? { ok: false, error: { code: 'ENGINE_FAIL', message: r.stderr || r.stdout || '绘图失败' } };
      }

      // P2 电网三图（visio-grid.ps1）/ P3 五类（visio-layouts.ps1）：
      // 符号库 → 模型/接线校验 → 展开连线 → 引擎（payload 同构）
      if (args.kind === 'cableRoute') {
        // P3 §4.4⑥：schema 定稿但本期未实现（E2024，不静默）
        return loadSymbols(args.kind);
      }
      if (!GRID_KINDS.includes(args.kind)) {
        return {
          ok: false,
          error: {
            code: 'E1041',
            message: `kind 必须为 flowchart | ${GRID_KINDS.join(' | ')}，收到: ${JSON.stringify(args.kind)}`,
            hint: 'flowchart 使用 nodes/edges；电网图使用 model（P2 §4.1 / P3 §4.4）',
          },
        };
      }
      if (args.model == null || typeof args.model !== 'object' || Array.isArray(args.model)) {
        return { ok: false, error: { code: 'E2021', message: '电网图需要 model 对象', hint: '模型结构见 docs/P2-设计.md §4.1 与 docs/P3-设计.md §4.4' } };
      }
      const engineScript = P3_GRID_KINDS.includes(args.kind) ? 'visio-layouts.ps1' : 'visio-grid.ps1';
      const loaded = loadSymbols(args.kind);
      if (!loaded.ok) return loaded;
      const validated = validateGridModel(args.kind, args.model, { symbols: loaded.result.symbols });
      if (!validated.ok) return validated;
      const expanded = expandGridLinks(validated.result.model);
      const outputPath = args.outputPath
        ? await resolveInputPath(ctx, exec, args.outputPath)
        : join(cwd, 'output', 'diagrams', `${args.kind}-${Date.now()}.vsdx`);
      mkdirSync(dirnameOf(outputPath), { recursive: true });
      const payload = {
        kind: args.kind,
        title: typeof args.title === 'string' && args.title ? args.title : (validated.result.model.title ?? ''),
        model: validated.result.model,
        symbols: loaded.result.symbols,
        outputPath,
        validations: validated.result.validations,
      };
      const r = await runPs1(ctx, engineScript, payload, { signal: exec.signal, timeoutMs: 300_000, cwd });
      const parsed = parseJson(r.stdout);
      if (!parsed) {
        return { ok: false, error: { code: 'ENGINE_FAIL', message: r.stderr || r.stdout || '绘图失败' } };
      }
      if (!parsed.ok) return parsed;
      const result = parsed.result ?? {};
      const stats = { ...(result.stats ?? {}) };
      if (stats.links == null) stats.links = expanded.length;
      return {
        ok: true,
        result: {
          ...result,
          kind: result.kind ?? args.kind,
          engine: result.engine ?? engineScript,
          outputPath: result.outputPath ?? outputPath,
          stats,
          validations: Array.isArray(result.validations) ? result.validations : validated.result.validations,
        },
      };
    },
  }));

  // ── 工具：导出 PDF（office_export，P3 §5.3） ──
  disposers.push(ctx.tools.register({
    name: 'office_export',
    description:
      '把 Word/Excel/PowerPoint（doc/docx/xls/xlsx/ppt/pptx）导出为 PDF 文件：output 为目录或 .pdf 路径，缺省写到会话工作区 output\\pdf\\<原名>.pdf（重名自动加 -<时间戳>）；渲染由宿主 officeToPdf 服务完成，缺失时报 E2012。导出是新文件，不触发备份。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '源文件路径（绝对路径，或相对当前会话工作目录）' },
        output: { type: 'string', description: '导出目标：目录或 .pdf 文件路径；缺省 = 会话工作区 output\\pdf\\<原名>.pdf' },
        engine: { type: 'string', enum: ['auto', 'officeToPdf'], description: "渲染引擎；固定走 Host officeToPdf（缺省 'auto'）" },
      },
      required: ['path'],
    },
    output: jsonOutput,
    async execute(args, exec) {
      const saveAs = typeof args.output === 'string' && args.output.trim() ? args.output.trim() : true;
      return previewPdfOp(
        ctx,
        effectiveConfig(config),
        { path: args.path, saveAs, ...(args.engine ? { engine: args.engine } : {}) },
        exec,
      );
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

  // ── P1 增量：面板远程命令（suite.env / suite.read / suite.draw / suite.files /
  //    suite.cad.convert / suite.cad.status / suite.cad.cancel /
  //    suite.config.get / suite.config.set / suite.preview） ──
  disposers.push(registerPanelCommands(ctx, { config }));

  // ── 系统提示：指引模型使用本套件工具 ──
  disposers.push(ctx.systemPrompt.section({
    name: 'tool:grid-office-suite',
    order: 900,
    text: () =>
      'Office 与绘图任务优先使用本套件工具：suite_env 检测本机软件环境与版本；office_read 读取 doc/docx/xls/xlsx（大纲/概况/样例）；office_edit 结构化编辑 Word/Excel（段落/标题/表格/单元格/公式，修改前自动备份）；office_format 排版 Word（rules/样式集/模板，批量用面板命令 suite.format.batch）；office_export 把 doc/docx/xls/xlsx/ppt/pptx 导出为 PDF 文件（output 传目录或 .pdf 路径，缺省写到 output\\pdf\\<原名>.pdf）；visio_draw 调用本机 Visio 绘图并保存为 .vsdx：flowchart 画流程图（nodes/edges），singleLine/distribution/networkTopology 画电气主接线图/配电系统图/通信网络拓扑图（需 model，见 docs/P2-设计.md §4.1），protection/substationPlan/geoNetwork/powerFlow/terminalStrip 画二次回路图/变电站总平面布置示意图/地理接线图/潮流图/端子排图（需 model，见 docs/P3-设计.md §4.4）。',
  }));

  void config;
  return () => {
    for (const dispose of [...disposers].reverse()) {
      try { dispose(); } catch { /* 卸载时忽略个别清理失败 */ }
    }
  };
}
