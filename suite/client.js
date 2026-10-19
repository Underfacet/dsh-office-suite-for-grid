/**
 * DSH 办公·电网绘图套件 —— Client 半（P1 面板）
 *
 * ModuleLoader 单文件产物：工厂内部按 `//#region` 分区（顺序即依赖顺序）：
 *   1 i18n 字典 → 2 全局 CSS → 3 工具函数 → 4 remote 封装 → 5 PanelGlyph
 *   → 6 OfficeSuitePanel 外壳 → 7 DocsTab → 8 VisioTab → 9 CadTab → 10 DiagTab
 *   → 11 inject + apply
 *
 * 与设计文档 docs/P1-设计.md 的差异（均为有意修正，逐条注明）：
 *  - §1.2 草图的 inject 把 `remote.commands` / `remote.workspaceFiles` /
 *    `remote.directoryPicker` 列为硬依赖。本实现改为**仅**硬依赖
 *    ['slots','layout','locale','theme','sessions','uiWorkspace','remote']，
 *    三个命名空间一律 `ctx.get(...)` 判空防御获取（见 region 4），缺挂时对应功能
 *    禁用并提示 `common.remoteMissing`，避免命名空间未挂载导致整个插件 apply 崩溃。
 *  - §1.3/§0.5 说「浏览其他目录」模式用 uiWorkspace.listDirectory 即可。核实
 *    dsh-api-workspace-controller 的 directoryPicker/list 契约后确认：它只返回
 *    **子目录**（"Direct child directories"），不含文件。因此外部目录模式以
 *    uiWorkspace.listDirectory 导航目录，文件行改由 P1 远程命令 `suite.files`
 *    补齐（命令缺失时降级为「仅目录」并提示）。
 *  - §1.3 CadTab 的「选择文件（场景 A）/选择文件夹（场景 B）」与 §3.2/3.3 的场景
 *    定义相反；以 §3 为准：merge = 文件夹合并，split = 单文件分割。面板内没有原生
 *    文件选择器，单文件用面板内「浏览…」拾取（目录来自 listDirectory / 文件来自
 *    suite.files），同时保留可编辑路径输入框作为兜底。
 *  - 页签记忆改为「切换即写不 localStorage」，效果等同卸载保存且更简单。
 *
 * 约束：只 require('react')；UI 全自绘；颜色只用 §1.5 令牌（SVG 图标配色除外）；
 * 深色模式靠 body[data-ds-dark-theme] 自动切换，不写基于系统配色偏好的媒体查询。
 */
window.__ModuleLoader__.load({
  id: 'dsh-office-suite-for-grid',
  factory(require) {
    const React = require('react');
    const h = React.createElement;

    //#region 1. i18n 字典（zh 为键集来源，en 覆盖同一键集）
    /** 面板字典命名空间（ctx.locale.register 的 ns，同时作为 slot 条目的 locale）。 */
    const NS = 'officeSuite';
    /** 侧栏条目 id 与 main 插槽的 key：两者必须一致，否则 layout.selectPanel 抛错。 */
    const PANEL_ID = 'office-suite';

    const zh = {
      'panel.title': '办公·电网套件',
      'panel.tabs': '功能页签',
      'panel.session.current': '当前会话',
      'tab.docs': '文档中心',
      'tab.visio': 'Visio 绘图',
      'tab.cad': '图纸转换',
      'tab.diag': '设置与诊断',

      'docs.title': '文档中心',
      'docs.tree.root': '工作区文件',
      'docs.tree.browse': '浏览其他目录…',
      'docs.tree.back': '返回工作区',
      'docs.tree.external': '外部目录',
      'docs.tree.filterAll': '显示全部文件',
      'docs.tree.filterHint': '默认仅显示 doc/docx/xls/xlsx/vsdx/dwg/pdf',
      'docs.tree.empty': '当前目录没有匹配文件',
      'docs.tree.loading': '加载中…',
      'docs.tree.up': '上级目录',
      'docs.tree.truncated': '目录条目过多，仅显示前一部分',
      'docs.tree.dirsOnly': '该目录仅能浏览子目录（文件行来自远程命令 suite.files）',
      'docs.tree.noSession': '需要活动会话才能浏览工作区文件',
      'docs.read': '读取',
      'docs.reading': '读取中…',
      'docs.read.prompt': '点击“读取”加载该文件内容',
      'docs.read.notSupported': '该类型暂不支持预览',
      'docs.read.empty': '点击左侧文件后点“读取”',
      'docs.preview.hint.pdf': 'PDF 预览将在后续版本提供',
      'docs.preview.hint.other': '该类型预览将在后续版本提供',
      'docs.meta.size': '大小',
      'docs.meta.version': '版本',
      'docs.meta.path': '路径',
      'docs.meta.name': '文件名',
      'docs.actions.refresh': '刷新',
      'docs.actions.copyPath': '复制路径',
      'docs.actions.copied': '已复制',
      'docs.word.paragraphs': '段落',
      'docs.word.tables': '表格',
      'docs.word.headings': '标题大纲',
      'docs.word.preview': '正文预览',
      'docs.excel.sheet': '工作表',
      'docs.excel.sheets': '工作表数',
      'docs.excel.rows': '行',
      'docs.excel.cols': '列',
      'docs.excel.sample': '数据样例',
      'docs.file.none': '未选择文件',

      'visio.title': 'Visio 绘图',
      'visio.kind': '图类型',
      'visio.kind.flowchart': '流程图（P1）',
      'visio.titleField': '图表标题',
      'visio.nodes': '节点 JSON（id/label/kind: start|process|decision|end）',
      'visio.edges': '连线 JSON（from/to）',
      'visio.template': '插入模板',
      'visio.validate.ok': 'JSON 校验通过',
      'visio.json.invalid': 'JSON 无效：{message}',
      'visio.node.invalid': '节点格式无效（需要 id/label）：第 {index} 项',
      'visio.node.kindInvalid': '节点 kind 只能是 start|process|decision|end：第 {index} 项',
      'visio.node.duplicate': '节点 id 重复：{id}',
      'visio.edge.invalid': '连线格式无效（需要 from/to）：第 {index} 项',
      'visio.edge.missing': '连线引用了不存在的节点：{id}',
      'visio.nl': '自然语言描述（P2 接入）',
      'visio.outputPath': '输出路径（留空 = 自动命名）',
      'visio.outputHint': '留空时保存到会话工作区 output\\diagrams\\ 并自动命名',
      'visio.chooseDir': '选择目录…',
      'visio.run': '绘图',
      'visio.running': '绘图运行中…',
      'visio.done.saved': '已保存：{path}',
      'visio.done.stats': '节点 {nodes} · 连线 {edges}',
      'visio.result': '运行结果',
      'visio.history': '本次运行历史',
      'visio.history.empty': '暂无运行记录',

      'cad.title': '图纸转换',
      'cad.mode': '转换场景',
      'cad.mode.merge': '合并（文件夹多图 → 一个 PDF）',
      'cad.mode.split': '分割（单文件多图框 → 每图一页）',
      'cad.input': '输入',
      'cad.input.file': '选择 DWG 文件…',
      'cad.input.dir': '选择文件夹…',
      'cad.input.placeholder': '未选择',
      'cad.input.path': '输入路径',
      'cad.input.browse': '浏览…',
      'cad.picker.title': '选择图纸文件',
      'cad.picker.hint': '仅显示 .dwg/.dxf/.pdf 文件',
      'cad.picker.empty': '该目录没有可用的图纸文件',
      'cad.picker.up': '上级目录',
      'cad.picker.close': '关闭',
      'cad.files.count': '共 {count} 个文件',
      'cad.files.counting': '统计中…',
      'cad.paper': 'PDF 纸型',
      'cad.paper.a4landscape': 'A4 横向（默认）',
      'cad.paper.keep': '保持原图幅',
      'cad.sort': '排序',
      'cad.sort.drawingNo': '图号优先（回退文件名）',
      'cad.sort.filename': '文件名自然排序',
      'cad.parallel': '并行度',
      'cad.parallel.auto': '自动',
      'cad.start': '开始转换',
      'cad.cancel': '取消',
      'cad.cancel.requested': '已请求取消…',
      'cad.state.idle': '未开始',
      'cad.state.queued': '排队',
      'cad.state.running': '转换中',
      'cad.state.converting': '转换中',
      'cad.state.splitting': '分割中',
      'cad.state.merging': '合并中',
      'cad.state.done': '完成',
      'cad.state.failed': '失败',
      'cad.state.cancelled': '已取消',
      'cad.progress.files': '文件 {done}/{total}',
      'cad.progress.frames': '图框 {done}/{total}',
      'cad.progress.current': '当前：{file}',
      'cad.log': '运行日志',
      'cad.log.empty': '暂无日志',
      'cad.result.title': '转换结果',
      'cad.result.output': '输出：{path}',
      'cad.result.openInDocs': '在文档中心查看',

      'diag.title': '设置与诊断',
      'diag.overview': '引擎概览',
      'diag.overview.available': '可用引擎：{list}',
      'diag.overview.none': '未检测到可用的办公/绘图引擎',
      'diag.overview.fallback': '未检测到 CAD：图纸转换走 ODA / PDF 级链路',
      'diag.engine.available': '可用',
      'diag.engine.missing': '缺失',
      'diag.engine.degraded': '降级',
      'diag.table.engine': '引擎',
      'diag.table.status': '状态',
      'diag.table.version': '版本 / ProgID',
      'diag.table.capability': '能力说明',
      'diag.probe': '重新探测',
      'diag.probing': '探测中…',
      'diag.copy': '复制诊断文本',
      'diag.copy.done': '诊断文本已复制',
      'diag.config': '运行配置',
      'diag.config.note': '修改将在下次转换时生效',
      'diag.config.empty': '尚未读取配置',
      'diag.config.show': '展开',
      'diag.config.hide': '收起',
      'diag.settings': '设置项',
      'diag.save': '保存',
      'diag.saving': '保存中…',
      'diag.saved': '已保存',
      'diag.key.outputRoot': '输出根目录 outputRoot',
      'diag.key.parallel': '并行度 parallel',
      'diag.key.allowExternalPaths': '允许 2026 之外读取 allowExternalPaths',
      'diag.key.odaPath': 'ODA 转换工具路径 odaPath',
      'diag.key.pdfPaper': '默认纸型 pdfPaper',
      'diag.key.sortMode': '默认排序 sortMode',
      'diag.value.auto': '自动',
      'diag.invalidParallel': '并行度只能是 auto 或 1–8',
      'diag.cap.word': 'Word 文档读取',
      'diag.cap.excel': 'Excel 表格读取',
      'diag.cap.powerpoint': 'PowerPoint 演示文稿',
      'diag.cap.visio': 'Visio 绘图（.vsdx）',
      'diag.cap.autocad': 'AutoCAD COM 打印 → PDF',
      'diag.cap.zwcad': '中望 CAD COM 打印 → PDF',
      'diag.cap.gstarcad': '浩辰 CAD COM 打印 → PDF',
      'diag.cap.libreoffice': 'LibreOffice 兜底',
      'diag.cap.python': 'Python 引擎（P2 预留）',
      'diag.cap.oda': 'ODAFileConverter DWG → PDF',

      'common.ok': '确定',
      'common.cancel': '取消',
      'common.retry': '重试',
      'common.close': '关闭',
      'common.error.title': '操作失败',
      'common.error.code': '错误码：{code}',
      'common.error.hint': '建议：{hint}',
      'common.busy': '处理中…',
      'common.session.missing': '请先在左侧工作区打开一个会话',
      'common.command.missing': '命令不可用，请刷新页面（插件可能已升级）',
      'common.remoteMissing': '远程服务 {name} 未挂载，相关功能不可用',
      'common.cancelled': '已取消',
      'common.unknownError': '未知错误：{message}',
    };

    const en = {
      'panel.title': 'Office & Grid Suite',
      'panel.tabs': 'Panel tabs',
      'panel.session.current': 'Current session',
      'tab.docs': 'Documents',
      'tab.visio': 'Visio drawing',
      'tab.cad': 'Drawing conversion',
      'tab.diag': 'Settings & diagnostics',

      'docs.title': 'Document center',
      'docs.tree.root': 'Workspace files',
      'docs.tree.browse': 'Browse another folder…',
      'docs.tree.back': 'Back to workspace',
      'docs.tree.external': 'External folder',
      'docs.tree.filterAll': 'Show all files',
      'docs.tree.filterHint': 'Showing doc/docx/xls/xlsx/vsdx/dwg/pdf only',
      'docs.tree.empty': 'No matching files in this folder',
      'docs.tree.loading': 'Loading…',
      'docs.tree.up': 'Parent folder',
      'docs.tree.truncated': 'Too many entries; only the first part is shown',
      'docs.tree.dirsOnly': 'Only subfolders can be browsed here (file rows come from the suite.files command)',
      'docs.tree.noSession': 'An active session is required to browse workspace files',
      'docs.read': 'Read',
      'docs.reading': 'Reading…',
      'docs.read.prompt': 'Press “Read” to load this file',
      'docs.read.notSupported': 'Preview is not supported for this type',
      'docs.read.empty': 'Pick a file on the left, then press “Read”',
      'docs.preview.hint.pdf': 'PDF preview arrives in a later version',
      'docs.preview.hint.other': 'Preview for this type arrives in a later version',
      'docs.meta.size': 'Size',
      'docs.meta.version': 'Version',
      'docs.meta.path': 'Path',
      'docs.meta.name': 'File name',
      'docs.actions.refresh': 'Refresh',
      'docs.actions.copyPath': 'Copy path',
      'docs.actions.copied': 'Copied',
      'docs.word.paragraphs': 'Paragraphs',
      'docs.word.tables': 'Tables',
      'docs.word.headings': 'Headings',
      'docs.word.preview': 'Text preview',
      'docs.excel.sheet': 'Sheet',
      'docs.excel.sheets': 'Sheets',
      'docs.excel.rows': 'Rows',
      'docs.excel.cols': 'Columns',
      'docs.excel.sample': 'Sample rows',
      'docs.file.none': 'No file selected',

      'visio.title': 'Visio drawing',
      'visio.kind': 'Diagram type',
      'visio.kind.flowchart': 'Flowchart (P1)',
      'visio.titleField': 'Title',
      'visio.nodes': 'Nodes JSON (id/label/kind: start|process|decision|end)',
      'visio.edges': 'Edges JSON (from/to)',
      'visio.template': 'Insert template',
      'visio.validate.ok': 'JSON is valid',
      'visio.json.invalid': 'Invalid JSON: {message}',
      'visio.node.invalid': 'Invalid node (id/label required): item {index}',
      'visio.node.kindInvalid': 'Node kind must be start|process|decision|end: item {index}',
      'visio.node.duplicate': 'Duplicate node id: {id}',
      'visio.edge.invalid': 'Invalid edge (from/to required): item {index}',
      'visio.edge.missing': 'Edge references an unknown node: {id}',
      'visio.nl': 'Natural-language description (P2)',
      'visio.outputPath': 'Output path (empty = auto name)',
      'visio.outputHint': 'Empty saves to output\\diagrams\\ in the session workspace with an automatic name',
      'visio.chooseDir': 'Choose folder…',
      'visio.run': 'Draw',
      'visio.running': 'Drawing…',
      'visio.done.saved': 'Saved: {path}',
      'visio.done.stats': 'Nodes {nodes} · Edges {edges}',
      'visio.result': 'Result',
      'visio.history': 'Run history (this session)',
      'visio.history.empty': 'No runs yet',

      'cad.title': 'Drawing conversion',
      'cad.mode': 'Mode',
      'cad.mode.merge': 'Merge (folder of drawings → one PDF)',
      'cad.mode.split': 'Split (one multi-frame file → one page per frame)',
      'cad.input': 'Input',
      'cad.input.file': 'Choose DWG file…',
      'cad.input.dir': 'Choose folder…',
      'cad.input.placeholder': 'Not selected',
      'cad.input.path': 'Path',
      'cad.input.browse': 'Browse…',
      'cad.picker.title': 'Choose a drawing file',
      'cad.picker.hint': 'Only .dwg/.dxf/.pdf files are listed',
      'cad.picker.empty': 'No drawing files in this folder',
      'cad.picker.up': 'Parent folder',
      'cad.picker.close': 'Close',
      'cad.files.count': '{count} files',
      'cad.files.counting': 'Counting…',
      'cad.paper': 'PDF paper',
      'cad.paper.a4landscape': 'A4 landscape (default)',
      'cad.paper.keep': 'Keep original frame size',
      'cad.sort': 'Sort',
      'cad.sort.drawingNo': 'Drawing number first (fallback file name)',
      'cad.sort.filename': 'File name (natural order)',
      'cad.parallel': 'Parallelism',
      'cad.parallel.auto': 'Auto',
      'cad.start': 'Start conversion',
      'cad.cancel': 'Cancel',
      'cad.cancel.requested': 'Cancellation requested…',
      'cad.state.idle': 'Idle',
      'cad.state.queued': 'Queued',
      'cad.state.running': 'Running',
      'cad.state.converting': 'Converting',
      'cad.state.splitting': 'Splitting',
      'cad.state.merging': 'Merging',
      'cad.state.done': 'Done',
      'cad.state.failed': 'Failed',
      'cad.state.cancelled': 'Cancelled',
      'cad.progress.files': 'Files {done}/{total}',
      'cad.progress.frames': 'Frames {done}/{total}',
      'cad.progress.current': 'Current: {file}',
      'cad.log': 'Log',
      'cad.log.empty': 'No log yet',
      'cad.result.title': 'Conversion result',
      'cad.result.output': 'Output: {path}',
      'cad.result.openInDocs': 'Show in Document center',

      'diag.title': 'Settings & diagnostics',
      'diag.overview': 'Engine overview',
      'diag.overview.available': 'Available engines: {list}',
      'diag.overview.none': 'No office or drawing engine detected',
      'diag.overview.fallback': 'No CAD detected: conversion uses the ODA / PDF pipeline',
      'diag.engine.available': 'Available',
      'diag.engine.missing': 'Missing',
      'diag.engine.degraded': 'Degraded',
      'diag.table.engine': 'Engine',
      'diag.table.status': 'Status',
      'diag.table.version': 'Version / ProgID',
      'diag.table.capability': 'Capability',
      'diag.probe': 'Probe again',
      'diag.probing': 'Probing…',
      'diag.copy': 'Copy diagnostics',
      'diag.copy.done': 'Diagnostics copied',
      'diag.config': 'Effective configuration',
      'diag.config.note': 'Changes apply to the next conversion',
      'diag.config.empty': 'Configuration not loaded yet',
      'diag.config.show': 'Show',
      'diag.config.hide': 'Hide',
      'diag.settings': 'Settings',
      'diag.save': 'Save',
      'diag.saving': 'Saving…',
      'diag.saved': 'Saved',
      'diag.key.outputRoot': 'Output root (outputRoot)',
      'diag.key.parallel': 'Parallelism (parallel)',
      'diag.key.allowExternalPaths': 'Allow reads outside 2026 (allowExternalPaths)',
      'diag.key.odaPath': 'ODA converter path (odaPath)',
      'diag.key.pdfPaper': 'Default paper (pdfPaper)',
      'diag.key.sortMode': 'Default sort (sortMode)',
      'diag.value.auto': 'Auto',
      'diag.invalidParallel': 'Parallelism must be auto or 1–8',
      'diag.cap.word': 'Word document reading',
      'diag.cap.excel': 'Excel workbook reading',
      'diag.cap.powerpoint': 'PowerPoint decks',
      'diag.cap.visio': 'Visio drawing (.vsdx)',
      'diag.cap.autocad': 'AutoCAD COM print → PDF',
      'diag.cap.zwcad': 'ZWCAD COM print → PDF',
      'diag.cap.gstarcad': 'GstarCAD COM print → PDF',
      'diag.cap.libreoffice': 'LibreOffice fallback',
      'diag.cap.python': 'Python engine (P2 reserved)',
      'diag.cap.oda': 'ODAFileConverter DWG → PDF',

      'common.ok': 'OK',
      'common.cancel': 'Cancel',
      'common.retry': 'Retry',
      'common.close': 'Close',
      'common.error.title': 'Operation failed',
      'common.error.code': 'Error code: {code}',
      'common.error.hint': 'Suggestion: {hint}',
      'common.busy': 'Working…',
      'common.session.missing': 'Open a session in the left workspace first',
      'common.command.missing': 'Command unavailable — refresh the page (the plugin may have been upgraded)',
      'common.remoteMissing': 'Remote service {name} is not mounted; related features are disabled',
      'common.cancelled': 'Cancelled',
      'common.unknownError': 'Unknown error: {message}',
    };
    //#endregion

    //#region 2. 全局 CSS 文本（apply 时注入一次 <style data-plugin="dsh-office-suite-for-grid">）
    /** 面板样式。颜色/圆角/字体/动效一律走 §1.5 令牌；唯一例外是主按钮上的对比色白字。 */
    const STYLE_TEXT = [
      /* 布局骨架 */
      '.dso-panel{box-sizing:border-box;height:100%;min-height:0;display:flex;flex-direction:column;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-family:var(--dsw-font-family);font-size:13px;line-height:20px;overflow:hidden}',
      '.dso-head{flex:none;padding:20px 20px 0;display:flex;flex-direction:column;gap:6px}',
      '[data-platform=darwin] .dso-head{padding-top:calc(20px + var(--dsh-frame-top-clearance,0px))}',
      '.dso-headRow{display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap}',
      '.dso-title{margin:0;font-size:17px;font-weight:600;line-height:24px;color:var(--dsw-alias-label-primary)}',
      '.dso-sub{margin:0;font-size:12px;line-height:18px;color:var(--dsw-alias-label-secondary)}',
      '.dso-status{display:inline-flex;align-items:center;gap:6px;padding:2px 10px;border-radius:var(--dsw-radius-sm);border:.5px solid var(--dsw-alias-border-l2);color:var(--dsw-alias-label-secondary);font-size:12px;line-height:18px;max-width:100%;overflow-wrap:anywhere}',
      '.dso-status--ok{border-color:color-mix(in srgb, var(--dsw-alias-state-success-primary) 45%, transparent);color:var(--dsw-alias-state-success-primary);background:color-mix(in srgb, var(--dsw-alias-state-success-primary) 8%, transparent)}',
      '.dso-status--warn{border-color:color-mix(in srgb, var(--dsw-alias-state-warn-primary) 45%, transparent);color:var(--dsw-alias-state-warn-primary);background:color-mix(in srgb, var(--dsw-alias-state-warn-primary) 8%, transparent)}',
      '.dso-status--error{border-color:color-mix(in srgb, var(--dsw-alias-state-error-primary) 45%, transparent);color:var(--dsw-alias-state-error-primary);background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 8%, transparent)}',
      '.dso-dot{width:6px;height:6px;border-radius:50%;background:currentColor;flex:none}',

      /* 页签 */
      '.dso-tabs{flex:none;display:flex;gap:2px;padding:12px 20px 0;border-bottom:.5px solid var(--dsw-alias-border-l1);overflow-x:auto}',
      '.dso-tab{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:13px;line-height:20px;padding:6px 12px;border-radius:var(--dsw-radius-sm) var(--dsw-radius-sm) 0 0;cursor:pointer;white-space:nowrap;transition:background var(--ds-transition-duration-fast) var(--ds-ease-in-out),color var(--ds-transition-duration-fast) var(--ds-ease-in-out)}',
      '.dso-tab:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}',
      '.dso-tab[aria-current=page]{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);box-shadow:inset 0 -2px 0 var(--dsw-alias-state-business-primary)}',
      '.dso-tab:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-state-business-primary));outline-offset:-2px}',

      /* 页体 */
      '.dso-body{flex:1;min-height:0;display:flex;overflow:hidden}',
      '.dso-pane{flex:1;min-width:0;min-height:0;display:flex;flex-direction:column;padding:16px 20px 20px;overflow:auto;gap:12px}',
      '.dso-pane[hidden]{display:none}',
      '.dso-split{flex:1;min-height:0;display:grid;grid-template-columns:240px minmax(0,1fr);gap:12px}',
      '@media (max-width:760px){.dso-split{grid-template-columns:1fr;grid-template-rows:minmax(120px,32%) minmax(0,1fr)}}',
      '.dso-stack{display:flex;flex-direction:column;gap:12px;min-width:0}',
      '.dso-row2{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;align-items:end}',
      '.dso-toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap}',
      '.dso-spacer{flex:1}',
      '.dso-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}',

      /* 卡片 / 表格 */
      '.dso-card{background:var(--dsw-alias-bg-layer-1);border:.5px solid var(--dsw-alias-border-l2);border-radius:var(--dsw-radius-md);min-width:0;display:flex;flex-direction:column;overflow:hidden}',
      '.dso-card--flush{overflow:visible}',
      '.dso-cardHead{display:flex;align-items:center;gap:8px;padding:8px 12px;border-bottom:.5px solid var(--dsw-alias-border-l1);flex:none;flex-wrap:wrap}',
      '.dso-cardTitle{margin:0;font-size:13px;font-weight:600;line-height:20px;color:var(--dsw-alias-label-primary)}',
      '.dso-cardBody{padding:12px;min-width:0;overflow:auto;flex:1;display:flex;flex-direction:column;gap:10px}',
      '.dso-table{width:100%;border-collapse:collapse;font-size:12px;line-height:18px}',
      '.dso-table th,.dso-table td{text-align:left;padding:5px 8px;border-bottom:.5px solid var(--dsw-alias-border-l1);vertical-align:top;overflow-wrap:anywhere}',
      '.dso-table th{color:var(--dsw-alias-label-secondary);font-weight:500;border-bottom-color:var(--dsw-alias-border-l2)}',
      '.dso-table tr:last-child td{border-bottom:0}',
      '.dso-table--sticky th{position:sticky;top:0;background:var(--dsw-alias-bg-layer-1);z-index:1}',

      /* 文件树 */
      '.dso-tree{flex:1;min-height:0;overflow:auto;padding:6px}',
      '.dso-row{display:flex;align-items:center;gap:4px;padding:3px 6px;border-radius:var(--dsw-radius-xs);cursor:pointer;color:var(--dsw-alias-label-primary);white-space:nowrap;overflow:hidden}',
      '.dso-row:hover{background:var(--dsw-alias-interactive-bg-hover)}',
      '.dso-row[aria-selected=true]{background:color-mix(in srgb, var(--dsw-alias-state-business-primary) 14%, transparent);box-shadow:inset 0 0 0 .5px color-mix(in srgb, var(--dsw-alias-state-business-primary) 45%, transparent)}',
      '.dso-row:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-state-business-primary));outline-offset:-2px}',
      '.dso-caret{width:14px;height:14px;flex:none;display:inline-flex;align-items:center;justify-content:center;color:var(--dsw-alias-label-tertiary);font-size:10px;border-radius:var(--dsw-radius-xs);cursor:pointer}',
      '.dso-caret:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}',
      '.dso-rowName{overflow:hidden;text-overflow:ellipsis;min-width:0}',
      '.dso-rowMeta{margin-left:auto;flex:none;color:var(--dsw-alias-label-tertiary);font-size:11px}',
      '.dso-treeNote{padding:6px 8px;color:var(--dsw-alias-label-tertiary);font-size:12px;line-height:18px;overflow-wrap:anywhere}',
      '.dso-crumbs{display:flex;align-items:center;gap:4px;flex-wrap:wrap;font-size:12px;color:var(--dsw-alias-label-secondary);min-width:0}',
      '.dso-crumb{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:12px;padding:1px 3px;border-radius:var(--dsw-radius-xs);cursor:pointer;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
      '.dso-crumb:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}',
      '.dso-crumbSep{color:var(--dsw-alias-label-tertiary)}',

      /* 按钮 */
      '.dso-btn{appearance:none;font:inherit;font-size:12px;line-height:18px;display:inline-flex;align-items:center;justify-content:center;gap:6px;padding:3px 10px;border-radius:var(--dsw-radius-sm);border:.5px solid var(--dsw-alias-border-l2);background:transparent;color:var(--dsw-alias-label-primary);cursor:pointer;transition:background var(--ds-transition-duration-fast) var(--ds-ease-in-out),border-color var(--ds-transition-duration-fast) var(--ds-ease-in-out);white-space:nowrap}',
      '.dso-btn:hover:not(:disabled){background:var(--dsw-alias-interactive-bg-hover)}',
      '.dso-btn:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-state-business-primary));outline-offset:1px}',
      '.dso-btn:disabled{opacity:.5;cursor:not-allowed}',
      '.dso-btn--primary{background:var(--dsw-alias-brand-primary);border-color:transparent;color:#fff}',
      '.dso-btn--primary:hover:not(:disabled){filter:brightness(.94)}',
      '.dso-btn--danger{color:var(--dsw-alias-state-error-primary);border-color:color-mix(in srgb, var(--dsw-alias-state-error-primary) 45%, transparent)}',
      '.dso-btn--danger:hover:not(:disabled){background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 10%, transparent)}',
      '.dso-btn--link{border-color:transparent;color:var(--dsw-alias-state-business-primary);padding:1px 4px}',
      '.dso-btn--sm{padding:1px 8px;font-size:11px;line-height:16px}',

      /* 表单 */
      '.dso-field{display:flex;flex-direction:column;gap:4px;min-width:0}',
      '.dso-label{font-size:12px;line-height:18px;color:var(--dsw-alias-label-secondary)}',
      '.dso-input,.dso-select,.dso-textarea{box-sizing:border-box;width:100%;font:inherit;font-size:12px;line-height:18px;color:var(--dsw-alias-label-primary);background:var(--dsw-alias-bg-base);border:.5px solid var(--dsw-alias-border-l2);border-radius:var(--dsw-radius-sm);padding:4px 8px;transition:border-color var(--ds-transition-duration-fast) var(--ds-ease-in-out)}',
      '.dso-input:hover,.dso-select:hover,.dso-textarea:hover{border-color:var(--dsw-alias-border-l3)}',
      '.dso-input:focus-visible,.dso-select:focus-visible,.dso-textarea:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-state-business-primary));outline-offset:1px;border-color:var(--dsw-alias-state-business-primary)}',
      '.dso-input:disabled,.dso-select:disabled,.dso-textarea:disabled{color:var(--dsw-alias-state-idle-primary);cursor:not-allowed}',
      '.dso-textarea{font-family:var(--ds-font-family-code);font-size:12px;line-height:18px;min-height:132px;resize:vertical;white-space:pre;overflow:auto}',
      '.dso-inline{display:flex;align-items:center;gap:8px;flex-wrap:wrap;min-width:0}',
      '.dso-check{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--dsw-alias-label-secondary);cursor:pointer;user-select:none}',
      '.dso-check input{accent-color:var(--dsw-alias-brand-primary);margin:0}',

      /* 文本与徽标 */
      '.dso-muted{color:var(--dsw-alias-label-secondary)}',
      '.dso-faint{color:var(--dsw-alias-label-tertiary)}',
      '.dso-error{color:var(--dsw-alias-state-error-primary)}',
      '.dso-warn{color:var(--dsw-alias-state-warn-primary)}',
      '.dso-ok{color:var(--dsw-alias-state-success-primary)}',
      '.dso-mono{font-family:var(--ds-font-family-code);font-size:12px;overflow-wrap:anywhere}',
      '.dso-pre{margin:0;font-family:var(--ds-font-family-code);font-size:11.5px;line-height:17px;background:var(--dsw-alias-bg-layer-2);border:.5px solid var(--dsw-alias-border-l1);border-radius:var(--dsw-radius-sm);padding:8px;overflow:auto;max-height:320px;white-space:pre-wrap;overflow-wrap:anywhere;color:var(--dsw-alias-label-secondary)}',
      '.dso-badge{display:inline-flex;align-items:center;gap:4px;padding:0 6px;border-radius:var(--dsw-radius-xs);font-size:11px;line-height:16px;border:.5px solid var(--dsw-alias-border-l2);color:var(--dsw-alias-label-secondary);white-space:nowrap}',
      '.dso-badge--ok{color:var(--dsw-alias-state-success-primary);border-color:color-mix(in srgb, var(--dsw-alias-state-success-primary) 45%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-success-primary) 8%, transparent)}',
      '.dso-badge--warn{color:var(--dsw-alias-state-warn-primary);border-color:color-mix(in srgb, var(--dsw-alias-state-warn-primary) 45%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-warn-primary) 8%, transparent)}',
      '.dso-badge--error{color:var(--dsw-alias-state-error-primary);border-color:color-mix(in srgb, var(--dsw-alias-state-error-primary) 45%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 8%, transparent)}',
      '.dso-badge--idle{color:var(--dsw-alias-state-idle-primary)}',
      '.dso-kv{display:grid;grid-template-columns:auto minmax(0,1fr);gap:2px 12px;font-size:12px;line-height:18px}',
      '.dso-kv dt{color:var(--dsw-alias-label-secondary);margin:0}',
      '.dso-kv dd{margin:0;color:var(--dsw-alias-label-primary);overflow-wrap:anywhere}',
      '.dso-empty{flex:1;min-height:120px;display:flex;align-items:center;justify-content:center;text-align:center;color:var(--dsw-alias-label-tertiary);font-size:12px;line-height:18px;padding:16px;overflow-wrap:anywhere}',
      '.dso-hint{font-size:11.5px;line-height:17px;color:var(--dsw-alias-label-tertiary);overflow-wrap:anywhere}',
      '.dso-errorBox{display:flex;flex-direction:column;gap:2px;border:.5px solid color-mix(in srgb, var(--dsw-alias-state-error-primary) 40%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 7%, transparent);border-radius:var(--dsw-radius-sm);padding:8px 10px;font-size:12px;line-height:18px;color:var(--dsw-alias-state-error-primary);overflow-wrap:anywhere}',
      '.dso-okBox{display:flex;flex-direction:column;gap:2px;border:.5px solid color-mix(in srgb, var(--dsw-alias-state-success-primary) 40%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-success-primary) 7%, transparent);border-radius:var(--dsw-radius-sm);padding:8px 10px;font-size:12px;line-height:18px;overflow-wrap:anywhere}',

      /* 进度条 / 日志 */
      '.dso-progress{height:6px;border-radius:999px;background:var(--dsw-alias-bg-layer-2);border:.5px solid var(--dsw-alias-border-l1);overflow:hidden}',
      '.dso-progressBar{height:100%;background:var(--dsw-alias-state-business-primary);transition:width var(--ds-transition-duration-slow) var(--ds-ease-in-out)}',
      '.dso-log{flex:1;min-height:120px;max-height:280px;overflow:auto;background:var(--dsw-alias-bg-layer-2);border:.5px solid var(--dsw-alias-border-l1);border-radius:var(--dsw-radius-sm);padding:6px 8px;font-family:var(--ds-font-family-code);font-size:11.5px;line-height:16px;color:var(--dsw-alias-label-secondary)}',
      '.dso-logLine{display:flex;gap:8px;white-space:pre-wrap;overflow-wrap:anywhere}',
      '.dso-logTime{color:var(--dsw-alias-label-tertiary);flex:none}',
      '.dso-logLine[data-level=warn]{color:var(--dsw-alias-state-warn-primary)}',
      '.dso-logLine[data-level=error]{color:var(--dsw-alias-state-error-primary)}',
      '.dso-logLine[data-level=success]{color:var(--dsw-alias-state-success-primary)}',

      /* 微件 */
      '.dso-spinner{width:12px;height:12px;flex:none;border-radius:50%;border:1.5px solid var(--dsw-alias-label-tertiary);border-top-color:transparent;animation:dso-spin .8s linear infinite}',
      '@keyframes dso-spin{to{transform:rotate(360deg)}}',
      '.dso-h1{font-size:14px;line-height:20px;font-weight:600;color:var(--dsw-alias-label-primary)}',
      '.dso-h2{font-size:13px;line-height:20px;font-weight:600;color:var(--dsw-alias-label-primary);padding-left:10px}',
      '.dso-h3{font-size:12px;line-height:18px;font-weight:600;color:var(--dsw-alias-label-secondary);padding-left:20px}',
      '.dso-list{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:2px}',
      '.dso-divider{height:0;border-top:.5px solid var(--dsw-alias-border-l1);margin:2px 0}',
      '.dso-sr{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0}',
    ].join('\n');
    //#endregion

    //#region 3. 工具函数（结果展开、命令 JSON 解析、自然排序、扩展名过滤、格式化）
    /** 文档中心默认显示的文件扩展名（§1.3）。 */
    const DEFAULT_EXTS = ['doc', 'docx', 'xls', 'xlsx', 'vsdx', 'dwg', 'pdf'];
    /** 可直接读取渲染的 Office 扩展名。 */
    const OFFICE_EXTS = ['doc', 'docx', 'xls', 'xlsx'];
    /** P1 只显示元信息的扩展名。 */
    const META_EXTS = ['vsdx', 'dwg', 'pdf'];
    /** 图纸转换可接受的输入扩展名（.pdf 供 PDF 级引擎链路）。 */
    const CAD_INPUT_EXTS = ['dwg', 'dxf', 'pdf'];
    /** CAD 任务终态。 */
    const CAD_TERMINAL = ['done', 'failed', 'cancelled'];
    /** 页签记忆键（§1.3）。 */
    const TAB_STORAGE_KEY = 'dsh.officeSuite.tab';
    /** 页签定义（顺序即渲染顺序）。 */
    const TABS = [
      { id: 'docs', label: 'tab.docs' },
      { id: 'visio', label: 'tab.visio' },
      { id: 'cad', label: 'tab.cad' },
      { id: 'diag', label: 'tab.diag' },
    ];

    /** `{name}` 占位符插值（字典里不允许写正则）。 */
    function fmt(text, vars) {
      if (typeof text !== 'string') return String(text ?? '');
      if (!vars) return text;
      return text.replace(/\{(\w+)\}/g, (match, key) =>
        Object.prototype.hasOwnProperty.call(vars, key) && vars[key] !== undefined && vars[key] !== null
          ? String(vars[key])
          : match);
    }

    /** 从任意文本中取首个 JSON 对象（Host 结果容忍前导警告行，与 host/ops.js 的 parseJson 同义）。 */
    function parseJsonLoose(text) {
      if (typeof text !== 'string' || text.length === 0) return undefined;
      try { return JSON.parse(text); } catch { /* 继续宽松解析 */ }
      const start = text.indexOf('{');
      const end = text.lastIndexOf('}');
      if (start < 0 || end <= start) return undefined;
      try { return JSON.parse(text.slice(start, end + 1)); } catch { return undefined; }
    }

    /** 小写扩展名（不含点）；无扩展名返回 ''。 */
    function extOf(name) {
      const text = String(name ?? '');
      const dot = text.lastIndexOf('.');
      return dot < 0 ? '' : text.slice(dot + 1).toLowerCase();
    }

    /** 拼接路径：按父路径的分隔符风格（Windows 反斜杠 / POSIX 正斜杠）。 */
    function joinPath(parent, name) {
      const base = String(parent ?? '');
      if (!base) return String(name ?? '');
      const sep = base.includes('\\') ? '\\' : '/';
      return base.endsWith(sep) ? base + name : base + sep + name;
    }

    /** 取父目录（保留盘符根）。 */
    function parentPath(path) {
      const text = String(path ?? '');
      if (!text) return '';
      const sep = text.includes('\\') ? '\\' : '/';
      const trimmed = text.endsWith(sep) ? text.slice(0, -1) : text;
      const cut = trimmed.lastIndexOf(sep);
      if (cut < 0) return '';
      if (cut === 0) return sep;
      const head = trimmed.slice(0, cut);
      return /^[A-Za-z]:$/.test(head) ? head + sep : head;
    }

    /** 自然排序比较器（数字段按数值比较）。 */
    function naturalCompare(a, b) {
      return String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' });
    }

    /** 字节数格式化。 */
    function formatBytes(value) {
      if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) return '—';
      if (value < 1024) return value + ' B';
      const units = ['KB', 'MB', 'GB', 'TB'];
      let size = value / 1024;
      let index = 0;
      while (size >= 1024 && index < units.length - 1) { size /= 1024; index += 1; }
      return (size >= 10 ? size.toFixed(0) : size.toFixed(1)) + ' ' + units[index];
    }

    /** 时钟格式化（HH:MM:SS）。 */
    function formatClock(ms) {
      const date = new Date(typeof ms === 'number' ? ms : Date.now());
      const pad = (n) => (n < 10 ? '0' + n : String(n));
      return pad(date.getHours()) + ':' + pad(date.getMinutes()) + ':' + pad(date.getSeconds());
    }

    /** 文件时间戳（yyyyMMdd-HHmmss，用于自动命名）。 */
    function formatStamp(ms) {
      const date = new Date(typeof ms === 'number' ? ms : Date.now());
      const pad = (n) => (n < 10 ? '0' + n : String(n));
      return String(date.getFullYear()) + pad(date.getMonth() + 1) + pad(date.getDate()) + '-' +
        pad(date.getHours()) + pad(date.getMinutes()) + pad(date.getSeconds());
    }

    /** 会话短标识（只展示尾部 8 位，避免面板过宽）。 */
    function shortId(id) {
      const text = String(id ?? '');
      return text.length > 8 ? '…' + text.slice(-8) : text;
    }

    /** 把任意抛出物规整成面板错误 `{code, message, hint, cancelled}`。 */
    function toPanelError(error, t) {
      const tr = typeof t === 'function' ? t : (key) => key;
      if (error && typeof error === 'object') {
        const code = typeof error.code === 'string' && error.code ? error.code : 'UNKNOWN';
        if (code === 'gateway/cancelled') {
          return { code, cancelled: true, message: tr('common.cancelled'), hint: undefined };
        }
        if (code === 'COMMAND_MISSING') {
          return { code, cancelled: false, message: tr('common.command.missing'), hint: undefined };
        }
        if (code === 'REMOTE_MISSING') {
          const name = typeof error.name === 'string' && error.name ? error.name : 'remote';
          return { code, cancelled: false, message: fmt(tr('common.remoteMissing'), { name }), hint: undefined };
        }
        const rawMessage = typeof error.message === 'string' ? error.message : '';
        const message = rawMessage || tr('common.unknownError', { message: code });
        return {
          code,
          cancelled: false,
          message,
          hint: typeof error.hint === 'string' && error.hint ? error.hint : undefined,
        };
      }
      return { code: 'UNKNOWN', cancelled: false, message: tr('common.unknownError', { message: String(error) }), hint: undefined };
    }

    /** 复制文本（navigator.clipboard 优先，失败回退隐藏 textarea）。 */
    async function copyText(text) {
      const value = String(text ?? '');
      try {
        if (typeof navigator !== 'undefined' && navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
          await navigator.clipboard.writeText(value);
          return true;
        }
      } catch { /* 回退 */ }
      try {
        const area = document.createElement('textarea');
        area.value = value;
        area.setAttribute('readonly', 'readonly');
        area.style.position = 'fixed';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        const ok = document.execCommand('copy');
        area.remove();
        return ok;
      } catch { return false; }
    }

    /** 读取页签记忆；无有效值时回落 docs。 */
    function readStoredTab() {
      try {
        const stored = window.localStorage.getItem(TAB_STORAGE_KEY);
        return TABS.some((tab) => tab.id === stored) ? stored : 'docs';
      } catch { return 'docs'; }
    }

    /** 当前主视图会话 id（§0.5：retainedBy.mainView > 0 的那一行）。 */
    function currentSessionIdFrom(snapshot) {
      const rows = snapshot && snapshot.byId ? snapshot.byId : {};
      const hit = Object.values(rows).find((row) => (row && row.retainedBy && (row.retainedBy.mainView ?? 0) > 0));
      return hit && typeof hit.id === 'string' ? hit.id : null;
    }

    /** 订阅会话 store，返回当前会话 id（无会话为 null）。 */
    function useCurrentSessionId(api) {
      const [sessionId, setSessionId] = React.useState(() => {
        try { return api.currentSessionId(); } catch { return null; }
      });
      React.useEffect(() => {
        const read = () => {
          let next = null;
          try { next = api.currentSessionId(); } catch { next = null; }
          setSessionId((prev) => (prev === next ? prev : next));
        };
        read();
        let dispose = () => {};
        try { dispose = api.subscribeSessions(read) || (() => {}); } catch { dispose = () => {}; }
        return () => { try { dispose(); } catch { /* 忽略 */ } };
      }, [api]);
      return sessionId;
    }

    /** 连接复位订阅（面板在重连后刷新自身状态）。 */
    function useConnectionReset(api, refresh) {
      const handler = React.useRef(refresh);
      handler.current = refresh;
      React.useEffect(() => {
        let dispose = () => {};
        try {
          dispose = api.subscribeReset(() => handler.current()) || (() => {});
        } catch { dispose = () => {}; }
        return () => { try { dispose(); } catch { /* 忽略 */ } };
      }, [api]);
    }

    /** 统一错误块。 */
    function ErrorBox({ t, error, onRetry }) {
      if (!error) return null;
      return h('div', { className: 'dso-errorBox', role: 'alert' },
        h('div', { style: { fontWeight: 600 } }, t('common.error.title')),
        h('div', null, error.message),
        error.code && error.code !== 'UNKNOWN' ? h('div', { className: 'dso-hint' }, fmt(t('common.error.code'), { code: error.code })) : null,
        error.hint ? h('div', { className: 'dso-hint' }, fmt(t('common.error.hint'), { hint: error.hint })) : null,
        onRetry ? h('div', null, h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: onRetry }, t('common.retry'))) : null);
    }

    /** 空态提示。 */
    function Empty({ text }) {
      return h('div', { className: 'dso-empty' }, text);
    }

    /** 行内 spinner + 文本。 */
    function Busy({ text }) {
      return h('span', { className: 'dso-inline' }, h('span', { className: 'dso-spinner', 'aria-hidden': 'true' }), h('span', { className: 'dso-muted' }, text));
    }
    //#endregion

    //#region 4. remote 封装（callCommand / listWorkspace / externalList / subscribeProgress）
    /**
     * 防御式读取命名空间服务。`remote.commands` / `remote.workspaceFiles` /
     * `remote.directoryPicker` 不写进 inject 硬依赖，缺挂时返回 null 由调用方降级。
     */
    function safeService(ctx, key) {
      try {
        const service = ctx.get(key);
        return service === undefined || service === null ? null : service;
      } catch { return null; }
    }

    /** 把抛出物规整为 `{code, message, hint?}`（保留 gateway/* 框架码）。 */
    function normalizeThrown(error) {
      if (error && typeof error === 'object') {
        const code = typeof error.code === 'string' && error.code ? error.code : 'UNKNOWN';
        const message = typeof error.message === 'string' && error.message
          ? error.message
          : (typeof error.details === 'object' && error.details && typeof error.details.message === 'string' ? error.details.message : 'unknown remote error');
        const hint = typeof error.hint === 'string' && error.hint ? error.hint : undefined;
        return hint ? { code, message, hint } : { code, message };
      }
      return { code: 'UNKNOWN', message: String(error) };
    }

    /**
     * 面板 → Host：line = '/' + name + ' ' + JSON.stringify(args)，
     * 经 remote.commands.execute(line, [], signal) 执行并展开
     * `{ok:true,value}|{ok:false,error}`（§0.4），再解出 Host 的 `{ok,result}|{ok:false,error}`（§2）。
     */
    async function executeCommand(ctx, name, args, signal) {
      const commands = safeService(ctx, 'remote.commands');
      if (!commands || typeof commands.execute !== 'function') {
        throw { code: 'COMMAND_MISSING', message: 'remote.commands is not mounted' };
      }
      const line = '/' + name + ' ' + JSON.stringify(args === undefined ? {} : args);
      let wrapped;
      try {
        wrapped = await commands.execute(line, [], signal);
      } catch (error) {
        throw normalizeThrown(error);
      }
      if (!wrapped || wrapped.ok !== true) throw normalizeThrown(wrapped && wrapped.error);
      const execution = wrapped.value;
      if (execution === undefined || execution === null) {
        // §1.4 步骤 4：命令解析不到（升级插件后旧页面）
        throw { code: 'COMMAND_MISSING', message: 'command not found: ' + name };
      }
      const settled = execution.result ?? execution;
      if (settled && settled.kind === 'error') {
        const parsed = parseJsonLoose(settled.text);
        if (parsed && parsed.error) throw normalizeThrown(parsed.error);
        throw { code: 'COMMAND_FAILED', message: typeof settled.text === 'string' && settled.text ? settled.text : 'command failed' };
      }
      const text = settled && typeof settled.text === 'string' ? settled.text : '';
      const parsed = parseJsonLoose(text);
      if (!parsed) {
        throw { code: 'BAD_RESPONSE', message: (name + ': command returned non-JSON output') };
      }
      if (parsed.ok === false) throw normalizeThrown(parsed.error);
      return parsed.result;
    }

    /** 会话工作区目录列举（remote.workspaceFiles.list，已展开结果包装）。 */
    async function workspaceList(ctx, path, signal) {
      const files = safeService(ctx, 'remote.workspaceFiles');
      if (!files || typeof files.list !== 'function') {
        throw { code: 'REMOTE_MISSING', name: 'remote.workspaceFiles', message: 'remote.workspaceFiles is not mounted' };
      }
      let wrapped;
      try {
        wrapped = await files.list(path, signal);
      } catch (error) {
        throw normalizeThrown(error);
      }
      if (!wrapped || wrapped.ok !== true) throw normalizeThrown(wrapped && wrapped.error);
      return wrapped.value;
    }

    /** 会话工作区文件元信息（remote.workspaceFiles.stat，已展开结果包装）。 */
    async function workspaceStat(ctx, path, signal) {
      const files = safeService(ctx, 'remote.workspaceFiles');
      if (!files || typeof files.stat !== 'function') {
        throw { code: 'REMOTE_MISSING', name: 'remote.workspaceFiles', message: 'remote.workspaceFiles is not mounted' };
      }
      const wrapped = await files.stat(path, signal);
      if (!wrapped || wrapped.ok !== true) throw normalizeThrown(wrapped && wrapped.error);
      return wrapped.value;
    }

    /** 任意目录列举（uiWorkspace.listDirectory，已是解开结果包装的接口，失败抛错）。 */
    function browseDir(ctx, path, signal) {
      const workspace = safeService(ctx, 'uiWorkspace');
      if (!workspace || typeof workspace.listDirectory !== 'function') {
        return Promise.reject({ code: 'REMOTE_MISSING', name: 'uiWorkspace', message: 'uiWorkspace is not mounted' });
      }
      return workspace.listDirectory(path, signal);
    }

    /** 系统目录选择器（uiWorkspace.pickDirectory，返回绝对路径或 null）。 */
    function pickDirectory(ctx) {
      const workspace = safeService(ctx, 'uiWorkspace');
      if (!workspace || typeof workspace.pickDirectory !== 'function') {
        return Promise.reject({ code: 'REMOTE_MISSING', name: 'uiWorkspace', message: 'uiWorkspace is not mounted' });
      }
      return workspace.pickDirectory();
    }

    /**
     * 外部目录列举（§1.3「浏览其他目录」模式）。
     * 首选 P1 远程命令 suite.files（含目录 + 文件 + 大小）；命令未实现/失败时
     * 降级为 uiWorkspace.listDirectory —— 该接口只返回**子目录**，故标记 dirsOnly。
     */
    async function externalList(ctx, path, signal) {
      let suiteError = null;
      try {
        const result = await executeCommand(ctx, 'suite.files', path ? { path } : {}, signal);
        const base = result && typeof result.path === 'string' && result.path ? result.path : (path || '');
        const entries = (Array.isArray(result && result.entries) ? result.entries : []).map((raw) => {
          const name = String(raw && raw.name !== undefined ? raw.name : '');
          const type = raw && raw.type === 'directory' ? 'dir' : raw && raw.type === 'file' ? 'file' : 'other';
          return {
            name,
            type,
            size: raw && typeof raw.size === 'number' ? raw.size : undefined,
            path: joinPath(base, name),
          };
        });
        return { entries, truncated: !!(result && result.truncated), dirsOnly: false, base: base || (path || ''), crumbs: [] };
      } catch (error) {
        suiteError = error;
      }
      let listing;
      try {
        listing = await browseDir(ctx, path || undefined, signal);
      } catch (error) {
        throw normalizeThrown(suiteError || error);
      }
      const base = listing && typeof listing.path === 'string' && listing.path ? listing.path : (path || '');
      const entries = (Array.isArray(listing && listing.entries) ? listing.entries : []).map((raw) => {
        const name = String(raw && raw.name !== undefined ? raw.name : '');
        return {
          name,
          type: 'dir',
          size: undefined,
          path: raw && typeof raw.path === 'string' && raw.path ? raw.path : joinPath(base, name),
        };
      });
      return {
        entries,
        truncated: !!(listing && listing.truncated),
        dirsOnly: true,
        base,
        crumbs: Array.isArray(listing && listing.crumbs) ? listing.crumbs : [],
      };
    }

    /** 订阅会话列表变化。 */
    function subscribeSessions(ctx, listener) {
      const store = ctx.sessions && ctx.sessions.list;
      if (!store || typeof store.subscribe !== 'function') return () => {};
      const dispose = store.subscribe(listener);
      return typeof dispose === 'function' ? dispose : () => {};
    }

    /** 订阅连接复位（重连后刷新面板状态，§0.4）。 */
    function subscribeReset(ctx, listener) {
      if (typeof ctx.on !== 'function') return () => {};
      const dispose = ctx.on('connection/reset', listener);
      return typeof dispose === 'function' ? dispose : () => {};
    }

    /**
     * 订阅 Host 任务进度事件（§0.4 ①/③：P1 以轮询为主，事件仅作加速）。
     * 事件源未注册时 `$on` 仍可安全订阅，收不到帧即自然退化为纯轮询。
     */
    function subscribeProgress(ctx, listener) {
      const remote = safeService(ctx, 'remote');
      if (!remote || typeof remote.$on !== 'function') return () => {};
      try {
        const dispose = remote.$on('office-suite/task-progress', listener);
        return typeof dispose === 'function' ? dispose : () => {};
      } catch { return () => {}; }
    }

    /** 组装注入给面板组件的 api 袋（组件不直接摸 ctx）。 */
    function buildApi(ctx) {
      return {
        currentSessionId: () => {
          const store = ctx.sessions && ctx.sessions.list;
          if (!store || typeof store.getSnapshot !== 'function') return null;
          return currentSessionIdFrom(store.getSnapshot());
        },
        callCommand: (name, args, signal) => executeCommand(ctx, name, args, signal),
        listWorkspace: (path, signal) => workspaceList(ctx, path, signal),
        statWorkspace: (path, signal) => workspaceStat(ctx, path, signal),
        externalList: (path, signal) => externalList(ctx, path, signal),
        browseDir: (path, signal) => browseDir(ctx, path, signal),
        pickDir: () => pickDirectory(ctx),
        copy: copyText,
        subscribeSessions: (listener) => subscribeSessions(ctx, listener),
        subscribeReset: (listener) => subscribeReset(ctx, listener),
        subscribeProgress: (listener) => subscribeProgress(ctx, listener),
      };
    }

    /** 把两种列举结果统一成 FsEntry：{name, type:'dir'|'file'|'other', path, size}。 */
    function toFsEntry(raw, mode, base) {
      const name = String(raw && raw.name !== undefined ? raw.name : '');
      const type = raw && raw.type === 'directory' ? 'dir' : raw && raw.type === 'file' ? 'file' : 'other';
      const path = mode === 'workspace'
        ? joinPath(base === '.' ? '' : base, name)
        : (raw && typeof raw.path === 'string' && raw.path ? raw.path : joinPath(base, name));
      return { name, type, path, size: raw && typeof raw.size === 'number' ? raw.size : undefined };
    }

    /** 目录优先、随后自然序。 */
    function compareEntries(left, right) {
      if (left.type === 'dir' && right.type !== 'dir') return -1;
      if (right.type === 'dir' && left.type !== 'dir') return 1;
      return naturalCompare(left.name, right.name);
    }
    //#endregion

    //#region 5. 组件：PanelGlyph（侧栏图标，只渲染字形，props {size, active}）
    /** 办公文件夹 + 电网折线组合字形；图标配色自有（§1.2 允许）。 */
    function PanelGlyph(props) {
      const size = typeof props.size === 'number' ? props.size : 18;
      const active = props.active === true;
      return h('svg', {
        viewBox: '0 0 20 20',
        width: size,
        height: size,
        'aria-hidden': 'true',
        focusable: 'false',
        style: { display: 'block' },
      },
      h('path', {
        d: 'M2.5 5A1.5 1.5 0 0 1 4 3.5h3.1l1.4 1.9H16A1.5 1.5 0 0 1 17.5 6.9v7.6A1.5 1.5 0 0 1 16 16H4a1.5 1.5 0 0 1-1.5-1.5V5Z',
        fill: 'none',
        stroke: active ? '#2f66d0' : '#4176e6',
        strokeWidth: 1.4,
        strokeLinejoin: 'round',
      }),
      h('path', {
        d: 'M6 13.6 9.3 9.4l2.4 2.1L15 8',
        fill: 'none',
        stroke: '#22c55e',
        strokeWidth: 1.4,
        strokeLinecap: 'round',
        strokeLinejoin: 'round',
      }),
      h('circle', { cx: 15, cy: 8, r: 1.15, fill: '#22c55e' }));
    }
    //#endregion

    //#region 6. 组件：OfficeSuitePanel（页签壳 + 会话状态条）
    /** 面板外壳：标题 / 会话状态 / 页签 / 四个页签（全部常驻挂载，隐藏时不卸载以保住轮询与状态）。 */
    function OfficeSuitePanel(props) {
      const api = props.api;
      const t = typeof props.t === 'function' ? props.t : props.tf;
      const [tab, setTab] = React.useState(readStoredTab);
      const [openDir, setOpenDir] = React.useState(null);
      const sessionId = useCurrentSessionId(api);

      // 页签记忆：切换即写入（等价于「卸载时保存」，且刷新后也能恢复）。
      React.useEffect(() => {
        try { window.localStorage.setItem(TAB_STORAGE_KEY, tab); } catch { /* 隐私模式忽略 */ }
      }, [tab]);

      // CadTab 完成后「在文档中心查看」：切页签 + 把输出目录设为浏览目录。
      const openInDocs = React.useCallback((path) => {
        if (!path) return;
        setOpenDir({ path: parentPath(path) || path, nonce: Date.now() });
        setTab('docs');
      }, []);

      const tabs = TABS.map((item) => {
        const active = item.id === tab;
        return h('button', {
          key: item.id,
          type: 'button',
          role: 'tab',
          id: 'dso-tab-' + item.id,
          'aria-selected': active ? 'true' : 'false',
          'aria-current': active ? 'page' : undefined,
          'aria-controls': 'dso-pane-' + item.id,
          className: 'dso-tab',
          onClick: () => setTab(item.id),
        }, t(item.label));
      });

      const sessionStatus = sessionId
        ? h('span', { className: 'dso-status dso-status--ok' },
          h('span', { className: 'dso-dot', 'aria-hidden': 'true' }),
          t('panel.session.current') + ' ' + shortId(sessionId))
        : h('span', { className: 'dso-status dso-status--warn' },
          h('span', { className: 'dso-dot', 'aria-hidden': 'true' }),
          t('common.session.missing'));

      return h('div', { className: 'dso-panel' },
        h('header', { className: 'dso-head' },
          h('div', { className: 'dso-headRow' },
            h('div', null,
              h('h1', { className: 'dso-title' }, t('panel.title')),
              h('p', { className: 'dso-sub' }, t('docs.title') + ' · ' + t('visio.title') + ' · ' + t('cad.title') + ' · ' + t('diag.title'))),
            sessionStatus)),
        h('nav', { className: 'dso-tabs', role: 'tablist', 'aria-label': t('panel.tabs') }, tabs),
        h('div', { className: 'dso-body' },
          h('div', {
            className: 'dso-pane',
            id: 'dso-pane-docs',
            role: 'tabpanel',
            'aria-labelledby': 'dso-tab-docs',
            hidden: tab !== 'docs',
          }, h(DocsTab, { api, t, sessionId, active: tab === 'docs', openDir })),
          h('div', {
            className: 'dso-pane',
            id: 'dso-pane-visio',
            role: 'tabpanel',
            'aria-labelledby': 'dso-tab-visio',
            hidden: tab !== 'visio',
          }, h(VisioTab, { api, t, sessionId, active: tab === 'visio' })),
          h('div', {
            className: 'dso-pane',
            id: 'dso-pane-cad',
            role: 'tabpanel',
            'aria-labelledby': 'dso-tab-cad',
            hidden: tab !== 'cad',
          }, h(CadTab, { api, t, sessionId, active: tab === 'cad', openInDocs })),
          h('div', {
            className: 'dso-pane',
            id: 'dso-pane-diag',
            role: 'tabpanel',
            'aria-labelledby': 'dso-tab-diag',
            hidden: tab !== 'diag',
          }, h(DiagTab, { api, t, sessionId, active: tab === 'diag' }))));
    }
    //#endregion

    //#region 7. 组件：DocsTab（文档中心：工作区文件树 / 外部目录双模式 + 读取渲染）
    /** 文档中心：左树右阅读区。 */
    function DocsTab({ api, t, sessionId, active, openDir }) {
      const [mode, setMode] = React.useState('workspace');
      const [rootPath, setRootPath] = React.useState('');
      const [showAll, setShowAll] = React.useState(false);
      const [expanded, setExpanded] = React.useState({});
      const [tree, setTree] = React.useState({});
      const [loading, setLoading] = React.useState({});
      const [treeErrors, setTreeErrors] = React.useState({});
      const [selected, setSelected] = React.useState(null);
      const [reading, setReading] = React.useState(false);
      const [readResult, setReadResult] = React.useState(null);
      const [readError, setReadError] = React.useState(null);
      const [meta, setMeta] = React.useState(null);
      const [copied, setCopied] = React.useState(false);

      const modeRef = React.useRef(mode);
      modeRef.current = mode;
      const treeRef = React.useRef(tree);
      const expandedRef = React.useRef(expanded);
      expandedRef.current = expanded;
      const busyRef = React.useRef({});
      const abortRef = React.useRef(null);
      const tRef = React.useRef(t);
      tRef.current = t;

      React.useEffect(() => () => {
        if (abortRef.current) abortRef.current.abort();
      }, []);

      const putTree = React.useCallback((path, node) => {
        const next = Object.assign({}, treeRef.current, { [path]: node });
        treeRef.current = next;
        setTree(next);
      }, []);

      const loadDir = React.useCallback(async (path, force) => {
        if (busyRef.current[path]) return;
        if (!force && treeRef.current[path]) return;
        busyRef.current[path] = true;
        setLoading((prev) => Object.assign({}, prev, { [path]: true }));
        setTreeErrors((prev) => {
          if (!prev[path]) return prev;
          const next = Object.assign({}, prev);
          delete next[path];
          return next;
        });
        try {
          let node;
          if (modeRef.current === 'workspace') {
            const listing = await api.listWorkspace(path);
            const base = listing && typeof listing.path === 'string' ? listing.path : path;
            node = {
              rows: (Array.isArray(listing && listing.entries) ? listing.entries : []).map((raw) => toFsEntry(raw, 'workspace', base)),
              truncated: !!(listing && listing.truncated),
              dirsOnly: false,
              base,
              crumbs: [],
            };
          } else {
            const listing = await api.externalList(path);
            node = {
              rows: listing.entries,
              truncated: listing.truncated,
              dirsOnly: listing.dirsOnly,
              base: listing.base,
              crumbs: listing.crumbs || [],
            };
          }
          node.rows.sort(compareEntries);
          putTree(path, node);
        } catch (error) {
          const normalized = toPanelError(error, tRef.current);
          setTreeErrors((prev) => Object.assign({}, prev, { [path]: normalized.message }));
        } finally {
          delete busyRef.current[path];
          setLoading((prev) => {
            const next = Object.assign({}, prev);
            delete next[path];
            return next;
          });
        }
      }, [api, putTree]);

      // 打开页签 / 切模式 / 换目录 / 换会话时加载根层。
      React.useEffect(() => {
        if (!active) return;
        if (mode === 'workspace' && !sessionId) return;
        const target = mode === 'workspace' ? '' : rootPath;
        if (mode === 'external' && !target) return;
        if (!expandedRef.current[target]) {
          const next = Object.assign({}, expandedRef.current, { [target]: true });
          expandedRef.current = next;
          setExpanded(next);
        }
        loadDir(target, false);
      }, [active, mode, rootPath, sessionId, loadDir]);

      // 会话切换：工作区根变了，清空目录缓存后重新加载（外部目录模式不受影响）。
      const lastSessionRef = React.useRef(sessionId);
      React.useEffect(() => {
        if (lastSessionRef.current === sessionId) return;
        lastSessionRef.current = sessionId;
        treeRef.current = {};
        setTree({});
        setExpanded({});
        expandedRef.current = {};
        if (sessionId && modeRef.current === 'workspace') loadDir('', true);
      }, [sessionId, loadDir]);

      // CadTab 的「在文档中心查看」：切到外部目录模式并定位到输出目录。
      React.useEffect(() => {
        if (!openDir || !openDir.path) return;
        setMode('external');
        setRootPath(openDir.path);
        setSelected(null);
        setReadResult(null);
        setReadError(null);
        setMeta(null);
      }, [openDir]);

      const selectRow = React.useCallback((row) => {
        setSelected(row);
        setReadResult(null);
        setReadError(null);
        setCopied(false);
        setMeta(null);
        if (row && row.type === 'file' && META_EXTS.includes(extOf(row.name))) {
          const requested = row;
          const info = { path: requested.path, size: requested.size, version: undefined, absolutePath: undefined };
          setMeta(info);
          if (modeRef.current === 'workspace') {
            api.statWorkspace(requested.path).then((stat) => {
              if (!stat) return;
              setMeta((prev) => {
                if (!prev || prev.path !== requested.path) return prev;
                return {
                  path: prev.path,
                  size: typeof stat.bytes === 'number' ? stat.bytes : prev.size,
                  version: typeof stat.version === 'string' ? stat.version : undefined,
                  absolutePath: typeof stat.absolutePath === 'string' ? stat.absolutePath : undefined,
                };
              });
            }).catch(() => { /* 元信息失败不影响选中 */ });
          }
        }
      }, [api]);

      const toggleDir = React.useCallback((row) => {
        const open = !expandedRef.current[row.path];
        const next = Object.assign({}, expandedRef.current, { [row.path]: open });
        expandedRef.current = next;
        setExpanded(next);
        if (open) loadDir(row.path, false);
      }, [loadDir]);

      const browseOther = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          setMode('external');
          setRootPath(dir);
          setSelected(null);
          setReadResult(null);
          setReadError(null);
          setMeta(null);
        } catch (error) {
          setReadError(toPanelError(error, t));
        }
      }, [api, t]);

      const backToWorkspace = React.useCallback(() => {
        setMode('workspace');
        setRootPath('');
        setSelected(null);
        setReadResult(null);
        setReadError(null);
        setMeta(null);
      }, []);

      const goUp = React.useCallback(() => {
        if (mode !== 'external' || !rootPath) return;
        const parent = parentPath(rootPath);
        if (!parent || parent === rootPath) return;
        setRootPath(parent);
        setSelected(null);
        setReadResult(null);
        setMeta(null);
      }, [mode, rootPath]);

      const refresh = React.useCallback(() => {
        const target = mode === 'workspace' ? '' : rootPath;
        loadDir(target, true);
        Object.keys(expanded).forEach((path) => { if (expanded[path]) loadDir(path, true); });
        if (selected) selectRow(selected);
      }, [mode, rootPath, expanded, selected, loadDir, selectRow]);

      const readFile = React.useCallback(async () => {
        if (!selected || selected.type !== 'file') return;
        if (!sessionId) {
          setReadError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        if (abortRef.current) abortRef.current.abort();
        const controller = new AbortController();
        abortRef.current = controller;
        setReading(true);
        setReadError(null);
        setReadResult(null);
        try {
          const result = await api.callCommand('suite.read', { path: selected.path, mode: 'full' }, controller.signal);
          setReadResult({ path: selected.path, result });
        } catch (error) {
          setReadError(toPanelError(error, t));
        } finally {
          setReading(false);
          abortRef.current = null;
        }
      }, [api, selected, sessionId, t]);

      const copyPath = React.useCallback(async () => {
        if (!selected) return;
        const ok = await api.copy(selected.path);
        setCopied(ok);
        if (ok) window.setTimeout(() => setCopied(false), 1600);
      }, [api, selected]);

      const selectedExt = selected && selected.type === 'file' ? extOf(selected.name) : '';

      // ── 左树 ──
      const rows = [];
      if (mode === 'workspace') {
        if (!sessionId) {
          rows.push(h('div', { key: 'no-session', className: 'dso-treeNote' }, t('docs.tree.noSession')));
        } else {
          collectRows('', 0, rows, { tree, loading, treeErrors, expanded, selected, showAll, t, selectRow, toggleDir });
        }
      } else if (!rootPath) {
        rows.push(h('div', { key: 'no-root', className: 'dso-treeNote' }, t('docs.tree.empty')));
      } else {
        collectRows(rootPath, 0, rows, { tree, loading, treeErrors, expanded, selected, showAll, t, selectRow, toggleDir });
      }

      const rootNode = tree[mode === 'workspace' ? '' : rootPath];
      const dirsOnly = mode === 'external' && rootNode && rootNode.dirsOnly;

      const treeCard = h('section', { className: 'dso-card' },
        h('div', { className: 'dso-cardHead' },
          h('h2', { className: 'dso-cardTitle' }, mode === 'workspace' ? t('docs.tree.root') : t('docs.tree.external')),
          h('span', { className: 'dso-spacer' }),
          h('button', {
            type: 'button',
            className: 'dso-btn dso-btn--sm',
            onClick: refresh,
            disabled: mode === 'workspace' ? !sessionId : !rootPath,
            title: t('docs.actions.refresh'),
          }, t('docs.actions.refresh'))),
        h('div', { className: 'dso-toolbar', style: { padding: '6px 10px', borderBottom: '.5px solid var(--dsw-alias-border-l1)' } },
          mode === 'workspace'
            ? h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: browseOther }, t('docs.tree.browse'))
            : h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: backToWorkspace }, t('docs.tree.back')),
          mode === 'external'
            ? h('button', {
              type: 'button',
              className: 'dso-btn dso-btn--sm',
              onClick: goUp,
              disabled: !parentPath(rootPath),
              title: t('docs.tree.up'),
            }, '↑ ' + t('docs.tree.up'))
            : null,
          h('label', { className: 'dso-check' },
            h('input', { type: 'checkbox', checked: showAll, onChange: (event) => setShowAll(event.target.checked) }),
            t('docs.tree.filterAll'))),
        showAll ? null : h('div', { className: 'dso-treeNote' }, t('docs.tree.filterHint')),
        mode === 'external' && rootPath
          ? h('div', { className: 'dso-crumbs', style: { padding: '4px 10px' } }, h('span', { className: 'dso-mono' }, rootPath))
          : null,
        dirsOnly ? h('div', { className: 'dso-treeNote' }, t('docs.tree.dirsOnly')) : null,
        rootNode && rootNode.truncated ? h('div', { className: 'dso-treeNote' }, t('docs.tree.truncated')) : null,
        h('div', { className: 'dso-tree', role: 'tree', 'aria-label': t('docs.tree.root') }, rows));

      // ── 右阅读区 ──
      let body = null;
      if (!selected) {
        body = h(Empty, { text: t('docs.read.empty') });
      } else if (selected.type !== 'file') {
        body = h(Empty, { text: t('docs.read.notSupported') });
      } else if (OFFICE_EXTS.includes(selectedExt)) {
        if (readError) {
          body = h('div', { className: 'dso-stack' }, h(ErrorBox, { t, error: readError, onRetry: readFile }));
        } else if (reading) {
          body = h('div', { className: 'dso-cardBody' }, h(Busy, { text: t('docs.reading') }));
        } else if (readResult && readResult.path === selected.path) {
          body = h('div', { className: 'dso-cardBody' }, renderOfficeResult(t, readResult.result));
        } else {
          body = h(Empty, { text: t('docs.read.prompt') });
        }
      } else if (META_EXTS.includes(selectedExt)) {
        const hint = selectedExt === 'pdf' ? t('docs.preview.hint.pdf') : t('docs.preview.hint.other');
        const info = meta && meta.path === selected.path ? meta : null;
        body = h('div', { className: 'dso-cardBody', 'data-preview-hook': selectedExt === 'pdf' ? 'officeToPdf' : undefined },
          h('dl', { className: 'dso-kv' },
            h('dt', null, t('docs.meta.name')),
            h('dd', null, selected.name),
            h('dt', null, t('docs.meta.size')),
            h('dd', null, info && typeof info.size === 'number' ? formatBytes(info.size) : '—'),
            h('dt', null, t('docs.meta.version')),
            h('dd', null, info && info.version ? info.version : '—'),
            h('dt', null, t('docs.meta.path')),
            h('dd', { className: 'dso-mono' }, info && info.absolutePath ? info.absolutePath : selected.path)),
          h('div', { className: 'dso-hint' }, hint));
      } else {
        body = h(Empty, { text: t('docs.read.notSupported') });
      }

      const readerCard = h('section', { className: 'dso-card' },
        h('div', { className: 'dso-cardHead' },
          h('h2', { className: 'dso-cardTitle' }, selected ? selected.name : t('docs.file.none')),
          h('span', { className: 'dso-spacer' }),
          h('div', { className: 'dso-actions' },
            selected && selected.type === 'file' && OFFICE_EXTS.includes(selectedExt)
              ? h('button', {
                type: 'button',
                className: 'dso-btn dso-btn--primary',
                onClick: readFile,
                disabled: reading || !sessionId,
                title: sessionId ? t('docs.read') : t('common.session.missing'),
              }, reading ? t('docs.reading') : t('docs.read'))
              : null,
            h('button', {
              type: 'button',
              className: 'dso-btn',
              onClick: copyPath,
              disabled: !selected,
            }, copied ? t('docs.actions.copied') : t('docs.actions.copyPath')))),
        h('div', { className: 'dso-cardBody' }, body));

      return h('div', { className: 'dso-split' }, treeCard, readerCard);
    }

    /** 递归收集文件树行（扁平数组 + 唯一 key，避免嵌套数组与 hooks 限制）。 */
    function collectRows(path, depth, out, state) {
      const node = state.tree[path];
      if (!node) {
        if (state.loading[path]) out.push(h('div', { key: 'l:' + path, className: 'dso-treeNote' }, state.t('docs.tree.loading')));
        else if (state.treeErrors[path]) out.push(h('div', { key: 'e:' + path, className: 'dso-treeNote dso-error' }, state.treeErrors[path]));
        return;
      }
      const visible = node.rows.filter((row) => row.type !== 'file' || state.showAll || DEFAULT_EXTS.includes(extOf(row.name)));
      if (visible.length === 0) {
        out.push(h('div', { key: '0:' + path, className: 'dso-treeNote' }, state.t('docs.tree.empty')));
        return;
      }
      for (const row of visible) {
        const isDir = row.type === 'dir';
        const isOpen = !!state.expanded[row.path];
        const isSelected = !!state.selected && state.selected.path === row.path;
        out.push(h('div', {
          key: row.path,
          className: 'dso-row',
          role: 'treeitem',
          tabIndex: 0,
          'aria-selected': isSelected ? 'true' : 'false',
          'aria-expanded': isDir ? (isOpen ? 'true' : 'false') : undefined,
          title: row.path,
          style: { paddingLeft: (6 + depth * 12) + 'px' },
          onClick: () => state.selectRow(row),
          onDoubleClick: isDir ? () => state.toggleDir(row) : undefined,
          onKeyDown: (event) => {
            if (event.key === 'Enter' || event.key === ' ') {
              event.preventDefault();
              state.selectRow(row);
            } else if (event.key === 'ArrowRight' && isDir && !isOpen) {
              state.toggleDir(row);
            } else if (event.key === 'ArrowLeft' && isDir && isOpen) {
              state.toggleDir(row);
            }
          },
        },
        isDir
          ? h('span', {
            className: 'dso-caret',
            'aria-hidden': 'true',
            onClick: (event) => { event.stopPropagation(); state.toggleDir(row); },
          }, isOpen ? '▾' : '▸')
          : h('span', { className: 'dso-caret', 'aria-hidden': 'true' }, '·'),
        h('span', { className: 'dso-rowName' }, row.name),
        row.type === 'file' && typeof row.size === 'number'
          ? h('span', { className: 'dso-rowMeta' }, formatBytes(row.size))
          : null));
        if (isDir && isOpen) collectRows(row.path, depth + 1, out, state);
      }
    }

    /** 渲染 suite.read（office_read 同款）结构：Word 大纲/预览，Excel 工作表清单 + 样例。 */
    function renderOfficeResult(t, payload) {
      const info = payload && payload.info ? payload.info : payload;
      if (!info || typeof info !== 'object') {
        return h('pre', { className: 'dso-pre' }, JSON.stringify(payload, null, 2));
      }
      if (info.kind === 'docx') {
        const headings = Array.isArray(info.headings) ? info.headings : [];
        const preview = Array.isArray(info.preview) ? info.preview : [];
        return h('div', { className: 'dso-stack' },
          h('dl', { className: 'dso-kv' },
            h('dt', null, t('docs.word.paragraphs')), h('dd', null, String(info.paragraphs ?? '—')),
            h('dt', null, t('docs.word.tables')), h('dd', null, String(info.tables ?? '—')),
            h('dt', null, t('docs.word.headings')), h('dd', null, String(headings.length))),
          h('div', { className: 'dso-divider' }),
          h('div', { className: 'dso-label' }, t('docs.word.headings')),
          headings.length === 0
            ? h('div', { className: 'dso-hint' }, t('docs.tree.empty'))
            : h('div', { className: 'dso-stack', style: { gap: '2px' } }, headings.map((row, index) => {
              const style = String(row && row.style ? row.style : '');
              const match = style.match(/(\d+)/);
              const level = match ? Math.min(3, Math.max(1, Number(match[1]))) : 1;
              return h('div', { key: 'h' + index, className: 'dso-h' + level }, String(row && row.text !== undefined ? row.text : ''));
            })),
          h('div', { className: 'dso-divider' }),
          h('div', { className: 'dso-label' }, t('docs.word.preview')),
          preview.length === 0
            ? h('div', { className: 'dso-hint' }, t('docs.tree.empty'))
            : h('div', { className: 'dso-stack', style: { gap: '2px' } }, preview.map((line, index) =>
              h('div', { key: 'p' + index, className: 'dso-muted' }, String(line)))),
          h('div', { className: 'dso-label' }, t('docs.preview.hint.other')));
      }
      if (info.kind === 'xlsx') {
        const sheets = Array.isArray(info.sheets) ? info.sheets : [];
        return h('div', { className: 'dso-stack' },
          h('dl', { className: 'dso-kv' },
            h('dt', null, t('docs.excel.sheets')), h('dd', null, String(info.sheetCount ?? sheets.length))),
          sheets.length === 0 ? h('div', { className: 'dso-hint' }, t('docs.tree.empty')) : null,
          sheets.map((sheet, index) => {
            const sample = Array.isArray(sheet && sheet.sample) ? sheet.sample : [];
            return h('div', { key: 's' + index, className: 'dso-stack', style: { gap: '4px' } },
              h('div', { className: 'dso-inline' },
                h('span', { className: 'dso-badge' }, t('docs.excel.sheet')),
                h('span', { style: { fontWeight: 600 } }, String(sheet && sheet.name !== undefined ? sheet.name : '')),
                h('span', { className: 'dso-faint' },
                  t('docs.excel.rows') + ' ' + String(sheet && sheet.rows !== undefined ? sheet.rows : '—') + ' · ' +
                  t('docs.excel.cols') + ' ' + String(sheet && sheet.cols !== undefined ? sheet.cols : '—'))),
              h('div', { className: 'dso-label' }, t('docs.excel.sample')),
              sample.length === 0
                ? h('div', { className: 'dso-hint' }, t('docs.tree.empty'))
                : h('table', { className: 'dso-table' },
                  h('tbody', null, sample.map((line, rowIndex) => {
                    const cells = String(line).split(' | ');
                    return h('tr', { key: 'r' + rowIndex }, cells.map((cell, cellIndex) =>
                      h('td', { key: 'c' + cellIndex }, cell)));
                  }))));
          }),
          h('div', { className: 'dso-label' }, t('docs.preview.hint.other')));
      }
      return h('pre', { className: 'dso-pre' }, JSON.stringify(info, null, 2));
    }
    //#endregion

    //#region 8. 组件：VisioTab（Visio 绘图：flowchart JSON 编辑 + 校验 + 运行 + 历史）
    /** flowchart 模板（与 host/engines/visio-flowchart.ps1 的节点类型一致）。 */
    const FLOWCHART_TEMPLATE_NODES = [
      { id: 'n1', label: '开始', kind: 'start' },
      { id: 'n2', label: '读取图纸', kind: 'process' },
      { id: 'n3', label: '校验是否通过？', kind: 'decision' },
      { id: 'n4', label: '生成报告', kind: 'process' },
      { id: 'n5', label: '结束', kind: 'end' },
    ];
    const FLOWCHART_TEMPLATE_EDGES = [
      { from: 'n1', to: 'n2' },
      { from: 'n2', to: 'n3' },
      { from: 'n3', to: 'n4' },
      { from: 'n4', to: 'n5' },
    ];
    /** 允许的节点类型（§1.3）。 */
    const NODE_KINDS = ['start', 'process', 'decision', 'end'];

    /** 解析并校验节点/连线 JSON；错误信息已本地化，带 field 以便落到对应输入框下方。 */
    function parseFlowchart(t, nodesText, edgesText) {
      let nodes;
      try {
        nodes = JSON.parse(nodesText);
      } catch (error) {
        return { ok: false, field: 'nodes', message: fmt(t('visio.json.invalid'), { message: error.message }) };
      }
      if (!Array.isArray(nodes)) {
        return { ok: false, field: 'nodes', message: fmt(t('visio.json.invalid'), { message: 'array required' }) };
      }
      const ids = {};
      for (let index = 0; index < nodes.length; index += 1) {
        const raw = nodes[index];
        const id = raw && typeof raw.id === 'string' ? raw.id.trim() : '';
        const label = raw && typeof raw.label === 'string' ? raw.label : '';
        const kind = raw && typeof raw.kind === 'string' ? raw.kind : 'process';
        if (!id || !label) return { ok: false, field: 'nodes', message: fmt(t('visio.node.invalid'), { index: index + 1 }) };
        if (!NODE_KINDS.includes(kind)) return { ok: false, field: 'nodes', message: fmt(t('visio.node.kindInvalid'), { index: index + 1 }) };
        if (ids[id]) return { ok: false, field: 'nodes', message: fmt(t('visio.node.duplicate'), { id }) };
        ids[id] = true;
        nodes[index] = { id, label, kind };
      }
      let edges;
      try {
        edges = JSON.parse(edgesText);
      } catch (error) {
        return { ok: false, field: 'edges', message: fmt(t('visio.json.invalid'), { message: error.message }) };
      }
      if (!Array.isArray(edges)) {
        return { ok: false, field: 'edges', message: fmt(t('visio.json.invalid'), { message: 'array required' }) };
      }
      for (let index = 0; index < edges.length; index += 1) {
        const raw = edges[index];
        const from = raw && typeof raw.from === 'string' ? raw.from.trim() : '';
        const to = raw && typeof raw.to === 'string' ? raw.to.trim() : '';
        if (!from || !to) return { ok: false, field: 'edges', message: fmt(t('visio.edge.invalid'), { index: index + 1 }) };
        if (!ids[from]) return { ok: false, field: 'edges', message: fmt(t('visio.edge.missing'), { id: from }) };
        if (!ids[to]) return { ok: false, field: 'edges', message: fmt(t('visio.edge.missing'), { id: to }) };
        edges[index] = { from, to };
      }
      return { ok: true, nodes, edges };
    }

    /** Visio 绘图页。 */
    function VisioTab({ api, t, sessionId }) {
      const [title, setTitle] = React.useState('');
      const [nodesText, setNodesText] = React.useState(() => JSON.stringify(FLOWCHART_TEMPLATE_NODES, null, 2));
      const [edgesText, setEdgesText] = React.useState(() => JSON.stringify(FLOWCHART_TEMPLATE_EDGES, null, 2));
      const [nlText, setNlText] = React.useState('');
      const [outputPath, setOutputPath] = React.useState('');
      const [running, setRunning] = React.useState(false);
      const [result, setResult] = React.useState(null);
      const [error, setError] = React.useState(null);
      const [history, setHistory] = React.useState([]);
      const abortRef = React.useRef(null);

      React.useEffect(() => () => {
        if (abortRef.current) abortRef.current.abort();
      }, []);

      const validation = React.useMemo(() => parseFlowchart(t, nodesText, edgesText), [t, nodesText, edgesText]);

      const pushHistory = React.useCallback((entry) => {
        setHistory((prev) => [Object.assign({ at: Date.now() }, entry)].concat(prev).slice(0, 5));
      }, []);

      const insertTemplate = React.useCallback(() => {
        setNodesText(JSON.stringify(FLOWCHART_TEMPLATE_NODES, null, 2));
        setEdgesText(JSON.stringify(FLOWCHART_TEMPLATE_EDGES, null, 2));
      }, []);

      const chooseDir = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          setOutputPath(joinPath(dir, 'flowchart-' + formatStamp(Date.now()) + '.vsdx'));
        } catch (picked) {
          setError(toPanelError(picked, t));
        }
      }, [api, t]);

      const run = React.useCallback(async () => {
        if (running) return;
        if (!sessionId) {
          setError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        if (!validation.ok) return;
        if (abortRef.current) abortRef.current.abort();
        const controller = new AbortController();
        abortRef.current = controller;
        setRunning(true);
        setError(null);
        setResult(null);
        try {
          const args = { kind: 'flowchart', title: title.trim(), nodes: validation.nodes, edges: validation.edges };
          const trimmed = outputPath.trim();
          if (trimmed) args.outputPath = trimmed;
          const drawn = await api.callCommand('suite.draw', args, controller.signal);
          setResult(drawn);
          pushHistory({ ok: true, outputPath: drawn && drawn.outputPath ? String(drawn.outputPath) : '', stats: drawn && drawn.stats });
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) {
            setError(normalized);
            pushHistory({ ok: false, code: normalized.code, message: normalized.message });
          }
        } finally {
          setRunning(false);
          abortRef.current = null;
        }
      }, [api, outputPath, pushHistory, running, sessionId, t, title, validation]);

      const abort = React.useCallback(() => {
        if (abortRef.current) abortRef.current.abort();
      }, []);

      const field = (label, node) => h('div', { className: 'dso-field' }, h('span', { className: 'dso-label' }, label), node);

      const historyItems = history.map((entry, index) => {
        const detail = entry.ok
          ? (entry.outputPath || '—') + (entry.stats
            ? ' · ' + fmt(t('visio.done.stats'), { nodes: entry.stats.nodes, edges: entry.stats.edges })
            : '')
          : (entry.message || '—');
        return h('li', { key: 'h' + index, className: 'dso-inline' },
          h('span', { className: 'dso-faint dso-mono' }, formatClock(entry.at)),
          h('span', { className: entry.ok ? 'dso-badge dso-badge--ok' : 'dso-badge dso-badge--error' },
            entry.ok ? t('cad.state.done') : t('cad.state.failed')),
          h('span', { className: 'dso-mono', style: { overflowWrap: 'anywhere' } }, detail));
      });

      return h('div', { className: 'dso-stack' },
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('visio.title'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-row2' },
              field(t('visio.kind'),
                h('select', { className: 'dso-select', value: 'flowchart', disabled: true, onChange: () => {} },
                  h('option', { value: 'flowchart' }, t('visio.kind.flowchart')))),
              field(t('visio.titleField'),
                h('input', {
                  className: 'dso-input',
                  type: 'text',
                  value: title,
                  placeholder: t('panel.title'),
                  onChange: (event) => setTitle(event.target.value),
                }))),
            field(t('visio.nodes'),
              h('textarea', {
                className: 'dso-textarea',
                value: nodesText,
                spellCheck: false,
                onChange: (event) => setNodesText(event.target.value),
              })),
            validation.ok || validation.field !== 'nodes'
              ? null
              : h('div', { className: 'dso-error', role: 'alert' }, validation.message),
            field(t('visio.edges'),
              h('textarea', {
                className: 'dso-textarea',
                value: edgesText,
                spellCheck: false,
                style: { minHeight: '104px' },
                onChange: (event) => setEdgesText(event.target.value),
              })),
            validation.ok || validation.field !== 'edges'
              ? null
              : h('div', { className: 'dso-error', role: 'alert' }, validation.message),
            h('div', { className: 'dso-actions' },
              h('button', { type: 'button', className: 'dso-btn', onClick: insertTemplate }, t('visio.template')),
              h('span', { className: validation.ok ? 'dso-badge dso-badge--ok' : 'dso-badge dso-badge--error' },
                validation.ok ? t('visio.validate.ok') : t('visio.json.invalid'))),
            field(t('visio.nl'),
              h('input', {
                className: 'dso-input',
                type: 'text',
                value: nlText,
                disabled: true,
                placeholder: t('visio.nl'),
                onChange: (event) => setNlText(event.target.value),
              })),
            field(t('visio.outputPath'),
              h('div', { className: 'dso-inline' },
                h('input', {
                  className: 'dso-input',
                  type: 'text',
                  value: outputPath,
                  placeholder: t('visio.outputHint'),
                  style: { flex: '1 1 260px', width: 'auto' },
                  onChange: (event) => setOutputPath(event.target.value),
                }),
                h('button', { type: 'button', className: 'dso-btn', onClick: chooseDir, disabled: running }, t('visio.chooseDir')))),
            h('div', { className: 'dso-hint' }, t('visio.outputHint')),
            h('div', { className: 'dso-actions' },
              h('button', {
                type: 'button',
                className: 'dso-btn dso-btn--primary',
                onClick: run,
                disabled: running || !validation.ok || !sessionId,
                title: sessionId ? t('visio.run') : t('common.session.missing'),
              }, running ? h('span', { className: 'dso-inline' }, h('span', { className: 'dso-spinner', 'aria-hidden': 'true' }), t('visio.running')) : t('visio.run')),
              running ? h('button', { type: 'button', className: 'dso-btn dso-btn--danger', onClick: abort }, t('common.cancel')) : null))),
        error
          ? h('div', { className: 'dso-card' }, h('div', { className: 'dso-cardBody' }, h(ErrorBox, { t, error })))
          : null,
        result
          ? h('div', { className: 'dso-card' },
            h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('visio.result'))),
            h('div', { className: 'dso-cardBody' },
              h('div', { className: 'dso-okBox' },
                h('div', null, fmt(t('visio.done.saved'), { path: result.outputPath || '—' })),
                result.stats
                  ? h('div', { className: 'dso-hint' }, fmt(t('visio.done.stats'), { nodes: result.stats.nodes, edges: result.stats.edges }))
                  : null)))
          : null,
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('visio.history'))),
          h('div', { className: 'dso-cardBody' },
            historyItems.length === 0
              ? h('div', { className: 'dso-hint' }, t('visio.history.empty'))
              : h('ul', { className: 'dso-list' }, historyItems))));
    }
    //#endregion

    //#region 9. 组件：CadTab（图纸转换：场景 + 参数 + 任务进度轮询 + 取消）
    /** 已知任务状态（未知状态原样显示）。 */
    const CAD_STATES = ['idle', 'queued', 'running', 'converting', 'splitting', 'merging', 'done', 'failed', 'cancelled'];

    /** 任务状态徽标样式。 */
    function cadStateClass(state) {
      if (state === 'done') return 'dso-badge dso-badge--ok';
      if (state === 'failed') return 'dso-badge dso-badge--error';
      if (state === 'cancelled' || state === 'idle') return 'dso-badge dso-badge--idle';
      if (state === 'queued') return 'dso-badge dso-badge--warn';
      return 'dso-badge';
    }

    /** 图纸转换页。 */
    function CadTab({ api, t, sessionId, openInDocs }) {
      const [mode, setMode] = React.useState('merge');
      const [inputPath, setInputPath] = React.useState('');
      const [paper, setPaper] = React.useState('A4-landscape');
      const [sortMode, setSortMode] = React.useState('drawingNo');
      const [parallel, setParallel] = React.useState('auto');
      const [starting, setStarting] = React.useState(false);
      const [taskId, setTaskId] = React.useState('');
      const [status, setStatus] = React.useState(null);
      const [startError, setStartError] = React.useState(null);
      const [pollError, setPollError] = React.useState(null);
      const [cancelling, setCancelling] = React.useState(false);
      const [eventDriven, setEventDriven] = React.useState(false);
      const [count, setCount] = React.useState(null);
      const [counting, setCounting] = React.useState(false);
      const [picker, setPicker] = React.useState(null);
      const logRef = React.useRef(null);
      const startAbortRef = React.useRef(null);
      const tRef = React.useRef(t);
      tRef.current = t;

      React.useEffect(() => () => {
        if (startAbortRef.current) startAbortRef.current.abort();
      }, []);

      // 输入计数预览（合并场景 = 目录内图纸数；分割场景 = 单文件）。
      React.useEffect(() => {
        let alive = true;
        const target = inputPath.trim();
        setCount(null);
        if (!target) { setCounting(false); return () => { alive = false; }; }
        if (mode === 'split') {
          setCount(1);
          setCounting(false);
          return () => { alive = false; };
        }
        setCounting(true);
        api.externalList(target).then((listing) => {
          if (!alive) return;
          if (listing.dirsOnly) {
            setCount(null);
            setCounting(false);
            return;
          }
          const files = (listing.entries || []).filter((row) => row.type === 'file' && CAD_INPUT_EXTS.includes(extOf(row.name)));
          setCount(files.length);
          setCounting(false);
        }).catch(() => {
          if (!alive) return;
          setCount(null);
          setCounting(false);
        });
        return () => { alive = false; };
      }, [api, inputPath, mode]);

      // 轮询 suite.cad.status（800ms）；终态停止；收到进度事件后停轮询（§0.4）。
      React.useEffect(() => {
        if (!taskId || eventDriven) return undefined;
        let alive = true;
        let timer = 0;
        let inFlight = false;
        const stop = () => {
          if (timer) {
            window.clearInterval(timer);
            timer = 0;
          }
        };
        const tick = async () => {
          if (inFlight || !alive) return;
          inFlight = true;
          try {
            const snapshot = await api.callCommand('suite.cad.status', { taskId });
            if (!alive) return;
            setStatus(snapshot);
            setPollError(null);
            const state = snapshot && typeof snapshot.state === 'string' ? snapshot.state : '';
            if (CAD_TERMINAL.includes(state)) stop();
          } catch (error) {
            if (!alive) return;
            const normalized = toPanelError(error, tRef.current);
            setPollError(normalized);
            if (normalized.code === 'TASK_NOT_FOUND' || normalized.cancelled) stop();
          } finally {
            inFlight = false;
          }
        };
        tick();
        timer = window.setInterval(tick, 800);
        return () => {
          alive = false;
          stop();
        };
      }, [api, eventDriven, taskId]);

      // 事件加速：Host 注册了 office-suite/task-progress 时首帧即停轮询。
      React.useEffect(() => {
        if (!taskId) return undefined;
        let dispose = () => {};
        try {
          dispose = api.subscribeProgress((frame) => {
            if (!frame || typeof frame !== 'object') return;
            setEventDriven(true);
            setStatus((prev) => {
              if (!prev) return prev;
              if (prev.taskId && frame.taskId && prev.taskId !== frame.taskId) return prev;
              const next = Object.assign({}, prev);
              if (frame.state) next.state = frame.state;
              if (frame.progress) next.progress = Object.assign({}, prev.progress || {}, frame.progress);
              if (Array.isArray(frame.log)) next.log = frame.log;
              if (frame.output) next.output = frame.output;
              if (frame.error) next.error = frame.error;
              return next;
            });
          }) || (() => {});
        } catch { dispose = () => {}; }
        return () => { try { dispose(); } catch { /* 忽略 */ } };
      }, [api, taskId]);

      // 日志自动滚到底部。
      React.useEffect(() => {
        const element = logRef.current;
        if (element) element.scrollTop = element.scrollHeight;
      }, [status]);

      const pickInputDir = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          setInputPath(dir);
          setStartError(null);
        } catch (error) {
          setStartError(toPanelError(error, t));
        }
      }, [api, t]);

      const loadPicker = React.useCallback(async (path) => {
        setPicker((prev) => ({ path, entries: (prev && prev.entries) || [], busy: true, error: null, dirsOnly: false }));
        try {
          const listing = await api.externalList(path);
          setPicker({
            path: listing.base || path,
            entries: listing.entries,
            busy: false,
            error: null,
            dirsOnly: listing.dirsOnly,
          });
        } catch (error) {
          setPicker((prev) => Object.assign({}, prev || { path, entries: [] }, { busy: false, error: toPanelError(error, t) }));
        }
      }, [api, t]);

      const openPicker = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          await loadPicker(dir);
        } catch (error) {
          setStartError(toPanelError(error, t));
        }
      }, [api, loadPicker, t]);

      const start = React.useCallback(async () => {
        if (starting) return;
        if (!sessionId) {
          setStartError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        const target = inputPath.trim();
        if (!target) {
          setStartError({ code: 'PARAM_INVALID', message: t('cad.input.placeholder') });
          return;
        }
        if (startAbortRef.current) startAbortRef.current.abort();
        const controller = new AbortController();
        startAbortRef.current = controller;
        setStarting(true);
        setStartError(null);
        setPollError(null);
        setStatus(null);
        setTaskId('');
        setEventDriven(false);
        try {
          const args = {
            mode,
            input: target,
            paper,
            sort: sortMode,
            parallel: parallel === 'auto' ? 'auto' : Number(parallel),
          };
          const started = await api.callCommand('suite.cad.convert', args, controller.signal);
          const id = started && typeof started.taskId === 'string' ? started.taskId : '';
          if (!id) throw { code: 'BAD_RESPONSE', message: 'suite.cad.convert returned no taskId' };
          setTaskId(id);
        } catch (error) {
          const normalized = toPanelError(error, t);
          setStartError(normalized.cancelled ? null : normalized);
        } finally {
          setStarting(false);
          startAbortRef.current = null;
        }
      }, [api, inputPath, mode, paper, parallel, sessionId, sortMode, starting, t]);

      const cancel = React.useCallback(async () => {
        if (!taskId || cancelling) return;
        setCancelling(true);
        try {
          await api.callCommand('suite.cad.cancel', { taskId });
        } catch (error) {
          setPollError(toPanelError(error, t));
        } finally {
          setCancelling(false);
        }
      }, [api, cancelling, t, taskId]);

      const progress = (status && status.progress) || {};
      const filesDone = Number(progress.filesDone) || 0;
      const filesTotal = Number(progress.filesTotal) || 0;
      const framesDone = Number(progress.framesDone) || 0;
      const framesTotal = Number(progress.framesTotal) || 0;
      const useFrames = framesTotal > 0;
      const doneCount = useFrames ? framesDone : filesDone;
      const totalCount = useFrames ? framesTotal : filesTotal;
      const state = status && typeof status.state === 'string' ? status.state : (taskId ? 'queued' : 'idle');
      const percent = totalCount > 0
        ? Math.min(100, Math.round((doneCount / totalCount) * 100))
        : (state === 'done' ? 100 : 0);
      const logLines = status && Array.isArray(status.log) ? status.log.slice(-200) : [];
      const outputPath = status && typeof status.output === 'string' ? status.output : '';
      const running = !!taskId && !CAD_TERMINAL.includes(state);

      const stateLabel = CAD_STATES.includes(state) ? t('cad.state.' + state) : state;

      const modeButton = (id, label) => h('button', {
        key: id,
        type: 'button',
        className: mode === id ? 'dso-btn dso-btn--primary' : 'dso-btn',
        'aria-pressed': mode === id ? 'true' : 'false',
        onClick: () => { setMode(id); setStartError(null); },
        disabled: running,
      }, label);

      return h('div', { className: 'dso-stack' },
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('cad.title'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-field' },
              h('span', { className: 'dso-label' }, t('cad.mode')),
              h('div', { className: 'dso-actions' },
                modeButton('merge', t('cad.mode.merge')),
                modeButton('split', t('cad.mode.split')))),
            h('div', { className: 'dso-field' },
              h('span', { className: 'dso-label' }, t('cad.input')),
              h('div', { className: 'dso-inline' },
                h('input', {
                  className: 'dso-input',
                  type: 'text',
                  value: inputPath,
                  placeholder: mode === 'merge' ? t('cad.input.dir') : t('cad.input.file'),
                  style: { flex: '1 1 300px', width: 'auto' },
                  onChange: (event) => setInputPath(event.target.value),
                }),
                mode === 'merge'
                  ? h('button', { type: 'button', className: 'dso-btn', onClick: pickInputDir, disabled: running }, t('cad.input.dir'))
                  : h('button', { type: 'button', className: 'dso-btn', onClick: openPicker, disabled: running }, t('cad.input.browse')),
                counting
                  ? h(Busy, { text: t('cad.files.counting') })
                  : (typeof count === 'number'
                    ? h('span', { className: 'dso-badge' }, fmt(t('cad.files.count'), { count }))
                    : null))),
            h('div', { className: 'dso-hint' }, t('cad.input.path') + ' · ' + t('cad.picker.hint')),
            h('div', { className: 'dso-row2' },
              h('div', { className: 'dso-field' },
                h('span', { className: 'dso-label' }, t('cad.paper')),
                h('select', { className: 'dso-select', value: paper, onChange: (event) => setPaper(event.target.value) },
                  h('option', { value: 'A4-landscape' }, t('cad.paper.a4landscape')),
                  h('option', { value: 'keep' }, t('cad.paper.keep')))),
              h('div', { className: 'dso-field' },
                h('span', { className: 'dso-label' }, t('cad.sort')),
                h('select', { className: 'dso-select', value: sortMode, onChange: (event) => setSortMode(event.target.value) },
                  h('option', { value: 'drawingNo' }, t('cad.sort.drawingNo')),
                  h('option', { value: 'filename' }, t('cad.sort.filename')))),
              h('div', { className: 'dso-field' },
                h('span', { className: 'dso-label' }, t('cad.parallel')),
                h('select', { className: 'dso-select', value: parallel, onChange: (event) => setParallel(event.target.value) },
                  h('option', { value: 'auto' }, t('cad.parallel.auto')),
                  [1, 2, 3, 4, 5, 6, 7, 8].map((value) => h('option', { key: 'p' + value, value: String(value) }, String(value)))))),
            h('div', { className: 'dso-actions' },
              h('button', {
                type: 'button',
                className: 'dso-btn dso-btn--primary',
                onClick: start,
                disabled: starting || running || !sessionId,
                title: sessionId ? t('cad.start') : t('common.session.missing'),
              }, starting
                ? h('span', { className: 'dso-inline' }, h('span', { className: 'dso-spinner', 'aria-hidden': 'true' }), t('common.busy'))
                : t('cad.start')),
              running
                ? h('button', { type: 'button', className: 'dso-btn dso-btn--danger', onClick: cancel, disabled: cancelling },
                  cancelling ? t('cad.cancel.requested') : t('cad.cancel'))
                : null))),
        startError ? h('div', { className: 'dso-card' }, h('div', { className: 'dso-cardBody' }, h(ErrorBox, { t, error: startError }))) : null,
        picker
          ? h('div', { className: 'dso-card' },
            h('div', { className: 'dso-cardHead' },
              h('h2', { className: 'dso-cardTitle' }, t('cad.picker.title')),
              h('span', { className: 'dso-spacer' }),
              h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => setPicker(null) }, t('cad.picker.close'))),
            h('div', { className: 'dso-cardBody' },
              h('div', { className: 'dso-inline' },
                h('button', {
                  type: 'button',
                  className: 'dso-btn dso-btn--sm',
                  onClick: () => loadPicker(parentPath(picker.path) || picker.path),
                  disabled: !parentPath(picker.path),
                }, '↑ ' + t('cad.picker.up')),
                h('span', { className: 'dso-mono dso-muted', style: { overflowWrap: 'anywhere' } }, picker.path)),
              picker.dirsOnly ? h('div', { className: 'dso-hint' }, t('docs.tree.dirsOnly')) : null,
              picker.busy ? h(Busy, { text: t('common.busy') }) : null,
              picker.error ? h(ErrorBox, { t, error: picker.error, onRetry: () => loadPicker(picker.path) }) : null,
              picker.busy || picker.error ? null : (picker.entries.length === 0
                ? h('div', { className: 'dso-hint' }, t('cad.picker.empty'))
                : h('div', { className: 'dso-tree', style: { maxHeight: '240px' } }, picker.entries
                  .filter((row) => row.type === 'dir' || CAD_INPUT_EXTS.includes(extOf(row.name)))
                  .map((row) => h('div', {
                    key: row.path,
                    className: 'dso-row',
                    role: 'button',
                    tabIndex: 0,
                    title: row.path,
                    onClick: () => (row.type === 'dir' ? loadPicker(row.path) : (setInputPath(row.path), setPicker(null))),
                    onKeyDown: (event) => {
                      if (event.key === 'Enter') {
                        event.preventDefault();
                        if (row.type === 'dir') loadPicker(row.path);
                        else { setInputPath(row.path); setPicker(null); }
                      }
                    },
                  },
                  h('span', { className: 'dso-caret', 'aria-hidden': 'true' }, row.type === 'dir' ? '▸' : '·'),
                  h('span', { className: 'dso-rowName' }, row.name),
                  row.type === 'file' && typeof row.size === 'number'
                    ? h('span', { className: 'dso-rowMeta' }, formatBytes(row.size))
                    : null))))))
          : null,
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' },
            h('h2', { className: 'dso-cardTitle' }, t('cad.log')),
            h('span', { className: 'dso-spacer' }),
            h('span', { className: cadStateClass(state) }, stateLabel),
            taskId ? h('span', { className: 'dso-faint dso-mono' }, taskId) : null),
          h('div', { className: 'dso-cardBody' },
            pollError ? h(ErrorBox, { t, error: pollError }) : null,
            h('div', { className: 'dso-inline' },
              h('span', { className: 'dso-muted' }, fmt(t('cad.progress.files'), { done: filesDone, total: filesTotal })),
              h('span', { className: 'dso-faint' }, fmt(t('cad.progress.frames'), { done: framesDone, total: framesTotal })),
              h('span', { className: 'dso-spacer' }),
              h('span', { className: 'dso-faint' }, percent + '%')),
            h('div', {
              className: 'dso-progress',
              role: 'progressbar',
              'aria-valuemin': 0,
              'aria-valuemax': 100,
              'aria-valuenow': percent,
            }, h('div', { className: 'dso-progressBar', style: { width: percent + '%' } })),
            progress.currentFile ? h('div', { className: 'dso-hint' }, fmt(t('cad.progress.current'), { file: progress.currentFile })) : null,
            logLines.length === 0
              ? h('div', { className: 'dso-hint' }, t('cad.log.empty'))
              : h('div', { className: 'dso-log', ref: logRef }, logLines.map((line, index) => h('div', {
                key: 'log' + index,
                className: 'dso-logLine',
                'data-level': line && typeof line.level === 'string' ? line.level : 'info',
              },
              h('span', { className: 'dso-logTime' }, formatClock(line && line.t ? line.t : Date.now())),
              h('span', null, line && line.msg !== undefined ? String(line.msg) : String(line))))))),
        state === 'done' && outputPath
          ? h('div', { className: 'dso-card' },
            h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('cad.result.title'))),
            h('div', { className: 'dso-cardBody' },
              h('div', { className: 'dso-okBox' }, fmt(t('cad.result.output'), { path: outputPath })),
              h('div', { className: 'dso-actions' },
                h('button', { type: 'button', className: 'dso-btn', onClick: () => api.copy(outputPath) }, t('docs.actions.copyPath')),
                h('button', { type: 'button', className: 'dso-btn dso-btn--primary', onClick: () => openInDocs(outputPath) }, t('cad.result.openInDocs')))))
          : null,
        state === 'failed' && status && status.error
          ? h('div', { className: 'dso-card' }, h('div', { className: 'dso-cardBody' },
            h(ErrorBox, { t, error: toPanelError(status.error, t) })))
          : null);
    }
    //#endregion

    //#region 10. 组件：DiagTab（设置与诊断：引擎探测表格 + 配置查看 + 白名单设置）
    /** PROGID 型引擎（probe.ps1 的键名即 ProgID）。 */
    const ENGINE_SPECS = [
      { label: 'Word', progId: 'Word.Application', cap: 'diag.cap.word' },
      { label: 'Excel', progId: 'Excel.Application', cap: 'diag.cap.excel' },
      { label: 'PowerPoint', progId: 'PowerPoint.Application', cap: 'diag.cap.powerpoint' },
      { label: 'Visio', progId: 'Visio.Application', cap: 'diag.cap.visio' },
      { label: 'AutoCAD', progId: 'AutoCAD.Application', cap: 'diag.cap.autocad' },
      { label: 'ZWCAD', progId: 'ZWCAD.Application', cap: 'diag.cap.zwcad' },
      { label: 'GstarCAD', progId: 'GstarCAD.Application', cap: 'diag.cap.gstarcad' },
    ];
    /** 可写配置白名单（§2 suite.config.set）。 */
    const SETTING_SPECS = [
      { key: 'outputRoot', kind: 'text' },
      { key: 'parallel', kind: 'text' },
      { key: 'allowExternalPaths', kind: 'boolean' },
      { key: 'odaPath', kind: 'text' },
      { key: 'pdfPaper', kind: 'select', options: ['A4-landscape', 'keep'] },
      { key: 'sortMode', kind: 'select', options: ['drawingNo', 'filename'] },
    ];

    /** 把 suite.env 结果整理成表格行（兼容 v0.1 顶层 ProgID 与 P1 的 cad/oda 结构化字段）。 */
    function buildEngineRows(env, t) {
      const rows = [];
      const platforms = env && env.cad && Array.isArray(env.cad.platforms) ? env.cad.platforms : [];
      const platformByProgId = {};
      for (const platform of platforms) {
        if (platform && typeof platform.progId === 'string') platformByProgId[platform.progId] = platform;
      }
      for (const spec of ENGINE_SPECS) {
        const legacy = (env && env[spec.progId]) || {};
        const platform = platformByProgId[spec.progId] || null;
        const installPaths = platform && Array.isArray(platform.installPaths) ? platform.installPaths : [];
        const present = platform ? platform.present === true : legacy.present === true;
        const version = (typeof legacy.curVer === 'string' && legacy.curVer)
          || (platform && platform.curVer ? String(platform.curVer) : '')
          || (installPaths.length > 0 ? String(installPaths[0]) : '');
        rows.push({ key: spec.label, label: spec.label, present, version, capability: t(spec.cap) });
      }
      const libre = (env && env.libreOffice) || {};
      const librePaths = Array.isArray(libre.paths) ? libre.paths : [];
      rows.push({
        key: 'LibreOffice',
        label: 'LibreOffice',
        present: libre.found === true,
        version: librePaths.length > 0 ? String(librePaths[0]) : '',
        capability: t('diag.cap.libreoffice'),
      });
      const python = (env && env.python) || {};
      rows.push({
        key: 'Python',
        label: 'Python',
        present: python.found === true,
        version: typeof python.version === 'string' ? python.version : '',
        capability: t('diag.cap.python'),
      });
      const odaConverter = (env && env.odaConverter) || {};
      const odaDirs = Array.isArray(odaConverter.dirs) ? odaConverter.dirs : [];
      const odaStructured = (env && env.oda) || {};
      const odaPath = typeof odaStructured.converter === 'string' && odaStructured.converter
        ? odaStructured.converter
        : (typeof odaConverter.converter === 'string' && odaConverter.converter ? odaConverter.converter : '');
      const odaPresent = odaConverter.found === true || odaStructured.found === true;
      const cadPresent = rows.some((row) => ['AutoCAD', 'ZWCAD', 'GstarCAD'].includes(row.key) && row.present);
      rows.push({
        key: 'ODA',
        label: 'ODAFileConverter',
        present: odaPresent,
        degraded: odaPresent && !cadPresent,
        version: odaPath || (odaDirs.length > 0 ? String(odaDirs[0]) : ''),
        capability: t('diag.cap.oda'),
      });
      return rows;
    }

    /** 设置与诊断页。 */
    function DiagTab({ api, t, active }) {
      const [env, setEnv] = React.useState(null);
      const [envError, setEnvError] = React.useState(null);
      const [probing, setProbing] = React.useState(false);
      const [copied, setCopied] = React.useState(false);
      const [config, setConfig] = React.useState(null);
      const [configError, setConfigError] = React.useState(null);
      const [configOpen, setConfigOpen] = React.useState(false);
      const [edits, setEdits] = React.useState({});
      const [savingKey, setSavingKey] = React.useState('');
      const [savedKey, setSavedKey] = React.useState('');
      const envRef = React.useRef(env);
      envRef.current = env;

      const probe = React.useCallback(async (refresh) => {
        setProbing(true);
        setEnvError(null);
        try {
          const snapshot = await api.callCommand('suite.env', refresh ? { refresh: true } : {});
          setEnv(snapshot);
        } catch (error) {
          setEnvError(toPanelError(error, t));
        } finally {
          setProbing(false);
        }
      }, [api, t]);

      const loadConfig = React.useCallback(async () => {
        setConfigError(null);
        try {
          const snapshot = await api.callCommand('suite.config.get', {});
          setConfig(snapshot);
          setEdits({});
        } catch (error) {
          setConfigError(toPanelError(error, t));
        }
      }, [api, t]);

      // 首次打开诊断页时自动探测 + 读取配置（失败后不再自动重试，交给按钮）。
      React.useEffect(() => {
        if (!active) return;
        if (env || envError || probing) return;
        probe(false);
      }, [active, env, envError, probing, probe]);

      React.useEffect(() => {
        if (!active) return;
        if (config || configError) return;
        loadConfig();
      }, [active, config, configError, loadConfig]);

      useConnectionReset(api, () => {
        probe(true);
        loadConfig();
      });

      const rows = React.useMemo(() => buildEngineRows(env, t), [env, t]);
      const available = rows.filter((row) => row.present).map((row) => row.label);
      const cadPresent = rows.some((row) => ['AutoCAD', 'ZWCAD', 'GstarCAD'].includes(row.key) && row.present);

      const copyDiagnostics = React.useCallback(async () => {
        const lines = [];
        lines.push(t('panel.title') + ' · ' + t('diag.title'));
        lines.push('time: ' + new Date().toISOString());
        lines.push('os: ' + (env && typeof env.os === 'string' ? env.os : '—'));
        lines.push('officeInstalledVersions: ' + (env && Array.isArray(env.officeInstalledVersions) ? env.officeInstalledVersions.join(', ') : '—'));
        lines.push('visioInstalledVersions: ' + (env && Array.isArray(env.visioInstalledVersions) ? env.visioInstalledVersions.join(', ') : '—'));
        for (const row of rows) {
          lines.push([row.label, row.present ? t('diag.engine.available') : t('diag.engine.missing'), row.version || '—', row.capability].join(' | '));
        }
        if (config) lines.push('config: ' + JSON.stringify(config));
        if (envError) lines.push('envError: ' + envError.message);
        if (configError) lines.push('configError: ' + configError.message);
        const ok = await api.copy(lines.join('\n'));
        setCopied(ok);
        if (ok) window.setTimeout(() => setCopied(false), 1600);
      }, [api, config, configError, env, envError, rows, t]);

      const valueOf = React.useCallback((key) => {
        // suite.config.get/set 返回 { effective, overrides, writableKeys, path, boundary } 信封
        const effective = config && typeof config === 'object' && config.effective && typeof config.effective === 'object'
          ? config.effective
          : config;
        if (Object.prototype.hasOwnProperty.call(edits, key)) return edits[key];
        if (effective && typeof effective === 'object' && Object.prototype.hasOwnProperty.call(effective, key)) return effective[key];
        return undefined;
      }, [config, edits]);

      const saveSetting = React.useCallback(async (spec) => {
        if (savingKey) return;
        setSavingKey(spec.key);
        setSavedKey('');
        setConfigError(null);
        try {
          const raw = valueOf(spec.key);
          let value = raw;
          if (spec.kind === 'boolean') {
            value = raw === true || raw === 'true';
          } else if (spec.key === 'parallel') {
            value = raw === 'auto' ? 'auto' : Number(raw);
            if (value !== 'auto' && !(Number.isInteger(value) && value >= 1 && value <= 8)) {
              throw { code: 'CONFIG_INVALID_VALUE', message: t('diag.invalidParallel') };
            }
          } else if (typeof raw === 'string') {
            value = raw;
          }
          const updated = await api.callCommand('suite.config.set', { key: spec.key, value });
          if (updated && typeof updated === 'object') setConfig(updated);
          setSavedKey(spec.key);
          window.setTimeout(() => setSavedKey((prev) => (prev === spec.key ? '' : prev)), 1800);
        } catch (error) {
          setConfigError(toPanelError(error, t));
        } finally {
          setSavingKey('');
        }
      }, [api, savingKey, t, valueOf]);

      const renderEditor = (spec) => {
        const value = valueOf(spec.key);
        const disabled = savingKey !== '' && savingKey !== spec.key;
        if (spec.kind === 'boolean') {
          return h('label', { className: 'dso-check' },
            h('input', {
              type: 'checkbox',
              checked: value === true || value === 'true',
              disabled,
              onChange: (event) => setEdits((prev) => Object.assign({}, prev, { [spec.key]: event.target.checked })),
            }),
            value === true || value === 'true' ? t('diag.engine.available') : t('diag.engine.missing'));
        }
        if (spec.kind === 'select') {
          return h('select', {
            className: 'dso-select',
            value: typeof value === 'string' ? value : '',
            disabled,
            onChange: (event) => setEdits((prev) => Object.assign({}, prev, { [spec.key]: event.target.value })),
          },
          spec.options.map((option) => h('option', { key: option, value: option },
            spec.key === 'pdfPaper'
              ? (option === 'keep' ? t('cad.paper.keep') : t('cad.paper.a4landscape'))
              : (option === 'drawingNo' ? t('cad.sort.drawingNo') : t('cad.sort.filename')))));
        }
        return h('input', {
          className: 'dso-input dso-mono',
          type: 'text',
          value: value === undefined || value === null ? '' : String(value),
          disabled,
          placeholder: spec.key === 'parallel' ? t('diag.value.auto') : '',
          onChange: (event) => setEdits((prev) => Object.assign({}, prev, { [spec.key]: event.target.value })),
        });
      };

      const settingRows = SETTING_SPECS.map((spec) => h('tr', { key: spec.key },
        h('td', { className: 'dso-mono' }, t('diag.key.' + spec.key)),
        h('td', null, renderEditor(spec)),
        h('td', { style: { whiteSpace: 'nowrap' } },
          h('button', {
            type: 'button',
            className: 'dso-btn dso-btn--sm',
            onClick: () => saveSetting(spec),
            disabled: savingKey === spec.key,
          }, savingKey === spec.key ? t('diag.saving') : t('diag.save')),
          savedKey === spec.key ? h('span', { className: 'dso-ok', style: { marginLeft: '8px' } }, t('diag.saved')) : null)));

      return h('div', { className: 'dso-stack' },
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' },
            h('h2', { className: 'dso-cardTitle' }, t('diag.overview')),
            h('span', { className: 'dso-spacer' }),
            h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => probe(true), disabled: probing },
              probing ? t('diag.probing') : t('diag.probe')),
            h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: copyDiagnostics },
              copied ? t('diag.copy.done') : t('diag.copy'))),
          h('div', { className: 'dso-cardBody' },
            envError ? h(ErrorBox, { t, error: envError, onRetry: () => probe(true) }) : null,
            probing && !env ? h(Busy, { text: t('diag.probing') }) : null,
            h('div', { className: available.length > 0 ? 'dso-okBox' : 'dso-errorBox' },
              h('div', null, available.length > 0
                ? fmt(t('diag.overview.available'), { list: available.join('、') })
                : t('diag.overview.none')),
              !cadPresent ? h('div', { className: 'dso-hint' }, t('diag.overview.fallback')) : null),
            rows.length === 0
              ? h('div', { className: 'dso-hint' }, t('diag.config.empty'))
              : h('table', { className: 'dso-table' },
                h('thead', null, h('tr', null,
                  h('th', null, t('diag.table.engine')),
                  h('th', null, t('diag.table.status')),
                  h('th', null, t('diag.table.version')),
                  h('th', null, t('diag.table.capability')))),
                h('tbody', null, rows.map((row) => h('tr', { key: row.key },
                  h('td', null, row.label),
                  h('td', null, h('span', {
                    className: row.present
                      ? (row.degraded ? 'dso-badge dso-badge--warn' : 'dso-badge dso-badge--ok')
                      : 'dso-badge dso-badge--idle',
                  }, row.present
                    ? (row.degraded ? t('diag.engine.degraded') : t('diag.engine.available'))
                    : t('diag.engine.missing'))),
                  h('td', { className: 'dso-mono' }, row.version || '—'),
                  h('td', { className: 'dso-muted' }, row.capability))))))),
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' },
            h('h2', { className: 'dso-cardTitle' }, t('diag.config')),
            h('span', { className: 'dso-spacer' }),
            h('button', {
              type: 'button',
              className: 'dso-btn dso-btn--sm',
              onClick: () => {
                const next = !configOpen;
                setConfigOpen(next);
                if (next && !config) loadConfig();
              },
            }, configOpen ? t('diag.config.hide') : t('diag.config.show'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-hint' }, t('diag.config.note')),
            configError ? h(ErrorBox, { t, error: configError, onRetry: loadConfig }) : null,
            configOpen
              ? (config
                ? h('div', { className: 'dso-stack', style: { gap: '6px' } },
                  h('pre', { className: 'dso-pre' }, JSON.stringify(config.effective !== undefined ? config.effective : config, null, 2)),
                  config.path ? h('div', { className: 'dso-hint dso-mono' }, String(config.path)) : null)
                : h('div', { className: 'dso-hint' }, t('diag.config.empty')))
              : h('div', { className: 'dso-hint' }, t('diag.config.empty')))),
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('diag.settings'))),
          h('div', { className: 'dso-cardBody' },
            h('table', { className: 'dso-table' },
              h('thead', null, h('tr', null,
                h('th', null, t('diag.table.engine')),
                h('th', null, t('diag.table.capability')),
                h('th', null, ''))),
              h('tbody', null, settingRows)))));
    }
    //#endregion

    //#region 11. exports: inject + apply（locale 注册、CSS 注入、main 与 sidebar.panellist 注册）
    /**
     * Client 插件服务依赖（§0.2）。
     * 关键修正：草图中的 `remote.commands` / `remote.workspaceFiles` /
     * `remote.directoryPicker` 属于可缺挂的命名空间服务，列入硬依赖会让插件在缺挂时
     * apply 不执行；这里只保留 7 个稳定服务，命名空间统一用 ctx.get() 防御读取（region 4）。
     */
    const inject = ['slots', 'layout', 'locale', 'theme', 'sessions', 'uiWorkspace', 'remote'];

    /** 注册字典 / 样式 / 主面板 / 侧栏图标。 */
    function apply(ctx) {
      // §5.1 降级：缺少面板基础设施时只记日志，不注册面板（工具与命令由 Host 半负责）。
      if (!ctx.get('slots') || typeof ctx.get('layout') !== 'object' || ctx.get('layout') === null) {
        console.warn('[dsh-office-suite-for-grid] slots/layout unavailable: panel not registered');
        return;
      }

      // 1) 字典（zh 为键集来源，en 覆盖同一键集）
      ctx.effect(() => ctx.locale.register(NS, { zh, en }), 'office-suite: dictionaries');
      const t = ctx.locale.bind(NS);

      // 2) 面板样式（一次性注入，卸载移除）
      ctx.effect(() => {
        if (typeof document === 'undefined') return () => {};
        const tag = document.createElement('style');
        tag.dataset.plugin = 'dsh-office-suite-for-grid';
        tag.dataset.pluginCss = 'dsh-office-suite-for-grid/panel.css';
        tag.textContent = STYLE_TEXT;
        document.head.appendChild(tag);
        return () => { tag.remove(); };
      }, 'office-suite: stylesheet');

      const api = buildApi(ctx);

      // 3) 主面板（options.key 必须与侧栏 id 相同；§1.2）
      try {
        ctx.slots.inject('main', () => ctx.slots.register({
          name: 'main',
          key: PANEL_ID,
          locale: NS,
          inject: () => ({ api, tf: t }),
        }, OfficeSuitePanel));
      } catch (error) {
        // R-10：id/key 冲突时只记日志，不拖垮插件其余能力。
        console.error('[dsh-office-suite-for-grid] main panel registration failed', error);
      }

      // 4) 侧栏图标（组件只渲染字形，props {size, active}；order=30）
      try {
        ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
          name: 'sidebar.panellist',
          id: PANEL_ID,
          order: 30,
          label: () => t('panel.title'),
          locale: NS,
        }, PanelGlyph));
      } catch (error) {
        console.error('[dsh-office-suite-for-grid] sidebar entry registration failed', error);
      }
    }

    return { inject, apply };
    //#endregion
  },
});
