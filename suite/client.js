/**
 * DSH 办公·电网绘图套件 —— Client 半（P1 面板）
 *
 * ModuleLoader 单文件产物：工厂内部按 `//#region` 分区（顺序即依赖顺序）：
 *   1 i18n 字典 → 2 全局 CSS → 3 工具函数 → 4 remote 封装 → 5 PanelGlyph
 *   → 6 OfficeSuitePanel 外壳 → 7 DocsTab → 8 VisioTab → 9 CadTab → 10 DiagTab
 *   → 11 inject + apply
 *
 * 与设计文档 docs/P1-设计.md 的差异（均为有意修正，逐条注明）：
 *  - 面板调用宿主命令改走插件自注册的同源 HTTP 路由 `/office-suite/rpc`（见 region 4）：
 *    typert 网关按 descriptor 校验 `remote.commands.execute` 参数时报
 *    missing "submittedAttachments"（框架层 descriptor 不一致），因此不再依赖
 *    `remote.commands` / `remote.workspaceFiles`；工作区文件浏览改走同一路由的
 *    workspace.list / workspace.stat，预览与诊断统一走 suite.preview。
 *  - §1.2 草图的 inject 把 `remote.commands` / `remote.workspaceFiles` /
 *    `remote.directoryPicker` 列为硬依赖。本实现改为**仅**硬依赖
 *    ['slots','layout','locale','theme','sessions','uiWorkspace','remote']，
 *    命名空间一律 `ctx.get(...)` 判空防御获取（见 region 4），缺挂时对应功能
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
      'docs.tree.filterHint': '默认仅显示 doc/docx/xls/xlsx/ppt/pptx/vsdx/dwg/pdf',
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

      /* P2/P3 增量：编辑 / 排版 / 预览 / 备份 / 导出 */
      'docs.sub.edit': '编辑',
      'docs.sub.format': '排版',
      'docs.sub.preview': '预览',
      'docs.edit.title.word': 'Word 编辑',
      'docs.edit.title.excel': 'Excel 编辑',
      'docs.edit.action': '指令类型',
      'docs.edit.op': '操作',
      'docs.edit.at': '位置（1 起）',
      'docs.edit.text': '文本',
      'docs.edit.style': '样式名',
      'docs.edit.level': '标题级别',
      'docs.edit.table': '表格序号',
      'docs.edit.row': '行号',
      'docs.edit.col': '列号',
      'docs.edit.sheet': '工作表名（留空 = 活动工作表）',
      'docs.edit.sheetName': '工作表名',
      'docs.edit.newName': '新工作表名',
      'docs.edit.value': '值',
      'docs.edit.formula': '公式（以 = 开头）',
      'docs.edit.add': '加入队列',
      'docs.edit.queue': '指令队列（{count} 条）',
      'docs.edit.queue.empty': '尚未加入指令',
      'docs.edit.remove': '移除',
      'docs.edit.clear': '清空',
      'docs.edit.apply': '应用编辑',
      'docs.edit.applying': '应用中…',
      'docs.edit.applied': '已应用 {count} 条指令',
      'docs.edit.partial': '部分指令失败（见下方逐条结果）',
      'docs.edit.engine': '引擎：{engine}',
      'docs.edit.backup': '备份：{path}',
      'docs.edit.result': '编辑结果',
      'docs.edit.formInvalid': '指令参数不完整或格式错误',
      'docs.edit.hint.word': 'Word 指令：段落 append/setText/setStyle/insert/delete · 标题 setStyle/setText · 表格 setCell/行列增删',
      'docs.edit.hint.excel': 'Excel 指令：单元格 value/formula 二选一 · 行列 insert/delete · 工作表 rename/add/delete',
      'docs.edit.action.paragraph': '段落',
      'docs.edit.action.heading': '标题',
      'docs.edit.action.table': '表格',
      'docs.edit.action.cell': '单元格',
      'docs.edit.action.row': '行',
      'docs.edit.action.column': '列',
      'docs.edit.action.sheet': '工作表',
      'docs.edit.op.append': '追加段落',
      'docs.edit.op.setText': '设置文本',
      'docs.edit.op.setStyle': '设置样式',
      'docs.edit.op.insert': '插入',
      'docs.edit.op.delete': '删除',
      'docs.edit.op.setCell': '设置单元格',
      'docs.edit.op.addRow': '添加行',
      'docs.edit.op.deleteRow': '删除行',
      'docs.edit.op.addColumn': '添加列',
      'docs.edit.op.deleteColumn': '删除列',
      'docs.edit.op.rename': '重命名',
      'docs.edit.op.add': '新增',
      'docs.format.rules': '排版规则',
      'docs.format.rule.add': '加入规则',
      'docs.format.rule.empty': '尚未加入规则',
      'docs.format.target': '规则目标',
      'docs.format.target.heading': '标题',
      'docs.format.target.body': '正文',
      'docs.format.target.page': '页面',
      'docs.format.target.header': '页眉',
      'docs.format.target.footer': '页脚',
      'docs.format.headingLevel': '级别',
      'docs.format.fontName': '字体',
      'docs.format.fontSize': '字号（pt）',
      'docs.format.bold': '加粗',
      'docs.format.paraLn': '行距（倍数）',
      'docs.format.spaceBefore': '段前（pt）',
      'docs.format.spaceAfter': '段后（pt）',
      'docs.format.align': '对齐',
      'docs.format.align.none': '不设置',
      'docs.format.numbering': '编号',
      'docs.format.numbering.chinese': '中文（一、）',
      'docs.format.numbering.arabic': '阿拉伯数字（1.）',
      'docs.format.numbering.decimalDot': '数字点（1.1）',
      'docs.format.numbering.none': '无编号',
      'docs.format.margins': '页边距（上/下/左/右，mm）',
      'docs.format.styleSet': '样式集',
      'docs.format.template': '模板',
      'docs.format.placeholders': '模板占位符',
      'docs.format.placeholder.add': '添加占位符',
      'docs.format.placeholder.key': '键',
      'docs.format.placeholder.value': '值',
      'docs.format.scope': '作用范围',
      'docs.format.scope.all': '全部',
      'docs.format.scope.headings': '仅指定标题级别',
      'docs.format.apply': '应用排版',
      'docs.format.applying': '排版中…',
      'docs.format.done': '排版完成：{path}',
      'docs.format.needRule': '请先加入至少一条规则，或选择样式集/模板',
      'docs.format.ruleEmpty': '该规则至少需要字体、段落或编号之一',
      'docs.format.batch': '批量排版',
      'docs.format.batch.dir': '目标目录',
      'docs.format.batch.note': '仅处理目录内 .doc/.docx（排除 ~$ 临时文件）',
      'docs.format.batch.start': '开始批量排版',
      'docs.format.batch.started': '批量排版任务已启动：{taskId}',
      'docs.format.batch.failures': '失败文件：{count}',
      'docs.preview.render': '渲染预览',
      'docs.preview.rendering': '预览渲染中…',
      'docs.preview.retry': '重新渲染',
      'docs.preview.unsupported': '预览服务不可用（E2012）',
      'docs.preview.unsupported.hint': '宿主未提供 officeToPdf 渲染服务；可改用系统应用打开文件',
      'docs.preview.missingFonts': '缺失字体 {count} 个',
      'docs.preview.empty': '点击“预览”生成页面预览',
      'docs.preview.scheme.direct': '渲染方案：直连（A）',
      'docs.preview.scheme.command': '渲染方案：服务端命令通道',
      'docs.preview.cached': '预览缓存 {count}/{max}',
      'docs.export': '导出 PDF',
      'docs.export.running': '导出中…',
      'docs.export.done': '已导出：{path}',
      'docs.backup.title': '自动备份',
      'docs.backup.none': '暂无备份（首次修改后生成）',
      'docs.backup.refresh': '刷新备份',
      'docs.backup.restore': '恢复',
      'docs.backup.confirm': '确认恢复该备份？当前版本会先自动备份',
      'docs.backup.restored': '已恢复：{path}',
      'docs.backup.id': '备份 ID',
      'docs.backup.time': '时间',
      'docs.backup.size': '大小',

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

      /* P2/P3 增量：电网九类图 + 模型模板 + 校验提示 */
      'visio.kind.singleLine': '电气主接线图（singleLine）',
      'visio.kind.distribution': '配电系统图（distribution）',
      'visio.kind.networkTopology': '通信网络拓扑图（networkTopology）',
      'visio.kind.protection': '二次回路/继保原理图（protection）',
      'visio.kind.substationPlan': '变电站总平面布置示意图（substationPlan）',
      'visio.kind.geoNetwork': '地理接线图（geoNetwork）',
      'visio.kind.powerFlow': '潮流图/系统接线图（powerFlow）',
      'visio.kind.terminalStrip': '端子排图（terminalStrip）',
      'visio.kind.cableRoute': '电缆路径图（cableRoute，P4）',
      'visio.kind.cableRoute.p4hint': '电缆路径图（cableRoute）在 P4 提供，本期仅定稿 schema',
      'visio.model': '模型 JSON（统一底座 + 类型专属字段）',
      'visio.model.notObject': '模型必须是 JSON 对象',
      'visio.model.kindMismatch': '模型 kind（{actual}）与所选图类型（{expected}）不一致',
      'visio.model.required': '缺少必填字段 {key}（{kind} 图要求）',
      'visio.template.model': '插入模型模板',
      'visio.validations': '模型校验提示',
      'visio.validations.empty': '无校验提示',
      'visio.validations.warn': '警告',
      'visio.validations.info': '提示',
      'visio.result.pages': '页数：{pages}',
      'visio.e2024': '电缆路径图（cableRoute）将在 P4 提供，本期请使用其他图类型',
      'visio.stats.symbols': '符号',
      'visio.stats.shapes': '图形',
      'visio.stats.links': '连线',
      'visio.stats.nodes': '节点',
      'visio.stats.edges': '连线',
      'visio.stats.items': '图元',
      'visio.stats.levels': '电压等级',
      'visio.stats.buses': '母线',
      'visio.stats.bays': '间隔',
      'visio.stats.groups': '组合',
      'visio.stats.circuits': '回路',
      'visio.stats.terminals': '端子',
      'visio.stats.strips': '端子排',
      'visio.stats.stations': '厂站',
      'visio.stats.lines': '线路',
      'visio.stats.flows': '潮流',
      'visio.stats.transformers': '变压器',
      'visio.stats.zones': '设备区',

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

      /* P3 增量：图纸目录表 + 批量队列 */
      'cad.catalog.title': '图纸目录表',
      'cad.catalog.input': '图纸文件或目录',
      'cad.catalog.run': '生成目录表',
      'cad.catalog.running': '正在提取目录字段…',
      'cad.catalog.done': '目录表已生成：{path}',
      'cad.catalog.rows': '目录行数：{rows}',
      'cad.catalog.sheets': '工作表数：{sheets}',
      'cad.catalog.engine': '写表引擎：{engine}',
      'cad.catalog.fields': '字段命中率',
      'cad.catalog.field.drawingNo': '图号',
      'cad.catalog.field.title': '图名',
      'cad.catalog.field.frameSize': '图幅',
      'cad.catalog.field.scale': '比例',
      'cad.catalog.field.count': '张数',
      'cad.catalog.hint': '支持 .dwg/.pdf 单文件或目录；规模上限由配置 catalogMaxFrames 控制',
      'cad.queue.title': '批量队列（多文件/多文件夹）',
      'cad.queue.inputs': '队列条目（{count}）',
      'cad.queue.addFile': '添加文件…',
      'cad.queue.addDir': '添加文件夹…',
      'cad.queue.remove': '移除',
      'cad.queue.clear': '清空',
      'cad.queue.empty': '队列为空',
      'cad.queue.note': '合并场景条目需为文件夹，分割场景条目需为单个 .dwg/.pdf；每项独立转换并汇总进度',
      'cad.queue.start': '开始批量转换',
      'cad.queue.started': '批量任务已启动：{taskId}',
      'cad.queue.progress': '条目 {done}/{total}',
      'cad.queue.failures': '失败条目：{count}',

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

      /* P2/P3 增量：venv / Office COM / 预览服务状态行 + previewCacheSize */
      'diag.engine.venv': 'Python venv',
      'diag.cap.venv': 'Python 编辑/排版引擎（python-docx/openpyxl）',
      'diag.venv.healthy': '健康',
      'diag.venv.unhealthy': '需重建',
      'diag.venv.missing': '尚未创建（首次编辑/排版时自动创建）',
      'diag.venv.packages': '依赖：docx {docx} · openpyxl {openpyxl}',
      'diag.engine.officeCom': 'Office COM',
      'diag.cap.officeCom': 'Office 编辑/排版 COM 引擎',
      'diag.com.usable': '可用',
      'diag.com.unusable': '不可用',
      'diag.com.unmeasured': '未实测（点击“重新探测”）',
      'diag.preview': '预览渲染 officeToPdf',
      'diag.cap.preview': '页面预览 / 导出 PDF 的渲染服务',
      'diag.preview.available': '可用',
      'diag.preview.unavailable': '不可用（E2012）',
      'diag.preview.unknown': '未知',
      'diag.preview.probing': '探测预览服务…',
      'diag.preview.direct': '直连（方案 A）',
      'diag.preview.command': '服务端命令通道',
      'diag.preview.note': '探测方式：suite.preview（E2012 判定为不可用）',
      'diag.key.previewCacheSize': '预览缓存条数 previewCacheSize（1–10）',
      'diag.invalidCacheSize': '预览缓存条数只能是 1–10 的整数',

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
      'common.none': '无',
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
      'docs.tree.filterHint': 'Showing doc/docx/xls/xlsx/ppt/pptx/vsdx/dwg/pdf only',
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

      /* P2/P3 additions: edit / format / preview / backup / export */
      'docs.sub.edit': 'Edit',
      'docs.sub.format': 'Format',
      'docs.sub.preview': 'Preview',
      'docs.edit.title.word': 'Word edit',
      'docs.edit.title.excel': 'Excel edit',
      'docs.edit.action': 'Command type',
      'docs.edit.op': 'Operation',
      'docs.edit.at': 'Position (1-based)',
      'docs.edit.text': 'Text',
      'docs.edit.style': 'Style name',
      'docs.edit.level': 'Heading level',
      'docs.edit.table': 'Table index',
      'docs.edit.row': 'Row',
      'docs.edit.col': 'Column',
      'docs.edit.sheet': 'Sheet name (empty = active sheet)',
      'docs.edit.sheetName': 'Sheet name',
      'docs.edit.newName': 'New sheet name',
      'docs.edit.value': 'Value',
      'docs.edit.formula': 'Formula (starts with =)',
      'docs.edit.add': 'Add to queue',
      'docs.edit.queue': 'Command queue ({count})',
      'docs.edit.queue.empty': 'No command queued yet',
      'docs.edit.remove': 'Remove',
      'docs.edit.clear': 'Clear',
      'docs.edit.apply': 'Apply edits',
      'docs.edit.applying': 'Applying…',
      'docs.edit.applied': '{count} command(s) applied',
      'docs.edit.partial': 'Some commands failed (see per-op results below)',
      'docs.edit.engine': 'Engine: {engine}',
      'docs.edit.backup': 'Backup: {path}',
      'docs.edit.result': 'Edit result',
      'docs.edit.formInvalid': 'Command arguments are incomplete or invalid',
      'docs.edit.hint.word': 'Word commands: paragraph append/setText/setStyle/insert/delete · heading setStyle/setText · table setCell/row-column add-remove',
      'docs.edit.hint.excel': 'Excel commands: cell value/formula (one of) · row/column insert/delete · sheet rename/add/delete',
      'docs.edit.action.paragraph': 'Paragraph',
      'docs.edit.action.heading': 'Heading',
      'docs.edit.action.table': 'Table',
      'docs.edit.action.cell': 'Cell',
      'docs.edit.action.row': 'Row',
      'docs.edit.action.column': 'Column',
      'docs.edit.action.sheet': 'Sheet',
      'docs.edit.op.append': 'Append paragraph',
      'docs.edit.op.setText': 'Set text',
      'docs.edit.op.setStyle': 'Set style',
      'docs.edit.op.insert': 'Insert',
      'docs.edit.op.delete': 'Delete',
      'docs.edit.op.setCell': 'Set cell',
      'docs.edit.op.addRow': 'Add row',
      'docs.edit.op.deleteRow': 'Delete row',
      'docs.edit.op.addColumn': 'Add column',
      'docs.edit.op.deleteColumn': 'Delete column',
      'docs.edit.op.rename': 'Rename',
      'docs.edit.op.add': 'Add',
      'docs.format.rules': 'Format rules',
      'docs.format.rule.add': 'Add rule',
      'docs.format.rule.empty': 'No rule added yet',
      'docs.format.target': 'Rule target',
      'docs.format.target.heading': 'Heading',
      'docs.format.target.body': 'Body',
      'docs.format.target.page': 'Page',
      'docs.format.target.header': 'Header',
      'docs.format.target.footer': 'Footer',
      'docs.format.headingLevel': 'Level',
      'docs.format.fontName': 'Font',
      'docs.format.fontSize': 'Size (pt)',
      'docs.format.bold': 'Bold',
      'docs.format.paraLn': 'Line spacing (multiple)',
      'docs.format.spaceBefore': 'Space before (pt)',
      'docs.format.spaceAfter': 'Space after (pt)',
      'docs.format.align': 'Alignment',
      'docs.format.align.none': 'Not set',
      'docs.format.numbering': 'Numbering',
      'docs.format.numbering.chinese': 'Chinese (一、)',
      'docs.format.numbering.arabic': 'Arabic (1.)',
      'docs.format.numbering.decimalDot': 'Decimal dot (1.1)',
      'docs.format.numbering.none': 'No numbering',
      'docs.format.margins': 'Margins (top/bottom/left/right, mm)',
      'docs.format.styleSet': 'Style set',
      'docs.format.template': 'Template',
      'docs.format.placeholders': 'Template placeholders',
      'docs.format.placeholder.add': 'Add placeholder',
      'docs.format.placeholder.key': 'Key',
      'docs.format.placeholder.value': 'Value',
      'docs.format.scope': 'Scope',
      'docs.format.scope.all': 'All',
      'docs.format.scope.headings': 'Selected heading levels only',
      'docs.format.apply': 'Apply formatting',
      'docs.format.applying': 'Formatting…',
      'docs.format.done': 'Formatted: {path}',
      'docs.format.needRule': 'Add at least one rule, or select a style set / template',
      'docs.format.ruleEmpty': 'A rule needs at least font, paragraph or numbering',
      'docs.format.batch': 'Batch format',
      'docs.format.batch.dir': 'Target folder',
      'docs.format.batch.note': 'Only .doc/.docx inside the folder are processed (~$ temp files excluded)',
      'docs.format.batch.start': 'Start batch formatting',
      'docs.format.batch.started': 'Batch format task started: {taskId}',
      'docs.format.batch.failures': 'Failed files: {count}',
      'docs.preview.render': 'Render preview',
      'docs.preview.rendering': 'Rendering preview…',
      'docs.preview.retry': 'Render again',
      'docs.preview.unsupported': 'Preview service unavailable (E2012)',
      'docs.preview.unsupported.hint': 'The host did not provide the officeToPdf render service; open the file with a system app instead',
      'docs.preview.missingFonts': '{count} missing font(s)',
      'docs.preview.empty': 'Press “Preview” to render this file',
      'docs.preview.scheme.direct': 'Render scheme: direct (A)',
      'docs.preview.scheme.command': 'Render scheme: server command channel',
      'docs.preview.cached': 'Preview cache {count}/{max}',
      'docs.export': 'Export PDF',
      'docs.export.running': 'Exporting…',
      'docs.export.done': 'Exported: {path}',
      'docs.backup.title': 'Automatic backups',
      'docs.backup.none': 'No backups yet (created after the first edit)',
      'docs.backup.refresh': 'Refresh backups',
      'docs.backup.restore': 'Restore',
      'docs.backup.confirm': 'Restore this backup? The current version will be backed up first.',
      'docs.backup.restored': 'Restored: {path}',
      'docs.backup.id': 'Backup ID',
      'docs.backup.time': 'Time',
      'docs.backup.size': 'Size',

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

      /* P2/P3 additions: nine grid diagram kinds + model templates + validations */
      'visio.kind.singleLine': 'Single-line diagram (singleLine)',
      'visio.kind.distribution': 'Distribution system (distribution)',
      'visio.kind.networkTopology': 'Telecom topology (networkTopology)',
      'visio.kind.protection': 'Protection circuit (protection)',
      'visio.kind.substationPlan': 'Substation plan (substationPlan)',
      'visio.kind.geoNetwork': 'Geographic network (geoNetwork)',
      'visio.kind.powerFlow': 'Power flow (powerFlow)',
      'visio.kind.terminalStrip': 'Terminal strip (terminalStrip)',
      'visio.kind.cableRoute': 'Cable route (cableRoute, P4)',
      'visio.kind.cableRoute.p4hint': 'cableRoute arrives in P4; this release only finalizes its schema',
      'visio.model': 'Model JSON (shared base + kind-specific fields)',
      'visio.model.notObject': 'The model must be a JSON object',
      'visio.model.kindMismatch': 'Model kind ({actual}) does not match the selected kind ({expected})',
      'visio.model.required': 'Missing required field {key} (required by {kind})',
      'visio.template.model': 'Insert model template',
      'visio.validations': 'Model validation notices',
      'visio.validations.empty': 'No validation notices',
      'visio.validations.warn': 'Warning',
      'visio.validations.info': 'Notice',
      'visio.result.pages': 'Pages: {pages}',
      'visio.e2024': 'The cable route diagram (cableRoute) arrives in P4; use another kind for now',
      'visio.stats.symbols': 'Symbols',
      'visio.stats.shapes': 'Shapes',
      'visio.stats.links': 'Links',
      'visio.stats.nodes': 'Nodes',
      'visio.stats.edges': 'Edges',
      'visio.stats.items': 'Items',
      'visio.stats.levels': 'Voltage levels',
      'visio.stats.buses': 'Buses',
      'visio.stats.bays': 'Bays',
      'visio.stats.groups': 'Groups',
      'visio.stats.circuits': 'Circuits',
      'visio.stats.terminals': 'Terminals',
      'visio.stats.strips': 'Strips',
      'visio.stats.stations': 'Stations',
      'visio.stats.lines': 'Lines',
      'visio.stats.flows': 'Flows',
      'visio.stats.transformers': 'Transformers',
      'visio.stats.zones': 'Zones',

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

      /* P3 additions: drawing catalog + batch queue */
      'cad.catalog.title': 'Drawing catalog',
      'cad.catalog.input': 'Drawing file or folder',
      'cad.catalog.run': 'Build catalog',
      'cad.catalog.running': 'Extracting catalog fields…',
      'cad.catalog.done': 'Catalog created: {path}',
      'cad.catalog.rows': 'Catalog rows: {rows}',
      'cad.catalog.sheets': 'Sheets: {sheets}',
      'cad.catalog.engine': 'Writer engine: {engine}',
      'cad.catalog.fields': 'Field hit rate',
      'cad.catalog.field.drawingNo': 'Drawing no.',
      'cad.catalog.field.title': 'Title',
      'cad.catalog.field.frameSize': 'Frame size',
      'cad.catalog.field.scale': 'Scale',
      'cad.catalog.field.count': 'Sheet count',
      'cad.catalog.hint': 'Accepts a .dwg/.pdf file or a folder; limits follow catalogMaxFrames',
      'cad.queue.title': 'Batch queue (multiple files/folders)',
      'cad.queue.inputs': 'Queue items ({count})',
      'cad.queue.addFile': 'Add file…',
      'cad.queue.addDir': 'Add folder…',
      'cad.queue.remove': 'Remove',
      'cad.queue.clear': 'Clear',
      'cad.queue.empty': 'The queue is empty',
      'cad.queue.note': 'Merge items must be folders; split items must be single .dwg/.pdf files; each item converts independently with aggregated progress',
      'cad.queue.start': 'Start batch conversion',
      'cad.queue.started': 'Batch task started: {taskId}',
      'cad.queue.progress': 'Items {done}/{total}',
      'cad.queue.failures': 'Failed items: {count}',

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

      /* P2/P3 additions: venv / Office COM / preview status rows + previewCacheSize */
      'diag.engine.venv': 'Python venv',
      'diag.cap.venv': 'Python edit/format engine (python-docx/openpyxl)',
      'diag.venv.healthy': 'Healthy',
      'diag.venv.unhealthy': 'Needs rebuild',
      'diag.venv.missing': 'Not created yet (created automatically on first edit/format)',
      'diag.venv.packages': 'Packages: docx {docx} · openpyxl {openpyxl}',
      'diag.engine.officeCom': 'Office COM',
      'diag.cap.officeCom': 'Office edit/format COM engine',
      'diag.com.usable': 'Usable',
      'diag.com.unusable': 'Not usable',
      'diag.com.unmeasured': 'Not measured (press “Probe again”)',
      'diag.preview': 'Preview renderer officeToPdf',
      'diag.cap.preview': 'Render service for page preview / PDF export',
      'diag.preview.available': 'Available',
      'diag.preview.unavailable': 'Unavailable (E2012)',
      'diag.preview.unknown': 'Unknown',
      'diag.preview.probing': 'Probing the preview service…',
      'diag.preview.direct': 'Direct (scheme A)',
      'diag.preview.command': 'Server command channel',
      'diag.preview.note': 'Probed via suite.preview (E2012 means unavailable)',
      'diag.key.previewCacheSize': 'Preview cache entries previewCacheSize (1–10)',
      'diag.invalidCacheSize': 'Preview cache entries must be an integer between 1 and 10',

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
      'common.none': 'None',
    };
    //#endregion

    //#region 2. 全局 CSS 文本（apply 时注入一次 <style data-plugin="dsh-office-suite-for-grid">）
    /** 面板样式。颜色/圆角/字体/动效一律走 §1.5 令牌；唯一例外是主按钮上的对比色白字。 */
    const STYLE_TEXT = [
      /* 布局骨架 */
      '.dso-panel{box-sizing:border-box;height:100%;min-height:0;display:flex;flex-direction:column;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-family:var(--dsw-font-family);font-size:13px;line-height:20px;overflow:hidden;color-scheme:light dark}',
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
      '.dso-tab:hover{background:var(--dsw-alias-interactive-bg-hover, var(--dsw-alias-bg-layer-2));color:var(--dsw-alias-label-primary)}',
      '.dso-tab[aria-current=page]{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);box-shadow:inset 0 -2px 0 var(--dsw-alias-brand-primary)}',
      '.dso-tab:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-brand-primary));outline-offset:-2px}',
      /* 子页签（P2/P3：编辑 / 排版 / 预览） */
      '.dso-subtabs{display:flex;gap:2px;border-bottom:.5px solid var(--dsw-alias-border-l1);overflow-x:auto;padding-bottom:0}',
      '.dso-subtab{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:12px;line-height:18px;padding:4px 10px;border-radius:var(--dsw-radius-sm) var(--dsw-radius-sm) 0 0;cursor:pointer;white-space:nowrap;transition:background var(--ds-transition-duration-fast) var(--ds-ease-in-out),color var(--ds-transition-duration-fast) var(--ds-ease-in-out)}',
      '.dso-subtab:hover{background:var(--dsw-alias-interactive-bg-hover, var(--dsw-alias-bg-layer-2));color:var(--dsw-alias-label-primary)}',
      '.dso-subtab[aria-current=page]{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);box-shadow:inset 0 -2px 0 var(--dsw-alias-brand-primary)}',
      '.dso-subtab:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-brand-primary));outline-offset:-2px}',
      '.dso-previewFrame{width:100%;height:60vh;min-height:320px;border:.5px solid var(--dsw-alias-border-l2);border-radius:var(--dsw-radius-md);background:#fff}',
      'body[data-ds-dark-theme] .dso-previewFrame{background:var(--dsw-alias-bg-layer-2)}',

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
      '.dso-row:hover{background:var(--dsw-alias-interactive-bg-hover, var(--dsw-alias-bg-layer-2))}',
      '.dso-row[aria-selected=true]{background:color-mix(in srgb, var(--dsw-alias-brand-primary) 14%, transparent);box-shadow:inset 0 0 0 .5px color-mix(in srgb, var(--dsw-alias-brand-primary) 45%, transparent)}',
      '.dso-row:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-brand-primary));outline-offset:-2px}',
      '.dso-caret{width:14px;height:14px;flex:none;display:inline-flex;align-items:center;justify-content:center;color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));font-size:10px;border-radius:var(--dsw-radius-xs);cursor:pointer}',
      '.dso-caret:hover{background:var(--dsw-alias-interactive-bg-hover, var(--dsw-alias-bg-layer-2));color:var(--dsw-alias-label-primary)}',
      '.dso-rowName{overflow:hidden;text-overflow:ellipsis;min-width:0}',
      '.dso-rowMeta{margin-left:auto;flex:none;color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));font-size:11px}',
      '.dso-treeNote{padding:6px 8px;color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));font-size:12px;line-height:18px;overflow-wrap:anywhere}',
      '.dso-crumbs{display:flex;align-items:center;gap:4px;flex-wrap:wrap;font-size:12px;color:var(--dsw-alias-label-secondary);min-width:0}',
      '.dso-crumb{appearance:none;border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:12px;padding:1px 3px;border-radius:var(--dsw-radius-xs);cursor:pointer;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
      '.dso-crumb:hover{background:var(--dsw-alias-interactive-bg-hover, var(--dsw-alias-bg-layer-2));color:var(--dsw-alias-label-primary)}',
      '.dso-crumbSep{color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary))}',

      /* 按钮 */
      '.dso-btn{appearance:none;font:inherit;font-size:12px;line-height:18px;display:inline-flex;align-items:center;justify-content:center;gap:6px;padding:3px 10px;border-radius:var(--dsw-radius-sm);border:.5px solid var(--dsw-alias-border-l2);background:transparent;color:var(--dsw-alias-label-primary);cursor:pointer;transition:background var(--ds-transition-duration-fast) var(--ds-ease-in-out),border-color var(--ds-transition-duration-fast) var(--ds-ease-in-out);white-space:nowrap}',
      '.dso-btn:hover:not(:disabled){background:var(--dsw-alias-interactive-bg-hover, var(--dsw-alias-bg-layer-2))}',
      '.dso-btn:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-brand-primary));outline-offset:1px}',
      '.dso-btn:disabled{opacity:.5;cursor:not-allowed}',
      '.dso-btn--primary{background:var(--dsw-alias-brand-primary);border-color:transparent;color:#fff}',
      '.dso-btn--primary:hover:not(:disabled){filter:brightness(.94)}',
      '.dso-btn--danger{color:var(--dsw-alias-state-error-primary);border-color:color-mix(in srgb, var(--dsw-alias-state-error-primary) 45%, transparent)}',
      '.dso-btn--danger:hover:not(:disabled){background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 10%, transparent)}',
      '.dso-btn--link{border-color:transparent;color:var(--dsw-alias-brand-primary);padding:1px 4px}',
      '.dso-btn--sm{padding:1px 8px;font-size:11px;line-height:16px}',

      /* 表单 */
      '.dso-field{display:flex;flex-direction:column;gap:4px;min-width:0}',
      '.dso-label{font-size:12px;line-height:18px;color:var(--dsw-alias-label-secondary)}',
      '.dso-input,.dso-select,.dso-textarea{box-sizing:border-box;width:100%;font:inherit;font-size:12px;line-height:18px;color:var(--dsw-alias-label-primary);background:var(--dsw-alias-bg-base);border:.5px solid var(--dsw-alias-border-l2);border-radius:var(--dsw-radius-sm);padding:4px 8px;transition:border-color var(--ds-transition-duration-fast) var(--ds-ease-in-out)}',
      '.dso-input:hover,.dso-select:hover,.dso-textarea:hover{border-color:var(--dsw-alias-border-l3, var(--dsw-alias-border-l2))}',
      '.dso-input:focus-visible,.dso-select:focus-visible,.dso-textarea:focus-visible{outline:var(--dsw-focus-ring-width) solid var(--dsw-focus-ring-color,var(--dsw-alias-brand-primary));outline-offset:1px;border-color:var(--dsw-alias-brand-primary)}',
      '.dso-input:disabled,.dso-select:disabled,.dso-textarea:disabled{color:var(--dsw-alias-state-idle-primary);cursor:not-allowed}',
      '.dso-textarea{font-family:var(--ds-font-family-code);font-size:12px;line-height:18px;min-height:132px;resize:vertical;white-space:pre;overflow:auto}',
      '.dso-inline{display:flex;align-items:center;gap:8px;flex-wrap:wrap;min-width:0}',
      '.dso-check{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--dsw-alias-label-secondary);cursor:pointer;user-select:none}',
      '.dso-check input{accent-color:var(--dsw-alias-brand-primary);margin:0}',

      /* 文本与徽标 */
      '.dso-muted{color:var(--dsw-alias-label-secondary)}',
      '.dso-faint{color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary))}',
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
      '.dso-empty{flex:1;min-height:120px;display:flex;align-items:center;justify-content:center;text-align:center;color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));font-size:12px;line-height:18px;padding:16px;overflow-wrap:anywhere}',
      '.dso-hint{font-size:11.5px;line-height:17px;color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));overflow-wrap:anywhere}',
      '.dso-errorBox{display:flex;flex-direction:column;gap:2px;border:.5px solid color-mix(in srgb, var(--dsw-alias-state-error-primary) 40%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-error-primary) 7%, transparent);border-radius:var(--dsw-radius-sm);padding:8px 10px;font-size:12px;line-height:18px;color:var(--dsw-alias-state-error-primary);overflow-wrap:anywhere}',
      '.dso-okBox{display:flex;flex-direction:column;gap:2px;border:.5px solid color-mix(in srgb, var(--dsw-alias-state-success-primary) 40%, transparent);background:color-mix(in srgb, var(--dsw-alias-state-success-primary) 7%, transparent);border-radius:var(--dsw-radius-sm);padding:8px 10px;font-size:12px;line-height:18px;overflow-wrap:anywhere}',

      /* 进度条 / 日志 */
      '.dso-progress{height:6px;border-radius:999px;background:var(--dsw-alias-bg-layer-2);border:.5px solid var(--dsw-alias-border-l1);overflow:hidden}',
      '.dso-progressBar{height:100%;background:var(--dsw-alias-brand-primary);transition:width var(--ds-transition-duration-slow) var(--ds-ease-in-out)}',
      '.dso-log{flex:1;min-height:120px;max-height:280px;overflow:auto;background:var(--dsw-alias-bg-layer-2);border:.5px solid var(--dsw-alias-border-l1);border-radius:var(--dsw-radius-sm);padding:6px 8px;font-family:var(--ds-font-family-code);font-size:11.5px;line-height:16px;color:var(--dsw-alias-label-secondary)}',
      '.dso-logLine{display:flex;gap:8px;white-space:pre-wrap;overflow-wrap:anywhere}',
      '.dso-logTime{color:var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));flex:none}',
      '.dso-logLine[data-level=warn]{color:var(--dsw-alias-state-warn-primary)}',
      '.dso-logLine[data-level=error]{color:var(--dsw-alias-state-error-primary)}',
      '.dso-logLine[data-level=success]{color:var(--dsw-alias-state-success-primary)}',

      /* 微件 */
      '.dso-spinner{width:12px;height:12px;flex:none;border-radius:50%;border:1.5px solid var(--dsw-alias-label-tertiary, var(--dsw-alias-label-secondary));border-top-color:transparent;animation:dso-spin .8s linear infinite}',
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
    /** 文档中心默认显示的文件扩展名（§1.3；P3 增补 ppt/pptx 以便页面预览）。 */
    const DEFAULT_EXTS = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'vsdx', 'dwg', 'pdf'];
    /** 可直接读取渲染的 Office 扩展名。 */
    const OFFICE_EXTS = ['doc', 'docx', 'xls', 'xlsx'];
    /** P1 只显示元信息的扩展名。 */
    const META_EXTS = ['vsdx', 'dwg', 'pdf'];
    /** 图纸转换可接受的输入扩展名（.pdf 供 PDF 级引擎链路）。 */
    const CAD_INPUT_EXTS = ['dwg', 'dxf', 'pdf'];
    /** 可页面预览 / 导出 PDF 的扩展名（P3 §5.1：officeToPdf 的 OfficeExtension）。 */
    const PREVIEW_EXTS = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'];
    /** Word 编辑 / 排版支持的扩展名。 */
    const WORD_EXTS = ['doc', 'docx'];
    /** Excel 编辑支持的扩展名。 */
    const EXCEL_EXTS = ['xls', 'xlsx'];
    /** 通用任务终态（与 CadTab 的 CAD_TERMINAL 同义，供通用轮询 Hook 使用）。 */
    const TASK_TERMINAL = ['done', 'failed', 'cancelled'];
    /** 预览 blob 缓存默认容量（P3 §5.4 previewCacheSize 默认 5）。 */
    const PREVIEW_CACHE_DEFAULT = 5;
    /** P2/P3 电网图类型（flowchart 之外的 8 类；cableRoute 属 P4，命令层 E2024）。 */
    const GRID_MODEL_KINDS = ['singleLine', 'distribution', 'networkTopology', 'protection', 'substationPlan', 'geoNetwork', 'powerFlow', 'terminalStrip'];
    /** P4 预留的图类型（VisioTab 置灰提示）。 */
    const GRID_KIND_PENDING = 'cableRoute';
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

    // ───────── P2/P3 增量：编辑 / 排版 / 电网模型 纯函数（便于冒烟测试与复用） ─────────

    /** 宽松 plain object 判定。 */
    function isPlainObjectLoose(value) {
      return !!value && typeof value === 'object' && !Array.isArray(value);
    }

    /** 正整数解析（1 起）；非法返回 null。 */
    function parsePositiveInt(value) {
      const text = String(value ?? '').trim();
      if (!/^\d+$/.test(text)) return null;
      const number = Number(text);
      return Number.isInteger(number) && number >= 1 ? number : null;
    }

    /** 有限数字解析（空串 → undefined；非法 → null）。 */
    function parseOptionalNumber(value) {
      const text = String(value ?? '').trim();
      if (!text) return undefined;
      const number = Number(text);
      return Number.isFinite(number) ? number : null;
    }

    /** 表单字段（标签 + 控件）。 */
    function formField(label, node) {
      return h('div', { className: 'dso-field' }, h('span', { className: 'dso-label' }, label), node);
    }

    /** base64 → Uint8Array（按 4 的倍数分块 atob，避免超大 base64 一次性解码卡顿；P3 §5.2）。 */
    function base64ToBytes(text) {
      const clean = String(text ?? '').replace(/^data:[^,]*,/, '').replace(/\s+/g, '');
      if (!clean) return new Uint8Array(0);
      const chunkChars = 4 * 65536;
      const chunks = [];
      let total = 0;
      for (let offset = 0; offset < clean.length; offset += chunkChars) {
        const binary = window.atob(clean.slice(offset, offset + chunkChars));
        const bytes = new Uint8Array(binary.length);
        for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
        chunks.push(bytes);
        total += bytes.length;
      }
      const output = new Uint8Array(total);
      let cursor = 0;
      for (const chunk of chunks) {
        output.set(chunk, cursor);
        cursor += chunk.length;
      }
      return output;
    }

    /**
     * 构造单条编辑指令（P2 §3.2）。
     * @param {Function} t i18n
     * @param {'word'|'excel'} family
     * @param {object} form 表单值（字符串为主）
     * @returns {{ok:true, op:object}|{ok:false, message:string}}
     */
    function buildEditOp(t, family, form) {
      const invalid = () => ({ ok: false, message: t('docs.edit.formInvalid') });
      const f = form || {};
      if (family === 'word') {
        const action = typeof f.action === 'string' ? f.action : 'paragraph';
        if (action === 'paragraph') {
          const op = typeof f.op === 'string' ? f.op : 'append';
          if (!['append', 'setText', 'setStyle', 'insert', 'delete'].includes(op)) return invalid();
          if (op === 'append') return { ok: true, op: { action: 'paragraph', op, text: String(f.text ?? '') } };
          const at = parsePositiveInt(f.at);
          if (at === null) return invalid();
          const built = { action: 'paragraph', op, at };
          if (op === 'setText' || op === 'insert') built.text = String(f.text ?? '');
          if (op === 'setStyle') {
            const style = String(f.style ?? '').trim();
            if (!style) return invalid();
            built.style = style;
          }
          return { ok: true, op: built };
        }
        if (action === 'heading') {
          const op = typeof f.op === 'string' ? f.op : 'setStyle';
          if (!['setText', 'setStyle', 'insert', 'delete'].includes(op)) return invalid();
          const at = parsePositiveInt(f.at);
          const level = parsePositiveInt(f.level);
          if (at === null || level === null || level > 3) return invalid();
          const built = { action: 'heading', op, level, at };
          if (op === 'setText' || op === 'insert') built.text = String(f.text ?? '');
          if (op === 'setStyle') {
            const style = String(f.style ?? '').trim();
            if (style) built.style = style;
          }
          return { ok: true, op: built };
        }
        if (action === 'table') {
          const op = typeof f.op === 'string' ? f.op : 'setCell';
          if (!['setCell', 'addRow', 'deleteRow', 'addColumn', 'deleteColumn'].includes(op)) return invalid();
          const table = parsePositiveInt(f.table);
          if (table === null) return invalid();
          const built = { action: 'table', table, op };
          if (op === 'setCell') {
            const row = parsePositiveInt(f.row);
            const col = parsePositiveInt(f.col);
            if (row === null || col === null) return invalid();
            built.row = row;
            built.col = col;
            built.text = String(f.text ?? '');
          } else if (op === 'deleteRow') {
            const row = parsePositiveInt(f.row);
            if (row === null) return invalid();
            built.row = row;
          } else if (op === 'addRow') {
            const at = String(f.at ?? '').trim();
            if (at && at !== 'end') {
              const row = parsePositiveInt(at);
              if (row === null) return invalid();
              built.at = row;
            } else {
              built.at = 'end';
            }
          } else if (op === 'addColumn') {
            const col = parsePositiveInt(f.col);
            if (col !== null) built.col = col;
            else built.at = 'end';
          } else if (op === 'deleteColumn') {
            const col = parsePositiveInt(f.col);
            if (col === null) return invalid();
            built.col = col;
          }
          return { ok: true, op: built };
        }
        return invalid();
      }
      if (family === 'excel') {
        const action = typeof f.action === 'string' ? f.action : 'cell';
        const sheet = String(f.sheet ?? '').trim();
        if (action === 'cell') {
          const at = String(f.at ?? '').trim().toUpperCase();
          if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(at)) return invalid();
          const formula = String(f.formula ?? '').trim();
          const built = { action: 'cell', at };
          if (sheet) built.sheet = sheet;
          if (formula) {
            if (!formula.startsWith('=')) return invalid();
            built.formula = formula;
          } else {
            const raw = String(f.value ?? '');
            if (/^-?\d+(\.\d+)?$/.test(raw.trim())) built.value = Number(raw.trim());
            else if (raw.trim() === 'true') built.value = true;
            else if (raw.trim() === 'false') built.value = false;
            else built.value = raw;
          }
          return { ok: true, op: built };
        }
        if (action === 'row') {
          const op = typeof f.op === 'string' ? f.op : 'insert';
          if (!['insert', 'delete'].includes(op)) return invalid();
          const at = parsePositiveInt(f.at);
          if (at === null) return invalid();
          const built = { action: 'row', op, at };
          if (sheet) built.sheet = sheet;
          return { ok: true, op: built };
        }
        if (action === 'column') {
          const op = typeof f.op === 'string' ? f.op : 'insert';
          if (!['insert', 'delete'].includes(op)) return invalid();
          const raw = String(f.at ?? '').trim().toUpperCase();
          if (!/^[A-Z]{1,3}$/.test(raw) && parsePositiveInt(raw) === null) return invalid();
          const built = { action: 'column', op, at: /^[A-Z]+$/.test(raw) ? raw : parsePositiveInt(raw) };
          if (sheet) built.sheet = sheet;
          return { ok: true, op: built };
        }
        if (action === 'sheet') {
          const op = typeof f.op === 'string' ? f.op : 'rename';
          if (!['rename', 'add', 'delete'].includes(op)) return invalid();
          const name = String(f.name ?? '').trim();
          if (!name) return invalid();
          const built = { action: 'sheet', op, name };
          if (op === 'rename') {
            const newName = String(f.newName ?? '').trim();
            if (!newName) return invalid();
            built.newName = newName;
          }
          return { ok: true, op: built };
        }
        return invalid();
      }
      return invalid();
    }

    /**
     * 构造一条排版规则（P2 §3.3）。
     * @returns {{ok:true, rule:object}|{ok:false, message:string}}
     */
    function buildFormatRule(t, form) {
      const f = form || {};
      const invalid = () => ({ ok: false, message: t('docs.edit.formInvalid') });
      const target = typeof f.target === 'string' ? f.target : 'body';
      if (!['heading', 'body', 'page', 'header', 'footer'].includes(target)) return invalid();
      const rule = { target };
      if (target === 'heading') {
        const level = parsePositiveInt(f.level);
        if (level === null || level > 3) return invalid();
        rule.level = level;
      }
      if (target === 'header' || target === 'footer') {
        rule[target] = { text: String(target === 'header' ? (f.headerText ?? '') : (f.footerText ?? '')) };
        return { ok: true, rule };
      }
      if (target === 'page') {
        const margins = {};
        for (const side of ['top', 'bottom', 'left', 'right']) {
          const value = parseOptionalNumber(f['margin.' + side]);
          if (value === null || (value !== undefined && (value < 1 || value > 500))) return invalid();
          if (value !== undefined) margins[side] = value;
        }
        if (!Object.keys(margins).length) return invalid();
        rule.page = { marginsMm: margins };
        return { ok: true, rule };
      }
      const font = {};
      const fontName = String(f.fontName ?? '').trim();
      if (fontName) font.name = fontName;
      const size = parseOptionalNumber(f.fontSize);
      if (size === null || (size !== undefined && (size < 1 || size > 200))) return invalid();
      if (size !== undefined) font.sizePt = size;
      if (f.bold === true) font.bold = true;
      if (Object.keys(font).length) rule.font = font;
      const paragraph = {};
      const lineSpacing = parseOptionalNumber(f.lineSpacing);
      if (lineSpacing === null || (lineSpacing !== undefined && (lineSpacing < 0.5 || lineSpacing > 10))) return invalid();
      if (lineSpacing !== undefined) paragraph.lineSpacing = lineSpacing;
      const before = parseOptionalNumber(f.spaceBefore);
      if (before === null || (before !== undefined && (before < 0 || before > 500))) return invalid();
      if (before !== undefined) paragraph.spaceBeforePt = before;
      const after = parseOptionalNumber(f.spaceAfter);
      if (after === null || (after !== undefined && (after < 0 || after > 500))) return invalid();
      if (after !== undefined) paragraph.spaceAfterPt = after;
      if (typeof f.align === 'string' && f.align) {
        if (!['left', 'center', 'right', 'justify'].includes(f.align)) return invalid();
        paragraph.align = f.align;
      }
      if (Object.keys(paragraph).length) rule.paragraph = paragraph;
      if (typeof f.numbering === 'string' && f.numbering) {
        if (!['chinese', 'arabic', 'decimalDot', 'none'].includes(f.numbering)) return invalid();
        rule.numbering = f.numbering;
      }
      if (!rule.font && !rule.paragraph && rule.numbering === undefined) return { ok: false, message: t('docs.format.ruleEmpty') };
      return { ok: true, rule };
    }

    /**
     * 构造 suite.format / suite.format.batch 参数（rules / styleSet / template / placeholders / scope）。
     * @returns {{ok:true, args:object}|{ok:false, message:string}}
     */
    function buildFormatSpec(t, state) {
      const s = state || {};
      const args = {};
      const rules = Array.isArray(s.rules) ? s.rules : [];
      if (rules.length) args.rules = rules;
      const styleSet = String(s.styleSet ?? '').trim();
      if (styleSet) args.styleSet = styleSet;
      const template = String(s.template ?? '').trim();
      if (template) args.template = template;
      if (template) {
        const placeholders = {};
        for (const row of Array.isArray(s.placeholders) ? s.placeholders : []) {
          const key = String(row && row.key !== undefined ? row.key : '').trim();
          if (!key) continue;
          placeholders[key] = String(row.value ?? '');
        }
        if (Object.keys(placeholders).length) args.placeholders = placeholders;
      }
      if (s.scope === 'headings') {
        const levels = (Array.isArray(s.headingLevels) ? s.headingLevels : [1, 2, 3])
          .map((value) => Number(value))
          .filter((value) => Number.isInteger(value) && value >= 1 && value <= 3);
        if (!levels.length) return { ok: false, message: t('docs.format.needRule') };
        args.scope = { headings: [...new Set(levels)].sort((a, b) => a - b) };
      } else {
        args.scope = 'all';
      }
      if (!args.rules && !args.styleSet && !args.template) return { ok: false, message: t('docs.format.needRule') };
      return { ok: true, args };
    }

    /** 每个电网图类型的必填字段（与 host/ops/grid-model.js 的校验口径对齐）。 */
    const GRID_REQUIRED_FIELDS = {
      singleLine: ['buses', 'bays'],
      distribution: ['sources', 'feeders'],
      networkTopology: ['rings'],
      protection: ['circuits'],
      substationPlan: ['plot', 'zones'],
      geoNetwork: ['stations'],
      powerFlow: ['levels', 'buses'],
      terminalStrip: ['strips'],
    };

    /**
     * 轻量模型校验（P2 §6.2）：JSON 解析 + kind 匹配 + 必填字段类型。
     * 结构细节与符号引用由 Host（E2021/E2022）给出，Client 不重复实现完整 schema。
     */
    function parseGridModel(t, kind, text) {
      let model;
      try {
        model = JSON.parse(text);
      } catch (error) {
        return { ok: false, message: fmt(t('visio.json.invalid'), { message: error.message }) };
      }
      if (!isPlainObjectLoose(model)) return { ok: false, message: t('visio.model.notObject') };
      if (model.kind !== undefined && model.kind !== null && model.kind !== kind) {
        return { ok: false, message: fmt(t('visio.model.kindMismatch'), { actual: String(model.kind), expected: kind }) };
      }
      const required = GRID_REQUIRED_FIELDS[kind] || [];
      for (const key of required) {
        const value = model[key];
        const expectedObject = key === 'plot';
        if (value === undefined || value === null || (expectedObject ? !isPlainObjectLoose(value) : !Array.isArray(value))) {
          return { ok: false, message: fmt(t('visio.model.required'), { key, kind }) };
        }
      }
      return { ok: true, model };
    }

    /** 电网图类型 → i18n 键（dynamic 使用）。 */
    function gridKindLabel(t, kind) {
      return t('visio.kind.' + kind);
    }

    /** 统计字段 → i18n 键（未收录的字段原样展示）。 */
    const STAT_LABEL_KEYS = {
      symbols: 'visio.stats.symbols',
      shapes: 'visio.stats.shapes',
      links: 'visio.stats.links',
      nodes: 'visio.stats.nodes',
      edges: 'visio.stats.edges',
      items: 'visio.stats.items',
      levels: 'visio.stats.levels',
      buses: 'visio.stats.buses',
      bays: 'visio.stats.bays',
      groups: 'visio.stats.groups',
      circuits: 'visio.stats.circuits',
      terminals: 'visio.stats.terminals',
      strips: 'visio.stats.strips',
      stations: 'visio.stats.stations',
      lines: 'visio.stats.lines',
      flows: 'visio.stats.flows',
      transformers: 'visio.stats.transformers',
      zones: 'visio.stats.zones',
    };

    /**
     * 电网图模型模板（与 docs/P2-设计.md §4.1 / docs/P3-设计.md §4.4 的 schema 一致的最小示例）。
     * 说明：P3 §4.5 的 KCOIL/KA/KB/XT 符号尚未随包（属符号库交付），protection 模板暂用
     * 现有符号 TA/QF/J 表达同类图元，保证「插入模板 → 运行」链路可用。
     */
    const GRID_TEMPLATES = {
      singleLine: {
        kind: 'singleLine',
        title: '110kV 白云变电气主接线图（示例）',
        meta: { paper: 'A3-landscape', standard: 'GB/T 4728-2022' },
        buses: [
          { id: 'B1', label: '110kV I 母', voltage: 110, section: 'I' },
          { id: 'B2', label: '10kV I 母', voltage: 10, section: 'I' },
        ],
        nodes: [],
        bays: [
          { id: 'X1', bus: 'B1', label: '1# 主变高压侧', devices: [
            { id: 'QS11', symbolRef: 'gb4728.7.QS', state: 'closed' },
            { id: 'QF11', symbolRef: 'gb4728.7.QF', state: 'closed' },
            { id: 'QS12', symbolRef: 'gb4728.7.QS', state: 'closed' },
          ] },
          { id: 'X5', bus: 'B2', label: '1# 主变低压侧', devices: [
            { id: 'QF21', symbolRef: 'gb4728.7.QF', state: 'closed' },
            { id: 'QS21', symbolRef: 'gb4728.7.QS', state: 'closed' },
          ] },
        ],
        transformers: [
          { id: 'T1', symbolRef: 'gb4728.6.T2', label: '1# 主变（110/10kV）', fromBay: 'X1', toBay: 'X5', voltage: [110, 10] },
        ],
      },
      distribution: {
        kind: 'distribution',
        title: '10kV 配电系统图（示例）',
        meta: { paper: 'A3-landscape', standard: 'GB/T 4728-2022' },
        sources: [
          { id: 'S1', symbolRef: 'gb4728.3.BUS', label: '110kV 站 10kV 母线', tags: { source: true } },
        ],
        nodes: [
          { id: 'R1', symbolRef: 'gb4728.7.RING-MAIN', label: '1# 环网柜' },
          { id: 'T1', symbolRef: 'gb4728.6.T2', label: '1# 台区变压器' },
        ],
        feeders: [
          { id: 'F1', label: '馈线 101', from: 'S1', to: 'R1', kind: 'feeder' },
          { id: 'F2', label: '馈线 102', from: 'R1', to: 'T1', kind: 'feeder' },
        ],
        links: [],
      },
      networkTopology: {
        kind: 'networkTopology',
        title: '光纤环网拓扑图（示例）',
        meta: { paper: 'A3-landscape', standard: 'GB/T 4728-2022' },
        nodes: [
          { id: 'C1', symbolRef: 'gb4728.9.NODE', label: '变电站 A' },
          { id: 'C2', symbolRef: 'gb4728.9.NODE', label: '变电站 B' },
          { id: 'C3', symbolRef: 'gb4728.9.NODE', label: '变电站 C' },
        ],
        rings: [{ id: 'RING1', label: '光纤环 1', nodes: ['C1', 'C2', 'C3'] }],
        links: [
          { id: 'K1', from: 'C1', to: 'C2', kind: 'fiber', label: '光纤 1' },
          { id: 'K2', from: 'C2', to: 'C3', kind: 'fiber' },
          { id: 'K3', from: 'C3', to: 'C1', kind: 'fiber' },
        ],
      },
      protection: {
        kind: 'protection',
        title: '110kV 线路保护电流回路图（示例）',
        meta: { paper: 'A3-landscape', standard: 'GB/T 4728-2022' },
        circuits: [
          { id: 'C1', label: '电流回路', voltage: 'DC110V', devices: [
            { id: 'TAa', symbolRef: 'gb4728.6.TA', label: '1TAa' },
            { id: 'QF1', symbolRef: 'gb4728.7.QF', label: '保护出口' },
            { id: 'XT1', symbolRef: 'gb4728.3.J', label: 'X1:1' },
          ] },
          { id: 'C2', label: '控制回路', voltage: 'DC110V', devices: [
            { id: 'QF2', symbolRef: 'gb4728.7.QF', label: '控制开关' },
            { id: 'XT2', symbolRef: 'gb4728.3.J', label: 'X1:2' },
          ] },
        ],
        loops: [{ id: 'L1', from: 'QF1', to: 'QF2', kind: 'signal', label: 'KA1-1' }],
        terminalStrips: [
          { id: 'TP1', label: '端子排 1X', side: 'top', terminals: [{ no: 1, ref: 'A411' }, { no: 2, ref: 'B411' }] },
        ],
      },
      substationPlan: {
        kind: 'substationPlan',
        title: '110kV 白云变总平面布置示意图（示例）',
        meta: { paper: 'A3-landscape', compass: true },
        plot: { widthMm: 800, heightMm: 600, gridMm: 50 },
        zones: [
          { id: 'Z1', label: '110kV 配电装置区', xMm: 0, yMm: 0, widthMm: 300, heightMm: 200 },
          { id: 'Z2', label: '主变区', xMm: 300, yMm: 0, widthMm: 200, heightMm: 200 },
          { id: 'Z3', label: '综合楼及生活区', xMm: 0, yMm: 200, widthMm: 500, heightMm: 200 },
        ],
        items: [
          { id: 'T1', label: '1# 主变', symbolRef: 'gb4728.6.T2', xMm: 350, yMm: 60, rot: 0 },
          { id: 'RD1', label: '进站道路', xMm: 0, yMm: 400, widthMm: 500, heightMm: 24, primitive: 'rect' },
          { id: 'F1', label: '站区围墙', xMm: 0, yMm: 0, widthMm: 800, heightMm: 600, primitive: 'rect', dashed: true },
        ],
        links: [],
      },
      geoNetwork: {
        kind: 'geoNetwork',
        title: '××供电区 35kV 及以上地理接线示意图（示例）',
        meta: { paper: 'A3-landscape' },
        viewport: { west: 112.0, south: 30.0, east: 114.0, north: 32.0 },
        stations: [
          { id: 'S1', label: '白云变', symbolRef: 'gb4728.9.NODE', lng: 113.2, lat: 31.1, voltage: 110 },
          { id: 'S2', label: '金亭变', symbolRef: 'gb4728.9.NODE', lng: 113.6, lat: 31.4, voltage: 110 },
        ],
        lines: [
          { id: 'LN1', label: '白-金 I 回', from: 'S1', to: 'S2', voltage: 110, points: [['S1', null], [113.4, 31.3], ['S2', null]] },
        ],
        labels: [{ id: 'LBL1', text: '××工业园', lng: 113.1, lat: 30.9 }],
      },
      powerFlow: {
        kind: 'powerFlow',
        title: '××电网 110kV 系统接线潮流图（示例）',
        meta: { paper: 'A3-landscape' },
        levels: [
          { id: 'LV110', voltage: 110, label: '110kV' },
          { id: 'LV35', voltage: 35, label: '35kV' },
        ],
        buses: [
          { id: 'B1', level: 'LV110', label: '白云 110kV 母线' },
          { id: 'B2', level: 'LV35', label: '白云 35kV 母线' },
        ],
        nodes: [
          { id: 'G1', symbolRef: 'gb4728.6.G', label: '金亭电厂', level: 'LV110', tags: { source: true } },
          { id: 'S3', symbolRef: 'gb4728.9.NODE', label: '外部系统', level: 'LV110', tags: { source: true } },
        ],
        transformers: [{ id: 'T1', symbolRef: 'gb4728.6.T2', label: '白 1# 主变', from: 'B1', to: 'B2' }],
        flows: [
          { id: 'F1', from: 'B1', to: 'B2', kind: 'flow', direction: 'B1→B2', label: 'P=85.4MW Q=12.1Mvar', style: 'solid' },
          { id: 'F2', from: 'G1', to: 'B1', kind: 'flow', direction: 'G1→B1', label: 'P=120MW' },
        ],
        links: [{ id: 'L1', from: 'B1', to: 'S3' }],
      },
      terminalStrip: {
        kind: 'terminalStrip',
        title: '端子排图 1X（示例）',
        meta: { paper: 'A4-landscape', terminalsPerPage: 20 },
        strips: [
          { id: 'XT1', label: '端子排 1X（保护屏）', columns: 1, terminals: [
            { no: 1, ref: 'A411', desc: '1TAa 电流', type: 'normal', link: null },
            { no: 2, ref: 'B411', desc: '1TAb 电流', type: 'shorting', link: 1 },
            { no: 3, ref: 'N411', desc: '1TAn 中性', type: 'ground', link: null },
          ] },
        ],
      },
    };

    /** 模板 JSON 文本。 */
    function gridTemplateText(kind) {
      const template = GRID_TEMPLATES[kind];
      return template ? JSON.stringify(template, null, 2) : '';
    }

    /** 判断文本是否仍等于某个内置模板（用于切换类型时安全地刷新模板）。 */
    function isBuiltinTemplateText(text) {
      const trimmed = String(text ?? '').trim();
      if (!trimmed) return true;
      return GRID_MODEL_KINDS.some((kind) => gridTemplateText(kind).trim() === trimmed);
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

    //#region 4. 命令路由封装（callCommand / listWorkspace / externalList / subscribeProgress）
    /**
     * 防御式读取命名空间服务。`remote` / `uiWorkspace` 不写进 inject 硬依赖，
     * 缺挂时返回 null 由调用方降级。
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

    /** 面板 → Host 的 RPC 路由（插件在 Host 侧自注册，与命令注册通道解耦）。 */
    const RPC_PATH = '/office-suite/rpc';

    /** 同源相对地址优先；必要时回退到 location.origin 拼出的绝对地址。 */
    function rpcUrls() {
      const urls = [RPC_PATH];
      try {
        if (typeof location !== 'undefined' && location && typeof location.origin === 'string' && location.origin) {
          const absolute = location.origin + RPC_PATH;
          if (absolute !== RPC_PATH) urls.push(absolute);
        }
      } catch { /* location 不可用时只保留相对地址 */ }
      return urls;
    }

    /** 从路由错误响应中尽量取出可读消息（响应体可能不是 JSON）。 */
    function rpcErrorMessage(status, body) {
      const error = body && typeof body === 'object' ? body.error : undefined;
      if (error && typeof error.message === 'string' && error.message) return error.message;
      if (body && typeof body.text === 'string') {
        const parsed = parseJsonLoose(body.text);
        const text = parsed && parsed.error && typeof parsed.error.message === 'string' ? parsed.error.message : '';
        if (text) return text;
      }
      if (body && typeof body.message === 'string' && body.message) return body.message;
      return 'RPC HTTP ' + status;
    }

    /** fetch 取消统一映射为旧网关的 gateway/cancelled（toPanelError 已识别）。 */
    function cancelledError() {
      return { code: 'gateway/cancelled', message: 'aborted' };
    }

    /**
     * 面板 → Host：POST `/office-suite/rpc`，body = `{command, args}`；
     * 响应信封沿用命令层现状 `{kind:'success'|'error', text}`（§2）：
     *   - success：按既有 parseJsonLoose + `{ok,result}` 逻辑展开；
     *   - error：沿用既有 normalizeThrown / COMMAND_FAILED 逻辑；
     *   - 404/405：保留 COMMAND_MISSING（路由未注册 / 旧页面）。
     * 不再依赖 remote.commands（typert descriptor 不一致）。
     */
    async function executeCommand(_ctx, name, args, signal) {
      const payload = JSON.stringify({ command: name, args: args === undefined ? {} : args });
      let response = null;
      let lastError = null;
      for (const url of rpcUrls()) {
        try {
          response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: payload,
            signal,
          });
          lastError = null;
          break;
        } catch (error) {
          if ((error && error.name === 'AbortError') || (signal && signal.aborted)) throw cancelledError();
          lastError = error;
        }
      }
      if (!response) {
        throw normalizeThrown({
          code: 'RPC_NETWORK',
          message: '无法连接面板 RPC 路由' + (lastError && lastError.message ? '：' + lastError.message : ''),
        });
      }

      // 路由未注册（如 Host 半未升级）时保持原命令缺失语义。
      if (response.status === 404 || response.status === 405) {
        throw { code: 'COMMAND_MISSING', message: 'command route not found: ' + name };
      }

      let body = null;
      try {
        body = await response.json();
      } catch (error) {
        if ((error && error.name === 'AbortError') || (signal && signal.aborted)) throw cancelledError();
        throw { code: 'RPC_HTTP_' + response.status, message: rpcErrorMessage(response.status, null) };
      }
      if (!response.ok) {
        throw { code: 'RPC_HTTP_' + response.status, message: rpcErrorMessage(response.status, body) };
      }
      if (!body || typeof body !== 'object' || (body.kind !== 'success' && body.kind !== 'error')) {
        throw { code: 'BAD_RESPONSE', message: name + ': rpc returned an invalid envelope' };
      }
      if (body.kind === 'error') {
        const parsed = parseJsonLoose(body.text);
        if (parsed && parsed.error) throw normalizeThrown(parsed.error);
        throw { code: 'COMMAND_FAILED', message: typeof body.text === 'string' && body.text ? body.text : 'command failed' };
      }
      const parsed = parseJsonLoose(body.text);
      if (!parsed) {
        throw { code: 'BAD_RESPONSE', message: name + ': command returned non-JSON output' };
      }
      if (parsed.ok === false) throw normalizeThrown(parsed.error);
      return parsed.result;
    }

    /** 工作区目录列举（经 RPC 路由 workspace.list；result 形状同 suite.files）。 */
    async function workspaceList(ctx, path, signal) {
      return executeCommand(ctx, 'workspace.list', path ? { path } : {}, signal);
    }

    /** 工作区文件元信息（经 RPC 路由 workspace.stat：bytes/version/absolutePath）。 */
    async function workspaceStat(ctx, path, signal) {
      return executeCommand(ctx, 'workspace.stat', { path }, signal);
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
          }, h(DocsTab, { api, t, sessionId, active: tab === 'docs', openDir, openInDocs })),
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

    //#region 6b. 通用任务轮询 Hook 与进度组件（P2/P3 批量：suite.office.status / suite.cad.status，800ms）
    /**
     * 轮询任务状态（与 CadTab 既有单任务轮询同款：800ms、终态停止、AbortController 联动卸载）。
     * @param {object} api 面板 api 袋
     * @param {string} taskId 任务 id（空串 = 不轮询）
     * @param {string} commandName 'suite.cad.status' | 'suite.office.status'
     * @param {Function} t i18n
     * @returns {{status:object|null, error:object|null}}
     */
    function useTaskPolling(api, taskId, commandName, t) {
      const [status, setStatus] = React.useState(null);
      const [error, setError] = React.useState(null);
      const tRef = React.useRef(t);
      tRef.current = t;

      React.useEffect(() => {
        setStatus(null);
        setError(null);
        if (!taskId) return undefined;
        let alive = true;
        let timer = 0;
        let inFlight = false;
        const controller = new AbortController();
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
            const snapshot = await api.callCommand(commandName, { taskId }, controller.signal);
            if (!alive) return;
            setStatus(snapshot);
            setError(null);
            const state = snapshot && typeof snapshot.state === 'string' ? snapshot.state : '';
            if (TASK_TERMINAL.includes(state)) stop();
          } catch (thrown) {
            if (!alive) return;
            const normalized = toPanelError(thrown, tRef.current);
            if (!normalized.cancelled) setError(normalized);
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
          controller.abort();
        };
      }, [api, commandName, taskId]);

      return { status, error };
    }

    /**
     * 通用任务进度块（状态徽标 + 进度条 + 当前项 + 日志 + 取消）。
     * @param {{t:Function,status:object|null,error:object|null,mode?:'files'|'items',
     *          onCancel?:Function,cancelling?:boolean,extra?:object|null}} props
     */
    function TaskProgress({ t, status, error, mode, onCancel, cancelling, extra }) {
      const logRef = React.useRef(null);
      React.useEffect(() => {
        const element = logRef.current;
        if (element) element.scrollTop = element.scrollHeight;
      }, [status]);
      if (!status && !error) return null;
      const progress = (status && status.progress) || {};
      const filesDone = Number(progress.filesDone) || 0;
      const filesTotal = Number(progress.filesTotal) || 0;
      const itemsDone = Number(progress.itemsDone) || 0;
      const itemsTotal = Number(progress.itemsTotal) || 0;
      const useItems = mode === 'items';
      const doneCount = useItems ? itemsDone : filesDone;
      const totalCount = useItems ? itemsTotal : filesTotal;
      const state = status && typeof status.state === 'string' ? status.state : 'queued';
      const percent = totalCount > 0
        ? Math.min(100, Math.round((doneCount / totalCount) * 100))
        : (state === 'done' ? 100 : 0);
      const logLines = status && Array.isArray(status.log) ? status.log.slice(-200) : [];
      const stateLabel = CAD_STATES.includes(state) ? t('cad.state.' + state) : state;
      const progressLabel = useItems
        ? fmt(t('cad.queue.progress'), { done: itemsDone, total: itemsTotal })
        : fmt(t('cad.progress.files'), { done: filesDone, total: filesTotal });
      const running = !TASK_TERMINAL.includes(state);
      return h('div', { className: 'dso-stack' },
        error ? h(ErrorBox, { t, error }) : null,
        h('div', { className: 'dso-inline' },
          h('span', { className: cadStateClass(state) }, stateLabel),
          h('span', { className: 'dso-muted' }, progressLabel),
          h('span', { className: 'dso-spacer' }),
          h('span', { className: 'dso-faint' }, percent + '%')),
        h('div', {
          className: 'dso-progress',
          role: 'progressbar',
          'aria-valuemin': 0,
          'aria-valuemax': 100,
          'aria-valuenow': percent,
        }, h('div', { className: 'dso-progressBar', style: { width: percent + '%' } })),
        progress.currentFile
          ? h('div', { className: 'dso-hint' }, fmt(t('cad.progress.current'), { file: progress.currentFile }))
          : null,
        extra || null,
        logLines.length === 0
          ? h('div', { className: 'dso-hint' }, t('cad.log.empty'))
          : h('div', { className: 'dso-log', ref: logRef }, logLines.map((line, index) => h('div', {
            key: 'tlog' + index,
            className: 'dso-logLine',
            'data-level': line && typeof line.level === 'string' ? line.level : 'info',
          },
          h('span', { className: 'dso-logTime' }, formatClock(line && line.t ? line.t : Date.now())),
          h('span', null, line && line.msg !== undefined ? String(line.msg) : String(line))))),
        onCancel && running
          ? h('div', { className: 'dso-actions' },
            h('button', {
              type: 'button',
              className: 'dso-btn dso-btn--danger',
              onClick: onCancel,
              disabled: cancelling === true,
            }, cancelling ? t('cad.cancel.requested') : t('cad.cancel')))
          : null);
    }
    //#endregion

    //#region 7. 组件：DocsTab（文档中心：工作区文件树 / 外部目录模式 + 读取渲染 + 编辑/排版/预览）
    /** 由 PDF 字节创建 blob URL（环境不支持时抛 E2012 语义错误）。 */
    function createPdfObjectUrl(bytes) {
      if (!bytes || !bytes.length) {
        throw { code: 'E2011', message: 'preview returned empty pdf' };
      }
      if (typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function' || typeof Blob === 'undefined') {
        throw { code: 'PREVIEW_UNSUPPORTED', message: 'URL.createObjectURL is unavailable' };
      }
      return URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
    }

    /**
     * 渲染预览（P3 §5.2 方案 B）：`suite.preview` 命令返回 pdfBase64 → Blob URL。
     * 统一走服务端命令通道（原 remote.officeToPdf 直连方案 A 已移除）；
     * E2012 由调用方提示不可用，其余错误按现状透传。
     * @returns {Promise<{url:string, missingFonts:string[], scheme:'command', generation:any}>}
     */
    async function renderPdfPreview(api, path, signal) {
      try {
        const result = await api.callCommand('suite.preview', { path }, signal);
        const bytes = base64ToBytes(result && result.pdfBase64);
        return {
          url: createPdfObjectUrl(bytes),
          missingFonts: result && Array.isArray(result.missingFonts) ? result.missingFonts.map((item) => String(item)) : [],
          scheme: 'command',
          generation: result && result.generation !== undefined ? result.generation : null,
        };
      } catch (error) {
        throw normalizeThrown(error);
      }
    }

    /** 文档中心：左树右阅读区。 */
    function DocsTab({ api, t, sessionId, active, openDir, openInDocs }) {
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

      /* P2/P3 增量：编辑 / 排版 / 预览子页签 + 预览 blob 缓存 + 导出 PDF */
      const [subTab, setSubTab] = React.useState('edit');
      const [preview, setPreview] = React.useState({ status: 'idle', url: '', missingFonts: [], scheme: '', generation: null, key: '', error: null });
      const [previewMax, setPreviewMax] = React.useState(PREVIEW_CACHE_DEFAULT);
      const [exporting, setExporting] = React.useState(false);
      const [exportResult, setExportResult] = React.useState(null);
      const [exportError, setExportError] = React.useState(null);
      const previewCacheRef = React.useRef(new Map());
      const previewAbortRef = React.useRef(null);
      const configLoadedRef = React.useRef(false);
      const previewMaxRef = React.useRef(PREVIEW_CACHE_DEFAULT);
      previewMaxRef.current = previewMax;

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
        if (previewAbortRef.current) previewAbortRef.current.abort();
        for (const entry of previewCacheRef.current.values()) {
          try { URL.revokeObjectURL(entry.url); } catch { /* 忽略 */ }
        }
        previewCacheRef.current.clear();
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
        // 预览 blob 缓存按文件路径键控；跨会话可能指向不同工作区的同名文件 → 清空。
        if (previewAbortRef.current) previewAbortRef.current.abort();
        for (const entry of previewCacheRef.current.values()) {
          try { URL.revokeObjectURL(entry.url); } catch { /* 忽略 */ }
        }
        previewCacheRef.current.clear();
        setPreview({ status: 'idle', url: '', missingFonts: [], scheme: '', generation: null, key: '', error: null });
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

      // 选中文件变化：重置子页签与预览状态（blob 缓存保留，切回时可秒开）。
      const selectedPath = selected && selected.type === 'file' ? selected.path : '';
      React.useEffect(() => {
        if (!selectedPath) return;
        const ext = extOf(selectedPath);
        setSubTab(WORD_EXTS.includes(ext) || EXCEL_EXTS.includes(ext) ? 'edit' : 'preview');
        setPreview((prev) => (prev.key === selectedPath ? prev : { status: 'idle', url: '', missingFonts: [], scheme: '', generation: null, key: '', error: null }));
        setExportResult(null);
        setExportError(null);
      }, [selectedPath]);

      // 预览缓存容量 previewCacheSize（1–10；配置不可读时回落默认 5）。
      React.useEffect(() => {
        if (!active || configLoadedRef.current) return;
        configLoadedRef.current = true;
        api.callCommand('suite.config.get', {}).then((snapshot) => {
          const effective = snapshot && isPlainObjectLoose(snapshot.effective) ? snapshot.effective : snapshot;
          const value = Number(effective && effective.previewCacheSize);
          if (Number.isInteger(value) && value >= 1 && value <= 10) setPreviewMax(value);
        }).catch(() => { /* 配置不可读时回落默认 5 */ });
      }, [active, api]);

      // 连接复位：清空预览缓存并释放 blob URL（P3 §5.2）。
      useConnectionReset(api, () => {
        if (previewAbortRef.current) previewAbortRef.current.abort();
        for (const entry of previewCacheRef.current.values()) {
          try { URL.revokeObjectURL(entry.url); } catch { /* 忽略 */ }
        }
        previewCacheRef.current.clear();
        setPreview({ status: 'idle', url: '', missingFonts: [], scheme: '', generation: null, key: '', error: null });
      });

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

      /**
       * 渲染预览（P3 §5.2 + LRU 缓存）：
       * 命中缓存直接复用 blob URL；未命中统一走 suite.preview 服务端命令通道（返回 pdfBase64）。
       */
      const renderPreview = React.useCallback(async (force) => {
        const file = selected;
        if (!file || file.type !== 'file' || !PREVIEW_EXTS.includes(extOf(file.name))) return;
        if (!sessionId) {
          setPreview({ status: 'error', url: '', missingFonts: [], scheme: '', generation: null, key: file.path, error: { code: 'NO_SESSION', message: t('common.session.missing') } });
          return;
        }
        const key = file.path;
        const cache = previewCacheRef.current;
        if (!force && cache.has(key)) {
          const entry = cache.get(key);
          cache.delete(key);
          cache.set(key, entry);
          setPreview({ status: 'ready', url: entry.url, missingFonts: entry.missingFonts, scheme: entry.scheme, generation: entry.generation, key, error: null });
          return;
        }
        if (previewAbortRef.current) previewAbortRef.current.abort();
        const controller = new AbortController();
        previewAbortRef.current = controller;
        setPreview({ status: 'busy', url: '', missingFonts: [], scheme: '', generation: null, key, error: null });
        try {
          const rendered = await renderPdfPreview(api, key, controller.signal);
          if (controller.signal.aborted) {
            try { URL.revokeObjectURL(rendered.url); } catch { /* 忽略 */ }
            return;
          }
          const existing = cache.get(key);
          if (existing) {
            try { URL.revokeObjectURL(existing.url); } catch { /* 忽略 */ }
            cache.delete(key);
          }
          const max = Math.max(1, Math.min(10, Number(previewMaxRef.current) || PREVIEW_CACHE_DEFAULT));
          while (cache.size >= max) {
            const oldestKey = cache.keys().next().value;
            const oldest = cache.get(oldestKey);
            if (oldest) {
              try { URL.revokeObjectURL(oldest.url); } catch { /* 忽略 */ }
            }
            cache.delete(oldestKey);
          }
          cache.set(key, { url: rendered.url, missingFonts: rendered.missingFonts, scheme: rendered.scheme, generation: rendered.generation });
          setPreview({ status: 'ready', url: rendered.url, missingFonts: rendered.missingFonts, scheme: rendered.scheme, generation: rendered.generation, key, error: null });
        } catch (thrown) {
          if (controller.signal.aborted) return;
          const normalized = toPanelError(thrown, t);
          if (normalized.cancelled) return;
          const unsupported = normalized.code === 'E2012' || normalized.code === 'PREVIEW_UNSUPPORTED';
          setPreview({ status: unsupported ? 'unsupported' : 'error', url: '', missingFonts: [], scheme: '', generation: null, key, error: normalized });
        } finally {
          if (previewAbortRef.current === controller) previewAbortRef.current = null;
        }
      }, [api, selected, sessionId, t]);

      /** 导出 PDF（P3 §5.3：suite.export，新文件不触发备份）。 */
      const exportPdf = React.useCallback(async () => {
        if (!selected || selected.type !== 'file' || exporting) return;
        if (!sessionId) {
          setExportError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        setExporting(true);
        setExportError(null);
        setExportResult(null);
        try {
          const result = await api.callCommand('suite.export', { path: selected.path });
          setExportResult(result && typeof result === 'object' ? result : {});
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setExportError(normalized);
        } finally {
          setExporting(false);
        }
      }, [api, exporting, selected, sessionId, t]);

      const selectedExt = selected && selected.type === 'file' ? extOf(selected.name) : '';
      const selectedIsWord = WORD_EXTS.includes(selectedExt);
      const selectedIsExcel = EXCEL_EXTS.includes(selectedExt);
      const selectedEditable = OFFICE_EXTS.includes(selectedExt);
      const selectedPreviewable = selected && selected.type === 'file' && PREVIEW_EXTS.includes(selectedExt);
      const previewForCurrent = preview.key === selectedPath ? preview : { status: 'idle', url: '', missingFonts: [], scheme: '', generation: null, key: selectedPath, error: null };

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
          body = h(Busy, { text: t('docs.reading') });
        } else if (readResult && readResult.path === selected.path) {
          body = renderOfficeResult(t, readResult.result);
        } else {
          body = h(Empty, { text: t('docs.read.prompt') });
        }
      } else if (PREVIEW_EXTS.includes(selectedExt)) {
        // ppt/pptx：无文本读取，仅页面预览（P3 §5.2）
        body = h('div', { className: 'dso-stack' },
          h('div', { className: 'dso-hint' }, t('docs.preview.hint.other')),
          h('div', { className: 'dso-hint' }, t('docs.preview.empty')));
      } else if (META_EXTS.includes(selectedExt)) {
        const hint = selectedExt === 'pdf' ? t('docs.preview.hint.pdf') : t('docs.preview.hint.other');
        const info = meta && meta.path === selected.path ? meta : null;
        body = h('div', { className: 'dso-stack', 'data-preview-hook': selectedExt === 'pdf' ? 'officeToPdf' : undefined },
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

      // ── 编辑 / 排版 / 预览子页签（P2 §6.1 / P3 §5.2） ──
      const subTabs = selectedPreviewable
        ? [
          selectedEditable ? { id: 'edit', label: 'docs.sub.edit' } : null,
          selectedIsWord ? { id: 'format', label: 'docs.sub.format' } : null,
          { id: 'preview', label: 'docs.sub.preview' },
        ].filter(Boolean)
        : [];
      const activeSub = subTabs.some((item) => item.id === subTab) ? subTab : (subTabs.length ? subTabs[0].id : '');
      const subTabBar = subTabs.length === 0 ? null : h('div', {
        className: 'dso-subtabs',
        role: 'tablist',
        'aria-label': t('docs.title'),
      }, subTabs.map((item) => h('button', {
        key: item.id,
        type: 'button',
        role: 'tab',
        className: 'dso-subtab',
        'aria-selected': activeSub === item.id ? 'true' : 'false',
        'aria-current': activeSub === item.id ? 'page' : undefined,
        onClick: () => {
          setSubTab(item.id);
          if (item.id === 'preview' && previewForCurrent.status === 'idle') renderPreview(false);
        },
      }, t(item.label))));

      let subPanel = null;
      if (selectedPreviewable) {
        if (activeSub === 'edit') {
          subPanel = h(DocsEditPanel, { api, t, path: selected.path, ext: selectedExt, sessionId });
        } else if (activeSub === 'format') {
          subPanel = h(DocsFormatPanel, { api, t, path: selected.path, sessionId });
        } else {
          subPanel = h(DocsPreviewPanel, {
            t,
            preview: previewForCurrent,
            cacheSize: previewMax,
            cacheCount: previewCacheRef.current.size,
            displayPath: selected.path,
            onRender: () => renderPreview(false),
            onRetry: () => renderPreview(true),
            onExport: exportPdf,
            exporting,
            missingSession: !sessionId,
          });
        }
      }
      const exportArea = exportError
        ? h(ErrorBox, { t, error: exportError, onRetry: exportPdf })
        : (exportResult
          ? h('div', { className: 'dso-okBox' },
            h('div', null, fmt(t('docs.export.done'), { path: exportResult.outputPath || '—' })),
            typeof exportResult.size === 'number' ? h('div', { className: 'dso-hint' }, formatBytes(exportResult.size)) : null,
            h('div', { className: 'dso-actions' },
              h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => api.copy(exportResult.outputPath) }, t('docs.actions.copyPath')),
              openInDocs && exportResult.outputPath
                ? h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => openInDocs(exportResult.outputPath) }, t('cad.result.openInDocs'))
                : null))
          : null);
      const subArea = selectedPreviewable
        ? h('div', { className: 'dso-stack' }, h('div', { className: 'dso-divider' }), subTabBar, subPanel, exportArea)
        : null;

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
            selectedPreviewable
              ? h('button', {
                type: 'button',
                className: 'dso-btn',
                onClick: () => {
                  setSubTab('preview');
                  if (previewForCurrent.status !== 'ready' && previewForCurrent.status !== 'busy') renderPreview(false);
                },
                disabled: !sessionId || previewForCurrent.status === 'unsupported' || previewForCurrent.status === 'busy',
                title: previewForCurrent.status === 'unsupported'
                  ? t('docs.preview.unsupported')
                  : (sessionId ? t('docs.preview.render') : t('common.session.missing')),
              }, previewForCurrent.status === 'busy' ? t('docs.preview.rendering') : t('docs.preview.render'))
              : null,
            selectedPreviewable
              ? h('button', {
                type: 'button',
                className: 'dso-btn',
                onClick: exportPdf,
                disabled: !sessionId || exporting,
                title: sessionId ? t('docs.export') : t('common.session.missing'),
              }, exporting ? t('docs.export.running') : t('docs.export'))
              : null,
            h('button', {
              type: 'button',
              className: 'dso-btn',
              onClick: copyPath,
              disabled: !selected,
            }, copied ? t('docs.actions.copied') : t('docs.actions.copyPath')))),
        h('div', { className: 'dso-cardBody' }, body, subArea));

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

    /** 编辑表单默认值（按家族）。 */
    function defaultEditForm(family) {
      return family === 'excel'
        ? { action: 'cell', op: 'insert', sheet: '', at: 'B3', value: '', formula: '', name: 'Sheet1', newName: '' }
        : { action: 'paragraph', op: 'append', at: '1', text: '', style: '正文', level: '1', table: '1', row: '1', col: '1' };
    }

    /** 排版规则表单默认值。 */
    function defaultFormatRuleForm() {
      return { target: 'body', level: '1', fontName: '宋体', fontSize: '12', bold: false, lineSpacing: '1.5', spaceBefore: '', spaceAfter: '', align: '', numbering: '', 'margin.top': '', 'margin.bottom': '', 'margin.left': '', 'margin.right': '', headerText: '', footerText: '' };
    }

    /**
     * 编辑子页（P2 §6.1）：Word / Excel 表单 → 指令队列 → suite.edit；
     * 底部「自动备份」列表 + 恢复（恢复前二次确认）。
     */
    function DocsEditPanel({ api, t, path, ext, sessionId }) {
      const family = WORD_EXTS.includes(ext) ? 'word' : (EXCEL_EXTS.includes(ext) ? 'excel' : '');
      const [form, setForm] = React.useState(() => defaultEditForm(family));
      const [queue, setQueue] = React.useState([]);
      const [formError, setFormError] = React.useState(null);
      const [applying, setApplying] = React.useState(false);
      const [result, setResult] = React.useState(null);
      const [error, setError] = React.useState(null);
      const [backups, setBackups] = React.useState(null);
      const [backupError, setBackupError] = React.useState(null);
      const [backupBusy, setBackupBusy] = React.useState(false);
      const [restoreId, setRestoreId] = React.useState('');
      const [restoring, setRestoring] = React.useState('');
      const [restoreMsg, setRestoreMsg] = React.useState('');
      const abortRef = React.useRef(null);
      const tRef = React.useRef(t);
      tRef.current = t;

      React.useEffect(() => () => {
        if (abortRef.current) abortRef.current.abort();
      }, []);

      // 切换文件：清空指令队列与结果（备份列表重新拉取）。
      React.useEffect(() => {
        setForm(defaultEditForm(family));
        setQueue([]);
        setFormError(null);
        setResult(null);
        setError(null);
        setBackups(null);
        setBackupError(null);
        setRestoreId('');
        setRestoreMsg('');
      }, [path, family]);

      const loadBackups = React.useCallback(async () => {
        setBackupBusy(true);
        setBackupError(null);
        try {
          const listing = await api.callCommand('suite.office.backup.list', { path });
          setBackups(Array.isArray(listing && listing.backups) ? listing.backups : []);
        } catch (thrown) {
          const normalized = toPanelError(thrown, tRef.current);
          if (normalized.code === 'BACKUP_NONE') setBackups([]);
          else setBackupError(normalized);
        } finally {
          setBackupBusy(false);
        }
      }, [api, path]);

      React.useEffect(() => {
        if (!family) return;
        loadBackups();
      }, [family, loadBackups]);

      const setValue = (key) => (event) => {
        const value = event.target.type === 'checkbox' ? event.target.checked : event.target.value;
        setForm((prev) => Object.assign({}, prev, { [key]: value }));
      };

      const changeAction = (event) => {
        const action = event.target.value;
        const next = defaultEditForm(family);
        next.action = action;
        if (family === 'word') {
          next.op = action === 'paragraph' ? 'append' : (action === 'heading' ? 'setStyle' : 'setCell');
        } else if (action === 'sheet') {
          next.op = 'rename';
        } else if (action === 'cell') {
          next.op = '';
        } else {
          next.op = 'insert';
        }
        setForm((prev) => Object.assign({}, prev, next));
        setFormError(null);
      };

      const addOp = () => {
        const built = buildEditOp(t, family, form);
        if (!built.ok) {
          setFormError(built.message);
          return;
        }
        setQueue((prev) => prev.concat([built.op]));
        setFormError(null);
      };

      const applyEdits = async () => {
        if (!queue.length || applying) {
          if (!queue.length) setFormError(t('docs.edit.queue.empty'));
          return;
        }
        if (abortRef.current) abortRef.current.abort();
        const controller = new AbortController();
        abortRef.current = controller;
        setApplying(true);
        setError(null);
        setResult(null);
        try {
          const applied = await api.callCommand('suite.edit', { path, ops: queue.slice() }, controller.signal);
          setResult(applied || {});
          if (!applied || !Array.isArray(applied.applied) || applied.applied.every((row) => !row || row.ok !== false)) {
            setQueue([]);
          }
          loadBackups();
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setError(normalized);
        } finally {
          setApplying(false);
          abortRef.current = null;
        }
      };

      const restoreBackup = async (id) => {
        if (restoring) return;
        setRestoring(id);
        setBackupError(null);
        setRestoreMsg('');
        try {
          const restored = await api.callCommand('suite.office.backup.restore', { path, id });
          setRestoreId('');
          setRestoreMsg(restored && typeof restored.restoredTo === 'string' ? restored.restoredTo : path);
          loadBackups();
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setBackupError(normalized);
        } finally {
          setRestoring('');
        }
      };

      if (!family) return h('div', { className: 'dso-hint' }, t('docs.read.notSupported'));

      const isWord = family === 'word';
      const actionOptions = isWord ? ['paragraph', 'heading', 'table'] : ['cell', 'row', 'column', 'sheet'];
      let opOptions = [];
      if (isWord) {
        if (form.action === 'paragraph' || form.action === 'heading') opOptions = ['append', 'setText', 'setStyle', 'insert', 'delete'];
        else opOptions = ['setCell', 'addRow', 'deleteRow', 'addColumn', 'deleteColumn'];
        if (form.action === 'heading') opOptions = ['setText', 'setStyle', 'insert', 'delete'];
        if (form.action === 'table') opOptions = ['setCell', 'addRow', 'deleteRow', 'addColumn', 'deleteColumn'];
      } else if (form.action === 'sheet') {
        opOptions = ['rename', 'add', 'delete'];
      } else if (form.action !== 'cell') {
        opOptions = ['insert', 'delete'];
      }

      const actionSelect = h('select', { className: 'dso-select', value: form.action, onChange: changeAction },
        actionOptions.map((action) => h('option', { key: action, value: action }, t('docs.edit.action.' + action))));
      const opSelect = opOptions.length
        ? h('select', { className: 'dso-select', value: opOptions.includes(form.op) ? form.op : opOptions[0], onChange: (event) => setForm((prev) => Object.assign({}, prev, { op: event.target.value })) },
          opOptions.map((op) => h('option', { key: op, value: op }, t('docs.edit.op.' + op))))
        : null;

      const fields = [];
      if (isWord) {
        if (form.action === 'paragraph') {
          if (form.op !== 'append' && form.op !== 'insert') fields.push(formField(t('docs.edit.at'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.at, onChange: setValue('at') })));
          if (form.op === 'append' || form.op === 'setText' || form.op === 'insert') fields.push(formField(t('docs.edit.text'), h('input', { className: 'dso-input', value: form.text, onChange: setValue('text') })));
          if (form.op === 'setStyle') fields.push(formField(t('docs.edit.style'), h('input', { className: 'dso-input', value: form.style, onChange: setValue('style') })));
          if (form.op === 'insert') fields.push(formField(t('docs.edit.style'), h('input', { className: 'dso-input', value: form.style, onChange: setValue('style') })));
        } else if (form.action === 'heading') {
          fields.push(formField(t('docs.edit.level'), h('select', { className: 'dso-select', value: form.level, onChange: setValue('level') },
            [1, 2, 3].map((level) => h('option', { key: level, value: String(level) }, String(level))))));
          fields.push(formField(t('docs.edit.at'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.at, onChange: setValue('at') })));
          if (form.op === 'setText' || form.op === 'insert') fields.push(formField(t('docs.edit.text'), h('input', { className: 'dso-input', value: form.text, onChange: setValue('text') })));
          if (form.op === 'setStyle') fields.push(formField(t('docs.edit.style'), h('input', { className: 'dso-input', value: form.style, onChange: setValue('style') })));
        } else {
          fields.push(formField(t('docs.edit.table'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.table, onChange: setValue('table') })));
          if (form.op === 'setCell') {
            fields.push(formField(t('docs.edit.row'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.row, onChange: setValue('row') })));
            fields.push(formField(t('docs.edit.col'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.col, onChange: setValue('col') })));
            fields.push(formField(t('docs.edit.text'), h('input', { className: 'dso-input', value: form.text, onChange: setValue('text') })));
          } else if (form.op === 'deleteRow') {
            fields.push(formField(t('docs.edit.row'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.row, onChange: setValue('row') })));
          } else if (form.op === 'addRow') {
            fields.push(formField(t('docs.edit.at'), h('input', { className: 'dso-input', value: form.at, placeholder: 'end', onChange: setValue('at') })));
          } else if (form.op === 'deleteColumn') {
            fields.push(formField(t('docs.edit.col'), h('input', { className: 'dso-input', type: 'number', min: 1, value: form.col, onChange: setValue('col') })));
          } else {
            fields.push(formField(t('docs.edit.col'), h('input', { className: 'dso-input', type: 'text', value: form.col, onChange: setValue('col') })));
          }
        }
      } else if (form.action === 'cell') {
        fields.push(formField(t('docs.edit.sheet'), h('input', { className: 'dso-input', value: form.sheet, onChange: setValue('sheet') })));
        fields.push(formField(t('docs.edit.at'), h('input', { className: 'dso-input', value: form.at, placeholder: 'B3', onChange: setValue('at') })));
        fields.push(formField(t('docs.edit.value'), h('input', { className: 'dso-input', value: form.value, onChange: setValue('value') })));
        fields.push(formField(t('docs.edit.formula'), h('input', { className: 'dso-input', value: form.formula, placeholder: '=SUM(B1:B2)', onChange: setValue('formula') })));
      } else if (form.action === 'row' || form.action === 'column') {
        fields.push(formField(t('docs.edit.sheet'), h('input', { className: 'dso-input', value: form.sheet, onChange: setValue('sheet') })));
        fields.push(formField(t('docs.edit.at'), h('input', {
          className: 'dso-input',
          value: form.at,
          placeholder: form.action === 'row' ? '3' : 'B',
          onChange: setValue('at'),
        })));
      } else {
        fields.push(formField(t('docs.edit.sheetName'), h('input', { className: 'dso-input', value: form.name, onChange: setValue('name') })));
        if (form.op === 'rename') fields.push(formField(t('docs.edit.newName'), h('input', { className: 'dso-input', value: form.newName, onChange: setValue('newName') })));
      }

      const appliedRows = result && Array.isArray(result.applied) ? result.applied : [];
      const partial = appliedRows.some((row) => row && row.ok === false);
      const resultArea = result
        ? h('div', { className: 'dso-stack' },
          h('div', { className: 'dso-inline' },
            h('span', { className: partial ? 'dso-badge dso-badge--warn' : 'dso-badge dso-badge--ok' },
              fmt(t('docs.edit.applied'), { count: appliedRows.length })),
            result.engine ? h('span', { className: 'dso-faint' }, fmt(t('docs.edit.engine'), { engine: String(result.engine) })) : null,
            result.outputPath ? h('span', { className: 'dso-mono dso-faint', style: { overflowWrap: 'anywhere' } }, result.outputPath) : null),
          partial ? h('div', { className: 'dso-warn' }, t('docs.edit.partial')) : null,
          appliedRows.length === 0
            ? null
            : h('table', { className: 'dso-table' },
              h('thead', null, h('tr', null, h('th', null, '#'), h('th', null, t('docs.edit.op')), h('th', null, 'OK'), h('th', null, ''))),
              h('tbody', null, appliedRows.map((row, index) => h('tr', { key: 'ap' + index },
                h('td', null, String(row && row.index !== undefined ? row.index + 1 : index + 1)),
                h('td', { className: 'dso-mono' }, String(row && (row.op || row.action) ? (row.op || row.action) : '')),
                h('td', null, h('span', { className: row && row.ok === false ? 'dso-badge dso-badge--error' : 'dso-badge dso-badge--ok' },
                  row && row.ok === false ? t('common.error.title') : t('diag.engine.available'))),
                h('td', { className: 'dso-hint' }, row && row.message ? String(row.message) : ''))))),
          result.backup && (result.backup.path || result.backup.id)
            ? h('div', { className: 'dso-hint dso-mono' }, fmt(t('docs.edit.backup'), { path: String(result.backup.path || result.backup.id) }))
            : null)
        : null;

      const backupRows = Array.isArray(backups) ? backups : [];
      const backupArea = h('div', { className: 'dso-card' },
        h('div', { className: 'dso-cardHead' },
          h('h3', { className: 'dso-cardTitle' }, t('docs.backup.title')),
          h('span', { className: 'dso-spacer' }),
          h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: loadBackups, disabled: backupBusy },
            backupBusy ? t('common.busy') : t('docs.backup.refresh'))),
        h('div', { className: 'dso-cardBody' },
          backupError ? h(ErrorBox, { t, error: backupError, onRetry: loadBackups }) : null,
          restoreMsg ? h('div', { className: 'dso-okBox' }, fmt(t('docs.backup.restored'), { path: restoreMsg })) : null,
          backupRows.length === 0
            ? h('div', { className: 'dso-hint' }, t('docs.backup.none'))
            : h('table', { className: 'dso-table' },
              h('thead', null, h('tr', null,
                h('th', null, t('docs.backup.time')),
                h('th', null, t('docs.backup.id')),
                h('th', null, t('docs.backup.size')),
                h('th', null, ''))),
              h('tbody', null, backupRows.map((entry, index) => {
                const id = String(entry && entry.id !== undefined ? entry.id : '');
                const time = entry && entry.time !== undefined ? entry.time : '';
                return h('tr', { key: 'bk' + index + ':' + id },
                  h('td', null, typeof time === 'number' ? formatClock(time) : String(time || '—')),
                  h('td', { className: 'dso-mono' }, id || '—'),
                  h('td', null, typeof (entry && entry.size) === 'number' ? formatBytes(entry.size) : '—'),
                  h('td', null, restoreId === id && id
                    ? h('span', { className: 'dso-inline' },
                      h('span', { className: 'dso-hint' }, t('docs.backup.confirm')),
                      h('button', { type: 'button', className: 'dso-btn dso-btn--danger dso-btn--sm', onClick: () => restoreBackup(id), disabled: restoring === id },
                        restoring === id ? t('common.busy') : t('common.ok')),
                      h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => setRestoreId('') }, t('common.cancel')))
                    : h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => setRestoreId(id), disabled: !id || !!restoring },
                      t('docs.backup.restore'))));
              })))));

      return h('div', { className: 'dso-stack' },
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' },
            h('h3', { className: 'dso-cardTitle' }, isWord ? t('docs.edit.title.word') : t('docs.edit.title.excel'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-row2' },
              formField(t('docs.edit.action'), actionSelect),
              opSelect ? formField(t('docs.edit.op'), opSelect) : null),
            fields.length ? h('div', { className: 'dso-row2' }, ...fields) : null,
            h('div', { className: 'dso-hint' }, isWord ? t('docs.edit.hint.word') : t('docs.edit.hint.excel')),
            formError ? h('div', { className: 'dso-error', role: 'alert' }, formError) : null,
            h('div', { className: 'dso-actions' },
              h('button', { type: 'button', className: 'dso-btn', onClick: addOp }, t('docs.edit.add')),
              h('button', { type: 'button', className: 'dso-btn', onClick: () => setQueue([]), disabled: queue.length === 0 }, t('docs.edit.clear')),
              h('span', { className: 'dso-spacer' }),
              h('button', {
                type: 'button',
                className: 'dso-btn dso-btn--primary',
                onClick: applyEdits,
                disabled: applying || queue.length === 0 || !sessionId,
                title: sessionId ? t('docs.edit.apply') : t('common.session.missing'),
              }, applying ? t('docs.edit.applying') : t('docs.edit.apply'))),
            h('div', { className: 'dso-label' }, fmt(t('docs.edit.queue'), { count: queue.length })),
            queue.length === 0
              ? h('div', { className: 'dso-hint' }, t('docs.edit.queue.empty'))
              : h('ul', { className: 'dso-list' }, queue.map((op, index) => h('li', { key: 'q' + index, className: 'dso-inline' },
                h('span', { className: 'dso-mono', style: { flex: '1 1 auto', overflowWrap: 'anywhere' } }, JSON.stringify(op)),
                h('button', {
                  type: 'button',
                  className: 'dso-btn dso-btn--sm dso-btn--danger',
                  onClick: () => setQueue((prev) => prev.filter((_, item) => item !== index)),
                }, t('docs.edit.remove'))))),
            error ? h(ErrorBox, { t, error }) : null)),
        result
          ? h('div', { className: 'dso-card' },
            h('div', { className: 'dso-cardHead' }, h('h3', { className: 'dso-cardTitle' }, t('docs.edit.result'))),
            h('div', { className: 'dso-cardBody' }, resultArea))
          : null,
        backupArea);
    }

    /**
     * 排版子页（P2 §6.1）：rules 编辑器 + 样式集/模板/占位符 + scope → suite.format；
     * 批量排版 → suite.format.batch（复用 800ms 任务轮询与取消）。
     */
    function DocsFormatPanel({ api, t, path, sessionId }) {
      const [ruleForm, setRuleForm] = React.useState(defaultFormatRuleForm);
      const [rules, setRules] = React.useState([]);
      const [ruleError, setRuleError] = React.useState(null);
      const [styleSet, setStyleSet] = React.useState('');
      const [template, setTemplate] = React.useState('');
      const [placeholders, setPlaceholders] = React.useState([{ key: '标题', value: '' }, { key: '日期', value: '' }]);
      const [scope, setScope] = React.useState('all');
      const [headingLevels, setHeadingLevels] = React.useState([1, 2, 3]);
      const [catalog, setCatalog] = React.useState(null);
      const [catalogError, setCatalogError] = React.useState(null);
      const [applying, setApplying] = React.useState(false);
      const [result, setResult] = React.useState(null);
      const [error, setError] = React.useState(null);
      const [batchDir, setBatchDir] = React.useState('');
      const [batchTaskId, setBatchTaskId] = React.useState('');
      const [batchStarting, setBatchStarting] = React.useState(false);
      const [batchError, setBatchError] = React.useState(null);
      const [batchCancelling, setBatchCancelling] = React.useState(false);
      const [batchParallel, setBatchParallel] = React.useState('1');
      const abortRef = React.useRef(null);
      const batchPoll = useTaskPolling(api, batchTaskId, 'suite.office.status', t);

      React.useEffect(() => () => {
        if (abortRef.current) abortRef.current.abort();
      }, []);

      React.useEffect(() => {
        setRules([]);
        setRuleError(null);
        setResult(null);
        setError(null);
        setBatchTaskId('');
        setBatchError(null);
        setBatchDir('');
      }, [path]);

      React.useEffect(() => {
        let alive = true;
        api.callCommand('suite.templates', {}).then((data) => {
          if (!alive) return;
          setCatalog({
            styles: Array.isArray(data && data.styles) ? data.styles : [],
            templates: Array.isArray(data && data.templates) ? data.templates : [],
            hint: data && typeof data.hint === 'string' ? data.hint : '',
          });
        }).catch((thrown) => {
          if (!alive) return;
          setCatalogError(toPanelError(thrown, t));
        });
        return () => { alive = false; };
      }, [api, t]);

      const setRuleValue = (key) => (event) => {
        const value = event.target.type === 'checkbox' ? event.target.checked : event.target.value;
        setRuleForm((prev) => Object.assign({}, prev, { [key]: value }));
      };

      const addRule = () => {
        const built = buildFormatRule(t, ruleForm);
        if (!built.ok) {
          setRuleError(built.message);
          return;
        }
        setRules((prev) => prev.concat([built.rule]));
        setRuleError(null);
      };

      const buildArgs = () => buildFormatSpec(t, { rules, styleSet, template, placeholders, scope, headingLevels });

      const applyFormat = async () => {
        if (applying) return;
        if (!sessionId) {
          setError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        const spec = buildArgs();
        if (!spec.ok) {
          setError({ code: 'PARAM_INVALID', message: spec.message });
          return;
        }
        if (abortRef.current) abortRef.current.abort();
        const controller = new AbortController();
        abortRef.current = controller;
        setApplying(true);
        setError(null);
        setResult(null);
        try {
          const formatted = await api.callCommand('suite.format', Object.assign({ path }, spec.args), controller.signal);
          setResult(formatted || {});
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setError(normalized);
        } finally {
          setApplying(false);
          abortRef.current = null;
        }
      };

      const chooseBatchDir = async () => {
        try {
          const dir = await api.pickDir();
          if (dir) setBatchDir(dir);
        } catch (thrown) {
          setBatchError(toPanelError(thrown, t));
        }
      };

      const startBatch = async () => {
        if (batchStarting) return;
        if (!sessionId) {
          setBatchError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        const dir = batchDir.trim();
        if (!dir) {
          setBatchError({ code: 'PARAM_INVALID', message: t('cad.input.placeholder') });
          return;
        }
        const spec = buildArgs();
        if (!spec.ok) {
          setBatchError({ code: 'PARAM_INVALID', message: spec.message });
          return;
        }
        setBatchStarting(true);
        setBatchError(null);
        setBatchTaskId('');
        try {
          const args = Object.assign({ dir, parallel: batchParallel === 'auto' ? 'auto' : Number(batchParallel) }, spec.args);
          const started = await api.callCommand('suite.format.batch', args);
          const id = started && typeof started.taskId === 'string' ? started.taskId : '';
          if (!id) throw { code: 'BAD_RESPONSE', message: 'suite.format.batch returned no taskId' };
          setBatchTaskId(id);
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setBatchError(normalized);
        } finally {
          setBatchStarting(false);
        }
      };

      const cancelBatch = async () => {
        if (!batchTaskId || batchCancelling) return;
        setBatchCancelling(true);
        try {
          await api.callCommand('suite.office.cancel', { taskId: batchTaskId });
        } catch (thrown) {
          setBatchError(toPanelError(thrown, t));
        } finally {
          setBatchCancelling(false);
        }
      };

      const styles = catalog && Array.isArray(catalog.styles) ? catalog.styles : [];
      const templates = catalog && Array.isArray(catalog.templates) ? catalog.templates : [];
      const ruleSummary = (rule) => {
        const parts = [t('docs.format.target.' + rule.target)];
        if (rule.level) parts.push('L' + rule.level);
        if (rule.font) parts.push([rule.font.name, rule.font.sizePt ? rule.font.sizePt + 'pt' : '', rule.font.bold ? 'B' : ''].filter(Boolean).join(' '));
        if (rule.paragraph) parts.push([rule.paragraph.lineSpacing ? rule.paragraph.lineSpacing + 'x' : '', rule.paragraph.spaceBeforePt !== undefined ? '↑' + rule.paragraph.spaceBeforePt : '', rule.paragraph.spaceAfterPt !== undefined ? '↓' + rule.paragraph.spaceAfterPt : '', rule.paragraph.align || ''].filter(Boolean).join(' '));
        if (rule.numbering) parts.push(t('docs.format.numbering.' + rule.numbering));
        if (rule.page) parts.push(Object.entries(rule.page.marginsMm).map(([side, value]) => side + ' ' + value + 'mm').join(' '));
        if (rule.header) parts.push(rule.header.text);
        if (rule.footer) parts.push(rule.footer.text);
        return parts.filter(Boolean).join(' · ');
      };

      const target = ruleForm.target;
      const scopeLevels = [1, 2, 3];
      const batchStatus = batchPoll.status;
      const batchOutput = batchStatus && batchStatus.output && typeof batchStatus.output === 'object' ? batchStatus.output : null;
      const batchFailures = batchOutput && Array.isArray(batchOutput.failures) ? batchOutput.failures : [];
      const batchExtra = batchFailures.length
        ? h('div', { className: 'dso-stack' },
          h('div', { className: 'dso-warn' }, fmt(t('docs.format.batch.failures'), { count: batchFailures.length })),
          h('ul', { className: 'dso-list' }, batchFailures.map((row, index) => h('li', { key: 'bf' + index, className: 'dso-hint dso-mono' },
            String(row && row.file) + ' · ' + String(row && row.code) + ' ' + String(row && row.message || '')))))
        : null;

      return h('div', { className: 'dso-stack' },
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h3', { className: 'dso-cardTitle' }, t('docs.format.rules'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-row2' },
              formField(t('docs.format.target'), h('select', { className: 'dso-select', value: target, onChange: setRuleValue('target') },
                ['heading', 'body', 'page', 'header', 'footer'].map((item) => h('option', { key: item, value: item }, t('docs.format.target.' + item))))),
              target === 'heading'
                ? formField(t('docs.format.headingLevel'), h('select', { className: 'dso-select', value: ruleForm.level, onChange: setRuleValue('level') },
                  [1, 2, 3].map((level) => h('option', { key: level, value: String(level) }, String(level)))))
                : null),
            target === 'body' || target === 'heading'
              ? h('div', { className: 'dso-row2' },
                formField(t('docs.format.fontName'), h('input', { className: 'dso-input', value: ruleForm.fontName, onChange: setRuleValue('fontName') })),
                formField(t('docs.format.fontSize'), h('input', { className: 'dso-input', type: 'number', min: 1, max: 200, value: ruleForm.fontSize, onChange: setRuleValue('fontSize') })),
                h('label', { className: 'dso-check', style: { alignSelf: 'end', paddingBottom: '4px' } },
                  h('input', { type: 'checkbox', checked: ruleForm.bold === true, onChange: setRuleValue('bold') }),
                  t('docs.format.bold')))
              : null,
            target === 'body' || target === 'heading'
              ? h('div', { className: 'dso-row2' },
                formField(t('docs.format.paraLn'), h('input', { className: 'dso-input', type: 'number', step: '0.1', min: 0.5, max: 10, value: ruleForm.lineSpacing, onChange: setRuleValue('lineSpacing') })),
                formField(t('docs.format.spaceBefore'), h('input', { className: 'dso-input', type: 'number', min: 0, max: 500, value: ruleForm.spaceBefore, onChange: setRuleValue('spaceBefore') })),
                formField(t('docs.format.spaceAfter'), h('input', { className: 'dso-input', type: 'number', min: 0, max: 500, value: ruleForm.spaceAfter, onChange: setRuleValue('spaceAfter') })),
                formField(t('docs.format.align'), h('select', { className: 'dso-select', value: ruleForm.align, onChange: setRuleValue('align') },
                  h('option', { value: '' }, t('docs.format.align.none')),
                  ['left', 'center', 'right', 'justify'].map((item) => h('option', { key: item, value: item }, item)))),
                formField(t('docs.format.numbering'), h('select', { className: 'dso-select', value: ruleForm.numbering, onChange: setRuleValue('numbering') },
                  h('option', { value: '' }, t('docs.format.align.none')),
                  ['chinese', 'arabic', 'decimalDot', 'none'].map((item) => h('option', { key: item, value: item }, t('docs.format.numbering.' + item))))))
              : null,
            target === 'page'
              ? h('div', { className: 'dso-row2' },
                formField(t('docs.format.margins'), h('div', { className: 'dso-inline' },
                  ['top', 'bottom', 'left', 'right'].map((side) => h('input', {
                    key: side,
                    className: 'dso-input',
                    type: 'number',
                    min: 1,
                    max: 500,
                    placeholder: side,
                    title: side,
                    value: ruleForm['margin.' + side],
                    onChange: setRuleValue('margin.' + side),
                    style: { width: '76px', flex: '0 0 auto' },
                  })))))
              : null,
            target === 'header' || target === 'footer'
              ? formField(t('docs.format.target.' + target), h('input', { className: 'dso-input', value: target === 'header' ? ruleForm.headerText : ruleForm.footerText, onChange: setRuleValue(target === 'header' ? 'headerText' : 'footerText') }))
              : null,
            ruleError ? h('div', { className: 'dso-error', role: 'alert' }, ruleError) : null,
            h('div', { className: 'dso-actions' },
              h('button', { type: 'button', className: 'dso-btn', onClick: addRule }, t('docs.format.rule.add')),
              h('button', { type: 'button', className: 'dso-btn', onClick: () => setRules([]), disabled: rules.length === 0 }, t('docs.edit.clear'))),
            rules.length === 0
              ? h('div', { className: 'dso-hint' }, t('docs.format.rule.empty'))
              : h('ul', { className: 'dso-list' }, rules.map((rule, index) => h('li', { key: 'rule' + index, className: 'dso-inline' },
                h('span', { className: 'dso-muted', style: { flex: '1 1 auto', overflowWrap: 'anywhere' } }, ruleSummary(rule)),
                h('button', {
                  type: 'button',
                  className: 'dso-btn dso-btn--sm dso-btn--danger',
                  onClick: () => setRules((prev) => prev.filter((_, item) => item !== index)),
                }, t('docs.edit.remove')))))),
            catalogError ? h('div', { className: 'dso-hint dso-error' }, catalogError.message) : null,
            h('div', { className: 'dso-row2' },
              formField(t('docs.format.styleSet'), h('select', { className: 'dso-select', value: styleSet, onChange: (event) => setStyleSet(event.target.value) },
                h('option', { value: '' }, t('common.none')),
                styles.map((entry) => h('option', { key: 'st' + entry.id, value: entry.id },
                  String(entry.name || entry.id) + (entry.builtin ? '' : ' *') + (entry.invalid ? ' (!)' : '') + ' · ' + String(entry.rules ?? 0))))),
              formField(t('docs.format.template'), h('select', {
                className: 'dso-select',
                value: template,
                onChange: (event) => {
                  const id = event.target.value;
                  setTemplate(id);
                  const entry = templates.find((item) => item && item.id === id);
                  if (entry && Array.isArray(entry.placeholders) && entry.placeholders.length) {
                    setPlaceholders(entry.placeholders.map((item) => {
                      if (typeof item === 'string') return { key: item, value: '' };
                      const key = item && (item.key ?? item.name ?? item.id);
                      return { key: key === undefined || key === null ? '' : String(key), value: item && item.value !== undefined ? String(item.value) : '' };
                    }));
                  }
                },
              },
              h('option', { value: '' }, t('common.none')),
              templates.map((entry) => h('option', { key: 'tp' + entry.id, value: entry.id },
                String(entry.name || entry.id) + (entry.builtin ? '' : ' *')))))),
            template
              ? h('div', { className: 'dso-stack' },
                h('div', { className: 'dso-label' }, t('docs.format.placeholders')),
                placeholders.map((row, index) => h('div', { key: 'ph' + index, className: 'dso-inline' },
                  h('input', {
                    className: 'dso-input',
                    value: row.key,
                    placeholder: t('docs.format.placeholder.key'),
                    style: { flex: '1 1 120px', width: 'auto' },
                    onChange: (event) => setPlaceholders((prev) => prev.map((item, itemIndex) => (itemIndex === index ? Object.assign({}, item, { key: event.target.value }) : item))),
                  }),
                  h('input', {
                    className: 'dso-input',
                    value: row.value,
                    placeholder: t('docs.format.placeholder.value'),
                    style: { flex: '2 1 200px', width: 'auto' },
                    onChange: (event) => setPlaceholders((prev) => prev.map((item, itemIndex) => (itemIndex === index ? Object.assign({}, item, { value: event.target.value }) : item))),
                  }),
                  h('button', {
                    type: 'button',
                    className: 'dso-btn dso-btn--sm',
                    onClick: () => setPlaceholders((prev) => prev.filter((_, itemIndex) => itemIndex !== index)),
                  }, t('docs.edit.remove')))),
                h('div', null, h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => setPlaceholders((prev) => prev.concat([{ key: '', value: '' }])) }, t('docs.format.placeholder.add'))))
              : null,
            h('div', { className: 'dso-row2' },
              formField(t('docs.format.scope'), h('select', { className: 'dso-select', value: scope, onChange: (event) => setScope(event.target.value) },
                h('option', { value: 'all' }, t('docs.format.scope.all')),
                h('option', { value: 'headings' }, t('docs.format.scope.headings')))),
              scope === 'headings'
                ? h('div', { className: 'dso-inline', style: { alignSelf: 'end', paddingBottom: '4px' } },
                  scopeLevels.map((level) => h('label', { key: 'lv' + level, className: 'dso-check' },
                    h('input', {
                      type: 'checkbox',
                      checked: headingLevels.includes(level),
                      onChange: (event) => setHeadingLevels((prev) => (event.target.checked
                        ? prev.concat([level])
                        : prev.filter((value) => value !== level))),
                    }),
                    String(level))))
                : null),
            error ? h(ErrorBox, { t, error }) : null,
            result
              ? h('div', { className: 'dso-okBox' },
                h('div', null, fmt(t('docs.format.done'), { path: result.outputPath || path })),
                result.engine ? h('div', { className: 'dso-hint' }, fmt(t('docs.edit.engine'), { engine: String(result.engine) })) : null,
                result.backup && result.backup.path ? h('div', { className: 'dso-hint dso-mono' }, fmt(t('docs.edit.backup'), { path: String(result.backup.path) })) : null)
              : null,
            h('div', { className: 'dso-actions' },
              h('button', {
                type: 'button',
                className: 'dso-btn dso-btn--primary',
                onClick: applyFormat,
                disabled: applying || !sessionId,
                title: sessionId ? t('docs.format.apply') : t('common.session.missing'),
              }, applying ? t('docs.format.applying') : t('docs.format.apply')))),
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h3', { className: 'dso-cardTitle' }, t('docs.format.batch'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-hint' }, t('docs.format.batch.note')),
            h('div', { className: 'dso-inline' },
              h('input', {
                className: 'dso-input',
                value: batchDir,
                placeholder: t('docs.format.batch.dir'),
                style: { flex: '1 1 260px', width: 'auto' },
                onChange: (event) => setBatchDir(event.target.value),
              }),
              h('button', { type: 'button', className: 'dso-btn', onClick: chooseBatchDir, disabled: batchStarting }, t('cad.input.dir'))),
            h('div', { className: 'dso-row2' },
              formField(t('cad.parallel'), h('select', { className: 'dso-select', value: batchParallel, onChange: (event) => setBatchParallel(event.target.value) },
                h('option', { value: 'auto' }, t('cad.parallel.auto')),
                [1, 2, 3, 4].map((value) => h('option', { key: 'bp' + value, value: String(value) }, String(value))))),
              h('div', { className: 'dso-inline', style: { alignSelf: 'end' } },
                h('button', {
                  type: 'button',
                  className: 'dso-btn dso-btn--primary',
                  onClick: startBatch,
                  disabled: batchStarting || !sessionId || !batchDir.trim(),
                }, batchStarting ? t('common.busy') : t('docs.format.batch.start')))),
            batchError ? h(ErrorBox, { t, error: batchError }) : null,
            batchTaskId
              ? h('div', { className: 'dso-hint dso-mono' }, fmt(t('docs.format.batch.started'), { taskId: batchTaskId }))
              : null,
            batchTaskId
              ? h(TaskProgress, {
                t,
                status: batchPoll.status,
                error: batchPoll.error,
                mode: 'files',
                cancelling: batchCancelling,
                onCancel: cancelBatch,
                extra: batchExtra,
              })
              : null)));
    }

    /** 预览子页（P3 §5.2）：按钮 / iframe / missingFonts / E2012 提示 / 导出 PDF。 */
    function DocsPreviewPanel({ t, preview, cacheSize, cacheCount, displayPath, onRender, onRetry, onExport, exporting, missingSession }) {
      const status = preview && typeof preview.status === 'string' ? preview.status : 'idle';
      const missingFonts = preview && Array.isArray(preview.missingFonts) ? preview.missingFonts : [];
      const head = h('div', { className: 'dso-inline' },
        h('span', {
          className: status === 'ready'
            ? 'dso-badge dso-badge--ok'
            : (status === 'error' || status === 'unsupported' ? 'dso-badge dso-badge--error' : 'dso-badge dso-badge--idle'),
        }, status === 'ready'
          ? t('docs.preview.render')
          : (status === 'busy'
            ? t('docs.preview.rendering')
            : (status === 'unsupported' ? t('docs.preview.unsupported') : (status === 'error' ? t('common.error.title') : t('docs.preview.empty'))))),
        h('span', { className: 'dso-faint' }, t('docs.preview.scheme.command')),
        missingFonts.length
          ? h('span', { className: 'dso-badge dso-badge--warn', title: missingFonts.join(', ') }, fmt(t('docs.preview.missingFonts'), { count: missingFonts.length }))
          : null,
        h('span', { className: 'dso-spacer' }),
        h('span', { className: 'dso-faint' }, fmt(t('docs.preview.cached'), { count: cacheCount, max: cacheSize })),
        h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: onRender, disabled: missingSession || status === 'busy' }, t('docs.preview.retry')),
        h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: onExport, disabled: missingSession || exporting }, exporting ? t('docs.export.running') : t('docs.export')));

      let content = null;
      if (status === 'busy') content = h(Busy, { text: t('docs.preview.rendering') });
      else if (status === 'ready' && preview.url) content = h('iframe', { className: 'dso-previewFrame', src: preview.url, title: displayPath });
      else if (status === 'unsupported') {
        content = h('div', { className: 'dso-errorBox' },
          h('div', null, t('docs.preview.unsupported')),
          h('div', { className: 'dso-hint' }, t('docs.preview.unsupported.hint')));
      } else if (status === 'error') {
        content = h('div', { className: 'dso-stack' },
          h(ErrorBox, { t, error: preview.error, onRetry }),
          h('div', { className: 'dso-hint' }, t('docs.preview.unsupported.hint')));
      } else content = h(Empty, { text: t('docs.preview.empty') });

      return h('div', { className: 'dso-stack' }, head, content);
    }
    //#endregion

    //#region 8. 组件：VisioTab（flowchart 节点/边 + P2 三类/P3 五类电网模型 + 校验 + 运行 + 历史）
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
    /** Visio 绘图页：flowchart + P2/P3 电网八类（模型 JSON + 模板 + 校验 + 结果 validations）。 */
    function VisioTab({ api, t, sessionId }) {
      const [kind, setKind] = React.useState('flowchart');
      const [title, setTitle] = React.useState('');
      const [nodesText, setNodesText] = React.useState(() => JSON.stringify(FLOWCHART_TEMPLATE_NODES, null, 2));
      const [edgesText, setEdgesText] = React.useState(() => JSON.stringify(FLOWCHART_TEMPLATE_EDGES, null, 2));
      const [modelText, setModelText] = React.useState(() => gridTemplateText(GRID_MODEL_KINDS[0]));
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

      const isFlowchart = kind === 'flowchart';
      const validation = React.useMemo(
        () => (isFlowchart ? parseFlowchart(t, nodesText, edgesText) : parseGridModel(t, kind, modelText)),
        [isFlowchart, kind, modelText, nodesText, edgesText, t],
      );

      const pushHistory = React.useCallback((entry) => {
        setHistory((prev) => [Object.assign({ at: Date.now() }, entry)].concat(prev).slice(0, 5));
      }, []);

      const changeKind = React.useCallback((next) => {
        if (!next || next === GRID_KIND_PENDING) return;
        setKind(next);
        setResult(null);
        setError(null);
        if (next === 'flowchart') return;
        // 切换图类型：仅当 textarea 仍是内置模板（或为空）时刷新模板，避免覆盖用户编辑。
        setModelText((prev) => (isBuiltinTemplateText(prev) ? gridTemplateText(next) : prev));
      }, []);

      const insertTemplate = React.useCallback(() => {
        if (isFlowchart) {
          setNodesText(JSON.stringify(FLOWCHART_TEMPLATE_NODES, null, 2));
          setEdgesText(JSON.stringify(FLOWCHART_TEMPLATE_EDGES, null, 2));
          return;
        }
        setModelText(gridTemplateText(kind));
        setError(null);
      }, [isFlowchart, kind]);

      const chooseDir = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          const stem = isFlowchart ? 'flowchart' : kind;
          setOutputPath(joinPath(dir, stem + '-' + formatStamp(Date.now()) + '.vsdx'));
        } catch (picked) {
          setError(toPanelError(picked, t));
        }
      }, [api, isFlowchart, kind, t]);

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
          const args = { kind };
          if (title.trim()) args.title = title.trim();
          if (isFlowchart) {
            args.nodes = validation.nodes;
            args.edges = validation.edges;
          } else {
            args.model = validation.model;
          }
          const trimmed = outputPath.trim();
          if (trimmed) args.outputPath = trimmed;
          const drawn = await api.callCommand('suite.draw', args, controller.signal);
          setResult(drawn || {});
          pushHistory({
            ok: true,
            kind,
            outputPath: drawn && drawn.outputPath ? String(drawn.outputPath) : '',
            stats: drawn && drawn.stats,
            pages: drawn && drawn.pages,
          });
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (normalized.code === 'E2024') normalized.hint = t('visio.e2024');
          if (!normalized.cancelled) {
            setError(normalized);
            pushHistory({ ok: false, kind, code: normalized.code, message: normalized.message });
          }
        } finally {
          setRunning(false);
          abortRef.current = null;
        }
      }, [api, isFlowchart, kind, outputPath, pushHistory, running, sessionId, t, title, validation]);

      const abort = React.useCallback(() => {
        if (abortRef.current) abortRef.current.abort();
      }, []);

      const field = (label, node) => h('div', { className: 'dso-field' }, h('span', { className: 'dso-label' }, label), node);

      const statLabel = (key) => (STAT_LABEL_KEYS[key] ? t(STAT_LABEL_KEYS[key]) : key);
      const stats = result && isPlainObjectLoose(result.stats) ? result.stats : null;
      const statEntries = stats
        ? Object.entries(stats).filter(([key, value]) => key !== 'pages' && (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string'))
        : [];
      const validationRows = result && Array.isArray(result.validations) ? result.validations : [];
      const historyItems = history.map((entry, index) => {
        let statsText = '';
        if (entry.stats) {
          if (entry.kind === 'flowchart') {
            statsText = fmt(t('visio.done.stats'), { nodes: entry.stats.nodes, edges: entry.stats.edges });
          } else {
            statsText = Object.entries(entry.stats)
              .filter(([key, value]) => typeof value === 'number' && key !== 'pages')
              .slice(0, 4)
              .map(([key, value]) => statLabel(key) + ' ' + value)
              .join(' · ');
          }
        }
        const detail = entry.ok
          ? (entry.outputPath || '—')
            + (statsText ? ' · ' + statsText : '')
            + (entry.pages != null ? ' · ' + fmt(t('visio.result.pages'), { pages: entry.pages }) : '')
          : (entry.message || '—');
        return h('li', { key: 'h' + index, className: 'dso-inline' },
          h('span', { className: 'dso-faint dso-mono' }, formatClock(entry.at)),
          h('span', { className: entry.ok ? 'dso-badge dso-badge--ok' : 'dso-badge dso-badge--error' },
            entry.ok ? t('cad.state.done') : t('cad.state.failed')),
          entry.ok && entry.kind && entry.kind !== 'flowchart'
            ? h('span', { className: 'dso-badge' }, gridKindLabel(t, entry.kind))
            : null,
          h('span', { className: 'dso-mono', style: { overflowWrap: 'anywhere' } }, detail));
      });

      return h('div', { className: 'dso-stack' },
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('visio.title'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-row2' },
              field(t('visio.kind'),
                h('select', {
                  className: 'dso-select',
                  value: kind,
                  disabled: running,
                  onChange: (event) => changeKind(event.target.value),
                },
                ['flowchart'].concat(GRID_MODEL_KINDS).map((item) => h('option', { key: item, value: item }, gridKindLabel(t, item))),
                h('option', { key: GRID_KIND_PENDING, value: GRID_KIND_PENDING, disabled: true }, t('visio.kind.cableRoute')))),
              field(t('visio.titleField'),
                h('input', {
                  className: 'dso-input',
                  type: 'text',
                  value: title,
                  placeholder: t('panel.title'),
                  onChange: (event) => setTitle(event.target.value),
                }))),
            h('div', { className: 'dso-hint' }, t('visio.kind.cableRoute.p4hint')),
            isFlowchart
              ? h('div', { className: 'dso-stack' },
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
                  : h('div', { className: 'dso-error', role: 'alert' }, validation.message))
              : h('div', { className: 'dso-stack' },
                field(t('visio.model'),
                  h('textarea', {
                    className: 'dso-textarea',
                    value: modelText,
                    spellCheck: false,
                    style: { minHeight: '260px' },
                    onChange: (event) => setModelText(event.target.value),
                  })),
                validation.ok ? null : h('div', { className: 'dso-error', role: 'alert' }, validation.message)),
            h('div', { className: 'dso-actions' },
              h('button', { type: 'button', className: 'dso-btn', onClick: insertTemplate, disabled: running },
                isFlowchart ? t('visio.template') : t('visio.template.model')),
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
                statEntries.length
                  ? h('div', { className: 'dso-inline' }, statEntries.map(([key, value]) => h('span', { key: 'st' + key, className: 'dso-badge' },
                    statLabel(key) + ' ' + String(value))))
                  : null,
                result.pages != null
                  ? h('div', { className: 'dso-hint' }, fmt(t('visio.result.pages'), { pages: result.pages }))
                  : null),
              h('div', { className: 'dso-label' }, t('visio.validations')),
              validationRows.length === 0
                ? h('div', { className: 'dso-hint' }, t('visio.validations.empty'))
                : h('ul', { className: 'dso-list' }, validationRows.map((row, index) => {
                  const level = row && row.level === 'warn' ? 'warn' : 'info';
                  return h('li', { key: 'val' + index, className: 'dso-inline' },
                    h('span', { className: level === 'warn' ? 'dso-badge dso-badge--warn' : 'dso-badge' },
                      level === 'warn' ? t('visio.validations.warn') : t('visio.validations.info')),
                    h('span', { className: 'dso-muted', style: { overflowWrap: 'anywhere' } },
                      String(row && row.message !== undefined ? row.message : '')),
                    row && row.code ? h('span', { className: 'dso-faint dso-mono' }, String(row.code)) : null);
                }))))
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

      /* P3 增量：图纸目录表 + 批量队列（父任务 kind='cad.queue'） */
      const [catalogInput, setCatalogInput] = React.useState('');
      const [catalogSort, setCatalogSort] = React.useState('drawingNo');
      const [catalogBusy, setCatalogBusy] = React.useState(false);
      const [catalogResult, setCatalogResult] = React.useState(null);
      const [catalogError, setCatalogError] = React.useState(null);
      const [queue, setQueue] = React.useState([]);
      const [queueTaskId, setQueueTaskId] = React.useState('');
      const [queueStarting, setQueueStarting] = React.useState(false);
      const [queueError, setQueueError] = React.useState(null);
      const [queueCancelling, setQueueCancelling] = React.useState(false);
      const [pickerTarget, setPickerTarget] = React.useState('input');
      const catalogAbortRef = React.useRef(null);
      const queueAbortRef = React.useRef(null);
      const queuePoll = useTaskPolling(api, queueTaskId, 'suite.cad.status', t);

      const logRef = React.useRef(null);
      const startAbortRef = React.useRef(null);
      const tRef = React.useRef(t);
      tRef.current = t;

      React.useEffect(() => () => {
        if (startAbortRef.current) startAbortRef.current.abort();
        if (catalogAbortRef.current) catalogAbortRef.current.abort();
        if (queueAbortRef.current) queueAbortRef.current.abort();
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

      const openPicker = React.useCallback(async (target) => {
        setPickerTarget(typeof target === 'string' ? target : 'input');
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

      /* ── P3 增量：目录表 / 批量队列操作 ── */
      const addQueuePath = React.useCallback((path) => {
        const value = String(path ?? '').trim();
        if (!value) return;
        setQueue((prev) => (prev.includes(value) ? prev : prev.concat([value])));
        setQueueError(null);
      }, []);

      const chooseCatalogInput = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          setCatalogInput(dir);
          setCatalogError(null);
        } catch (error) {
          setCatalogError(toPanelError(error, t));
        }
      }, [api, t]);

      const addQueueDir = React.useCallback(async () => {
        try {
          const dir = await api.pickDir();
          if (!dir) return;
          addQueuePath(dir);
        } catch (error) {
          setQueueError(toPanelError(error, t));
        }
      }, [addQueuePath, api, t]);

      const runCatalog = React.useCallback(async () => {
        if (catalogBusy) return;
        if (!sessionId) {
          setCatalogError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        const input = catalogInput.trim();
        if (!input) {
          setCatalogError({ code: 'PARAM_INVALID', message: t('cad.catalog.input') });
          return;
        }
        if (catalogAbortRef.current) catalogAbortRef.current.abort();
        const controller = new AbortController();
        catalogAbortRef.current = controller;
        setCatalogBusy(true);
        setCatalogError(null);
        setCatalogResult(null);
        try {
          const result = await api.callCommand('suite.cad.catalog', { input, sort: catalogSort }, controller.signal);
          setCatalogResult(result && typeof result === 'object' ? result : {});
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setCatalogError(normalized);
        } finally {
          setCatalogBusy(false);
          catalogAbortRef.current = null;
        }
      }, [api, catalogBusy, catalogInput, catalogSort, sessionId, t]);

      const startQueue = React.useCallback(async () => {
        if (queueStarting) return;
        if (!sessionId) {
          setQueueError({ code: 'NO_SESSION', message: t('common.session.missing') });
          return;
        }
        if (queue.length === 0) {
          setQueueError({ code: 'PARAM_INVALID', message: t('cad.queue.empty') });
          return;
        }
        if (queueAbortRef.current) queueAbortRef.current.abort();
        const controller = new AbortController();
        queueAbortRef.current = controller;
        setQueueStarting(true);
        setQueueError(null);
        setQueueTaskId('');
        try {
          const started = await api.callCommand('suite.cad.convert', {
            inputs: queue.slice(),
            mode,
            paper,
            sort: sortMode,
            parallel: parallel === 'auto' ? 'auto' : Number(parallel),
          }, controller.signal);
          const id = started && typeof started.taskId === 'string' ? started.taskId : '';
          if (!id) throw { code: 'BAD_RESPONSE', message: 'suite.cad.convert returned no taskId' };
          setQueueTaskId(id);
        } catch (thrown) {
          const normalized = toPanelError(thrown, t);
          if (!normalized.cancelled) setQueueError(normalized);
        } finally {
          setQueueStarting(false);
          queueAbortRef.current = null;
        }
      }, [api, mode, paper, parallel, queue, queueStarting, sessionId, sortMode, t]);

      const cancelQueue = React.useCallback(async () => {
        if (!queueTaskId || queueCancelling) return;
        setQueueCancelling(true);
        try {
          await api.callCommand('suite.cad.cancel', { taskId: queueTaskId });
        } catch (thrown) {
          setQueueError(toPanelError(thrown, t));
        } finally {
          setQueueCancelling(false);
        }
      }, [api, queueCancelling, queueTaskId, t]);

      const pickPickerRow = React.useCallback((row) => {
        if (!row) return;
        if (row.type === 'dir') {
          loadPicker(row.path);
          return;
        }
        if (pickerTarget === 'catalog') setCatalogInput(row.path);
        else if (pickerTarget === 'queue') addQueuePath(row.path);
        else setInputPath(row.path);
        setPicker(null);
      }, [addQueuePath, loadPicker, pickerTarget]);

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

      // 目录表字段命中统计（P3 §2.1 result.fields）
      const catalogFields = catalogResult && isPlainObjectLoose(catalogResult.fields) ? catalogResult.fields : null;
      const catalogFieldRows = catalogFields
        ? ['drawingNo', 'title', 'frameSize', 'scale', 'count'].map((key) => ({ key, stat: catalogFields[key] }))
        : [];
      // 批量队列父任务输出（完成时：succeeded/failed/failures/outputs）
      const queueStatus = queuePoll.status;
      const queueOutput = queueStatus && queueStatus.output && typeof queueStatus.output === 'object' ? queueStatus.output : null;
      const queueRuntimeState = queueStatus && typeof queueStatus.state === 'string' ? queueStatus.state : (queueTaskId ? 'queued' : 'idle');
      const queueRunning = !!queueTaskId && !CAD_TERMINAL.includes(queueRuntimeState);
      const queueFailures = queueOutput && Array.isArray(queueOutput.failures) ? queueOutput.failures : [];
      const queueOutputs = queueOutput && Array.isArray(queueOutput.outputs) ? queueOutput.outputs : [];
      const queueExtra = (queueFailures.length || queueOutputs.length)
        ? h('div', { className: 'dso-stack' },
          queueFailures.length
            ? h('div', { className: 'dso-stack' },
              h('div', { className: 'dso-warn' }, fmt(t('cad.queue.failures'), { count: queueFailures.length })),
              h('ul', { className: 'dso-list' }, queueFailures.map((row, index) => h('li', { key: 'qf' + index, className: 'dso-hint dso-mono' },
                '[' + String(row && row.index !== undefined ? row.index : index + 1) + '] ' + String(row && row.input) + ' · ' +
                String(row && row.code) + ' ' + String(row && row.message || '')))))
            : null,
          queueOutputs.length
            ? h('div', { className: 'dso-stack' },
              h('div', { className: 'dso-label' }, t('cad.result.title')),
              h('ul', { className: 'dso-list' }, queueOutputs.map((row, index) => h('li', { key: 'qo' + index, className: 'dso-hint dso-mono' },
                '[' + String(row && row.index !== undefined ? row.index : index + 1) + '] ' + String(row && row.input) + ' → ' + String(row && row.output || '—')))))
            : null)
        : null;

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
                  : h('button', { type: 'button', className: 'dso-btn', onClick: () => openPicker('input'), disabled: running }, t('cad.input.browse')),
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
        // ── P3 图纸目录表 ──
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' }, h('h2', { className: 'dso-cardTitle' }, t('cad.catalog.title'))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-inline' },
              h('input', {
                className: 'dso-input',
                type: 'text',
                value: catalogInput,
                placeholder: t('cad.catalog.input'),
                style: { flex: '1 1 300px', width: 'auto' },
                onChange: (event) => setCatalogInput(event.target.value),
              }),
              h('button', { type: 'button', className: 'dso-btn', onClick: chooseCatalogInput, disabled: catalogBusy }, t('cad.input.dir')),
              h('button', { type: 'button', className: 'dso-btn', onClick: () => openPicker('catalog'), disabled: catalogBusy }, t('cad.input.browse'))),
            h('div', { className: 'dso-row2' },
              formField(t('cad.sort'), h('select', { className: 'dso-select', value: catalogSort, onChange: (event) => setCatalogSort(event.target.value) },
                h('option', { value: 'drawingNo' }, t('cad.sort.drawingNo')),
                h('option', { value: 'filename' }, t('cad.sort.filename')))),
              h('div', { className: 'dso-inline', style: { alignSelf: 'end' } },
                h('button', {
                  type: 'button',
                  className: 'dso-btn dso-btn--primary',
                  onClick: runCatalog,
                  disabled: catalogBusy || !sessionId || !catalogInput.trim(),
                }, catalogBusy ? t('cad.catalog.running') : t('cad.catalog.run')))),
            h('div', { className: 'dso-hint' }, t('cad.catalog.hint')),
            catalogError ? h(ErrorBox, { t, error: catalogError, onRetry: runCatalog }) : null,
            catalogResult
              ? h('div', { className: 'dso-stack' },
                h('div', { className: 'dso-okBox' },
                  h('div', null, fmt(t('cad.catalog.done'), { path: catalogResult.outputPath || '—' })),
                  h('div', { className: 'dso-inline' },
                    h('span', { className: 'dso-badge' }, fmt(t('cad.catalog.rows'), { rows: catalogResult.rows ?? '—' })),
                    h('span', { className: 'dso-badge' }, fmt(t('cad.catalog.sheets'), { sheets: catalogResult.sheets ?? '—' })),
                    h('span', { className: 'dso-badge' }, fmt(t('cad.catalog.engine'), { engine: String(catalogResult.engine || '—') })))),
                catalogFieldRows.length === 0
                  ? null
                  : h('div', { className: 'dso-stack' },
                    h('div', { className: 'dso-label' }, t('cad.catalog.fields')),
                    h('table', { className: 'dso-table' },
                      h('thead', null, h('tr', null,
                        h('th', null, t('cad.catalog.fields')),
                        h('th', null, 'hit/total'),
                        h('th', null, '%'))),
                      h('tbody', null, catalogFieldRows.map((row) => {
                        const stat = row.stat && typeof row.stat === 'object' ? row.stat : {};
                        const rate = Number(stat.rate);
                        return h('tr', { key: 'cf' + row.key },
                          h('td', null, t('cad.catalog.field.' + row.key)),
                          h('td', { className: 'dso-mono' }, String(stat.hit ?? 0) + '/' + String(stat.total ?? 0)),
                          h('td', { className: 'dso-mono' }, Number.isFinite(rate) ? (Math.round(rate * 1000) / 10) + '%' : '—'));
                      })))),
                h('div', { className: 'dso-actions' },
                  h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => api.copy(catalogResult.outputPath) }, t('docs.actions.copyPath')),
                  catalogResult.outputPath
                    ? h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => openInDocs(catalogResult.outputPath) }, t('cad.result.openInDocs'))
                    : null))
              : null)),
        // ── P3 批量队列（inputs → 父任务 kind='cad.queue'） ──
        h('div', { className: 'dso-card' },
          h('div', { className: 'dso-cardHead' },
            h('h2', { className: 'dso-cardTitle' }, t('cad.queue.title')),
            h('span', { className: 'dso-spacer' }),
            h('span', { className: 'dso-badge' }, fmt(t('cad.queue.inputs'), { count: queue.length }))),
          h('div', { className: 'dso-cardBody' },
            h('div', { className: 'dso-hint' }, t('cad.queue.note')),
            h('div', { className: 'dso-actions' },
              h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => openPicker('queue'), disabled: queueStarting }, t('cad.queue.addFile')),
              h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: addQueueDir, disabled: queueStarting }, t('cad.queue.addDir')),
              h('button', { type: 'button', className: 'dso-btn dso-btn--sm', onClick: () => setQueue([]), disabled: queueStarting || queue.length === 0 }, t('cad.queue.clear'))),
            queue.length === 0
              ? h('div', { className: 'dso-hint' }, t('cad.queue.empty'))
              : h('ul', { className: 'dso-list' }, queue.map((item, index) => h('li', { key: 'qi' + index + ':' + item, className: 'dso-inline' },
                h('span', { className: 'dso-faint dso-mono' }, String(index + 1) + '.'),
                h('span', { className: 'dso-mono', style: { flex: '1 1 auto', overflowWrap: 'anywhere' } }, item),
                h('button', {
                  type: 'button',
                  className: 'dso-btn dso-btn--sm dso-btn--danger',
                  onClick: () => setQueue((prev) => prev.filter((_, itemIndex) => itemIndex !== index)),
                  disabled: queueStarting,
                }, t('cad.queue.remove'))))),
            h('div', { className: 'dso-actions' },
              h('button', {
                type: 'button',
                className: 'dso-btn dso-btn--primary',
                onClick: startQueue,
                disabled: queueStarting || queueRunning || !sessionId || queue.length === 0,
              }, queueStarting ? t('common.busy') : t('cad.queue.start'))),
            queueError ? h(ErrorBox, { t, error: queueError, onRetry: startQueue }) : null,
            queueTaskId
              ? h('div', { className: 'dso-stack' },
                h('div', { className: 'dso-hint dso-mono' }, fmt(t('cad.queue.started'), { taskId: queueTaskId })),
                h(TaskProgress, {
                  t,
                  status: queuePoll.status,
                  error: queuePoll.error,
                  mode: 'items',
                  cancelling: queueCancelling,
                  onCancel: cancelQueue,
                  extra: queueExtra,
                }))
              : null)),
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
                    onClick: () => pickPickerRow(row),
                    onKeyDown: (event) => {
                      if (event.key === 'Enter') {
                        event.preventDefault();
                        pickPickerRow(row);
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
    /** 可写配置白名单（§2 suite.config.set；P3 增 previewCacheSize）。 */
    const SETTING_SPECS = [
      { key: 'outputRoot', kind: 'text' },
      { key: 'parallel', kind: 'text' },
      { key: 'allowExternalPaths', kind: 'boolean' },
      { key: 'odaPath', kind: 'text' },
      { key: 'pdfPaper', kind: 'select', options: ['A4-landscape', 'keep'] },
      { key: 'sortMode', kind: 'select', options: ['drawingNo', 'filename'] },
      { key: 'previewCacheSize', kind: 'number', min: 1, max: 10 },
    ];

    /** 把 suite.env 结果整理成表格行（兼容 v0.1 顶层 ProgID 与 P1 的 cad/oda 结构化字段；P2/P3 增 venv/COM/预览行）。 */
    function buildEngineRows(env, t, previewState) {
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

      // P2 §6.3：Python venv 状态行（依赖版本 + 健康状态）
      const venv = (env && env.venv) || {};
      const packages = venv.packages && typeof venv.packages === 'object' ? venv.packages : {};
      const venvHealthy = venv.healthy === true;
      const venvPresent = venv.present === true;
      const venvVersion = [
        typeof venv.pythonVersion === 'string' && venv.pythonVersion ? 'Python ' + venv.pythonVersion : '',
        fmt(t('diag.venv.packages'), { docx: packages.docx || '—', openpyxl: packages.openpyxl || '—' }),
      ].filter(Boolean).join(' · ');
      rows.push({
        key: 'venv',
        label: t('diag.engine.venv'),
        present: venvHealthy,
        degraded: venvPresent && !venvHealthy,
        version: venvVersion || (venvPresent ? '' : t('diag.venv.missing')),
        capability: t('diag.cap.venv'),
        statusText: venvHealthy ? t('diag.venv.healthy') : (venvPresent ? t('diag.venv.unhealthy') : t('diag.engine.missing')),
        statusClass: venvHealthy ? 'dso-badge dso-badge--ok' : (venvPresent ? 'dso-badge dso-badge--warn' : 'dso-badge dso-badge--idle'),
      });

      // P2 §6.3：Office COM 可用性行（comUsable 仅在 refresh:true 时实测）
      const officeCom = (env && env.officeCom) || {};
      const wordCom = officeCom.word && typeof officeCom.word === 'object' ? officeCom.word : {};
      const excelCom = officeCom.excel && typeof officeCom.excel === 'object' ? officeCom.excel : {};
      const wordUsable = wordCom.comUsable;
      const excelUsable = excelCom.comUsable;
      const comMeasured = wordUsable === true || wordUsable === false || excelUsable === true || excelUsable === false;
      const comUsable = wordUsable === true || excelUsable === true;
      const comVersion = 'Word ' + (wordUsable === true ? '✓' : (wordUsable === false ? '✗' : '—'))
        + ' · Excel ' + (excelUsable === true ? '✓' : (excelUsable === false ? '✗' : '—'));
      rows.push({
        key: 'officeCom',
        label: t('diag.engine.officeCom'),
        present: comUsable,
        version: comMeasured ? comVersion : '',
        capability: t('diag.cap.officeCom'),
        statusText: comUsable ? t('diag.com.usable') : (comMeasured ? t('diag.com.unusable') : t('diag.com.unmeasured')),
        statusClass: comUsable ? 'dso-badge dso-badge--ok' : (comMeasured ? 'dso-badge dso-badge--error' : 'dso-badge dso-badge--idle'),
      });

      // P3 §5.2：预览渲染服务（suite.preview 探测一次；E2012 → 不可用）
      const preview = previewState && typeof previewState === 'object' ? previewState : {};
      const previewAvailable = preview.available === true;
      const previewKnown = preview.available === true || preview.available === false;
      const schemeLabel = preview.scheme === 'command' ? t('diag.preview.command') : '';
      rows.push({
        key: 'preview',
        label: t('diag.preview'),
        present: previewAvailable,
        version: schemeLabel || (preview.probing ? t('diag.preview.probing') : ''),
        capability: t('diag.cap.preview') + ' · ' + t('diag.preview.note'),
        statusText: previewAvailable ? t('diag.preview.available') : (previewKnown ? t('diag.preview.unavailable') : t('diag.preview.unknown')),
        statusClass: previewAvailable ? 'dso-badge dso-badge--ok' : (previewKnown ? 'dso-badge dso-badge--error' : 'dso-badge dso-badge--idle'),
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
      /* P3 增量：预览服务可用性（suite.preview 探测一次，E2012 → 不可用） */
      const [previewState, setPreviewState] = React.useState(null);
      const [probingPreview, setProbingPreview] = React.useState(false);
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

      /**
       * 预览服务探测：调用 suite.preview 一次（服务端命令通道）。
       * 说明：Host 先校验 path 再判 officeToPdf 服务（office-pdf.js），因此必须传一个
       * 非空但不存在的探针路径：服务缺失 → E2012（不可用）；服务存在 → E1023/E1041
       * 等路径错误（可用）；COMMAND_MISSING 表示路由/命令通道不可达。
       */
      const probePreview = React.useCallback(async () => {
        setProbingPreview(true);
        setPreviewState((prev) => Object.assign({}, prev || {}, { probing: true }));
        try {
          await api.callCommand('suite.preview', { path: '_preview_probe_.docx' });
          setPreviewState({ available: true, scheme: 'command' });
        } catch (error) {
          const normalized = toPanelError(error, t);
          if (normalized.code === 'E2012' || normalized.code === 'PREVIEW_UNSUPPORTED') {
            setPreviewState({ available: false, scheme: null, code: normalized.code, message: normalized.message });
          } else if (normalized.code === 'COMMAND_MISSING') {
            setPreviewState({ available: null, scheme: null, code: normalized.code, message: normalized.message });
          } else {
            // 命令可达（仅探测参数被拒）：说明渲染服务存在
            setPreviewState({ available: true, scheme: 'command', note: normalized.code });
          }
        } finally {
          setProbingPreview(false);
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

      // 首次打开诊断页时探测一次预览服务可用性。
      React.useEffect(() => {
        if (!active) return;
        if (previewState || probingPreview) return;
        probePreview();
      }, [active, previewState, probingPreview, probePreview]);

      useConnectionReset(api, () => {
        probe(true);
        loadConfig();
        probePreview();
      });

      const rows = React.useMemo(
        () => buildEngineRows(env, t, Object.assign({ probing: probingPreview }, previewState || {})),
        [env, probingPreview, previewState, t],
      );
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
          lines.push([row.label, row.statusText || (row.present ? t('diag.engine.available') : t('diag.engine.missing')), row.version || '—', row.capability].join(' | '));
        }
        if (previewState) lines.push('preview: ' + JSON.stringify(previewState));
        if (config) lines.push('config: ' + JSON.stringify(config));
        if (envError) lines.push('envError: ' + envError.message);
        if (configError) lines.push('configError: ' + configError.message);
        const ok = await api.copy(lines.join('\n'));
        setCopied(ok);
        if (ok) window.setTimeout(() => setCopied(false), 1600);
      }, [api, config, configError, env, envError, previewState, rows, t]);

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
          } else if (spec.key === 'previewCacheSize') {
            value = Number(raw);
            if (!Number.isInteger(value) || value < 1 || value > 10) {
              throw { code: 'CONFIG_INVALID_VALUE', message: t('diag.invalidCacheSize') };
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
        if (spec.kind === 'number') {
          return h('input', {
            className: 'dso-input dso-mono',
            type: 'number',
            min: spec.min,
            max: spec.max,
            step: 1,
            value: value === undefined || value === null ? '' : String(value),
            disabled,
            onChange: (event) => setEdits((prev) => Object.assign({}, prev, { [spec.key]: event.target.value })),
          });
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
            h('button', {
              type: 'button',
              className: 'dso-btn dso-btn--sm',
              onClick: () => { probe(true); probePreview(); },
              disabled: probing || probingPreview,
            }, probing || probingPreview ? t('diag.probing') : t('diag.probe')),
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
                    className: row.statusClass || (row.present
                      ? (row.degraded ? 'dso-badge dso-badge--warn' : 'dso-badge dso-badge--ok')
                      : 'dso-badge dso-badge--idle'),
                  }, row.statusText || (row.present
                    ? (row.degraded ? t('diag.engine.degraded') : t('diag.engine.available'))
                    : t('diag.engine.missing')))),
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
     * 宿主命令调用不再依赖 remote.commands，统一走 Host 自注册的 `/office-suite/rpc`。
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
