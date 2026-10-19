# P2 Office COM 编辑/排版引擎（Windows PowerShell 5.1 兼容，UTF-8 BOM）
# 输入：-Json '{"mode":"edit"|"format","path":"...","ext":".docx","app":"word","ops":[...]|"spec":{...}, ...}'
# 输出：{ok:true,result:{outputPath,applied,stats,warnings?}} | {ok:false,error:{code,message,hint?},result?}
# 备份由 Host 层（host/ops/backup.js）负责，脚本内不做备份。
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ───────────────────────── 结构化错误 ─────────────────────────
class OpFailure : System.Exception {
  [string]$Code
  [string]$Hint
  OpFailure([string]$code, [string]$message, [string]$hint) : base($message) {
    $this.Code = $code
    $this.Hint = $hint
  }
}

$script:applied = New-Object System.Collections.ArrayList
$script:warnings = New-Object System.Collections.ArrayList
$script:outputPath = ''
$script:saved = $false
$script:result = $null
$script:exitCode = 0
$script:installedFonts = $null

function Add-Applied([string]$opName, [int]$index, [bool]$ok, [string]$message) {
  $entry = [ordered]@{ op = $opName; index = $index; ok = $ok }
  if ($message) { $entry.message = $message }
  [void]$script:applied.Add($entry)
}

function Add-Warning([string]$message) {
  if ($message -and -not ($script:warnings -contains $message)) { [void]$script:warnings.Add($message) }
}

function Set-FailureResult($failure, $stats) {
  $errorHash = [ordered]@{ code = [string]$failure.code; message = [string]$failure.message }
  if ($failure.hint) { $errorHash['hint'] = [string]$failure.hint }
  $resultHash = [ordered]@{
    outputPath = $script:outputPath
    saved      = [bool]$script:saved
    applied    = @($script:applied)
  }
  if ($stats) { $resultHash['stats'] = $stats }
  if ($script:warnings.Count -gt 0) { $resultHash['warnings'] = @($script:warnings) }
  $script:result = [ordered]@{ ok = $false; error = $errorHash; result = $resultHash }
  $script:exitCode = 1
}

function Set-SuccessResult($stats, $extra) {
  $resultHash = [ordered]@{
    outputPath = $script:outputPath
    saved      = $true
    applied    = @($script:applied)
    stats      = $stats
  }
  if ($script:warnings.Count -gt 0) { $resultHash['warnings'] = @($script:warnings) }
  if ($extra) {
    foreach ($key in $extra.Keys) { $resultHash[$key] = $extra[$key] }
  }
  $script:result = [ordered]@{ ok = $true; result = $resultHash }
  $script:exitCode = 0
}

# ───────────────────────── 通用小工具 ─────────────────────────

function Get-Text([object]$value) {
  if ($null -eq $value) { return '' }
  return [string]$value
}

function Test-HasProperty($object, [string]$name) {
  if ($null -eq $object) { return $false }
  return ($object.PSObject.Properties.Name -contains $name)
}

function Get-InstalledFontNames() {
  if ($null -ne $script:installedFonts) { return $script:installedFonts }
  $names = @()
  try {
    $props = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' -ErrorAction Stop
    foreach ($prop in $props.PSObject.Properties) {
      if ($prop.Name -notlike 'PS*') { $names += [string]$prop.Name }
    }
  } catch { }
  $script:installedFonts = $names
  return $names
}

function Test-FontInstalled([string]$fontName) {
  if (-not $fontName) { return $true }
  $names = Get-InstalledFontNames
  if (-not $names -or $names.Count -eq 0) { return $true }
  $aliases = @{
    '黑体' = @('simhei', '黑体')
    '宋体' = @('simsun', '宋体')
    '仿宋' = @('simfang', '仿宋')
    '仿宋_gb2312' = @('fangsong_gb2312', 'simfang', '仿宋')
    '楷体' = @('simkai', '楷体')
    '微软雅黑' = @('msyh', '微软雅黑')
    '新宋体' = @('nsimsun', '新宋体')
  }
  $needles = New-Object System.Collections.ArrayList
  [void]$needles.Add($fontName.ToLowerInvariant())
  $key = $fontName.ToLowerInvariant()
  if ($aliases.ContainsKey($key)) {
    foreach ($alias in $aliases[$key]) { [void]$needles.Add($alias) }
  }
  foreach ($name in $names) {
    $lower = $name.ToLowerInvariant()
    foreach ($needle in $needles) {
      if ($lower.Contains($needle)) { return $true }
    }
  }
  return $false
}

function Get-FormatCode([string]$ext, [string]$app) {
  $e = $ext.ToLowerInvariant()
  if ($app -eq 'excel') {
    if ($e -eq '.xls') { return 56 }
    return 51
  }
  if ($e -eq '.doc') { return 0 }
  return 16
}

# ───────────────────────── Word：段落 / 样式 / 表格 ─────────────────────────

function Get-WordStyle($doc, [string]$name) {
  if (-not $name) { throw [OpFailure]::new('E2005', '样式名为空', '') }
  $candidates = New-Object System.Collections.ArrayList
  [void]$candidates.Add($name)
  if ($name -match '^标题\s*([1-9])$') { [void]$candidates.Add("Heading $($Matches[1])") }
  if ($name -match '^Heading\s*([1-9])$') { [void]$candidates.Add("标题 $($Matches[1])") }
  if ($name -eq '正文') { [void]$candidates.Add('Normal') }
  if ($name -eq 'Normal') { [void]$candidates.Add('正文') }
  foreach ($candidate in $candidates) {
    try { return $doc.Styles.Item($candidate) } catch { }
  }
  throw [OpFailure]::new('E2005', "样式不存在: $name", '请使用 Word 中显示的样式名（NameLocal），如「标题 1」「正文」')
}

function Get-WordHeadingLevel($paragraph) {
  $styleName = ''
  try { $styleName = [string]$paragraph.Range.Style.NameLocal } catch { }
  if ($styleName -match '^(?:标题|Heading)\s*([1-9])$') { return [int]$Matches[1] }
  try {
    $outline = [int]$paragraph.Range.ParagraphFormat.OutlineLevel
    if ($outline -ge 1 -and $outline -le 9) { return $outline }
  } catch { }
  return 0
}

function Get-ParaText($paragraph) {
  try {
    $range = $paragraph.Range
    [void]$range.MoveEnd(1, -1)
    return [string]$range.Text
  } catch {
    return ''
  }
}

function Set-ParaText($paragraph, [string]$text) {
  $range = $paragraph.Range
  try { [void]$range.MoveEnd(1, -1) } catch { }
  $range.Text = [string]$text
}

function Insert-WordParagraph($doc, [int]$at) {
  $count = [int]$doc.Paragraphs.Count
  if ($at -lt 1 -or $at -gt ($count + 1)) {
    throw [OpFailure]::new('E2005', "段落索引 $at 超出范围（当前共 $count 段）", '')
  }
  if ($at -eq ($count + 1)) {
    return $doc.Paragraphs.Add()
  }
  $target = $doc.Paragraphs.Item($at)
  [void]$target.Range.InsertParagraphBefore()
  return $doc.Paragraphs.Item($at)
}

function Apply-HeadingStyle($doc, $paragraph, [int]$level) {
  $style = $null
  try { $style = Get-WordStyle $doc "标题 $level" } catch { }
  if ($null -eq $style) {
    try { $style = Get-WordStyle $doc "Heading $level" } catch { }
  }
  if ($null -eq $style) {
    throw [OpFailure]::new('E2005', "找不到标题样式（级别 $level）", '文档缺少「标题 N」/「Heading N」样式')
  }
  $paragraph.Range.Style = $style
}

function Invoke-WordParagraphOp($doc, $op) {
  $sub = [string]$op.op
  switch ($sub) {
    'setText' {
      $at = [int]$op.at
      if ($at -lt 1 -or $at -gt [int]$doc.Paragraphs.Count) { throw [OpFailure]::new('E2005', "段落索引 $at 超出范围（当前共 $($doc.Paragraphs.Count) 段）", '') }
      Set-ParaText ($doc.Paragraphs.Item($at)) (Get-Text $op.text)
    }
    'setStyle' {
      $at = [int]$op.at
      if ($at -lt 1 -or $at -gt [int]$doc.Paragraphs.Count) { throw [OpFailure]::new('E2005', "段落索引 $at 超出范围", '') }
      $style = Get-WordStyle $doc (Get-Text $op.style)
      $doc.Paragraphs.Item($at).Range.Style = $style
    }
    'insert' {
      $paragraph = Insert-WordParagraph $doc ([int]$op.at)
      Set-ParaText $paragraph (Get-Text $op.text)
    }
    'delete' {
      $at = [int]$op.at
      if ($at -lt 1 -or $at -gt [int]$doc.Paragraphs.Count) { throw [OpFailure]::new('E2005', "段落索引 $at 超出范围", '') }
      [void]$doc.Paragraphs.Item($at).Range.Delete()
    }
    'append' {
      $paragraph = $doc.Paragraphs.Add()
      Set-ParaText $paragraph (Get-Text $op.text)
    }
    default { throw [OpFailure]::new('E2005', "未知段落操作: $sub", '') }
  }
}

function Invoke-WordHeadingOp($doc, $op) {
  $sub = [string]$op.op
  $level = [int]$op.level
  switch ($sub) {
    'setText' {
      $at = [int]$op.at
      if ($at -lt 1 -or $at -gt [int]$doc.Paragraphs.Count) { throw [OpFailure]::new('E2005', "段落索引 $at 超出范围", '') }
      $paragraph = $doc.Paragraphs.Item($at)
      Set-ParaText $paragraph (Get-Text $op.text)
      Apply-HeadingStyle $doc $paragraph $level
    }
    'setStyle' {
      $at = [int]$op.at
      if ($at -lt 1 -or $at -gt [int]$doc.Paragraphs.Count) { throw [OpFailure]::new('E2005', "段落索引 $at 超出范围", '') }
      $paragraph = $doc.Paragraphs.Item($at)
      if ((Test-HasProperty $op 'style') -and $op.style) {
        $paragraph.Range.Style = Get-WordStyle $doc (Get-Text $op.style)
      } else {
        Apply-HeadingStyle $doc $paragraph $level
      }
    }
    'insert' {
      $paragraph = Insert-WordParagraph $doc ([int]$op.at)
      Set-ParaText $paragraph (Get-Text $op.text)
      Apply-HeadingStyle $doc $paragraph $level
    }
    'delete' {
      $at = [int]$op.at
      if ($at -lt 1 -or $at -gt [int]$doc.Paragraphs.Count) { throw [OpFailure]::new('E2005', "段落索引 $at 超出范围", '') }
      [void]$doc.Paragraphs.Item($at).Range.Delete()
    }
    default { throw [OpFailure]::new('E2005', "未知标题操作: $sub", '') }
  }
}

function Get-WordTable($doc, [int]$tableIndex) {
  if ($tableIndex -lt 1 -or $tableIndex -gt [int]$doc.Tables.Count) {
    throw [OpFailure]::new('E2005', "表格索引 $tableIndex 超出范围（共 $($doc.Tables.Count) 个表格）", '')
  }
  return $doc.Tables.Item($tableIndex)
}

function Invoke-WordTableOp($doc, $op) {
  $table = Get-WordTable $doc ([int]$op.table)
  $sub = [string]$op.op
  switch ($sub) {
    'setCell' {
      $row = [int]$op.row
      $col = [int]$op.col
      if ($row -lt 1 -or $row -gt [int]$table.Rows.Count) { throw [OpFailure]::new('E2005', "行索引 $row 超出范围（共 $($table.Rows.Count) 行）", '') }
      if ($col -lt 1 -or $col -gt [int]$table.Columns.Count) { throw [OpFailure]::new('E2005', "列索引 $col 超出范围（共 $($table.Columns.Count) 列）", '') }
      $table.Cell($row, $col).Range.Text = (Get-Text $op.text)
    }
    'addRow' {
      if ((Test-HasProperty $op 'at') -and $op.at -ne $null -and [string]$op.at -ne 'end' -and [string]$op.at -ne '') {
        $at = [int]$op.at
        if ($at -lt 1 -or $at -gt [int]$table.Rows.Count) { throw [OpFailure]::new('E2005', "插入行位置 $at 超出范围（共 $($table.Rows.Count) 行）", '') }
        [void]$table.Rows.Add($table.Rows.Item($at))
      } else {
        [void]$table.Rows.Add()
      }
    }
    'deleteRow' {
      $row = [int]$op.row
      if ($row -lt 1 -or $row -gt [int]$table.Rows.Count) { throw [OpFailure]::new('E2005', "行索引 $row 超出范围", '') }
      $table.Rows.Item($row).Delete()
    }
    'addColumn' {
      $position = $null
      if ((Test-HasProperty $op 'at') -and $op.at -ne $null -and [string]$op.at -ne 'end' -and [string]$op.at -ne '') { $position = [int]$op.at }
      elseif ((Test-HasProperty $op 'col') -and $op.col -ne $null) { $position = [int]$op.col }
      if ($null -ne $position) {
        if ($position -lt 1 -or $position -gt [int]$table.Columns.Count) { throw [OpFailure]::new('E2005', "插入列位置 $position 超出范围（共 $($table.Columns.Count) 列）", '') }
        [void]$table.Columns.Add($table.Columns.Item($position))
      } else {
        [void]$table.Columns.Add()
      }
    }
    'deleteColumn' {
      $col = [int]$op.col
      if ($col -lt 1 -or $col -gt [int]$table.Columns.Count) { throw [OpFailure]::new('E2005', "列索引 $col 超出范围", '') }
      $table.Columns.Item($col).Delete()
    }
    default { throw [OpFailure]::new('E2005', "未知表格操作: $sub", '') }
  }
}

function Save-WordDocument($doc, [string]$targetPath, [string]$ext) {
  if (-not $targetPath) { throw [OpFailure]::new('E2010', '保存路径为空', '') }
  $dir = [System.IO.Path]::GetDirectoryName($targetPath)
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
  $format = Get-FormatCode $ext 'word'
  try {
    $doc.SaveAs2($targetPath, $format) | Out-Null
  } catch {
    $doc.SaveAs($targetPath, $format) | Out-Null
  }
}

function Get-WordStats($doc) {
  $stats = [ordered]@{}
  try { $stats['paragraphs'] = [int]$doc.Paragraphs.Count } catch { $stats['paragraphs'] = 0 }
  try { $stats['tables'] = [int]$doc.Tables.Count } catch { $stats['tables'] = 0 }
  try { $stats['sections'] = [int]$doc.Sections.Count } catch { $stats['sections'] = 0 }
  try { $stats['words'] = [int]$doc.Words.Count } catch { }
  return $stats
}

function Invoke-WordEdit($doc, $ops) {
  $failure = $null
  for ($i = 0; $i -lt $ops.Count; $i++) {
    $op = $ops[$i]
    $action = [string]$op.action
    try {
      switch ($action) {
        'create' { }
        'paragraph' { Invoke-WordParagraphOp $doc $op }
        'heading' { Invoke-WordHeadingOp $doc $op }
        'table' { Invoke-WordTableOp $doc $op }
        'saveAs' { $script:outputPath = [string]$op.path }
        default { throw [OpFailure]::new('E2005', "未知编辑指令: $action", '') }
      }
      Add-Applied $action $i $true ''
    } catch [OpFailure] {
      Add-Applied $action $i $false $_.Exception.Message
      $failure = @{ code = $_.Exception.Code; message = $_.Exception.Message; hint = $_.Exception.Hint }
      break
    } catch {
      Add-Applied $action $i $false $_.Exception.Message
      $failure = @{ code = 'E2005'; message = "执行 $action 失败: $($_.Exception.Message)" }
      break
    }
  }
  return $failure
}

# ───────────────────────── Word：排版 ─────────────────────────

function Apply-FontToRange($range, $font) {
  if ($null -eq $font) { return }
  if ((Test-HasProperty $font 'name') -and $font.name) {
    $name = [string]$font.name
    $range.Font.Name = $name
    try { $range.Font.NameFarEast = $name } catch { }
    try { $range.Font.NameAscii = $name } catch { }
    try { $range.Font.NameOther = $name } catch { }
    if (-not (Test-FontInstalled $name)) {
      Add-Warning "字体未安装，已按名称设置（显示可能回退）: $name"
    }
  }
  if ((Test-HasProperty $font 'sizePt') -and $null -ne $font.sizePt) { $range.Font.Size = [double]$font.sizePt }
  if ((Test-HasProperty $font 'bold') -and $null -ne $font.bold) {
    if ([bool]$font.bold) { $range.Font.Bold = -1 } else { $range.Font.Bold = 0 }
  }
}

function Apply-ParagraphFormat($range, $paragraph) {
  if ($null -eq $paragraph) { return }
  $pf = $range.ParagraphFormat
  if ((Test-HasProperty $paragraph 'lineSpacing') -and $null -ne $paragraph.lineSpacing) {
    $pf.LineSpacingRule = 5
    $pf.LineSpacing = 12.0 * [double]$paragraph.lineSpacing
  } elseif ((Test-HasProperty $paragraph 'lineSpacingPt') -and $null -ne $paragraph.lineSpacingPt) {
    $pf.LineSpacingRule = 4
    $pf.LineSpacing = [double]$paragraph.lineSpacingPt
  }
  if ((Test-HasProperty $paragraph 'spaceBeforePt') -and $null -ne $paragraph.spaceBeforePt) { $pf.SpaceBefore = [double]$paragraph.spaceBeforePt }
  if ((Test-HasProperty $paragraph 'spaceAfterPt') -and $null -ne $paragraph.spaceAfterPt) { $pf.SpaceAfter = [double]$paragraph.spaceAfterPt }
  if ((Test-HasProperty $paragraph 'align') -and $paragraph.align) {
    $map = @{ left = 0; center = 1; right = 2; justify = 3 }
    $key = [string]$paragraph.align
    if ($map.ContainsKey($key)) { $pf.Alignment = $map[$key] }
  }
}

function ConvertTo-ChineseNumber([int]$n) {
  if ($n -le 0) { return '零' }
  $digits = @('零', '一', '二', '三', '四', '五', '六', '七', '八', '九')
  if ($n -lt 10) { return $digits[$n] }
  if ($n -eq 10) { return '十' }
  if ($n -lt 20) { return '十' + $digits[$n - 10] }
  if ($n -lt 100) {
    $tens = [int][Math]::Floor($n / 10)
    $units = $n % 10
    $text = $digits[$tens] + '十'
    if ($units -gt 0) { $text += $digits[$units] }
    return $text
  }
  return [string]$n
}

function Format-NumberPrefix([int]$level, $counters, [string]$mode) {
  if ($mode -eq 'none' -or -not $mode) { return '' }
  if ($mode -eq 'chinese') {
    $cn = ConvertTo-ChineseNumber ([int]$counters[$level])
    if ($level -eq 1) { return $cn + '、' }
    if ($level -eq 2) { return '（' + $cn + '）' }
    return ([string][int]$counters[$level]) + '. '
  }
  if ($mode -eq 'arabic') { return ([string][int]$counters[$level]) + '. ' }
  if ($mode -eq 'decimalDot') {
    $parts = @()
    for ($i = 1; $i -le $level; $i++) { $parts += [string][int]$counters[$i] }
    return ($parts -join '.') + ' '
  }
  return ''
}

function Remove-NumberPrefix([string]$text) {
  if (-not $text) { return '' }
  return [regex]::Replace(
    $text,
    '^\s*(?:[一二三四五六七八九十百零〇]+[、.．]|（[一二三四五六七八九十百零〇]+）|\([0-9一二三四五六七八九十]+\)|[0-9]+(?:\.[0-9]+)*[、.．]?\s*)',
    ''
  )
}

function Apply-WordNumbering($doc, $rules, $scope) {
  $numberingByLevel = @{}
  foreach ($rule in $rules) {
    if ([string]$rule.target -ne 'heading') { continue }
    if (-not (Test-HasProperty $rule 'numbering')) { continue }
    $level = [int]$rule.level
    if ($level -ge 1 -and $level -le 3) { $numberingByLevel[$level] = [string]$rule.numbering }
  }
  if ($numberingByLevel.Count -eq 0) { return }
  if ($scope -and $scope.type -eq 'paragraphs') { return }

  $counters = @{ 1 = 0; 2 = 0; 3 = 0 }
  foreach ($paragraph in $doc.Paragraphs) {
    $level = Get-WordHeadingLevel $paragraph
    if ($level -lt 1 -or $level -gt 3) { continue }
    if ($scope -and $scope.type -eq 'headings' -and -not ($scope.headings -contains $level)) { continue }
    $counters[$level] = [int]$counters[$level] + 1
    for ($deeper = $level + 1; $deeper -le 3; $deeper++) { $counters[$deeper] = 0 }
    if (-not $numberingByLevel.ContainsKey($level)) { continue }
    $mode = [string]$numberingByLevel[$level]
    if ($mode -eq 'none') { continue }
    $prefix = Format-NumberPrefix $level $counters $mode
    $text = Get-ParaText $paragraph
    $stripped = Remove-NumberPrefix $text
    Set-ParaText $paragraph ($prefix + $stripped)
  }
}

function Apply-PageSetupRule($doc, $page) {
  if ($null -eq $page -or $null -eq $page.marginsMm) { return 0 }
  $mmToPt = 72.0 / 25.4
  $count = 0
  foreach ($section in $doc.Sections) {
    $setup = $section.PageSetup
    $margins = $page.marginsMm
    if (Test-HasProperty $margins 'top') { $setup.TopMargin = [double]$margins.top * $mmToPt }
    if (Test-HasProperty $margins 'bottom') { $setup.BottomMargin = [double]$margins.bottom * $mmToPt }
    if (Test-HasProperty $margins 'left') { $setup.LeftMargin = [double]$margins.left * $mmToPt }
    if (Test-HasProperty $margins 'right') { $setup.RightMargin = [double]$margins.right * $mmToPt }
    $count += 1
  }
  return $count
}

function Apply-HeaderFooterRule($doc, $rule) {
  $isHeader = ([string]$rule.target -eq 'header')
  $payload = if ($isHeader) { $rule.header } else { $rule.footer }
  if ($null -eq $payload) { return 0 }
  $count = 0
  foreach ($section in $doc.Sections) {
    if ($isHeader) { $container = $section.Headers.Item(1) } else { $container = $section.Footers.Item(1) }
    $range = $container.Range
    $range.Text = (Get-Text $payload.text)
    if ($null -ne $payload.font) { Apply-FontToRange $range $payload.font }
    if (-not $isHeader) {
      try { Add-PageFields $container $range } catch { }
    }
    $count += 1
  }
  return $count
}

function Add-PageFields($container, $range) {
  $text = [string]$range.Text
  if ($text -notmatch '[XY]') { return }
  foreach ($pair in @(@('X', 33), @('Y', 26))) {
    $find = $range.Find
    $find.ClearFormatting()
    $find.Text = [string]$pair[0]
    $find.Forward = $true
    $find.Wrap = 1
    $find.MatchCase = $true
    $guard = 0
    while ($guard -lt 5) {
      $guard += 1
      $found = $find.Execute([string]$pair[0], $true, $false, $false, $false, $false, $true, 1)
      if (-not $found) { break }
      [void]$range.Fields.Add($range, [int]$pair[1])
    }
  }
}

function Apply-WordRule($doc, $rule, $scope) {
  $target = [string]$rule.target
  if ($target -eq 'page') { return Apply-PageSetupRule $doc $rule.page }
  if ($target -eq 'header' -or $target -eq 'footer') { return Apply-HeaderFooterRule $doc $rule }

  $count = 0
  if ($scope -and $scope.type -eq 'paragraphs') {
    foreach ($idx in $scope.paragraphs) {
      if ([int]$idx -lt 1 -or [int]$idx -gt [int]$doc.Paragraphs.Count) {
        throw [OpFailure]::new('E2006', "scope.paragraphs 索引 $idx 超出范围（共 $($doc.Paragraphs.Count) 段）", '')
      }
      $paragraph = $doc.Paragraphs.Item([int]$idx)
      if ($null -ne $rule.font) { Apply-FontToRange $paragraph.Range $rule.font }
      if ($null -ne $rule.paragraph) { Apply-ParagraphFormat $paragraph.Range $rule.paragraph }
      $count += 1
    }
    return $count
  }

  if ($target -eq 'heading') {
    $level = [int]$rule.level
    if ($scope -and $scope.type -eq 'headings' -and -not ($scope.headings -contains $level)) { return 0 }
    foreach ($paragraph in $doc.Paragraphs) {
      if ((Get-WordHeadingLevel $paragraph) -ne $level) { continue }
      if ($null -ne $rule.font) { Apply-FontToRange $paragraph.Range $rule.font }
      if ($null -ne $rule.paragraph) { Apply-ParagraphFormat $paragraph.Range $rule.paragraph }
      $count += 1
    }
    return $count
  }

  if ($target -eq 'body') {
    foreach ($paragraph in $doc.Paragraphs) {
      if ((Get-WordHeadingLevel $paragraph) -ne 0) { continue }
      $inTable = $false
      try { $inTable = [bool]$paragraph.Range.Information(12) } catch { }
      if ($inTable) { continue }
      $text = Get-ParaText $paragraph
      if (-not $text -or -not $text.Trim()) { continue }
      if ($null -ne $rule.font) { Apply-FontToRange $paragraph.Range $rule.font }
      if ($null -ne $rule.paragraph) { Apply-ParagraphFormat $paragraph.Range $rule.paragraph }
      $count += 1
    }
    return $count
  }

  throw [OpFailure]::new('E2006', "未知排版目标: $target", '')
}

function Invoke-WordFormat($doc, $spec) {
  $rules = @()
  if ($spec -and (Test-HasProperty $spec 'rules') -and $null -ne $spec.rules) { $rules = @($spec.rules) }
  $scope = $null
  if ($spec -and (Test-HasProperty $spec 'scope')) { $scope = $spec.scope }

  Apply-WordNumbering $doc $rules $scope

  $failure = $null
  for ($i = 0; $i -lt $rules.Count; $i++) {
    $rule = $rules[$i]
    $target = [string]$rule.target
    try {
      $count = Apply-WordRule $doc $rule $scope
      $entry = [ordered]@{ rule = $i; target = $target; count = $count; ok = $true }
      if ($target -eq 'heading') { $entry['level'] = [int]$rule.level }
      [void]$script:applied.Add($entry)
    } catch [OpFailure] {
      $entry = [ordered]@{ rule = $i; target = $target; count = 0; ok = $false; message = $_.Exception.Message }
      [void]$script:applied.Add($entry)
      $failure = @{ code = $_.Exception.Code; message = $_.Exception.Message; hint = $_.Exception.Hint }
      break
    } catch {
      $entry = [ordered]@{ rule = $i; target = $target; count = 0; ok = $false; message = $_.Exception.Message }
      [void]$script:applied.Add($entry)
      $failure = @{ code = 'E2006'; message = "应用排版规则失败: $($_.Exception.Message)" }
      break
    }
  }
  return $failure
}

# ───────────────────────── Word：模板占位符 ─────────────────────────

function Expand-Placeholders($doc, $placeholders) {
  if ($null -eq $placeholders) { return 0 }
  $keys = @($placeholders.PSObject.Properties.Name)
  if ($keys.Count -eq 0) { return 0 }
  $count = 0
  foreach ($story in $doc.StoryRanges) {
    foreach ($key in $keys) {
      $value = [string]$placeholders.$key
      foreach ($token in @(('【' + $key + '】'), ('{{' + $key + '}}'))) {
        try {
          $find = $story.Find
          $find.ClearFormatting()
          $find.Replacement.ClearFormatting()
          $find.Text = $token
          $find.Replacement.Text = $value
          $find.Forward = $true
          $find.Wrap = 1
          $find.MatchCase = $false
          $ok = $find.Execute($token, $false, $false, $false, $false, $false, $true, 1, $false, $value, 2)
          if ($ok) { $count += 1 }
        } catch { }
      }
    }
  }
  return $count
}

# ───────────────────────── Excel：编辑 ─────────────────────────

function Get-Worksheet($workbook, [string]$name) {
  if (-not $name) { return $workbook.ActiveSheet }
  try {
    return $workbook.Worksheets.Item($name)
  } catch {
    throw [OpFailure]::new('E2005', "工作表不存在: $name", '')
  }
}

function Get-ExcelColumnIndex($worksheet, $value) {
  if ($value -is [int] -or $value -is [long] -or $value -is [double]) { return [int]$value }
  $text = [string]$value
  if ($text -match '^\d+$') { return [int]$text }
  if ($text -match '^[A-Za-z]{1,3}$') {
    return [int]$worksheet.Range(($text.ToUpperInvariant() + '1')).Column
  }
  throw [OpFailure]::new('E2005', "非法列标识: $value", '')
}

function Invoke-ExcelOp($workbook, $op) {
  $action = [string]$op.action
  switch ($action) {
    'create' { }
    'cell' {
      $worksheet = Get-Worksheet $workbook ([string]$op.sheet)
      $cell = $worksheet.Range([string]$op.at)
      if ((Test-HasProperty $op 'formula') -and $op.formula) {
        $cell.Formula = [string]$op.formula
      } else {
        $cell.Value2 = $op.value
      }
    }
    'row' {
      $worksheet = Get-Worksheet $workbook ([string]$op.sheet)
      $index = [int]$op.at
      if ([string]$op.op -eq 'insert') {
        $worksheet.Rows.Item($index).Insert() | Out-Null
      } elseif ([string]$op.op -eq 'delete') {
        $worksheet.Rows.Item($index).Delete() | Out-Null
      } else {
        throw [OpFailure]::new('E2005', "未知行操作: $($op.op)", '')
      }
    }
    'column' {
      $worksheet = Get-Worksheet $workbook ([string]$op.sheet)
      $index = Get-ExcelColumnIndex $worksheet $op.at
      if ([string]$op.op -eq 'insert') {
        $worksheet.Columns.Item($index).Insert() | Out-Null
      } elseif ([string]$op.op -eq 'delete') {
        $worksheet.Columns.Item($index).Delete() | Out-Null
      } else {
        throw [OpFailure]::new('E2005', "未知列操作: $($op.op)", '')
      }
    }
    'sheet' {
      $sub = [string]$op.op
      $name = [string]$op.name
      if ($sub -eq 'rename') {
        $worksheet = Get-Worksheet $workbook $name
        $worksheet.Name = [string]$op.newName
      } elseif ($sub -eq 'add') {
        $sheet = $workbook.Worksheets.Add()
        try { $sheet.Name = $name } catch { throw [OpFailure]::new('E2005', "新建工作表失败（名称可能重复）: $name", '') }
      } elseif ($sub -eq 'delete') {
        $worksheet = Get-Worksheet $workbook $name
        if ([int]$workbook.Worksheets.Count -le 1) { throw [OpFailure]::new('E2005', '不能删除工作簿中的唯一工作表', '') }
        $worksheet.Delete()
      } else {
        throw [OpFailure]::new('E2005', "未知工作表操作: $sub", '')
      }
    }
    'saveAs' { $script:outputPath = [string]$op.path }
    default { throw [OpFailure]::new('E2005', "未知 Excel 指令: $action", '') }
  }
}

function Save-Workbook($workbook, [string]$targetPath, [string]$ext) {
  if (-not $targetPath) { throw [OpFailure]::new('E2010', '保存路径为空', '') }
  $dir = [System.IO.Path]::GetDirectoryName($targetPath)
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
  $format = Get-FormatCode $ext 'excel'
  try {
    $workbook.SaveAs2($targetPath, $format) | Out-Null
  } catch {
    $workbook.SaveAs($targetPath, $format) | Out-Null
  }
}

function Get-ExcelStats($workbook) {
  $stats = [ordered]@{}
  try { $stats['sheets'] = [int]$workbook.Worksheets.Count } catch { $stats['sheets'] = 0 }
  try {
    $used = $workbook.ActiveSheet.UsedRange
    $stats['rows'] = [int]$used.Rows.Count
    $stats['cols'] = [int]$used.Columns.Count
  } catch { }
  return $stats
}

function Invoke-ExcelEdit($workbook, $ops) {
  $failure = $null
  for ($i = 0; $i -lt $ops.Count; $i++) {
    $op = $ops[$i]
    $action = [string]$op.action
    try {
      Invoke-ExcelOp $workbook $op
      Add-Applied $action $i $true ''
    } catch [OpFailure] {
      Add-Applied $action $i $false $_.Exception.Message
      $failure = @{ code = $_.Exception.Code; message = $_.Exception.Message; hint = $_.Exception.Hint }
      break
    } catch {
      Add-Applied $action $i $false $_.Exception.Message
      $failure = @{ code = 'E2005'; message = "执行 $action 失败: $($_.Exception.Message)" }
      break
    }
  }
  return $failure
}

# ───────────────────────── 主流程 ─────────────────────────

$word = $null
$excel = $null
$document = $null
$workbook = $null

try {
  $arg = $Json | ConvertFrom-Json
  $mode = [string]$arg.mode
  $path = [string]$arg.path
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
  $app = [string]$arg.app
  if (-not $app) { if ($ext -in @('.xls', '.xlsx')) { $app = 'excel' } else { $app = 'word' } }

  if (-not $path) { throw [OpFailure]::new('E1041', '缺少 path', '') }
  if ($mode -ne 'edit' -and $mode -ne 'format') { throw [OpFailure]::new('E1041', "mode 必须为 'edit' | 'format'", '') }
  if ($mode -eq 'edit' -and $null -eq $arg.ops) { throw [OpFailure]::new('E1041', 'edit 模式缺少 ops', '') }
  if ($mode -eq 'format' -and $null -eq $arg.spec) { throw [OpFailure]::new('E1041', 'format 模式缺少 spec', '') }

  if ((Test-HasProperty $arg 'outputPath') -and $arg.outputPath) { $script:outputPath = [string]$arg.outputPath } else { $script:outputPath = $path }
  $outExt = [System.IO.Path]::GetExtension($script:outputPath).ToLowerInvariant()
  if (-not $outExt) { $outExt = $ext }

  if ($app -eq 'excel' -or $ext -in @('.xls', '.xlsx')) {
    # ── Excel ──
    try {
      $excel = New-Object -ComObject Excel.Application
    } catch {
      throw [OpFailure]::new('E2004', "无法启动 Excel COM: $($_.Exception.Message)", '请确认 Excel 已安装且未被其它进程占用')
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      if ($arg.create -eq $true) {
        $workbook = $excel.Workbooks.Add()
      } else {
        if (-not (Test-Path -LiteralPath $path)) { throw [OpFailure]::new('E2004', "文件不存在: $path", '') }
        $workbook = $excel.Workbooks.Open($path, 0, $false)
      }
    } catch [OpFailure] {
      throw
    } catch {
      throw [OpFailure]::new('E2004', "打开工作簿失败: $($_.Exception.Message)", "请关闭 Excel 中已打开的文件后重试：$path")
    }
    if ($null -eq $workbook) { throw [OpFailure]::new('E2004', "无法打开工作簿: $path", '') }
    if ($mode -eq 'format') { throw [OpFailure]::new('E2006', 'Excel 排版不在 P2 范围（仅 Word）', '') }

    $failure = Invoke-ExcelEdit $workbook $arg.ops
    $saveError = $null
    try {
      Save-Workbook $workbook $script:outputPath $outExt
      $script:saved = $true
    } catch {
      $saveError = $_.Exception.Message
      $script:saved = $false
      if (-not $failure) { $failure = @{ code = 'E2010'; message = "保存工作簿失败: $saveError" } }
    }
    $stats = Get-ExcelStats $workbook
    if ($failure) { Set-FailureResult $failure $stats } else { Set-SuccessResult $stats $null }
  }
  else {
    # ── Word ──
    try {
      $word = New-Object -ComObject Word.Application
    } catch {
      throw [OpFailure]::new('E2004', "无法启动 Word COM: $($_.Exception.Message)", '请确认 Word 已安装且未被其它进程占用')
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $templatePath = ''
    if ((Test-HasProperty $arg 'templatePath') -and $arg.templatePath) { $templatePath = [string]$arg.templatePath }

    if ($mode -eq 'format' -and $templatePath) {
      if (-not (Test-Path -LiteralPath $templatePath)) { throw [OpFailure]::new('E2007', "模板不存在: $templatePath", '') }
      try {
        $document = $word.Documents.Open($templatePath, $false, $false)
      } catch {
        throw [OpFailure]::new('E2004', "打开模板失败: $($_.Exception.Message)", "模板：$templatePath")
      }
      if ($null -eq $document) { throw [OpFailure]::new('E2004', "无法打开模板: $templatePath", '') }
      $replaced = Expand-Placeholders $document $arg.placeholders
      try {
        Save-WordDocument $document $script:outputPath $outExt
        $script:saved = $true
      } catch {
        throw [OpFailure]::new('E2010', "保存排版结果失败: $($_.Exception.Message)", '')
      }
      [void]$script:applied.Add([ordered]@{ rule = 'template'; target = 'template'; count = $replaced; ok = $true })
      $stats = Get-WordStats $document
      Set-SuccessResult $stats ([ordered]@{ templatePlaceholders = $replaced })
    }
    else {
      try {
        if ($arg.create -eq $true) {
          $document = $word.Documents.Add()
        } else {
          if (-not (Test-Path -LiteralPath $path)) { throw [OpFailure]::new('E2004', "文件不存在: $path", '') }
          $document = $word.Documents.Open($path, $false, $false)
        }
      } catch [OpFailure] {
        throw
      } catch {
        throw [OpFailure]::new('E2004', "打开文档失败: $($_.Exception.Message)", "请关闭 Word 中已打开的文件后重试：$path")
      }
      if ($null -eq $document) { throw [OpFailure]::new('E2004', "无法打开文档: $path", '') }

      if ($mode -eq 'format') {
        $failure = Invoke-WordFormat $document $arg.spec
      } else {
        $failure = Invoke-WordEdit $document $arg.ops
      }

      $saveError = $null
      try {
        Save-WordDocument $document $script:outputPath $outExt
        $script:saved = $true
      } catch {
        $saveError = $_.Exception.Message
        $script:saved = $false
        if (-not $failure) { $failure = @{ code = 'E2010'; message = "保存文档失败: $saveError" } }
      }
      $stats = Get-WordStats $document
      if ($failure) { Set-FailureResult $failure $stats } else { Set-SuccessResult $stats $null }
    }
  }
}
catch [OpFailure] {
  $failure = @{ code = $_.Exception.Code; message = $_.Exception.Message; hint = $_.Exception.Hint }
  Set-FailureResult $failure $null
}
catch {
  $failure = @{ code = 'E1051'; message = "Office COM 引擎异常: $($_.Exception.Message)" }
  Set-FailureResult $failure $null
}
finally {
  if ($null -ne $document) {
    try { $document.Close(0) | Out-Null } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) } catch { }
    $document = $null
  }
  if ($null -ne $word) {
    try { $word.Quit() | Out-Null } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    $word = $null
  }
  if ($null -ne $workbook) {
    try { $workbook.Close($false) | Out-Null } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch { }
    $workbook = $null
  }
  if ($null -ne $excel) {
    try { $excel.Quit() | Out-Null } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    $excel = $null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

if ($null -eq $script:result) {
  $script:result = [ordered]@{
    ok    = $false
    error = [ordered]@{ code = 'E1051'; message = 'Office COM 引擎未产生结果' }
  }
  $script:exitCode = 1
}
$script:result | ConvertTo-Json -Depth 12 -Compress
exit $script:exitCode
