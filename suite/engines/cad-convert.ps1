[CmdletBinding()]
param([string]$Json = '{}')

# ============================================================================
# CAD → PDF COM 引擎骨架（docs/P1-设计.md §3.1 / §3.3）
#
# 场景 A（mode=split）：打开 DWG → 识别图框（块名 + 矩形尺寸双重依据）→
#                      逐图框 SetWindowToPlot → PlotToFile（DWG To PDF.pc3，A4 横向）。
# 场景 B（mode=merge）：打开 DWG → 模型空间 Layout → PlotType=extents →
#                      PlotToFile 单文件 PDF（供 Node 侧合并）。
#
# 输出：{ ok:true, result:{...} } | { ok:false, error:{ code, message, hint? } }
# 错误码：E1001 无 COM / E1003 打开失败 / E1004 打印失败 / E1012 未识别图框 /
#         E1023 输入不存在 / E1041 参数错误 / E1051 其它引擎失败。
#
# 说明：本机未安装 AutoCAD/中望/浩辰，本脚本未经真机联调（P3/D5 联调项）；
#       所有 COM 属性/方法均以 late binding 调用，逐步骤 try/catch 并映射错误码。
# ============================================================================

$ErrorActionPreference = 'Stop'

# ───────────────────────── 通用工具 ─────────────────────────

function New-CadError([string]$code, [string]$message, [string]$hint) {
  $exception = New-Object System.Exception($message)
  $exception.Data['code'] = $code
  if ($hint) { $exception.Data['hint'] = $hint }
  return $exception
}

function Test-ProcessName($names) {
  foreach ($name in @($names)) {
    if (Get-Process -Name $name -ErrorAction SilentlyContinue) { return $true }
  }
  return $false
}

function Test-BlockName([string]$name, $patterns) {
  foreach ($pattern in @($patterns)) {
    $wildcard = New-Object System.Management.Automation.WildcardPattern([string]$pattern, [System.Management.Automation.WildcardOptions]::IgnoreCase)
    if ($wildcard.IsMatch($name)) { return $true }
  }
  return $false
}

function Get-DefaultFrameSizes {
  return @(
    [ordered]@{ name = 'A4'; width = 297; height = 210 },
    [ordered]@{ name = 'A3'; width = 420; height = 297 },
    [ordered]@{ name = 'A2'; width = 594; height = 420 },
    [ordered]@{ name = 'A1'; width = 841; height = 594 },
    [ordered]@{ name = 'A0'; width = 1189; height = 841 },
    [ordered]@{ name = 'A4x3'; width = 891; height = 210 },
    [ordered]@{ name = 'A4x4'; width = 1189; height = 210 },
    [ordered]@{ name = 'A3x3'; width = 891; height = 420 },
    [ordered]@{ name = 'A3x4'; width = 1189; height = 420 },
    [ordered]@{ name = 'A2x3'; width = 1261; height = 594 },
    [ordered]@{ name = 'A2x4'; width = 1682; height = 594 },
    [ordered]@{ name = 'A2x5'; width = 2102; height = 594 },
    [ordered]@{ name = 'A1x3'; width = 1783; height = 841 },
    [ordered]@{ name = 'A1x4'; width = 2378; height = 841 },
    [ordered]@{ name = 'A0x2'; width = 1682; height = 1189 },
    [ordered]@{ name = 'A0x3'; width = 2523; height = 1189 }
  )
}

function Get-FrameSizeList($inputSizes) {
  $list = @()
  if ($null -ne $inputSizes) {
    foreach ($entry in @($inputSizes)) {
      if ($null -eq $entry) { continue }
      $name = [string]$entry.name
      $width = [double]$entry.width
      $height = [double]$entry.height
      if ($name -and $width -gt 0 -and $height -gt 0) {
        $list += [ordered]@{ name = $name; width = $width; height = $height }
      }
    }
  }
  if ($list.Count -eq 0) { $list = Get-DefaultFrameSizes }
  return $list
}

function Match-FrameSize([double]$width, [double]$height, $sizes, [double]$tolerance) {
  if ($width -le 0 -or $height -le 0) { return $null }
  $long = [Math]::Max($width, $height)
  $short = [Math]::Min($width, $height)
  foreach ($size in @($sizes)) {
    $sizeLong = [Math]::Max([double]$size.width, [double]$size.height)
    $sizeShort = [Math]::Min([double]$size.width, [double]$size.height)
    if ([Math]::Abs($long - $sizeLong) -gt $tolerance) { continue }
    if ([Math]::Abs($short - $sizeShort) -gt $tolerance) { continue }
    $expected = $sizeLong / $sizeShort
    $actual = $long / $short
    if (([Math]::Abs($actual - $expected) / $expected) -gt 0.005) { continue }
    return [string]$size.name
  }
  return $null
}

function Get-EntityBoundingBox($entity) {
  $min = New-Object 'double[]' 3
  $max = New-Object 'double[]' 3
  try {
    $null = $entity.GetBoundingBox($min, $max)
  } catch {
    try {
      $null = $entity.GetBoundingBox([ref]$min, [ref]$max)
    } catch {
      return $null
    }
  }
  if ($null -eq $min -or $null -eq $max) { return $null }
  if ($min.Length -lt 2 -or $max.Length -lt 2) { return $null }
  return [ordered]@{
    minX = [double]$min[0]; minY = [double]$min[1]
    maxX = [double]$max[0]; maxY = [double]$max[1]
  }
}

function Get-IoU($a, $b) {
  $left = [Math]::Max([double]$a.minX, [double]$b.minX)
  $right = [Math]::Min([double]$a.maxX, [double]$b.maxX)
  $bottom = [Math]::Max([double]$a.minY, [double]$b.minY)
  $top = [Math]::Min([double]$a.maxY, [double]$b.maxY)
  if ($right -le $left -or $top -le $bottom) { return 0.0 }
  $inter = ($right - $left) * ($top - $bottom)
  $areaA = ([double]$a.maxX - [double]$a.minX) * ([double]$a.maxY - [double]$a.minY)
  $areaB = ([double]$b.maxX - [double]$b.minX) * ([double]$b.maxY - [double]$b.minY)
  $union = $areaA + $areaB - $inter
  if ($union -le 0) { return 0.0 }
  return $inter / $union
}

function Test-Contains($outer, $inner) {
  return (([double]$inner.minX -ge [double]$outer.minX - 0.5) -and
          ([double]$inner.minY -ge [double]$outer.minY - 0.5) -and
          ([double]$inner.maxX -le [double]$outer.maxX + 0.5) -and
          ([double]$inner.maxY -le [double]$outer.maxY + 0.5))
}

function Merge-FrameCandidates($candidates) {
  $sorted = @($candidates | Sort-Object { ([double]$_.maxX - [double]$_.minX) * ([double]$_.maxY - [double]$_.minY) } -Descending)
  $kept = New-Object System.Collections.ArrayList
  foreach ($candidate in $sorted) {
    $drop = $false
    foreach ($existing in $kept) {
      if ((Get-IoU $candidate $existing) -gt 0.9) { $drop = $true; break }
      $candidateArea = ([double]$candidate.maxX - [double]$candidate.minX) * ([double]$candidate.maxY - [double]$candidate.minY)
      $existingArea = ([double]$existing.maxX - [double]$existing.minX) * ([double]$existing.maxY - [double]$existing.minY)
      if ((Test-Contains $existing $candidate) -and ($candidateArea -lt $existingArea * 0.5)) { $drop = $true; break }
    }
    if (-not $drop) { [void]$kept.Add($candidate) }
  }
  return @($kept)
}

function Get-FrameCandidates($doc, $frameBlockNames, $frameSizes, [double]$tolerance) {
  $candidates = @()
  foreach ($entity in $doc.ModelSpace) {
    $objectName = [string]$entity.ObjectName
    $isBlock = ($objectName -match 'BlockReference')
    if ($isBlock) {
      $blockName = ''
      try { $blockName = [string]$entity.EffectiveName } catch {
        try { $blockName = [string]$entity.Name } catch { $blockName = '' }
      }
      if (-not (Test-BlockName $blockName $frameBlockNames)) { continue }
    } elseif ($objectName -match 'AcDbPolyline') {
      $closed = $false
      try { $closed = [bool]$entity.Closed } catch { $closed = $false }
      if (-not $closed) { continue }
    } else {
      continue
    }
    $bbox = Get-EntityBoundingBox $entity
    if (-not $bbox) { continue }
    $width = [double]$bbox.maxX - [double]$bbox.minX
    $height = [double]$bbox.maxY - [double]$bbox.minY
    if ($width -le 1 -or $height -le 1) { continue }
    $sizeName = Match-FrameSize $width $height $frameSizes $tolerance
    $unitScale = 1.0
    if (-not $sizeName) {
      # 图幅按 m 绘制时尝试换算（1m = 1000mm）
      $sizeName = Match-FrameSize ($width * 1000) ($height * 1000) $frameSizes $tolerance
      if ($sizeName) { $unitScale = 1000.0 }
    }
    if (-not $sizeName) { continue }
    $candidates += [ordered]@{
      minX      = [double]$bbox.minX
      minY      = [double]$bbox.minY
      maxX      = [double]$bbox.maxX
      maxY      = [double]$bbox.maxY
      width     = $width
      height    = $height
      widthMm   = $width * $unitScale
      heightMm  = $height * $unitScale
      frameSize = $sizeName
      source    = $(if ($isBlock) { 'block' } else { 'polyline' })
      unitScale = $unitScale
    }
  }
  return @($candidates)
}

function Get-TextItems($doc) {
  $items = @()
  foreach ($entity in $doc.ModelSpace) {
    $objectName = [string]$entity.ObjectName
    if ($objectName -match 'AcDbText$|AcDbMText$') {
      try {
        $text = [string]$entity.TextString
        $point = $entity.InsertionPoint
        $height = 0.0
        try { $height = [double]$entity.Height } catch {
          try { $height = [double]$entity.CharacterHeight } catch { $height = 0.0 }
        }
        if ($text -and $null -ne $point) {
          $items += [ordered]@{ text = $text; x = [double]$point[0]; y = [double]$point[1]; height = $height }
        }
      } catch { }
    }
    if ($objectName -match 'BlockReference') {
      try {
        foreach ($attribute in $entity.GetAttributes()) {
          $text = [string]$attribute.TextString
          $point = $attribute.InsertionPoint
          $height = 0.0
          try { $height = [double]$attribute.Height } catch { $height = 0.0 }
          if ($text -and $null -ne $point) {
            $items += [ordered]@{ text = $text; x = [double]$point[0]; y = [double]$point[1]; height = $height }
          }
        }
      } catch { }
    }
  }
  return @($items)
}

function Get-DrawingNo($textItems, $frame, $zone, $patterns) {
  $x0 = [double]$zone[0]; $y0 = [double]$zone[1]
  $x1 = [double]$zone[2]; $y1 = [double]$zone[3]
  $frameWidth = [double]$frame.maxX - [double]$frame.minX
  $frameHeight = [double]$frame.maxY - [double]$frame.minY
  if ($frameWidth -le 0 -or $frameHeight -le 0) { return $null }
  $selected = @()
  foreach ($item in $textItems) {
    $relativeX = ([double]$item.x - [double]$frame.minX) / $frameWidth
    $relativeY = ([double]$item.y - [double]$frame.minY) / $frameHeight
    if ($relativeX -ge $x0 -and $relativeX -le $x1 -and $relativeY -ge $y0 -and $relativeY -le $y1) {
      $selected += [ordered]@{
        text = [string]$item.text
        x = $relativeX
        y = $relativeY
        height = ([double]$item.height / $frameHeight)
      }
    }
  }
  if ($selected.Count -eq 0) { return $null }

  # 行合并：y 差 < 0.6 × 行高
  $lines = New-Object System.Collections.ArrayList
  foreach ($item in @($selected | Sort-Object { [double]$_.y } -Descending)) {
    $target = $null
    foreach ($line in $lines) {
      $tolerance = 0.6 * [Math]::Max([double]$line.height, [double]$item.height)
      if ([Math]::Abs([double]$line.y - [double]$item.y) -le $tolerance) { $target = $line; break }
    }
    if (-not $target) {
      $target = @{ y = [double]$item.y; height = [double]$item.height; items = (New-Object System.Collections.ArrayList) }
      [void]$lines.Add($target)
    } else {
      $target.height = [Math]::Max([double]$target.height, [double]$item.height)
    }
    [void]$target.items.Add($item)
  }
  $lineTexts = @()
  foreach ($line in $lines) {
    $sortedItems = @($line.items | Sort-Object { [double]$_.x })
    $text = ($sortedItems | ForEach-Object { [string]$_.text }) -join ''
    $lineTexts += [ordered]@{ text = $text; y = [double]$line.y }
  }
  $orderedLines = @($lineTexts | Sort-Object { [double]$_.y } -Descending)

  foreach ($pattern in @($patterns)) {
    $regex = $null
    try {
      $regex = New-Object System.Text.RegularExpressions.Regex([string]$pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    } catch { continue }
    foreach ($line in $orderedLines) {
      $match = $regex.Match([string]$line.text)
      if ($match.Success) {
        $value = $match.Groups[1].Value
        if (-not $value) { $value = $match.Value }
        $value = ([string]$value).Trim([char[]]' :：-–—_。，;；')
        if ($value) { return $value }
      }
    }
  }
  return $null
}

function Compare-Natural($a, $b) {
  if ($null -eq $a -and $null -eq $b) { return 0 }
  if ($null -eq $a -or "$a" -eq '') { return 1 }   # 空图号排最后
  if ($null -eq $b -or "$b" -eq '') { return -1 }
  $tokensA = [regex]::Split([string]$a, '(\d+)')
  $tokensB = [regex]::Split([string]$b, '(\d+)')
  $count = [Math]::Min($tokensA.Count, $tokensB.Count)
  for ($i = 0; $i -lt $count; $i++) {
    $x = [string]$tokensA[$i]; $y = [string]$tokensB[$i]
    if ($x -eq $y) { continue }
    $numberX = [long]0; $numberY = [long]0
    $isNumberX = [long]::TryParse($x, [ref]$numberX)
    $isNumberY = [long]::TryParse($y, [ref]$numberY)
    if ($isNumberX -and $isNumberY) {
      if ($numberX -lt $numberY) { return -1 }
      if ($numberX -gt $numberY) { return 1 }
    } else {
      $compare = [string]::Compare($x, $y, $true)
      if ($compare -lt 0) { return -1 }
      if ($compare -gt 0) { return 1 }
    }
  }
  if ($tokensA.Count -lt $tokensB.Count) { return -1 }
  if ($tokensA.Count -gt $tokensB.Count) { return 1 }
  return 0
}

function Sort-Frames-Z($frames) {
  $list = @($frames)
  if ($list.Count -eq 0) { return @() }
  $widths = @($list | ForEach-Object { [double]$_.width } | Sort-Object)
  $medianWidth = [double]$widths[[int][Math]::Floor($widths.Count / 2)]
  $tolerance = [Math]::Max(1.0, $medianWidth * 0.5)
  $columns = New-Object System.Collections.ArrayList
  foreach ($frame in @($list | Sort-Object { [double]$_.minX })) {
    $x = [double]$frame.minX
    $column = $null
    foreach ($candidate in $columns) {
      if ([Math]::Abs([double]$candidate.x - $x) -le $tolerance) { $column = $candidate; break }
    }
    if (-not $column) {
      $column = @{ x = $x; frames = (New-Object System.Collections.ArrayList) }
      [void]$columns.Add($column)
    }
    [void]$column.frames.Add($frame)
  }
  $ordered = New-Object System.Collections.ArrayList
  foreach ($column in @($columns | Sort-Object { [double]$_.x })) {
    foreach ($frame in @($column.frames | Sort-Object { [double]$_.minY } -Descending)) {
      [void]$ordered.Add($frame)
    }
  }
  return @($ordered)
}

function Sort-Frames-ByDrawingNo($frames) {
  $list = New-Object System.Collections.ArrayList
  foreach ($frame in @($frames)) { [void]$list.Add($frame) }
  for ($i = 1; $i -lt $list.Count; $i++) {
    $current = $list[$i]
    $j = $i - 1
    while ($j -ge 0 -and (Compare-Natural $list[$j].drawingNo $current.drawingNo) -gt 0) {
      $list[$j + 1] = $list[$j]
      $j--
    }
    $list[$j + 1] = $current
  }
  return @($list)
}

# ───────────────────────── AutoCAD COM 打印操作 ─────────────────────────

function Set-PlotDevice($layout, [string]$pc3) {
  $name = $pc3
  if (-not $name) { $name = 'DWG To PDF.pc3' }
  try {
    $layout.ConfigName = $name
  } catch {
    throw (New-CadError 'E1004' "无法设置打印设备 $name : $($_.Exception.Message)" '确认 CAD 已安装 DWG To PDF.pc3（config.pc3 可指定）')
  }
}

function Set-A4LandscapeMedia($layout) {
  $candidates = @(
    'ISO_A4_(297.00_x_210.00_MM)',
    'ISO_A4_(210.00_x_297.00_MM)',
    'ISO A4 (297.00 x 210.00 MM)',
    'A4'
  )
  foreach ($candidate in $candidates) {
    try {
      $layout.CanonicalMediaName = $candidate
      return
    } catch { }
  }
  throw (New-CadError 'E1004' '无法设置 A4 纸型' '当前打印设备没有可用的 A4 媒体名，请检查 pc3')
}

function Initialize-ModelLayout($doc, [string]$paper, [string]$pc3) {
  try { $doc.ActiveSpace = 1 } catch { }   # acModelSpace
  $layout = $null
  try { $layout = $doc.ModelSpace.Layout } catch {
    try { $layout = $doc.ActiveLayout } catch { $layout = $null }
  }
  if (-not $layout) { throw (New-CadError 'E1004' '无法获取模型空间 Layout') }
  Set-PlotDevice $layout $pc3
  try { $layout.PaperUnits = 1 } catch { }          # mm
  try { $layout.CenterPlot = $true } catch { }
  try { $layout.PlotWithPlotStyles = $true } catch { }
  try { $layout.UseStandardScale = $true } catch { }
  try { $layout.StandardScale = 0 } catch { }       # acScaleToFit
  if ($paper -ne 'keep') {
    Set-A4LandscapeMedia $layout
    try { $layout.PlotRotation = 1 } catch { }      # 90° → 横向
  }
  return $layout
}

function Set-WindowToPlot($layout, $frame) {
  $lowerLeft = New-Object 'double[]' 3
  $lowerLeft[0] = [double]$frame.minX
  $lowerLeft[1] = [double]$frame.minY
  $lowerLeft[2] = 0
  $upperRight = New-Object 'double[]' 3
  $upperRight[0] = [double]$frame.maxX
  $upperRight[1] = [double]$frame.maxY
  $upperRight[2] = 0
  try {
    $layout.PlotType = 4   # acWindow
  } catch {
    throw (New-CadError 'E1004' "设置窗口打印模式失败: $($_.Exception.Message)")
  }
  try {
    $layout.SetWindowToPlot($lowerLeft, $upperRight)
  } catch {
    try {
      $layout.SetWindowToPlot([ref]$lowerLeft, [ref]$upperRight)
    } catch {
      throw (New-CadError 'E1004' "SetWindowToPlot 失败: $($_.Exception.Message)")
    }
  }
}

function Set-ExtentsPlot($layout) {
  try {
    $layout.PlotType = 1   # acExtents
    $layout.CenterPlot = $true
  } catch {
    throw (New-CadError 'E1004' "设置 extents 打印失败: $($_.Exception.Message)")
  }
}

function Invoke-PlotToFile($doc, [string]$outPath, [string]$pc3) {
  try {
    if ($pc3) { $doc.Plot.PlotToFile($outPath, $pc3) } else { $doc.Plot.PlotToFile($outPath) }
  } catch {
    throw (New-CadError 'E1004' "PlotToFile 失败: $($_.Exception.Message)" "输出: $outPath")
  }
  if (-not (Test-Path -LiteralPath $outPath)) {
    throw (New-CadError 'E1004' "PlotToFile 未产出文件: $outPath" '检查打印设备/权限/磁盘空间')
  }
}

function Open-Drawing($app, [string]$path) {
  $doc = $null
  try {
    $doc = $app.Documents.Open($path, $true)
  } catch {
    try { $doc = $app.Documents.Open($path) } catch {
      throw (New-CadError 'E1003' "打开图纸失败: $path — $($_.Exception.Message)" '文件可能损坏、版本过新或被占用')
    }
  }
  if (-not $doc) { throw (New-CadError 'E1003' "打开图纸失败: $path") }
  return $doc
}

# ───────────────────────── 主流程 ─────────────────────────

$exitCode = 0
$payload = $null
$app = $null
$doc = $null
$quitApp = $false
$platform = $null

try {
  $argObj = $null
  try { $argObj = $Json | ConvertFrom-Json } catch {
    throw (New-CadError 'E1041' "参数 JSON 解析失败: $($_.Exception.Message)")
  }
  if (-not $argObj) { throw (New-CadError 'E1041' '参数 JSON 为空') }

  $mode = [string]$argObj.mode
  if ($mode -ne 'split' -and $mode -ne 'merge') {
    throw (New-CadError 'E1041' "mode 必须为 'split' | 'merge'，收到: $mode")
  }
  $inputPath = [string]$argObj.input
  if (-not $inputPath -or -not (Test-Path -LiteralPath $inputPath)) {
    throw (New-CadError 'E1023' "输入文件不存在: $inputPath")
  }
  $paper = [string]$argObj.paper
  if (-not $paper) { $paper = 'A4-landscape' }
  $pc3 = [string]$argObj.pc3
  if (-not $pc3) { $pc3 = 'DWG To PDF.pc3' }
  $tolerance = 2.0
  if ($null -ne $argObj.frameToleranceMm) { $tolerance = [double]$argObj.frameToleranceMm }
  $frameSizes = Get-FrameSizeList $argObj.frameSizes
  $frameBlockNames = @($argObj.frameBlockNames)
  if ($frameBlockNames.Count -eq 0) { $frameBlockNames = @('*图框*', 'A0', 'A1', 'A2', 'A3', 'A4', '*TK*') }
  $zone = @($argObj.titleBlockZone)
  if ($zone.Count -lt 4) { $zone = @(0.6, 0, 1.0, 0.25) }
  $patterns = @($argObj.drawingNoPatterns)
  if ($patterns.Count -eq 0) {
    $patterns = @(
      '[/／](?:图号|编号|图别)[:：]?\s*([A-Za-z0-9\u4e00-\u9fa5_./-]+)',
      '\b([A-Z]{0,4}-?\d{2,5}(?:-\d{1,3})?(?:[A-Z]?\d{0,3})?)\b'
    )
  }

  # 1) CAD COM 探测（三平台 late binding）
  $platforms = @(
    [ordered]@{ id = 'autocad'; progId = 'AutoCAD.Application'; exeNames = @('acad') },
    [ordered]@{ id = 'zwcad'; progId = 'ZWCAD.Application'; exeNames = @('zwcad') },
    [ordered]@{ id = 'gstarcad'; progId = 'GstarCAD.Application'; exeNames = @('gcad', 'GstarCAD') }
  )
  $hadProcess = @{}
  foreach ($candidatePlatform in $platforms) {
    $hadProcess[$candidatePlatform.id] = Test-ProcessName $candidatePlatform.exeNames
  }
  foreach ($candidatePlatform in $platforms) {
    try {
      $candidateApp = New-Object -ComObject $candidatePlatform.progId -ErrorAction Stop
      if ($candidateApp) {
        $app = $candidateApp
        $platform = $candidatePlatform
        break
      }
    } catch { }
  }
  if (-not $app) {
    throw (New-CadError 'E1001' '未检测到 AutoCAD / 中望CAD / 浩辰CAD COM' '请安装 CAD，或使用 ODA/PDF 级引擎')
  }
  $quitApp = -not [bool]$hadProcess[$platform.id]

  # 2) 打开图纸
  $doc = Open-Drawing $app $inputPath
  $platformName = [string]$platform.id

  if ($mode -eq 'split') {
    $outputDir = [string]$argObj.outputDir
    if (-not $outputDir) { throw (New-CadError 'E1041' 'split 模式需要 outputDir') }
    if (-not (Test-Path -LiteralPath $outputDir)) { New-Item -ItemType Directory -Force -Path $outputDir | Out-Null }
    $layout = Initialize-ModelLayout $doc $paper $pc3
    $textItems = Get-TextItems $doc
    $candidates = Get-FrameCandidates $doc $frameBlockNames $frameSizes $tolerance
    $frames = @(Merge-FrameCandidates $candidates)
    if ($frames.Count -eq 0) {
      throw (New-CadError 'E1012' '未识别到任何图框' '检查 frameBlockNames / frameSizes 配置，或图框是否以块参照/闭合多段线绘制')
    }
    $frames = @(Sort-Frames-Z $frames)
    foreach ($frame in $frames) {
      $frame['drawingNo'] = Get-DrawingNo $textItems $frame $zone $patterns
    }
    $frames = @(Sort-Frames-ByDrawingNo $frames)

    $frameResults = @()
    $sequence = 0
    foreach ($frame in $frames) {
      $sequence++
      $drawingNo = [string]$frame.drawingNo
      $safeName = 'frame'
      if ($drawingNo) {
        $safeName = ($drawingNo -replace '[\\/:*?"<>|]', '_')
      } else {
        $safeName = "frame$sequence"
      }
      $outPath = Join-Path $outputDir ("frame-{0:d3}-{1}.pdf" -f $sequence, $safeName)
      Set-WindowToPlot $layout $frame
      Invoke-PlotToFile $doc $outPath $pc3
      $frameResults += [ordered]@{
        seq       = $sequence
        path      = $outPath
        drawingNo = $(if ($drawingNo) { $drawingNo } else { $null })
        frameSize = [string]$frame.frameSize
        source    = [string]$frame.source
        widthMm   = [double]$frame.widthMm
        heightMm  = [double]$frame.heightMm
        window    = [ordered]@{ minX = $frame.minX; minY = $frame.minY; maxX = $frame.maxX; maxY = $frame.maxY }
      }
    }
    $payload = [ordered]@{
      ok     = $true
      result = [ordered]@{
        mode       = 'split'
        platform   = $platformName
        input      = $inputPath
        frames     = @($frameResults)
        candidates = $candidates.Count
        kept       = $frames.Count
        paper      = $paper
        pc3        = $pc3
      }
    }
  }
  else {
    # 场景 B：单文件 extents 打印
    $outputPath = [string]$argObj.output
    if (-not $outputPath) { throw (New-CadError 'E1041' 'merge 模式需要 output（单文件 PDF 路径）') }
    $outputParent = Split-Path -Parent $outputPath
    if ($outputParent -and -not (Test-Path -LiteralPath $outputParent)) {
      New-Item -ItemType Directory -Force -Path $outputParent | Out-Null
    }
    $layout = Initialize-ModelLayout $doc $paper $pc3
    Set-ExtentsPlot $layout
    Invoke-PlotToFile $doc $outputPath $pc3
    $payload = [ordered]@{
      ok     = $true
      result = [ordered]@{
        mode     = 'merge'
        platform = $platformName
        input    = $inputPath
        output   = $outputPath
        paper    = $paper
        pc3      = $pc3
      }
    }
  }
}
catch {
  $code = 'E1099'
  $hint = $null
  $exception = $_.Exception
  if ($exception.Data -and $exception.Data.Contains('code')) {
    $code = [string]$exception.Data['code']
    if ($exception.Data.Contains('hint')) { $hint = [string]$exception.Data['hint'] }
  }
  $errorObj = [ordered]@{ code = $code; message = [string]$exception.Message }
  if ($hint) { $errorObj['hint'] = $hint }
  $payload = [ordered]@{ ok = $false; error = $errorObj }
  $exitCode = 2
}
finally {
  if ($doc) { try { $null = $doc.Close($false) } catch { } }
  if ($app -and $quitApp) { try { $null = $app.Quit() } catch { } }
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$payload | ConvertTo-Json -Depth 10
exit $exitCode
