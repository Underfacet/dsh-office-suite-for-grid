# visio-layouts.ps1 —— P3 五类电网图布局引擎（docs/P3-设计.md §4.4）
#
# 输入（-Json 单个 JSON 参数，与 visio-grid.ps1 同款）：
#   {
#     "kind": "protection" | "substationPlan" | "geoNetwork" | "powerFlow" | "terminalStrip",
#     "title": "110kV 线路保护电流回路图（示例）",
#     "model": { ...P3 §4.4 数据模型... },
#     "symbols": { "gb4728.6.TA": { ...符号 JSON... }, ... },   # Host 注入（已校验）
#     "outputPath": "D:\...\xxx.vsdx",
#     "validations": [ { level, code, message } ]                 # Host 已算好的接线提示（原样透传）
#   }
# 输出（stdout 单个 JSON）：
#   { "ok": true, "result": { "outputPath", "format", "kind", "paper", "pages",
#       "stats": { symbols, shapes, links, groups, pages, circuits, terminals, strips,
#                  stations, lines, flows, zones, items },
#       "warnings": [ ...引擎级提示... ], "validations": [...] } }
#   失败 { "ok": false, "error": { "code", "message", "hint"? } }，code=E2021/E2023
#
# 设计要点（与 visio-grid.ps1 同款 COM 风格；P2 三类仍由 visio-grid.ps1 绘制）：
# - 单位：符号 JSON 为 mm（原点左上、x 右 y 下），Visio 内部为英寸，渲染时统一换算；
# - COM 内存组装（不依赖本机模具/模板），每符号实例绘制后 Group（Window.Select + Selection.Group）；
# - 不设置 Connection Points：端子坐标由引擎内存表维护，连线一律 DrawLine；
# - 符号内文字不随实例旋转（图元线/圆按实例旋转，文字保持水平）；
# - 五类布局：protection=回路横向串接 + 端子排带 + 触点映射虚线；
#   substationPlan=场区/网格 + 设备区矩形 + 绝对坐标设备/图元 + 道路；
#   geoNetwork=经纬度等距圆柱投影 + 折线 + 双回线平行偏移；
#   powerFlow=电压等级分层（电源→母线→线路→负荷）+ 潮流箭头（EndArrow）；
#   terminalStrip=多行栅格端子表 + 每页 terminalsPerPage 行（多页输出，失败降级单页）；
# - 防御式 Close/Quit；[void] 抑制 SaveAs/Close/Quit；每个 COM 调用上下文写入 E2023。
#
# 兼容：Windows PowerShell 5.1（文件 UTF-8 带 BOM；不使用 pwsh 7 独有语法）。

[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ─────────────────────────── 常量（mm / 英寸） ───────────────────────────
$script:MM_PER_IN = 25.4
$script:PT_PER_MM = 72.0 / 25.4
$script:PAPERS = @{
  'A4-landscape' = @(11.69, 8.27)
  'A4-portrait'  = @(8.27, 11.69)
  'A3-landscape' = @(16.54, 11.69)
  'A3-portrait'  = @(11.69, 16.54)
}
$script:KIND_PAPERS = @{
  protection      = 'A3-landscape'
  substationPlan  = 'A3-landscape'
  geoNetwork      = 'A3-landscape'
  powerFlow       = 'A3-landscape'
  terminalStrip   = 'A4-landscape'
}
$script:KINDS = @('protection', 'substationPlan', 'geoNetwork', 'powerFlow', 'terminalStrip')
$script:MARGIN_MM = 15.0
$script:MAX_SYMBOLS = 2000
$script:GRID_GRAY = 'RGB(205,205,205)'

# ─────────────────────────── 全局状态 ───────────────────────────
$script:Model = $null
$script:Symbols = @{}
$script:Title = ''
$script:Kind = ''
$script:PaperName = 'A3-landscape'
$script:PaperW = 16.54
$script:PaperH = 11.69
$script:Scale = 1.0
$script:OffsetX = 0.0
$script:OffsetY = 0.0
$script:BBoxMinX = 0.0
$script:BBoxMinY = 0.0
$script:Step = '初始化'
$script:Stats = @{
  symbols = 0; shapes = 0; links = 0; groups = 0; pages = 0
  circuits = 0; terminals = 0; strips = 0
  stations = 0; lines = 0; flows = 0; zones = 0; items = 0; buses = 0
}
$script:EngineWarnings = New-Object System.Collections.Generic.List[string]
$script:GeoBounds = $null

# ─────────────────────────── 基础工具 ───────────────────────────
function Fail-Engine {
  param([string]$Code, [string]$Message, [string]$Hint)
  $err = @{ code = $Code; message = $Message }
  if ($Hint) { $err['hint'] = $Hint }
  @{ ok = $false; error = $err } | ConvertTo-Json -Depth 6
  exit 1
}

function Get-Prop {
  param($Object, [string]$Name, $Default = $null)
  if ($null -eq $Object) { return $Default }
  if ($Object -is [System.Collections.IDictionary]) {
    if ($Object.Contains($Name)) { return $Object[$Name] }
    return $Default
  }
  $prop = $Object.PSObject.Properties[$Name]
  if ($null -eq $prop) { return $Default }
  return $prop.Value
}

function Get-Items {
  param($Object, [string]$Name)
  $value = Get-Prop $Object $Name
  if ($null -eq $value) { return @() }
  if ($value -is [array]) { return $value }
  # PS 5.1 已知问题：@(List[T]) 会抛 "Argument types do not match"，改用 ToArray()
  if ($value -is [System.Collections.IList]) {
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($item in $value) { $items.Add($item) }
    return $items.ToArray()
  }
  return @($value)
}

function Get-Str {
  param($Object, [string]$Name, [string]$Default = '')
  $value = Get-Prop $Object $Name
  if ($null -eq $value) { return $Default }
  return [string]$value
}

function Get-Num {
  param($Object, [string]$Name, [double]$Default = 0.0)
  $value = Get-Prop $Object $Name
  if ($null -eq $value) { return $Default }
  try { return [double]$value } catch { return $Default }
}

# ─────────────────────────── 页面数据模型 ───────────────────────────
function New-PageData {
  param([string]$Name = '')
  @{
    Name       = $Name
    Items      = New-Object System.Collections.Generic.List[object]
    Placements = New-Object System.Collections.Generic.List[object]
    Groups     = New-Object System.Collections.Generic.List[object]
    Boxes      = @{}
    Terminals  = @{}
    BusAnchor  = @{}
    Bounds     = @{ minX = [double]::MaxValue; minY = [double]::MaxValue; maxX = [double]::MinValue; maxY = [double]::MinValue }
    Layout     = @{}
  }
}

function Update-PageBounds {
  param($Page, [double]$X, [double]$Y)
  $b = $Page.Bounds
  if ($X -lt $b.minX) { $b.minX = $X }
  if ($Y -lt $b.minY) { $b.minY = $Y }
  if ($X -gt $b.maxX) { $b.maxX = $X }
  if ($Y -gt $b.maxY) { $b.maxY = $Y }
}

function Add-LineItem {
  param($Page, $Points, [double]$Weight = 0.35, [bool]$Closed = $false, [bool]$Arrow = $false,
    [bool]$Dashed = $false, [string]$Color = '')
  # PS 5.1：@(List[object]) 在元素为数组时会抛 ArgumentException，逐项收集规避
  $pts = New-Object System.Collections.Generic.List[object]
  if ($null -ne $Points) { foreach ($point in $Points) { $pts.Add($point) } }
  if ($pts.Count -lt 2) { return }
  foreach ($pt in $pts) { Update-PageBounds $Page ([double]$pt[0]) ([double]$pt[1]) }
  $Page.Items.Add(@{ kind = 'line'; pts = $pts; weight = $Weight; closed = $Closed; arrow = $Arrow; dashed = $Dashed; color = $Color })
}

function Add-RectItem {
  param($Page, [double]$X, [double]$Y, [double]$W, [double]$H, [double]$Weight = 0.35,
    [bool]$Dashed = $false, [string]$Color = '')
  if ($W -le 0 -or $H -le 0) { return }
  $pts = New-Object System.Collections.Generic.List[object]
  $pts.Add(@($X, $Y))
  $pts.Add(@(($X + $W), $Y))
  $pts.Add(@(($X + $W), ($Y + $H)))
  $pts.Add(@($X, ($Y + $H)))
  Add-LineItem -Page $Page -Points $pts -Weight $Weight -Closed $true -Dashed $Dashed -Color $Color
}

function Add-EllipseItem {
  param($Page, [double]$Cx, [double]$Cy, [double]$Radius, [double]$Weight = 0.35)
  if ($Radius -le 0) { return }
  Update-PageBounds $Page ($Cx - $Radius) ($Cy - $Radius)
  Update-PageBounds $Page ($Cx + $Radius) ($Cy + $Radius)
  $Page.Items.Add(@{ kind = 'ellipse'; cx = $Cx; cy = $Cy; r = $Radius; weight = $Weight })
}

function Get-TextBox {
  param([double]$X, [double]$Y, [string]$Text, [double]$Size, [string]$Anchor)
  $length = $Text.Length
  if ($length -lt 1) { $length = 1 }
  $cjk = 0
  foreach ($ch in $Text.ToCharArray()) { if ([int]$ch -ge 128) { $cjk += 1 } }
  $ascii = $length - $cjk
  $w = ($ascii * $Size * 0.62) + ($cjk * $Size * 1.02) + ($Size * 0.8)
  if ($w -lt ($Size * 9.0)) { $w = $Size * 9.0 }
  $h = $Size * 2.0
  switch ($Anchor) {
    'top-center'    { return @(($X - $w / 2.0), $Y, $w, $h) }
    'center'        { return @(($X - $w / 2.0), ($Y - $h / 2.0), $w, $h) }
    'bottom-center' { return @(($X - $w / 2.0), ($Y - $h), $w, $h) }
    'bottom-left'   { return @($X, ($Y - $h), $w, $h) }
    default         { return @($X, $Y, $w, $h) }
  }
}

function Add-TextItem {
  param($Page, [double]$X, [double]$Y, [string]$Text, [double]$Size = 2.5, [string]$Anchor = 'top-left')
  if (-not $Text) { return }
  $box = Get-TextBox -X $X -Y $Y -Text $Text -Size $Size -Anchor $Anchor
  Update-PageBounds $Page $box[0] $box[1]
  Update-PageBounds $Page ($box[0] + $box[2]) ($box[1] + $box[3])
  $Page.Items.Add(@{ kind = 'text'; x = $X; y = $Y; text = $Text; size = $Size; anchor = $Anchor })
}

# ─────────────────────────── 几何换算 ───────────────────────────
function Convert-Point {
  param([double]$X, [double]$Y, [double]$W, [double]$H, [int]$Rot)
  switch ($Rot) {
    90  { return @(($H - $Y), $X) }
    180 { return @(($W - $X), ($H - $Y)) }
    270 { return @($Y, ($W - $X)) }
    default { return @($X, $Y) }
  }
}

function Get-InstanceSize {
  param($Symbol, [int]$Rot, [double]$Scale)
  $view = Get-Prop $Symbol 'view'
  $w = (Get-Num $view 'widthMm' 1.0) * $Scale
  $h = (Get-Num $view 'heightMm' 1.0) * $Scale
  if ($Rot -eq 90 -or $Rot -eq 270) { return @($h, $w) }
  return @($w, $h)
}

function Get-DeviceRot {
  param($Symbol)
  $terms = @(Get-Items $Symbol 'terminals')
  $t1 = $null
  $t2 = $null
  foreach ($terminal in $terms) {
    $tid = Get-Str $terminal 'id' ''
    if ($tid -eq 'T1') { $t1 = $terminal }
    if ($tid -eq 'T2') { $t2 = $terminal }
  }
  if ($null -eq $t1 -or $null -eq $t2) { return 0 }
  $a1 = @(Get-Items $t1 'at')
  $a2 = @(Get-Items $t2 'at')
  if ($a1.Count -lt 2 -or $a2.Count -lt 2) { return 0 }
  $dx = [Math]::Abs([double]$a2[0] - [double]$a1[0])
  $dy = [Math]::Abs([double]$a2[1] - [double]$a1[1])
  if ($dx -ge $dy) { return 90 }
  return 0
}

# ─────────────────────────── 符号实例 → 图元 ───────────────────────────
function Add-SymbolInstance {
  param(
    $Page,
    [string]$Id,
    [string]$SymbolRef,
    [double]$X,
    [double]$Y,
    [int]$Rot = 0,
    [double]$Scale = 1.0,
    [string]$Label = '',
    [int]$Depth = 0
  )
  $symbol = $script:Symbols[$SymbolRef]
  if ($null -eq $symbol) { throw "符号未注入：$SymbolRef（实例 $Id）" }
  if ($Depth -gt 2) { return }
  if ($Depth -eq 0) { $script:Stats.symbols += 1 }
  $rot = (($Rot % 360) + 360) % 360
  $view = Get-Prop $symbol 'view'
  $vw = (Get-Num $view 'widthMm' 1.0) * $Scale
  $vh = (Get-Num $view 'heightMm' 1.0) * $Scale
  $size = Get-InstanceSize $symbol $rot $Scale
  $ew = $size[0]
  $eh = $size[1]

  $Page.Boxes[$Id] = @{
    x = $X; y = $Y; w = $ew; h = $eh; rot = $rot
    center = @(($X + $ew / 2.0), ($Y + $eh / 2.0))
    symbolRef = $SymbolRef
  }
  Update-PageBounds $Page $X $Y
  Update-PageBounds $Page ($X + $ew) ($Y + $eh)

  # 端子绝对坐标（mm）
  $terminalMap = @{}
  foreach ($terminal in @(Get-Items $symbol 'terminals')) {
    $tid = Get-Str $terminal 'id' ''
    $at = @(Get-Items $terminal 'at')
    if (-not $tid -or $at.Count -lt 2) { continue }
    $local = Convert-Point ([double]$at[0] * $Scale) ([double]$at[1] * $Scale) $vw $vh $rot
    $terminalMap[$tid] = @(($X + [double]$local[0]), ($Y + [double]$local[1]))
  }
  $Page.Terminals[$Id] = $terminalMap

  # 图元（按数组顺序绘制，后覆盖先）
  foreach ($primitive in @(Get-Items $symbol 'primitives')) {
    $type = Get-Str $primitive 'type' ''
    $stroke = Get-Num $primitive 'stroke' 0.35
    if ($type -eq 'polyline' -or $type -eq 'line') {
      $points = New-Object System.Collections.Generic.List[object]
      foreach ($rawPoint in @(Get-Items $primitive 'points')) {
        $point = @($rawPoint)
        if ($point.Count -lt 2) { continue }
        $local = Convert-Point ([double]$point[0] * $Scale) ([double]$point[1] * $Scale) $vw $vh $rot
        $points.Add(@(($X + [double]$local[0]), ($Y + [double]$local[1])))
      }
      if ($points.Count -ge 2) {
        Add-LineItem -Page $Page -Points $points -Weight $stroke -Closed ([bool](Get-Prop $primitive 'closed' $false))
      }
    }
    elseif ($type -eq 'rect') {
      $origin = @(Get-Items $primitive 'origin')
      $rw = Get-Num $primitive 'width' 0
      $rh = Get-Num $primitive 'height' 0
      if ($origin.Count -ge 2 -and $rw -gt 0 -and $rh -gt 0) {
        $corners = @(
          @([double]$origin[0], [double]$origin[1]),
          @(([double]$origin[0] + $rw), [double]$origin[1]),
          @(([double]$origin[0] + $rw), ([double]$origin[1] + $rh)),
          @([double]$origin[0], ([double]$origin[1] + $rh))
        )
        $points = New-Object System.Collections.Generic.List[object]
        foreach ($corner in $corners) {
          $local = Convert-Point ([double]$corner[0] * $Scale) ([double]$corner[1] * $Scale) $vw $vh $rot
          $points.Add(@(($X + [double]$local[0]), ($Y + [double]$local[1])))
        }
        Add-LineItem -Page $Page -Points $points -Weight $stroke -Closed $true
      }
    }
    elseif ($type -eq 'circle') {
      $center = @(Get-Items $primitive 'center')
      $radius = Get-Num $primitive 'radius' 0
      if ($center.Count -ge 2 -and $radius -gt 0) {
        $local = Convert-Point ([double]$center[0] * $Scale) ([double]$center[1] * $Scale) $vw $vh $rot
        Add-EllipseItem -Page $Page -Cx ($X + [double]$local[0]) -Cy ($Y + [double]$local[1]) -Radius ($radius * $Scale) -Weight $stroke
      }
    }
    elseif ($type -eq 'arc') {
      $center = @(Get-Items $primitive 'center')
      $radius = Get-Num $primitive 'radius' 0
      $fromDeg = Get-Num $primitive 'fromDeg' 0
      $toDeg = Get-Num $primitive 'toDeg' 360
      if ($center.Count -ge 2 -and $radius -gt 0) {
        if ($toDeg -le $fromDeg) { $toDeg = $fromDeg + 360 }
        $points = New-Object System.Collections.Generic.List[object]
        $angle = $fromDeg
        while ($angle -le $toDeg) {
          $rad = $angle * [Math]::PI / 180.0
          $local = Convert-Point (([double]$center[0] + $radius * [Math]::Cos($rad)) * $Scale) (([double]$center[1] + $radius * [Math]::Sin($rad)) * $Scale) $vw $vh $rot
          $points.Add(@(($X + [double]$local[0]), ($Y + [double]$local[1])))
          $angle += 15.0
        }
        if ($points.Count -ge 2) { Add-LineItem -Page $Page -Points $points -Weight $stroke }
      }
    }
    elseif ($type -eq 'text') {
      $at = @(Get-Items $primitive 'at')
      $text = Get-Str $primitive 'text' ''
      if ($at.Count -lt 2 -or -not $text) { continue }
      # 文字只随实例平移（不随实例旋转），避免 Visio TxtAngle 跨版本差异
      $local = Convert-Point ([double]$at[0] * $Scale) ([double]$at[1] * $Scale) $vw $vh 0
      Add-TextItem -Page $Page -X ($X + [double]$local[0]) -Y ($Y + [double]$local[1]) -Text $text -Size ((Get-Num $primitive 'sizeMm' 2.5) * $Scale) -Anchor (Get-Str $primitive 'anchor' 'top-left')
    }
  }

  # 组合符号（compose，递归深度 ≤2）
  foreach ($entry in @(Get-Items $symbol 'compose')) {
    $childRef = Get-Str $entry 'symbolRef' ''
    if (-not $childRef) { continue }
    if ($null -eq $script:Symbols[$childRef]) { continue }
    $at = @(Get-Items $entry 'at')
    if ($at.Count -lt 2) { continue }
    $childScale = (Get-Num $entry 'scale' 1.0) * $Scale
    $local = Convert-Point ([double]$at[0] * $Scale) ([double]$at[1] * $Scale) $vw $vh $rot
    Add-SymbolInstance -Page $Page -Id ("$Id/$childRef") -SymbolRef $childRef -X ($X + [double]$local[0]) -Y ($Y + [double]$local[1]) -Rot $rot -Scale $childScale -Depth ($Depth + 1)
  }

  if ($Label) {
    Add-TextItem -Page $Page -X ($X + $ew / 2.0) -Y ($Y + $eh + 1.5) -Text $Label -Size 2.6 -Anchor 'top-center'
  }
}

function Add-Placement {
  param($Page, [string]$Id, [string]$SymbolRef, [double]$X, [double]$Y, [int]$Rot = 0, [double]$Scale = 1.0, [string]$Label = '')
  if (-not $Id -or -not $SymbolRef) { return }
  $Page.Placements.Add(@{ id = $Id; symbolRef = $SymbolRef; x = $X; y = $Y; rot = $Rot; scale = $Scale; label = $Label })
}

function Expand-Placements {
  param($Page)
  foreach ($placement in $Page.Placements) {
    $start = $Page.Items.Count
    Add-SymbolInstance -Page $Page -Id $placement.id -SymbolRef $placement.symbolRef -X $placement.x -Y $placement.y -Rot $placement.rot -Scale $placement.scale -Label $placement.label
    $count = $Page.Items.Count - $start
    if ($count -gt 0) {
      $Page.Groups.Add(@{ start = $start; count = $count; id = $placement.id })
    }
  }
}

function Add-BusLine {
  param($Page, [string]$Id, [double]$X1, [double]$X2, [double]$Y, [string]$Label)
  if ($X2 -le $X1) { $X2 = $X1 + 200.0 }
  $points = New-Object System.Collections.Generic.List[object]
  $points.Add(@($X1, $Y))
  $points.Add(@($X2, $Y))
  Add-LineItem -Page $Page -Points $points -Weight 0.8
  $terminalMap = @{}
  $busSymbol = $script:Symbols['gb4728.3.BUS']
  $fractions = @(0.0, (1.0 / 3.0), (2.0 / 3.0), 1.0)
  if ($null -ne $busSymbol) {
    $view = Get-Prop $busSymbol 'view'
    $widthMm = Get-Num $view 'widthMm' 120.0
    if ($widthMm -gt 0) {
      $terms = @(Get-Items $busSymbol 'terminals')
      $fractions = @()
      foreach ($terminal in $terms) {
        $at = @(Get-Items $terminal 'at')
        if ($at.Count -ge 2) { $fractions += ([double]$at[0] / $widthMm) }
      }
      if ($fractions.Count -eq 0) { $fractions = @(0.0, (1.0 / 3.0), (2.0 / 3.0), 1.0) }
    }
  }
  for ($index = 0; $index -lt $fractions.Count; $index++) {
    $fraction = [double]$fractions[$index]
    $terminalMap[('T' + ($index + 1))] = @(($X1 + ($X2 - $X1) * $fraction), $Y)
  }
  $Page.Terminals[$Id] = $terminalMap
  $Page.BusAnchor[$Id] = @{ x1 = $X1; x2 = $X2; y = $Y }
  if ($Label) { Add-TextItem -Page $Page -X $X1 -Y ($Y - 2.5) -Text $Label -Size 3.0 -Anchor 'bottom-left' }
}

function Get-PagePoint {
  param($Page, [string]$Ref, $OtherPoint)
  $text = [string]$Ref
  if (-not $text) { return $null }
  $base = $text
  $terminal = ''
  $dot = $text.LastIndexOf('.')
  if ($dot -gt 0) {
    $maybeBase = $text.Substring(0, $dot)
    if ($Page.Terminals.ContainsKey($maybeBase) -or $Page.Boxes.ContainsKey($maybeBase) -or $Page.BusAnchor.ContainsKey($maybeBase)) {
      $base = $maybeBase
      $terminal = $text.Substring($dot + 1)
    }
  }
  if ($terminal -and $Page.Terminals.ContainsKey($base)) {
    $terms = $Page.Terminals[$base]
    if ($terms.ContainsKey($terminal)) { return $terms[$terminal] }
  }
  if ($Page.BusAnchor.ContainsKey($base) -and -not $terminal) {
    $anchor = $Page.BusAnchor[$base]
    $x = ($anchor.x1 + $anchor.x2) / 2.0
    if ($null -ne $OtherPoint) {
      $x = [Math]::Min([Math]::Max([double]$OtherPoint[0], $anchor.x1), $anchor.x2)
    }
    return @($x, [double]$anchor.y)
  }
  if ($Page.Terminals.ContainsKey($base)) {
    $terms = $Page.Terminals[$base]
    if ($terms.Count -gt 0) {
      if ($null -ne $OtherPoint -and $Page.Boxes.ContainsKey($base) -and $terms.ContainsKey('T1') -and $terms.ContainsKey('T2')) {
        $p1 = $terms['T1']
        $p2 = $terms['T2']
        $d1 = ([double]$p1[0] - [double]$OtherPoint[0]) * ([double]$p1[0] - [double]$OtherPoint[0]) + ([double]$p1[1] - [double]$OtherPoint[1]) * ([double]$p1[1] - [double]$OtherPoint[1])
        $d2 = ([double]$p2[0] - [double]$OtherPoint[0]) * ([double]$p2[0] - [double]$OtherPoint[0]) + ([double]$p2[1] - [double]$OtherPoint[1]) * ([double]$p2[1] - [double]$OtherPoint[1])
        if ($d2 -lt $d1) { return $p2 }
        return $p1
      }
      if ($terms.ContainsKey('T1')) { return $terms['T1'] }
      foreach ($key in $terms.Keys) { return $terms[$key] }
    }
  }
  if ($Page.BusAnchor.ContainsKey($base)) {
    $anchor = $Page.BusAnchor[$base]
    $x = ($anchor.x1 + $anchor.x2) / 2.0
    if ($null -ne $OtherPoint) {
      $x = [Math]::Min([Math]::Max([double]$OtherPoint[0], $anchor.x1), $anchor.x2)
    }
    return @($x, [double]$anchor.y)
  }
  if ($Page.Boxes.ContainsKey($base)) { return $Page.Boxes[$base].center }
  return $null
}

function Get-CenterPoint {
  param($Page, [string]$Id)
  if ($Page.Boxes.ContainsKey($Id)) { return $Page.Boxes[$Id].center }
  if ($Page.BusAnchor.ContainsKey($Id)) {
    $anchor = $Page.BusAnchor[$Id]
    return @((($anchor.x1 + $anchor.x2) / 2.0), [double]$anchor.y)
  }
  return $null
}

# ─────────────────────────── 布局①：二次回路/继保原理图（§4.4①） ───────────────────────────
function Add-ProtectionStrip {
  param($Page, $Strip, [double]$X, [double]$Y)
  $label = Get-Str $Strip 'label' (Get-Str $Strip 'id' '端子排')
  Add-TextItem -Page $Page -X $X -Y $Y -Text $label -Size 2.6 -Anchor 'top-left'
  $terminals = @(Get-Items $Strip 'terminals')
  $baselineY = $Y + 12.0
  $startX = $X + 26.0
  $cursorX = $startX
  foreach ($terminal in $terminals) {
    $no = Get-Str $terminal 'no' ''
    Add-EllipseItem -Page $Page -Cx $cursorX -Cy $baselineY -Radius 4.0 -Weight 0.35
    Add-TextItem -Page $Page -X $cursorX -Y $baselineY -Text $no -Size 2.4 -Anchor 'center'
    $ref = Get-Str $terminal 'ref' ''
    if ($ref) { Add-TextItem -Page $Page -X $cursorX -Y ($baselineY + 5.0) -Text $ref -Size 2.2 -Anchor 'top-center' }
    $cursorX += 12.0
  }
  if ($terminals.Count -gt 0) {
    $band = New-Object System.Collections.Generic.List[object]
    $band.Add(@(($startX - 6.0), $baselineY))
    $band.Add(@(($cursorX - 6.0), $baselineY))
    Add-LineItem -Page $Page -Points $band -Weight 0.5
  }
}

function Invoke-ProtectionLayout {
  $page = New-PageData -Name '二次回路'
  $model = $script:Model
  $circuits = @(Get-Items $model 'circuits')
  $strips = @(Get-Items $model 'terminalStrips')
  $loops = @(Get-Items $model 'loops')
  $margin = $script:MARGIN_MM
  $rowGap = 40.0
  $deviceGap = 12.0

  $topStrips = @($strips | Where-Object { (Get-Str $_ 'side' 'top') -ne 'bottom' })
  $bottomStrips = @($strips | Where-Object { (Get-Str $_ 'side' 'top') -eq 'bottom' })
  $stripHeight = 22.0

  $stripY = $margin
  foreach ($strip in $topStrips) {
    Add-ProtectionStrip -Page $page -Strip $strip -X $margin -Y $stripY
    $stripY += $stripHeight
  }
  $circuitY = $margin
  if ($topStrips.Count -gt 0) { $circuitY = $stripY + 6.0 }

  $rows = New-Object System.Collections.Generic.List[object]
  $rowIndex = 0
  foreach ($circuit in $circuits) {
    $circuitId = Get-Str $circuit 'id' ("C$rowIndex")
    $devices = @(Get-Items $circuit 'devices')
    $sizes = New-Object System.Collections.Generic.List[object]
    $maxH = 0.0
    $totalW = 0.0
    foreach ($device in $devices) {
      $ref = Get-Str $device 'symbolRef' ''
      $symbol = $script:Symbols[$ref]
      if ($null -eq $symbol) { throw "符号未注入：$ref（设备 $((Get-Str $device 'id' ''))）" }
      # 回路横向串接：强制端子水平（Get-DeviceRot 的规则面向单线图竖列，需要取反）
      $verticalRot = Get-DeviceRot $symbol
      $rot = 0
      if ($verticalRot -eq 0) { $rot = 90 }
      $size = Get-InstanceSize $symbol $rot 1.0
      $sizes.Add(@{ w = $size[0]; h = $size[1]; rot = $rot; ref = $ref; device = $device })
      if ($size[1] -gt $maxH) { $maxH = $size[1] }
      $totalW += $size[0]
    }
    if ($devices.Count -gt 1) { $totalW += $deviceGap * ($devices.Count - 1) }
    $label = Get-Str $circuit 'label' $circuitId
    $voltage = Get-Str $circuit 'voltage' ''
    if ($voltage) { $label = "$label（$voltage）" }
    Add-TextItem -Page $page -X $margin -Y ($circuitY + 8.0) -Text $label -Size 2.8 -Anchor 'top-left'

    $cursorX = $margin + 30.0
    $deviceIds = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $devices.Count; $index++) {
      $entry = $sizes[$index]
      $deviceId = Get-Str $entry.device 'id' ''
      $deviceLabel = Get-Str $entry.device 'label' ''
      $offsetY = $circuitY + (($maxH - $entry.h) / 2.0)
      Add-Placement -Page $page -Id $deviceId -SymbolRef $entry.ref -X $cursorX -Y $offsetY -Rot $entry.rot -Scale 1.0 -Label $deviceLabel
      if ($deviceId) { $deviceIds.Add($deviceId) }
      $cursorX += $entry.w + $deviceGap
    }
    $rows.Add(@{ circuitId = $circuitId; deviceIds = $deviceIds; y = $circuitY; height = $maxH; width = $totalW })
    $circuitY += [Math]::Max($rowGap, $maxH + 22.0)
    $rowIndex += 1
  }

  if ($bottomStrips.Count -gt 0) {
    $stripY = $circuitY + 6.0
    foreach ($strip in $bottomStrips) {
      Add-ProtectionStrip -Page $page -Strip $strip -X $margin -Y $stripY
      $stripY += $stripHeight
    }
  }

  $page.Layout = @{ kind = 'protection'; rows = $rows; loops = $loops }
  return @($page)
}

function Draw-ProtectionLinks {
  param($Page)
  $rows = @(Get-Prop $Page.Layout 'rows' @())
  foreach ($row in $rows) {
    $ids = @($row.deviceIds)
    for ($index = 0; $index + 1 -lt $ids.Count; $index++) {
      $fromId = [string]$ids[$index]
      $toId = [string]$ids[$index + 1]
      $fromTerms = $null
      if ($Page.Terminals.ContainsKey($fromId)) { $fromTerms = $Page.Terminals[$fromId] }
      $toTerms = $null
      if ($Page.Terminals.ContainsKey($toId)) { $toTerms = $Page.Terminals[$toId] }
      $pointA = $null
      if ($null -ne $fromTerms -and $fromTerms.ContainsKey('T2')) { $pointA = $fromTerms['T2'] }
      if ($null -eq $pointA) { $pointA = Get-CenterPoint $Page $fromId }
      $pointB = $null
      if ($null -ne $toTerms -and $toTerms.ContainsKey('T1')) { $pointB = $toTerms['T1'] }
      if ($null -eq $pointB) { $pointB = Get-CenterPoint $Page $toId }
      if ($null -eq $pointA -or $null -eq $pointB) { continue }
      $points = New-Object System.Collections.Generic.List[object]
      $points.Add(@([double]$pointA[0], [double]$pointA[1]))
      $points.Add(@([double]$pointB[0], [double]$pointB[1]))
      Add-LineItem -Page $Page -Points $points -Weight 0.35
      $script:Stats.links += 1
    }
  }

  $loops = @(Get-Prop $Page.Layout 'loops' @())
  foreach ($loop in $loops) {
    $fromId = Get-Str $loop 'from' ''
    $toId = Get-Str $loop 'to' ''
    $pointA = Get-CenterPoint $Page $fromId
    $pointB = Get-CenterPoint $Page $toId
    if ($null -eq $pointA -or $null -eq $pointB) { continue }
    $points = New-Object System.Collections.Generic.List[object]
    $points.Add(@([double]$pointA[0], [double]$pointA[1]))
    $points.Add(@([double]$pointB[0], [double]$pointB[1]))
    Add-LineItem -Page $Page -Points $points -Weight 0.3 -Dashed $true
    $script:Stats.links += 1
    $label = Get-Str $loop 'label' ''
    if ($label) {
      Add-TextItem -Page $Page -X (([double]$pointA[0] + [double]$pointB[0]) / 2.0) -Y ((([double]$pointA[1] + [double]$pointB[1]) / 2.0) - 1.5) -Text $label -Size 2.3 -Anchor 'bottom-center'
    }
  }
}

# ─────────────────────────── 布局②：变电站总平面布置（§4.4②） ───────────────────────────
function Invoke-SubstationLayout {
  $page = New-PageData -Name '总平面布置'
  $model = $script:Model
  $plot = Get-Prop $model 'plot'
  $plotW = Get-Num $plot 'widthMm' 800.0
  $plotH = Get-Num $plot 'heightMm' 600.0
  $gridMm = Get-Num $plot 'gridMm' 0.0
  $margin = $script:MARGIN_MM
  $originX = $margin
  $originY = $margin

  # 1) 网格 + 场区外框
  if ($gridMm -gt 0) {
    $x = $gridMm
    while ($x -lt $plotW) {
      $points = New-Object System.Collections.Generic.List[object]
      $points.Add(@(($originX + $x), $originY))
      $points.Add(@(($originX + $x), ($originY + $plotH)))
      Add-LineItem -Page $page -Points $points -Weight 0.15 -Color $script:GRID_GRAY
      $x += $gridMm
    }
    $y = $gridMm
    while ($y -lt $plotH) {
      $points = New-Object System.Collections.Generic.List[object]
      $points.Add(@($originX, ($originY + $y)))
      $points.Add(@(($originX + $plotW), ($originY + $y)))
      Add-LineItem -Page $page -Points $points -Weight 0.15 -Color $script:GRID_GRAY
      $y += $gridMm
    }
  }
  Add-RectItem -Page $page -X $originX -Y $originY -W $plotW -H $plotH -Weight 0.6

  # 2) 指北针
  $meta = Get-Prop $model 'meta'
  if ((Get-Prop $meta 'compass' $false) -eq $true) {
    $cx = $originX + $plotW - 30.0
    $cy = $originY + 30.0
    $points = New-Object System.Collections.Generic.List[object]
    $points.Add(@($cx, ($cy + 18.0)))
    $points.Add(@($cx, ($cy - 8.0)))
    Add-LineItem -Page $page -Points $points -Weight 0.5 -Arrow $true
    Add-TextItem -Page $page -X $cx -Y ($cy + 20.0) -Text 'N' -Size 3.2 -Anchor 'top-center'
  }

  # 3) 设备区
  $zones = @(Get-Items $model 'zones')
  foreach ($zone in $zones) {
    $zx = $originX + (Get-Num $zone 'xMm' 0.0)
    $zy = $originY + (Get-Num $zone 'yMm' 0.0)
    $zw = Get-Num $zone 'widthMm' 0.0
    $zh = Get-Num $zone 'heightMm' 0.0
    Add-RectItem -Page $page -X $zx -Y $zy -W $zw -H $zh -Weight 0.3
    $label = Get-Str $zone 'label' (Get-Str $zone 'id' '')
    if ($label) { Add-TextItem -Page $page -X ($zx + $zw / 2.0) -Y ($zy + 3.0) -Text $label -Size 2.8 -Anchor 'top-center' }
  }

  # 4) 设备/图元落位（绝对坐标；缺坐标按 zone 网格吸附）
  $items = @(Get-Items $model 'items')
  $zoneById = @{}
  foreach ($zone in $zones) {
    $zoneId = Get-Str $zone 'id' ''
    if ($zoneId) { $zoneById[$zoneId] = $zone }
  }
  $autoCounter = @{}
  foreach ($item in $items) {
    $itemId = Get-Str $item 'id' ''
    $label = Get-Str $item 'label' ''
    $primitive = Get-Str $item 'primitive' ''
    $rawX = Get-Prop $item 'xMm'
    $rawY = Get-Prop $item 'yMm'
    $hasCoords = ($null -ne $rawX) -and ($null -ne $rawY)
    $x = 0.0
    $y = 0.0
    if ($hasCoords) {
      $x = $originX + [double]$rawX
      $y = $originY + [double]$rawY
    }
    else {
      $zoneId = Get-Str $item 'zone' ''
      $counter = 0
      if ($autoCounter.ContainsKey($zoneId)) { $counter = $autoCounter[$zoneId] }
      $autoCounter[$zoneId] = $counter + 1
      if ($zoneById.ContainsKey($zoneId)) {
        $zone = $zoneById[$zoneId]
        $zw = Get-Num $zone 'widthMm' 200.0
        $cols = [int][Math]::Floor(($zw - 20.0) / 60.0)
        if ($cols -lt 1) { $cols = 1 }
        $col = $counter % $cols
        $row = [int][Math]::Floor($counter / $cols)
        $x = $originX + (Get-Num $zone 'xMm' 0.0) + 10.0 + $col * 60.0
        $y = $originY + (Get-Num $zone 'yMm' 0.0) + 14.0 + $row * 30.0
      }
      else {
        $x = $originX + 20.0 + ($counter % 8) * 60.0
        $y = $originY + 20.0 + [int][Math]::Floor($counter / 8) * 30.0
      }
    }

    if ($primitive -eq 'rect') {
      $w = Get-Num $item 'widthMm' 0.0
      $h = Get-Num $item 'heightMm' 0.0
      $dashed = (Get-Prop $item 'dashed' $false) -eq $true
      Add-RectItem -Page $page -X $x -Y $y -W $w -H $h -Weight 0.35 -Dashed $dashed
      if ($label) { Add-TextItem -Page $page -X ($x + $w / 2.0) -Y ($y + $h / 2.0) -Text $label -Size 2.4 -Anchor 'center' }
      # 矩形图元也登记中心/包围盒，供 links 端点解析（道路/围墙等可作为连线端点）
      if ($itemId) {
        $page.Boxes[$itemId] = @{
          x = $x; y = $y; w = $w; h = $h; rot = 0
          center = @(($x + $w / 2.0), ($y + $h / 2.0))
          symbolRef = $null
        }
      }
      $script:Stats.items += 1
      continue
    }
    $ref = Get-Str $item 'symbolRef' ''
    $symbol = $script:Symbols[$ref]
    if ($null -eq $symbol) { throw "符号未注入：$ref（图元 $itemId）" }
    $rot = 0
    $rawRot = Get-Prop $item 'rot'
    if ($null -ne $rawRot) { $rot = [int]$rawRot }
    Add-Placement -Page $page -Id $itemId -SymbolRef $ref -X $x -Y $y -Rot $rot -Scale 1.0 -Label $label
    $script:Stats.items += 1
  }

  $links = @(Get-Items $model 'links')
  $page.Layout = @{ kind = 'substationPlan'; links = $links }
  return @($page)
}

function Draw-SubstationLinks {
  param($Page)
  $links = @(Get-Prop $Page.Layout 'links' @())
  foreach ($link in $links) {
    $fromId = Get-Str $link 'from' ''
    $toId = Get-Str $link 'to' ''
    $pointA = Get-CenterPoint $Page $fromId
    $pointB = Get-CenterPoint $Page $toId
    if ($null -eq $pointA -or $null -eq $pointB) { continue }
    $points = New-Object System.Collections.Generic.List[object]
    $points.Add(@([double]$pointA[0], [double]$pointA[1]))
    $points.Add(@([double]$pointB[0], [double]$pointB[1]))
    Add-LineItem -Page $Page -Points $points -Weight 0.35
    $script:Stats.links += 1
    $label = Get-Str $link 'label' ''
    if ($label) {
      Add-TextItem -Page $Page -X (([double]$pointA[0] + [double]$pointB[0]) / 2.0) -Y ((([double]$pointA[1] + [double]$pointB[1]) / 2.0) - 1.5) -Text $label -Size 2.4 -Anchor 'bottom-center'
    }
  }
}

# ─────────────────────────── 布局③：地理接线图（§4.4③） ───────────────────────────
function Convert-GeoPoint {
  param([double]$Lng, [double]$Lat)
  $bounds = $script:GeoBounds
  $spanX = [double]$bounds.east - [double]$bounds.west
  $spanY = [double]$bounds.north - [double]$bounds.south
  if ($spanX -le 0) { $spanX = 1.0 }
  if ($spanY -le 0) { $spanY = 1.0 }
  $x = [double]$bounds.ox + (($Lng - [double]$bounds.west) / $spanX) * [double]$bounds.w
  $y = [double]$bounds.oy + (([double]$bounds.north - $Lat) / $spanY) * [double]$bounds.h
  return @($x, $y)
}

function Get-GeoStationCenter {
  param($Station)
  $hasLng = $null -ne (Get-Prop $Station 'lng')
  $hasLat = $null -ne (Get-Prop $Station 'lat')
  if ($script:GeoBounds.useLng -and $hasLng -and $hasLat) {
    return Convert-GeoPoint (Get-Num $Station 'lng' 0.0) (Get-Num $Station 'lat' 0.0)
  }
  return Convert-GeoPoint (Get-Num $Station 'xMm' 0.0) (Get-Num $Station 'yMm' 0.0)
}

function Convert-GeoAnyPoint {
  param($Point, $StationsById)
  $text = [string]$Point[0]
  if ($text -and -not ($Point[0] -is [double]) -and -not ($Point[0] -is [int]) -and -not ($Point[0] -is [decimal]) -and -not ($Point[0] -is [single])) {
    if ($StationsById.ContainsKey($text)) { return Get-GeoStationCenter $StationsById[$text] }
    return $null
  }
  return Convert-GeoPoint ([double]$Point[0]) ([double]$Point[1])
}

function Invoke-GeoLayout {
  $page = New-PageData -Name '地理接线'
  $model = $script:Model
  $stations = @(Get-Items $model 'stations')
  $lines = @(Get-Items $model 'lines')
  $labels = @(Get-Items $model 'labels')
  $margin = $script:MARGIN_MM

  $paper = $script:PAPERS[$script:PaperName]
  $paperWmm = [double]$paper[0] * $script:MM_PER_IN
  $paperHmm = [double]$paper[1] * $script:MM_PER_IN
  $designW = [Math]::Max(120.0, $paperWmm - 50.0)
  $designH = $designW * ($paperHmm / $paperWmm)
  $designH = [Math]::Min($designH, [Math]::Max(120.0, $paperHmm - 50.0))

  $viewport = Get-Prop $model 'viewport'
  $useLng = $null -ne $viewport
  $west = 0.0
  $south = 0.0
  $east = 0.0
  $north = 0.0
  if ($useLng) {
    $west = Get-Num $viewport 'west' 112.0
    $south = Get-Num $viewport 'south' 30.0
    $east = Get-Num $viewport 'east' 114.0
    $north = Get-Num $viewport 'north' 32.0
    if (-not ($west -lt $east)) { $east = $west + 1.0 }
    if (-not ($south -lt $north)) { $north = $south + 1.0 }
  }
  else {
    # 无 viewport：以坐标包围盒自适配（平面坐标，x 右 y 下）
    $minX = [double]::MaxValue
    $minY = [double]::MaxValue
    $maxX = [double]::MinValue
    $maxY = [double]::MinValue
    foreach ($station in $stations) {
      $sx = Get-Num $station 'xMm' 0.0
      $sy = Get-Num $station 'yMm' 0.0
      if ($sx -lt $minX) { $minX = $sx }
      if ($sy -lt $minY) { $minY = $sy }
      if ($sx -gt $maxX) { $maxX = $sx }
      if ($sy -gt $maxY) { $maxY = $sy }
    }
    if ($minX -eq [double]::MaxValue) { $minX = 0.0; $minY = 0.0; $maxX = 100.0; $maxY = 100.0 }
    $padX = [Math]::Max(1.0, ($maxX - $minX) * 0.08)
    $padY = [Math]::Max(1.0, ($maxY - $minY) * 0.08)
    $west = $minX - $padX
    $south = $minY - $padY
    $east = $maxX + $padX
    $north = $maxY + $padY
    # 平面坐标与经纬投影同构：x 用 west/east，y 用 north/south（y 向下 = 南在下）
  }
  $script:GeoBounds = @{
    west = $west; south = $south; east = $east; north = $north
    ox = $margin; oy = $margin; w = $designW; h = $designH; useLng = $useLng
  }

  # 站
  $stationsById = @{}
  foreach ($station in $stations) {
    $id = Get-Str $station 'id' ''
    if ($id) { $stationsById[$id] = $station }
    $center = Get-GeoStationCenter $station
    $ref = Get-Str $station 'symbolRef' ''
    $symbol = $script:Symbols[$ref]
    if ($null -eq $symbol) { throw "符号未注入：$ref（厂站 $id）" }
    $size = Get-InstanceSize $symbol 0 1.0
    Add-Placement -Page $page -Id $id -SymbolRef $ref -X ($center[0] - $size[0] / 2.0) -Y ($center[1] - $size[1] / 2.0) -Rot 0 -Scale 1.0 -Label ''
    $label = Get-Str $station 'label' ''
    if ($label) {
      Add-TextItem -Page $page -X ($center[0] + $size[0] / 2.0 + 2.0) -Y ($center[1] - 2.0) -Text $label -Size 2.6 -Anchor 'top-left'
    }
    $script:Stats.stations += 1
  }

  # 线路：双回线（同 from/to）自动平行偏移 ±
  $pairCounts = @{}
  foreach ($line in $lines) {
    $fromId = Get-Str $line 'from' ''
    $toId = Get-Str $line 'to' ''
    if (-not $fromId -or -not $toId) { continue }
    $key = "$fromId|$toId"
    if ([string]::CompareOrdinal($fromId, $toId) -gt 0) { $key = "$toId|$fromId" }
    if ($pairCounts.ContainsKey($key)) { $pairCounts[$key] = $pairCounts[$key] + 1 } else { $pairCounts[$key] = 1 }
  }
  $pairSeen = @{}
  foreach ($line in $lines) {
    $fromId = Get-Str $line 'from' ''
    $toId = Get-Str $line 'to' ''
    if (-not $fromId -or -not $toId) { continue }
    $points = New-Object System.Collections.Generic.List[object]
    $rawPoints = @(Get-Items $line 'points')
    if ($rawPoints.Count -ge 2) {
      foreach ($rawPoint in $rawPoints) {
        $point = @($rawPoint)
        if ($point.Count -lt 2) { continue }
        $converted = Convert-GeoAnyPoint $point $stationsById
        if ($null -ne $converted) { $points.Add($converted) }
      }
    }
    if ($points.Count -lt 2) {
      $a = $null
      $b = $null
      if ($stationsById.ContainsKey($fromId)) { $a = Get-GeoStationCenter $stationsById[$fromId] }
      if ($stationsById.ContainsKey($toId)) { $b = Get-GeoStationCenter $stationsById[$toId] }
      if ($null -eq $a -or $null -eq $b) { continue }
      $points.Add($a)
      $points.Add($b)
    }

    # 首末点落在厂站符号框内时沿线段方向裁剪到框外，避免线路穿过站符号
    if ($stationsById.ContainsKey($fromId) -and $points.Count -ge 2) {
      $p0 = $points[0]
      $p1 = $points[1]
      $dx = [double]$p1[0] - [double]$p0[0]
      $dy = [double]$p1[1] - [double]$p0[1]
      $len = [Math]::Sqrt($dx * $dx + $dy * $dy)
      if ($len -gt 0.001) {
        $trim = 8.0
        if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) { $trim = 12.0 }
        $points[0] = @(([double]$p0[0] + $dx / $len * $trim), ([double]$p0[1] + $dy / $len * $trim))
      }
    }
    if ($stationsById.ContainsKey($toId) -and $points.Count -ge 2) {
      $lastIndex = $points.Count - 1
      $pl = $points[$lastIndex]
      $pp = $points[$lastIndex - 1]
      $dx = [double]$pl[0] - [double]$pp[0]
      $dy = [double]$pl[1] - [double]$pp[1]
      $len = [Math]::Sqrt($dx * $dx + $dy * $dy)
      if ($len -gt 0.001) {
        $trim = 8.0
        if ([Math]::Abs($dx) -ge [Math]::Abs($dy)) { $trim = 12.0 }
        $points[$lastIndex] = @(([double]$pl[0] - $dx / $len * $trim), ([double]$pl[1] - $dy / $len * $trim))
      }
    }

    $key = "$fromId|$toId"
    if ([string]::CompareOrdinal($fromId, $toId) -gt 0) { $key = "$toId|$fromId" }
    $count = 1
    if ($pairCounts.ContainsKey($key)) { $count = $pairCounts[$key] }
    $seen = 0
    if ($pairSeen.ContainsKey($key)) { $seen = $pairSeen[$key] }
    $pairSeen[$key] = $seen + 1
    if ($count -ge 2) {
      $offset = ($seen - (($count - 1) / 2.0)) * 3.0
      $first = $points[0]
      $last = $points[$points.Count - 1]
      $dx = [double]$last[0] - [double]$first[0]
      $dy = [double]$last[1] - [double]$first[1]
      $len = [Math]::Sqrt($dx * $dx + $dy * $dy)
      if ($len -gt 0.001) {
        $px = -$dy / $len
        $py = $dx / $len
        $shifted = New-Object System.Collections.Generic.List[object]
        foreach ($point in $points) {
          $shifted.Add(@(([double]$point[0] + $px * $offset), ([double]$point[1] + $py * $offset)))
        }
        $points = $shifted
      }
    }

    $weight = 0.45
    $voltage = Get-Num $line 'voltage' 0.0
    if ($voltage -ge 220) { $weight = 0.8 }
    elseif ($voltage -ge 110) { $weight = 0.65 }
    Add-LineItem -Page $page -Points $points -Weight $weight
    $script:Stats.lines += 1
    $label = Get-Str $line 'label' ''
    if ($label) {
      $midIndex = [int][Math]::Floor(($points.Count - 1) / 2)
      $p0 = $points[$midIndex]
      $p1 = $points[$midIndex + 1]
      Add-TextItem -Page $page -X (([double]$p0[0] + [double]$p1[0]) / 2.0) -Y ((([double]$p0[1] + [double]$p1[1]) / 2.0) - 1.5) -Text $label -Size 2.3 -Anchor 'bottom-center'
    }
  }

  # 独立地名标签
  foreach ($labelItem in $labels) {
    $text = Get-Str $labelItem 'text' ''
    if (-not $text) { continue }
    $hasLng = $null -ne (Get-Prop $labelItem 'lng')
    if ($script:GeoBounds.useLng -and $hasLng) {
      $point = Convert-GeoPoint (Get-Num $labelItem 'lng' 0.0) (Get-Num $labelItem 'lat' 0.0)
    }
    else {
      $point = Convert-GeoPoint (Get-Num $labelItem 'xMm' 0.0) (Get-Num $labelItem 'yMm' 0.0)
    }
    Add-TextItem -Page $page -X $point[0] -Y $point[1] -Text $text -Size 2.6 -Anchor 'top-left'
  }

  $script:GeoBounds = $null
  return @($page)
}

# ─────────────────────────── 布局④：潮流图/系统接线图（§4.4④） ───────────────────────────
function Invoke-PowerFlowLayout {
  $page = New-PageData -Name '潮流图'
  $model = $script:Model
  $levels = @(Get-Items $model 'levels')
  $buses = @(Get-Items $model 'buses')
  $nodes = @(Get-Items $model 'nodes')
  $transformers = @(Get-Items $model 'transformers')
  $margin = $script:MARGIN_MM
  $levelGap = 70.0

  # 分层：电压降序（同电压按模型顺序）
  $levelOrder = New-Object System.Collections.Generic.List[object]
  $index = 0
  foreach ($level in $levels) {
    $levelOrder.Add([pscustomobject]@{ Level = $level; Voltage = (-1.0 * (Get-Num $level 'voltage' 0.0)); Index = $index })
    $index += 1
  }
  $ordered = @($levelOrder | Sort-Object Voltage, Index)
  $rowOf = @{}
  $rowY = @{}
  for ($rowIndex = 0; $rowIndex -lt $ordered.Count; $rowIndex++) {
    $levelId = Get-Str $ordered[$rowIndex].Level 'id' ''
    if ($levelId) { $rowOf[$levelId] = $rowIndex }
    $rowY[$rowIndex] = $margin + $rowIndex * $levelGap
  }

  $busWidth = 140.0
  $busGap = 30.0
  $nodeGap = 40.0
  $maxRowIndex = 0
  for ($rowIndex = 0; $rowIndex -lt $ordered.Count; $rowIndex++) {
    $levelId = Get-Str $ordered[$rowIndex].Level 'id' ''
    $y = $rowY[$rowIndex]
    $cursorX = $margin
    $rowBuses = @($buses | Where-Object { (Get-Str $_ 'level' '') -eq $levelId })
    foreach ($bus in $rowBuses) {
      $busId = Get-Str $bus 'id' ''
      Add-BusLine -Page $page -Id $busId -X1 $cursorX -X2 ($cursorX + $busWidth) -Y $y -Label (Get-Str $bus 'label' '')
      $cursorX += $busWidth + $busGap
      $script:Stats.buses += 1
    }
    $rowNodes = @($nodes | Where-Object { (Get-Str $_ 'level' '') -eq $levelId })
    foreach ($node in $rowNodes) {
      $nodeId = Get-Str $node 'id' ''
      $ref = Get-Str $node 'symbolRef' ''
      $symbol = $script:Symbols[$ref]
      if ($null -eq $symbol) { throw "符号未注入：$ref（节点 $nodeId）" }
      $size = Get-InstanceSize $symbol 0 1.0
      Add-Placement -Page $page -Id $nodeId -SymbolRef $ref -X $cursorX -Y ($y - $size[1] / 2.0) -Rot 0 -Scale 1.0 -Label (Get-Str $node 'label' '')
      $cursorX += $size[0] + $nodeGap
    }
    if ($rowIndex -gt $maxRowIndex) { $maxRowIndex = $rowIndex }
  }

  # 未挂电压等级的节点：放在最后一行下方
  foreach ($node in $nodes) {
    $levelId = Get-Str $node 'level' ''
    if ($levelId -and $rowOf.ContainsKey($levelId)) { continue }
    $nodeId = Get-Str $node 'id' ''
    $ref = Get-Str $node 'symbolRef' ''
    $symbol = $script:Symbols[$ref]
    if ($null -eq $symbol) { throw "符号未注入：$ref（节点 $nodeId）" }
    $size = Get-InstanceSize $symbol 0 1.0
    Add-Placement -Page $page -Id $nodeId -SymbolRef $ref -X ($margin + 60.0) -Y ($rowY[$maxRowIndex] + $levelGap) -Rot 0 -Scale 1.0 -Label (Get-Str $node 'label' '')
  }

  # 变压器跨层居中
  foreach ($transformer in $transformers) {
    $transformerId = Get-Str $transformer 'id' ''
    $ref = Get-Str $transformer 'symbolRef' ''
    $symbol = $script:Symbols[$ref]
    if ($null -eq $symbol) { throw "符号未注入：$ref（变压器 $transformerId）" }
    $size = Get-InstanceSize $symbol 0 1.0
    $fromPoint = $null
    $toPoint = $null
    foreach ($bus in $buses) {
      if ((Get-Str $bus 'id' '') -eq (Get-Str $transformer 'from' '')) { $fromPoint = @(($margin + $busWidth / 2.0), $rowY[$rowOf[(Get-Str $bus 'level' '')]]) }
      if ((Get-Str $bus 'id' '') -eq (Get-Str $transformer 'to' '')) { $toPoint = @(($margin + $busWidth / 2.0), $rowY[$rowOf[(Get-Str $bus 'level' '')]]) }
    }
    $centerX = $margin + 120.0
    $centerY = $margin + 40.0
    if ($null -ne $fromPoint -and $null -ne $toPoint) {
      $centerX = ([double]$fromPoint[0] + [double]$toPoint[0]) / 2.0
      $centerY = ([double]$fromPoint[1] + [double]$toPoint[1]) / 2.0
    }
    elseif ($null -ne $fromPoint) { $centerX = [double]$fromPoint[0]; $centerY = [double]$fromPoint[1] + 30.0 }
    elseif ($null -ne $toPoint) { $centerX = [double]$toPoint[0]; $centerY = [double]$toPoint[1] - 30.0 }
    Add-Placement -Page $page -Id $transformerId -SymbolRef $ref -X ($centerX - $size[0] / 2.0) -Y ($centerY - $size[1] / 2.0) -Rot 0 -Scale 1.0 -Label (Get-Str $transformer 'label' '')
  }

  $flows = @(Get-Items $model 'flows')
  $links = @(Get-Items $model 'links')
  $page.Layout = @{ kind = 'powerFlow'; flows = $flows; links = $links }
  return @($page)
}

function Get-FlowEndpoints {
  param($Page, $Flow)
  $fromId = Get-Str $Flow 'from' ''
  $toId = Get-Str $Flow 'to' ''
  $direction = Get-Str $Flow 'direction' ''
  if ($direction) {
    $parts = $null
    foreach ($separator in @('→', '->')) {
      $at = $direction.IndexOf($separator)
      if ($at -gt 0) {
        $left = $direction.Substring(0, $at).Trim()
        $right = $direction.Substring($at + $separator.Length).Trim()
        if ($left -and $right) { $parts = @($left, $right) }
        break
      }
    }
    if ($null -ne $parts -and $parts[0] -eq $toId -and $parts[1] -eq $fromId) {
      return @($toId, $fromId)
    }
  }
  return @($fromId, $toId)
}

function Draw-PowerFlowLinks {
  param($Page)
  $links = @(Get-Prop $Page.Layout 'links' @())
  foreach ($link in $links) {
    $fromId = Get-Str $link 'from' ''
    $toId = Get-Str $link 'to' ''
    $pointA = Get-PagePoint $Page $fromId $null
    $pointB = Get-PagePoint $Page $toId $pointA
    if ($null -eq $pointA -or $null -eq $pointB) { continue }
    $points = New-Object System.Collections.Generic.List[object]
    $ax = [double]$pointA[0]
    $ay = [double]$pointA[1]
    $bx = [double]$pointB[0]
    $by = [double]$pointB[1]
    $points.Add(@($ax, $ay))
    if ([Math]::Abs($ay - $by) -lt 8.0 -and [Math]::Abs($ax - $bx) -gt 8.0) {
      # 同一层内的联线：绕行母线下方 22mm，避免穿过同行节点符号
      $routeY = $ay + 22.0
      $points.Add(@($ax, $routeY))
      $points.Add(@($bx, $routeY))
    }
    $points.Add(@($bx, $by))
    Add-LineItem -Page $Page -Points $points -Weight 0.35
    $script:Stats.links += 1
  }

  $flows = @(Get-Prop $Page.Layout 'flows' @())
  foreach ($flow in $flows) {
    $ends = Get-FlowEndpoints $Page $flow
    $fromRef = [string]$ends[0]
    $toRef = [string]$ends[1]
    $pointA = Get-PagePoint $Page $fromRef $null
    $pointB = Get-PagePoint $Page $toRef $pointA
    if ($null -eq $pointA -or $null -eq $pointB) { continue }
    $style = Get-Str $flow 'style' 'solid'
    $dashed = $style -eq 'dashed'
    $points = New-Object System.Collections.Generic.List[object]
    $ax = [double]$pointA[0]
    $ay = [double]$pointA[1]
    $bx = [double]$pointB[0]
    $by = [double]$pointB[1]
    $isBusPair = $Page.BusAnchor.ContainsKey($fromRef) -and $Page.BusAnchor.ContainsKey($toRef)
    if ($isBusPair -and [Math]::Abs($ax - $bx) -lt 8.0) {
      # 母线间潮流：沿母线向右偏 35mm 的竖直通道，避开跨层主变
      $anchor = $Page.BusAnchor[$fromRef]
      $offsetX = [Math]::Min([Math]::Max($ax + 35.0, [double]$anchor.x1 + 10.0), [double]$anchor.x2 - 10.0)
      $points.Add(@($offsetX, $ay))
      $points.Add(@($offsetX, $by))
    }
    else {
      $points.Add(@($ax, $ay))
      if (([Math]::Abs($ax - $bx) -gt 8.0) -and ([Math]::Abs($ay - $by) -gt 8.0)) {
        $midX = ($ax + $bx) / 2.0
        $points.Add(@($midX, $ay))
        $points.Add(@($midX, $by))
      }
      $points.Add(@($bx, $by))
    }
    Add-LineItem -Page $Page -Points $points -Weight 0.4 -Arrow $true -Dashed $dashed
    if ($style -eq 'double') {
      # 双线：在末段垂直方向 ±1mm 平行一条（不回退主箭头）
      $last = $points[$points.Count - 1]
      $prev = $points[$points.Count - 2]
      $dx = [double]$last[0] - [double]$prev[0]
      $dy = [double]$last[1] - [double]$prev[1]
      $len = [Math]::Sqrt($dx * $dx + $dy * $dy)
      if ($len -gt 0.001) {
        $ox = (-$dy / $len) * 1.0
        $oy = ($dx / $len) * 1.0
        $offsetPoints = New-Object System.Collections.Generic.List[object]
        foreach ($point in $points) { $offsetPoints.Add(@(([double]$point[0] + $ox), ([double]$point[1] + $oy))) }
        Add-LineItem -Page $Page -Points $offsetPoints -Weight 0.3
      }
    }
    $script:Stats.links += 1
    $script:Stats.flows += 1
    $label = Get-Str $flow 'label' ''
    if ($label) {
      # 标注取折线中段中点（母线间潮流会横向偏移，端点中点会压在主变上）
      $midIndex = [int][Math]::Floor(($points.Count - 1) / 2)
      $p0 = $points[$midIndex]
      $p1 = $points[$midIndex + 1]
      Add-TextItem -Page $Page -X ((([double]$p0[0] + [double]$p1[0]) / 2.0) + 1.5) -Y ((([double]$p0[1] + [double]$p1[1]) / 2.0) - 1.5) -Text $label -Size 2.3 -Anchor 'bottom-left'
    }
  }
}

# ─────────────────────────── 布局⑤：端子排图（§4.4⑤） ───────────────────────────
function Add-TerminalGround {
  param($Page, [double]$Cx, [double]$Cy)
  $stem = New-Object System.Collections.Generic.List[object]
  $stem.Add(@($Cx, ($Cy + 3.5)))
  $stem.Add(@($Cx, ($Cy + 5.0)))
  Add-LineItem -Page $Page -Points $stem -Weight 0.35
  $bars = @(
    @(($Cx - 3.5), ($Cy + 5.0), ($Cx + 3.5), ($Cy + 5.0)),
    @(($Cx - 2.2), ($Cy + 6.0), ($Cx + 2.2), ($Cy + 6.0)),
    @(($Cx - 1.0), ($Cy + 7.0), ($Cx + 1.0), ($Cy + 7.0))
  )
  foreach ($bar in $bars) {
    $points = New-Object System.Collections.Generic.List[object]
    $points.Add(@($bar[0], $bar[1]))
    $points.Add(@($bar[2], $bar[3]))
    Add-LineItem -Page $Page -Points $points -Weight 0.3
  }
}

function Add-TerminalTable {
  param($Page, $Strip, [double]$X, [double]$Y, [bool]$Continued)
  $terminals = @(Get-Items $Strip 'terminals')
  $columns = 1
  $rawColumns = Get-Prop $Strip 'columns'
  if ($null -ne $rawColumns) { $columns = [int]$rawColumns }
  if ($columns -lt 1) { $columns = 1 }
  if ($columns -gt 2) { $columns = 2 }
  $groupW = 120.0
  $groupGap = 6.0
  $tableW = $columns * $groupW + ($columns - 1) * $groupGap
  $titleH = 9.0
  $headerH = 9.0
  $rowH = 12.0
  $rowCount = [int][Math]::Ceiling($terminals.Count / [double]$columns)
  if ($rowCount -lt 1) { $rowCount = 1 }

  $label = Get-Str $Strip 'label' (Get-Str $Strip 'id' '端子排')
  if ($Continued) { $label = "$label（续）" }
  Add-TextItem -Page $Page -X ($X + $tableW / 2.0) -Y $Y -Text $label -Size 3.0 -Anchor 'top-center'

  $headerY = $Y + $titleH
  $headerTexts = @('端子号', '回路编号', '说明')
  for ($col = 0; $col -lt $columns; $col++) {
    $gx = $X + $col * ($groupW + $groupGap)
    Add-RectItem -Page $Page -X $gx -Y $headerY -W $groupW -H $headerH -Weight 0.35
    $cells = @(20.0, 35.0, 65.0)
    $cellX = $gx
    for ($cell = 0; $cell -lt 3; $cell++) {
      Add-TextItem -Page $Page -X ($cellX + $cells[$cell] / 2.0) -Y ($headerY + $headerH / 2.0) -Text $headerTexts[$cell] -Size 2.5 -Anchor 'center'
      $cellX += $cells[$cell]
      if ($cell -lt 2) {
        $divider = New-Object System.Collections.Generic.List[object]
        $divider.Add(@($cellX, $headerY))
        $divider.Add(@($cellX, ($headerY + $headerH)))
        Add-LineItem -Page $Page -Points $divider -Weight 0.25
      }
    }
  }

  $positions = @{}
  for ($row = 0; $row -lt $rowCount; $row++) {
    $rowY = $headerY + $headerH + $row * $rowH
    for ($col = 0; $col -lt $columns; $col++) {
      $terminalIndex = $col * $rowCount + $row
      if ($terminalIndex -ge $terminals.Count) { continue }
      $terminal = $terminals[$terminalIndex]
      $gx = $X + $col * ($groupW + $groupGap)
      Add-RectItem -Page $Page -X $gx -Y $rowY -W $groupW -H $rowH -Weight 0.3
      $cellX = $gx
      $cells = @(20.0, 35.0, 65.0)
      for ($cell = 0; $cell -lt 3; $cell++) {
        $cellX += $cells[$cell]
        if ($cell -lt 2) {
          $divider = New-Object System.Collections.Generic.List[object]
          $divider.Add(@($cellX, $rowY))
          $divider.Add(@($cellX, ($rowY + $rowH)))
          Add-LineItem -Page $Page -Points $divider -Weight 0.2
        }
      }
      $cx = $gx + 10.0
      $cy = $rowY + 5.0
      $no = Get-Str $terminal 'no' ''
      $type = Get-Str $terminal 'type' 'normal'
      Add-EllipseItem -Page $Page -Cx $cx -Cy $cy -Radius 4.0 -Weight 0.4
      Add-TextItem -Page $Page -X $cx -Y $cy -Text $no -Size 2.5 -Anchor 'center'
      if ($type -eq 'test') {
        Add-RectItem -Page $Page -X ($cx - 6.0) -Y ($cy - 5.0) -W 12.0 -H 10.0 -Weight 0.3 -Dashed $true
      }
      if ($type -eq 'ground') { Add-TerminalGround -Page $Page -Cx $cx -Cy $cy }
      $ref = Get-Str $terminal 'ref' ''
      if ($ref) { Add-TextItem -Page $Page -X ($gx + 37.5) -Y ($rowY + $rowH / 2.0) -Text $ref -Size 2.5 -Anchor 'center' }
      $desc = Get-Str $terminal 'desc' ''
      if ($desc) { Add-TextItem -Page $Page -X ($gx + 87.5) -Y ($rowY + $rowH / 2.0) -Text $desc -Size 2.3 -Anchor 'center' }
      if ($no -and -not $positions.ContainsKey($no)) { $positions[$no] = @($cx, $cy) }
    }
  }

  # 短接片：端子间竖短接线（link 指向同排另一端子号）
  foreach ($terminal in $terminals) {
    $linkValue = Get-Prop $terminal 'link'
    if ($null -eq $linkValue) { continue }
    $no = Get-Str $terminal 'no' ''
    $link = [string]$linkValue
    if (-not $positions.ContainsKey($no) -or -not $positions.ContainsKey($link)) { continue }
    $p1 = $positions[$no]
    $p2 = $positions[$link]
    $points = New-Object System.Collections.Generic.List[object]
    $points.Add(@([double]$p1[0], ([double]$p1[1] + 4.0)))
    if ([Math]::Abs([double]$p1[0] - [double]$p2[0]) -lt 0.1) {
      $points.Add(@([double]$p1[0], ([double]$p2[1] - 4.0)))
    }
    else {
      $midY = (([double]$p1[1] + 4.0) + ([double]$p2[1] - 4.0)) / 2.0
      $points.Add(@([double]$p1[0], $midY))
      $points.Add(@([double]$p2[0], $midY))
      $points.Add(@([double]$p2[0], ([double]$p2[1] - 4.0)))
    }
    Add-LineItem -Page $Page -Points $points -Weight 0.35
  }

  return ($titleH + $headerH + $rowCount * $rowH + 2.0)
}

function Split-TerminalChunks {
  param($List, [int]$ChunkSize)
  $chunks = New-Object System.Collections.Generic.List[object]
  $current = New-Object System.Collections.Generic.List[object]
  foreach ($item in $List) {
    $current.Add($item)
    if ($current.Count -ge $ChunkSize) {
      $chunks.Add($current)
      $current = New-Object System.Collections.Generic.List[object]
    }
  }
  if ($current.Count -gt 0) { $chunks.Add($current) }
  return $chunks
}

function Invoke-TerminalStripLayout {
  param([switch]$ForceSinglePage)
  $model = $script:Model
  $strips = @(Get-Items $model 'strips')
  $meta = Get-Prop $model 'meta'
  $rowsPerPage = 20
  if ($null -ne $meta) {
    $value = Get-Num $meta 'terminalsPerPage' 20.0
    if ($value -ge 1) { $rowsPerPage = [int]$value }
  }
  if ($ForceSinglePage) { $rowsPerPage = 100000 }

  $margin = $script:MARGIN_MM
  $pages = New-Object System.Collections.Generic.List[object]
  $page = New-PageData -Name '端子排'
  $pages.Add($page)
  $y = $margin
  $usedRows = 0

  foreach ($strip in $strips) {
    $terminals = @(Get-Items $strip 'terminals')
    $chunks = @(Split-TerminalChunks -List $terminals -ChunkSize $rowsPerPage)
    $chunkIndex = 0
    foreach ($chunk in $chunks) {
      $sectionRows = $chunk.Count + 2
      if ($usedRows -gt 0 -and ($usedRows + $sectionRows) -gt $rowsPerPage) {
        $page = New-PageData -Name '端子排'
        $pages.Add($page)
        $y = $margin
        $usedRows = 0
      }
      $continued = $chunkIndex -gt 0
      $height = Add-TerminalTable -Page $page -Strip @{ id = (Get-Str $strip 'id' ''); label = (Get-Str $strip 'label' ''); columns = (Get-Prop $strip 'columns') ; terminals = $chunk } -X $margin -Y $y -Continued $continued
      $y += $height + 18.0
      $usedRows += $sectionRows
      $chunkIndex += 1
    }
  }
  return $pages
}

# ─────────────────────────── 渲染（COM） ───────────────────────────
function Map-X {
  param([double]$X)
  return $script:OffsetX + (($X - $script:BBoxMinX) / $script:MM_PER_IN) * $script:Scale
}

function Map-Y {
  param([double]$Y)
  return $script:PaperH - $script:OffsetY - (($Y - $script:BBoxMinY) / $script:MM_PER_IN) * $script:Scale
}

function Render-Item {
  param($Page, $Item)
  $kind = [string]$Item.kind
  if ($kind -eq 'line') {
    $pts = $Item.pts
    $segments = $pts.Count - 1
    if ([bool]$Item.closed -and $pts.Count -ge 3) { $segments = $pts.Count }
    $shapes = New-Object System.Collections.Generic.List[object]
    for ($index = 0; $index -lt $segments; $index++) {
      $from = $pts[$index]
      $to = $pts[($index + 1) % $pts.Count]
      $x1 = Map-X ([double]$from[0])
      $y1 = Map-Y ([double]$from[1])
      $x2 = Map-X ([double]$to[0])
      $y2 = Map-Y ([double]$to[1])
      if ([Math]::Abs($x1 - $x2) -lt 1e-6 -and [Math]::Abs($y1 - $y2) -lt 1e-6) { continue }
      $shape = $Page.DrawLine($x1, $y1, $x2, $y2)
      $weightPt = [Math]::Max(0.5, [double]$Item.weight * $script:Scale * $script:PT_PER_MM)
      try { $shape.Cells('LineWeight').FormulaU = ("{0:0.##} pt" -f $weightPt) } catch { }
      if ([bool]$Item.dashed) { try { $shape.Cells('LinePattern').FormulaU = '2' } catch { } }
      if ([bool]$Item.color) { try { $shape.Cells('LineColor').FormulaU = [string]$Item.color } catch { } }
      if ([bool]$Item.arrow -and $index -eq ($segments - 1)) {
        try { $shape.Cells('EndArrow').FormulaU = '1' } catch { }
      }
      $script:Stats.shapes += 1
      $shapes.Add($shape)
    }
    return $shapes
  }
  if ($kind -eq 'ellipse') {
    $x1 = Map-X ([double]$Item.cx - [double]$Item.r)
    $x2 = Map-X ([double]$Item.cx + [double]$Item.r)
    $y1 = Map-Y ([double]$Item.cy - [double]$Item.r)
    $y2 = Map-Y ([double]$Item.cy + [double]$Item.r)
    $shape = $Page.DrawOval($x1, $y1, $x2, $y2)
    $weightPt = [Math]::Max(0.5, [double]$Item.weight * $script:Scale * $script:PT_PER_MM)
    try { $shape.Cells('LineWeight').FormulaU = ("{0:0.##} pt" -f $weightPt) } catch { }
    $script:Stats.shapes += 1
    return @($shape)
  }
  if ($kind -eq 'text') {
    $text = [string]$Item.text
    if (-not $text) { return @() }
    $box = Get-TextBox -X ([double]$Item.x) -Y ([double]$Item.y) -Text $text -Size ([double]$Item.size) -Anchor ([string]$Item.anchor)
    $left = Map-X $box[0]
    $right = Map-X ($box[0] + $box[2])
    $top = Map-Y $box[1]
    $bottom = Map-Y ($box[1] + $box[3])
    $x1 = [Math]::Min($left, $right)
    $x2 = [Math]::Max($left, $right)
    $y1 = [Math]::Min($top, $bottom)
    $y2 = [Math]::Max($top, $bottom)
    if (($x2 - $x1) -lt 1e-4 -or ($y2 - $y1) -lt 1e-4) { return @() }
    $shape = $Page.DrawRectangle($x1, $y1, $x2, $y2)
    try { $shape.Cells('LinePattern').FormulaU = '0' } catch { }
    try { $shape.Cells('FillPattern').FormulaU = '0' } catch { }
    $shape.Text = $text
    $sizePt = [Math]::Max(4.0, [double]$Item.size * $script:Scale * $script:PT_PER_MM)
    try { $shape.Cells('Char.Size').FormulaU = ("{0:0.##} pt" -f $sizePt) } catch { }
    try { $shape.Cells('Para.HorzAlign').FormulaU = '1' } catch { }
    $script:Stats.shapes += 1
    return @($shape)
  }
  return @()
}

function Group-Shapes {
  param($Visio, $Shapes)
  if ($null -eq $Shapes -or $Shapes.Count -lt 2) { return }
  try {
    $window = $Visio.ActiveWindow
    if ($null -eq $window) { return }
    $window.DeselectAll()
    # PS 5.1 late binding 取不到 Shape.Select，改用 Window.Select(shape, visSelect=2)
    foreach ($shape in $Shapes) { [void]$window.Select($shape, 2) }
    [void]$window.Selection.Group()
    $script:Stats.groups += 1
  }
  catch { }
}

function Get-FitScale {
  param($PageData)
  $b = $PageData.Bounds
  $bboxW = [Math]::Max(1.0, $b.maxX - $b.minX)
  $bboxH = [Math]::Max(1.0, $b.maxY - $b.minY)
  $marginIn = 0.39
  $availW = [Math]::Max(1.0, $script:PaperW - 2.0 * $marginIn)
  $availH = [Math]::Max(1.0, $script:PaperH - 2.0 * $marginIn)
  $scaleX = $availW / ($bboxW / $script:MM_PER_IN)
  $scaleY = $availH / ($bboxH / $script:MM_PER_IN)
  return [Math]::Min(1.0, [Math]::Min($scaleX, $scaleY))
}

function Render-PageData {
  param($Visio, $Doc, $PageData, [int]$Index, [double]$UniformScale)
  $b = $PageData.Bounds
  if ($b.minX -eq [double]::MaxValue) {
    $b.minX = 0.0; $b.minY = 0.0; $b.maxX = 100.0; $b.maxY = 100.0
  }
  $script:BBoxMinX = $b.minX
  $script:BBoxMinY = $b.minY
  $script:Scale = $UniformScale
  $bboxW = [Math]::Max(1.0, $b.maxX - $b.minX)
  $bboxH = [Math]::Max(1.0, $b.maxY - $b.minY)
  $drawW = ($bboxW / $script:MM_PER_IN) * $script:Scale
  $drawH = ($bboxH / $script:MM_PER_IN) * $script:Scale
  $script:OffsetX = ($script:PaperW - $drawW) / 2.0
  $script:OffsetY = ($script:PaperH - $drawH) / 2.0

  if ($Index -eq 0) {
    $page = $Doc.Pages.Item(1)
  }
  else {
    $page = $Doc.Pages.Add()
  }
  if ($null -eq $page) { throw '无法获取 Visio 页面' }
  try { $page.PageSheet.Cells('PageWidth').ResultIU = $script:PaperW } catch { }
  try { $page.PageSheet.Cells('PageHeight').ResultIU = $script:PaperH } catch { }
  $pageName = $PageData.Name
  if (-not $pageName) { $pageName = '页面' }
  if ($Index -gt 0) { $pageName = "$pageName-$($Index + 1)" }
  try { $page.Name = $pageName } catch { }
  try { $Visio.ActiveWindow.Page = $page } catch { }

  $items = $PageData.Items
  $groups = $PageData.Groups
  $groupIndex = 0
  $itemIndex = 0
  while ($itemIndex -lt $items.Count) {
    if ($groupIndex -lt $groups.Count -and $groups[$groupIndex].start -eq $itemIndex) {
      $group = $groups[$groupIndex]
      $groupIndex += 1
      $shapes = New-Object System.Collections.Generic.List[object]
      $end = $itemIndex + [int]$group.count
      while ($itemIndex -lt $end) {
        $rendered = Render-Item -Page $page -Item $items[$itemIndex]
        foreach ($shape in @($rendered)) {
          if ($null -ne $shape) { $shapes.Add($shape) }
        }
        $itemIndex += 1
      }
      Group-Shapes -Visio $Visio -Shapes $shapes
    }
    else {
      $null = Render-Item -Page $page -Item $items[$itemIndex]
      $itemIndex += 1
    }
  }
}

# ─────────────────────────── 主流程 ───────────────────────────
$script:Step = '解析参数'
try {
  $arg = $Json | ConvertFrom-Json
}
catch {
  Fail-Engine 'E2023' ("参数 JSON 解析失败：{0}" -f $_.Exception.Message)
}
if ($null -eq $arg) { Fail-Engine 'E2023' '缺少 -Json 参数内容' }

$script:Kind = Get-Str $arg 'kind' ''
$model = Get-Prop $arg 'model'
if ($null -eq $model) { Fail-Engine 'E2021' '缺少 model 参数' }
if (-not $script:Kind) { $script:Kind = Get-Str $model 'kind' '' }
if ($script:KINDS -notcontains $script:Kind) {
  Fail-Engine 'E2021' ("不支持的图类型 kind={0}（visio-layouts.ps1 仅处理 {1}；P2 三类走 visio-grid.ps1）" -f $script:Kind, ($script:KINDS -join ' / '))
}

$script:Model = $model
$script:Title = Get-Str $arg 'title' ''
if (-not $script:Title) { $script:Title = Get-Str $model 'title' '' }
$outPath = Get-Str $arg 'outputPath' ''
if (-not $outPath) { Fail-Engine 'E2023' 'outputPath 不能为空' }
$validations = @(Get-Items $arg 'validations')

$script:PaperName = Get-Str (Get-Prop $model 'meta') 'paper' ''
if (-not $script:PaperName -or -not $script:PAPERS.ContainsKey($script:PaperName)) {
  $defaultPaper = $script:KIND_PAPERS[$script:Kind]
  if ($defaultPaper) { $script:PaperName = $defaultPaper } else { $script:PaperName = 'A3-landscape' }
}
$paper = $script:PAPERS[$script:PaperName]
$script:PaperW = [double]$paper[0]
$script:PaperH = [double]$paper[1]

$symbolsArg = Get-Prop $arg 'symbols'
if ($null -ne $symbolsArg) {
  if ($symbolsArg -is [array]) {
    foreach ($symbol in @($symbolsArg)) {
      $symbolId = Get-Str $symbol 'id' ''
      if ($symbolId) { $script:Symbols[$symbolId] = $symbol }
    }
  }
  else {
    foreach ($property in $symbolsArg.PSObject.Properties) { $script:Symbols[$property.Name] = $property.Value }
  }
}

# 统计基数（引擎 result.stats 用；不影响绘制）
switch ($script:Kind) {
  'protection' {
    $script:Stats.circuits = (@(Get-Items $model 'circuits')).Count
    $script:Stats.strips = (@(Get-Items $model 'terminalStrips')).Count
    foreach ($strip in @(Get-Items $model 'terminalStrips')) { $script:Stats.terminals += (@(Get-Items $strip 'terminals')).Count }
  }
  'substationPlan' {
    $script:Stats.zones = (@(Get-Items $model 'zones')).Count
    $script:Stats.items = 0
  }
  'geoNetwork' {
    $script:Stats.stations = (@(Get-Items $model 'stations')).Count
    $script:Stats.lines = (@(Get-Items $model 'lines')).Count
  }
  'powerFlow' {
    $script:Stats.flows = (@(Get-Items $model 'flows')).Count
  }
  'terminalStrip' {
    $script:Stats.strips = (@(Get-Items $model 'strips')).Count
    foreach ($strip in @(Get-Items $model 'strips')) { $script:Stats.terminals += (@(Get-Items $strip 'terminals')).Count }
  }
}

$outDir = Split-Path -Parent $outPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

$script:Step = '布局'
$pages = @()
try {
  switch ($script:Kind) {
    'protection'      { $pages = @(Invoke-ProtectionLayout) }
    'substationPlan'  { $pages = @(Invoke-SubstationLayout) }
    'geoNetwork'      { $pages = @(Invoke-GeoLayout) }
    'powerFlow'       { $pages = @(Invoke-PowerFlowLayout) }
    'terminalStrip'   { $pages = @(Invoke-TerminalStripLayout) }
  }
}
catch {
  Fail-Engine 'E2023' ("布局失败（行 {0}）：{1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) '检查模型引用（节点/设备/厂站 symbolRef）是否都已在 symbols 中注入'
}
if ($pages.Count -eq 0) { Fail-Engine 'E2023' '布局未生成任何页面' }
$script:Stats.pages = $pages.Count

# 符号实例展开 + 每类连线（连线依赖展开后的端子坐标）
$script:Step = '展开符号图元'
try {
  foreach ($pageData in $pages) {
    Expand-Placements -Page $pageData
    switch ($script:Kind) {
      'protection'     { Draw-ProtectionLinks -Page $pageData }
      'substationPlan' { Draw-SubstationLinks -Page $pageData }
      'powerFlow'      { Draw-PowerFlowLinks -Page $pageData }
    }
  }
}
catch {
  Fail-Engine 'E2023' ("符号图元展开/连线失败（行 {0}）：{1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) '检查符号 JSON 的 primitives/view/terminals 字段与模型引用的 symbolRef'
}

$placementTotal = 0
foreach ($pageData in $pages) { $placementTotal += $pageData.Placements.Count }
if ($placementTotal -gt $script:MAX_SYMBOLS) {
  Fail-Engine 'E2021' ("符号实例数 {0} 超过上限 {1}" -f $placementTotal, $script:MAX_SYMBOLS)
}

# 每页计算布局适配比例（多页取统一最小值，保证打印一致）
$scales = @()
foreach ($pageData in $pages) { $scales += (Get-FitScale -PageData $pageData) }
$uniformScale = ($scales | Measure-Object -Minimum).Minimum

$script:Step = '启动 Visio'
$visio = $null
$doc = $null
$outputPath = $outPath
$format = 'vsdx'
try {
  $visio = New-Object -ComObject Visio.Application
  $visio.Visible = $false
  try { $visio.AlertResponse = 7 } catch { }
  $doc = $visio.Documents.Add('')

  $script:Step = '绘制图形'
  $renderFailed = $false
  try {
    for ($index = 0; $index -lt $pages.Count; $index++) {
      Render-PageData -Visio $visio -Doc $doc -PageData $pages[$index] -Index $index -UniformScale $uniformScale
    }
  }
  catch {
    $renderFailed = $true
    $firstError = $_.Exception.Message
  }

  if ($renderFailed -and $script:Kind -eq 'terminalStrip' -and $pages.Count -gt 1) {
    # T-04 防御：多页 Pages.Add 失败 → 降级单页整体缩放并 warn
    $script:EngineWarnings.Add('多页输出失败，已降级单页整体缩放（Visio Pages.Add 不可用）')
    try { if ($null -ne $doc) { [void]$doc.Close() } } catch { }
    $doc = $visio.Documents.Add('')
    $pages = @(Invoke-TerminalStripLayout -ForceSinglePage)
    $scales = @()
    foreach ($pageData in $pages) { $scales += (Get-FitScale -PageData $pageData) }
    $uniformScale = ($scales | Measure-Object -Minimum).Minimum
    $script:Stats.pages = $pages.Count
    for ($index = 0; $index -lt $pages.Count; $index++) {
      Render-PageData -Visio $visio -Doc $doc -PageData $pages[$index] -Index $index -UniformScale $uniformScale
    }
  }
  elseif ($renderFailed) {
    throw $firstError
  }

  $script:Step = '保存 vsdx'
  try {
    [void]$doc.SaveAs($outputPath)
  }
  catch {
    $legacyPath = [System.IO.Path]::ChangeExtension($outputPath, '.vsd')
    [void]$doc.SaveAs($legacyPath)
    $outputPath = $legacyPath
    $format = 'vsd'
  }
}
catch {
  Fail-Engine 'E2023' ("Visio 绘制失败（步骤：{0}）：{1}" -f $script:Step, $_.Exception.Message) '确认本机 Visio 可正常启动（COM 可用）；符号坐标为 mm 且 view.widthMm/heightMm 为正'
}
finally {
  if ($null -ne $doc) { try { [void]$doc.Close() } catch { } }
  if ($null -ne $visio) { try { [void]$visio.Quit() } catch { } }
}

if (-not (Test-Path -LiteralPath $outputPath)) {
  Fail-Engine 'E2023' ("保存失败：输出文件不存在 {0}" -f $outputPath)
}
$fileSize = 0
try { $fileSize = (Get-Item -LiteralPath $outputPath).Length } catch { $fileSize = 0 }
if ($fileSize -le 0) {
  Fail-Engine 'E2023' ("保存失败：输出文件为空 {0}" -f $outputPath)
}

@{
  ok = $true
  result = @{
    outputPath = $outputPath
    format = $format
    kind = $script:Kind
    paper = $script:PaperName
    pages = $script:Stats.pages
    bytes = $fileSize
    stats = @{
      symbols = $script:Stats.symbols
      shapes = $script:Stats.shapes
      links = $script:Stats.links
      groups = $script:Stats.groups
      pages = $script:Stats.pages
      circuits = $script:Stats.circuits
      terminals = $script:Stats.terminals
      strips = $script:Stats.strips
      stations = $script:Stats.stations
      lines = $script:Stats.lines
      flows = $script:Stats.flows
      zones = $script:Stats.zones
      items = $script:Stats.items
      buses = $script:Stats.buses
    }
    warnings = @($script:EngineWarnings)
    validations = @($validations)
  }
} | ConvertTo-Json -Depth 10
