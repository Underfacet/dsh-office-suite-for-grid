# visio-grid.ps1 —— 电网图绘制引擎（docs/P2-设计.md §4.3 / §4.4）
#
# 输入（-Json 单个 JSON 参数）：
#   {
#     "kind": "singleLine" | "distribution" | "networkTopology",
#     "title": "110kV 白云变主接线图",
#     "model": { ...§4.1 数据模型... },
#     "symbols": { "gb4728.7.QF": { ...符号 JSON... }, ... },   # Host 注入（已校验）
#     "outputPath": "D:\...\xxx.vsdx",
#     "validations": [ { level, code, message } ]                 # Host 已算好的接线提示（原样透传）
#   }
# 输出（stdout 单个 JSON）：
#   { "ok": true, "result": { "outputPath", "format", "kind", "paper",
#       "stats": { "symbols", "shapes", "links", "buses", "bays", "groups" },
#       "validations": [...] } }
#   失败 { "ok": false, "error": { "code", "message", "hint"? } }，code=E2021/E2023
#
# 设计要点（与 P1 visio-flowchart.ps1 同款 COM 风格，内存组装 master，不依赖本机模具）：
# - 单位：符号 JSON 为 mm（原点左上、x 右 y 下），Visio 内部为英寸，渲染时统一换算；
# - 三种布局（§4.4）：母线横排 + 间隔竖列；分层 BFS；环圆 + 径向树；
# - 绘制顺序：图元按数组顺序（后覆盖先）；连线 DrawLine + EndArrow；
# - 防御式 Close/Quit；每个 COM 调用上下文写入 E2023 的 message。
#
# 兼容：Windows PowerShell 5.1（文件 UTF-8 带 BOM；不使用 pwsh 7 独有语法）。

[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ─────────────────────────── 常量（mm） ───────────────────────────
$script:MM_PER_IN = 25.4
$script:PT_PER_MM = 72.0 / 25.4
$script:PAPERS = @{
  'A4-landscape' = @(11.69, 8.27)
  'A4-portrait'  = @(8.27, 11.69)
  'A3-landscape' = @(16.54, 11.69)
  'A3-portrait'  = @(11.69, 16.54)
}
$script:MARGIN_MM = 15.0
$script:BUS_ROW_GAP_MM = 60.0
$script:BAY_FIRST_MM = 12.0
$script:BAY_GAP_MM = 8.0
$script:OUT_LEAD_MM = 24.0
$script:COL_GAP_MM = 70.0
$script:ROW_GAP_MM = 40.0
$script:RING_MIN_R_MM = 45.0
$script:RING_LAYER_MM = 30.0
$script:MAX_SYMBOLS = 2000

# ─────────────────────────── 全局状态 ───────────────────────────
$script:Model = $null
$script:Symbols = @{}
$script:Items = New-Object System.Collections.Generic.List[object]
$script:Global = New-Object System.Collections.Generic.List[object]
$script:InstanceGroups = New-Object System.Collections.Generic.List[object]
$script:Placements = New-Object System.Collections.Generic.List[object]
$script:Terminals = @{}
$script:Boxes = @{}
$script:BusAnchor = @{}
$script:RingCenters = @{}
$script:ShapeRefs = New-Object System.Collections.Generic.List[object]
$script:Step = '初始化'
$script:Bounds = @{ minX = [double]::MaxValue; minY = [double]::MaxValue; maxX = [double]::MinValue; maxY = [double]::MinValue }
$script:Stats = @{ symbols = 0; shapes = 0; links = 0; buses = 0; bays = 0; groups = 0 }
$script:Scale = 1.0
$script:OffsetX = 0.0
$script:OffsetY = 0.0
$script:PaperW = 16.54
$script:PaperH = 11.69
$script:BBoxMinX = 0.0
$script:BBoxMinY = 0.0

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

function Update-Bounds {
  param([double]$X, [double]$Y)
  if ($X -lt $script:Bounds.minX) { $script:Bounds.minX = $X }
  if ($Y -lt $script:Bounds.minY) { $script:Bounds.minY = $Y }
  if ($X -gt $script:Bounds.maxX) { $script:Bounds.maxX = $X }
  if ($Y -gt $script:Bounds.maxY) { $script:Bounds.maxY = $Y }
}

function Add-LineItem {
  param($Points, [double]$Weight = 0.35, [bool]$Closed = $false, [bool]$Arrow = $false)
  # PS 5.1：@(List[object]) 在元素为数组时会抛 ArgumentException，逐项收集规避
  $pts = New-Object System.Collections.Generic.List[object]
  if ($null -ne $Points) { foreach ($point in $Points) { $pts.Add($point) } }
  if ($pts.Count -lt 2) { return }
  foreach ($pt in $pts) { Update-Bounds ([double]$pt[0]) ([double]$pt[1]) }
  $script:Items.Add(@{ kind = 'line'; pts = $pts; weight = $Weight; closed = $Closed; arrow = $Arrow })
}

function Add-EllipseItem {
  param([double]$Cx, [double]$Cy, [double]$Radius, [double]$Weight = 0.35)
  if ($Radius -le 0) { return }
  Update-Bounds ($Cx - $Radius) ($Cy - $Radius)
  Update-Bounds ($Cx + $Radius) ($Cy + $Radius)
  $script:Items.Add(@{ kind = 'ellipse'; cx = $Cx; cy = $Cy; r = $Radius; weight = $Weight })
}

function Get-TextBox {
  param([double]$X, [double]$Y, [string]$Text, [double]$Size, [string]$Anchor)
  $length = $Text.Length
  if ($length -lt 1) { $length = 1 }
  $cjk = 0
  foreach ($ch in $Text.ToCharArray()) { if ([int]$ch -ge 128) { $cjk += 1 } }
  $ascii = $length - $cjk
  $w = ($ascii * $Size * 0.62) + ($cjk * $Size * 1.02) + ($Size * 0.8)
  if ($w -lt ($Size * 6.0)) { $w = $Size * 6.0 }   # 过窄会触发 Visio 自动换行（QS/QF 折行）
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
  param([double]$X, [double]$Y, [string]$Text, [double]$Size = 2.5, [string]$Anchor = 'top-left')
  if (-not $Text) { return }
  $box = Get-TextBox -X $X -Y $Y -Text $Text -Size $Size -Anchor $Anchor
  Update-Bounds $box[0] $box[1]
  Update-Bounds ($box[0] + $box[2]) ($box[1] + $box[3])
  $script:Items.Add(@{ kind = 'text'; x = $X; y = $Y; text = $Text; size = $Size; anchor = $Anchor })
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
  $rot = (($Rot % 360) + 360) % 360
  $view = Get-Prop $symbol 'view'
  $vw = (Get-Num $view 'widthMm' 1.0) * $Scale
  $vh = (Get-Num $view 'heightMm' 1.0) * $Scale
  $size = Get-InstanceSize $symbol $rot $Scale
  $ew = $size[0]
  $eh = $size[1]

  $script:Boxes[$Id] = @{
    x = $X; y = $Y; w = $ew; h = $eh; rot = $rot
    center = @(($X + $ew / 2.0), ($Y + $eh / 2.0))
    symbolRef = $SymbolRef
  }
  Update-Bounds $X $Y
  Update-Bounds ($X + $ew) ($Y + $eh)

  # 端子绝对坐标（mm）
  $terminalMap = @{}
  foreach ($terminal in @(Get-Items $symbol 'terminals')) {
    $tid = Get-Str $terminal 'id' ''
    $at = @(Get-Items $terminal 'at')
    if (-not $tid -or $at.Count -lt 2) { continue }
    $local = Convert-Point ([double]$at[0] * $Scale) ([double]$at[1] * $Scale) $vw $vh $rot
    $terminalMap[$tid] = @(($X + [double]$local[0]), ($Y + [double]$local[1]))
  }
  $script:Terminals[$Id] = $terminalMap

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
        Add-LineItem -Points $points -Weight $stroke -Closed ([bool](Get-Prop $primitive 'closed' $false))
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
        Add-LineItem -Points $points -Weight $stroke -Closed $true
      }
    }
    elseif ($type -eq 'circle') {
      $center = @(Get-Items $primitive 'center')
      $radius = Get-Num $primitive 'radius' 0
      if ($center.Count -ge 2 -and $radius -gt 0) {
        $local = Convert-Point ([double]$center[0] * $Scale) ([double]$center[1] * $Scale) $vw $vh $rot
        Add-EllipseItem -Cx ($X + [double]$local[0]) -Cy ($Y + [double]$local[1]) -Radius ($radius * $Scale) -Weight $stroke
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
        if ($points.Count -ge 2) { Add-LineItem -Points $points -Weight $stroke }
      }
    }
    elseif ($type -eq 'text') {
      $at = @(Get-Items $primitive 'at')
      $text = Get-Str $primitive 'text' ''
      if ($at.Count -lt 2 -or -not $text) { continue }
      $local = Convert-Point ([double]$at[0] * $Scale) ([double]$at[1] * $Scale) $vw $vh $rot
      Add-TextItem -X ($X + [double]$local[0]) -Y ($Y + [double]$local[1]) -Text $text -Size ((Get-Num $primitive 'sizeMm' 2.5) * $Scale) -Anchor (Get-Str $primitive 'anchor' 'top-left')
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
    Add-SymbolInstance -Id ("$Id/$childRef") -SymbolRef $childRef -X ($X + [double]$local[0]) -Y ($Y + [double]$local[1]) -Rot $rot -Scale $childScale -Depth ($Depth + 1)
  }

  if ($Label) {
    Add-TextItem -X ($X + $ew / 2.0) -Y ($Y + $eh + 1.5) -Text $Label -Size 2.6 -Anchor 'top-center'
  }
}

function Add-Placement {
  param([string]$Id, [string]$SymbolRef, [double]$X, [double]$Y, [int]$Rot = 0, [double]$Scale = 1.0, [string]$Label = '')
  if (-not $Id -or -not $SymbolRef) { return }
  $script:Placements.Add(@{ id = $Id; symbolRef = $SymbolRef; x = $X; y = $Y; rot = $Rot; scale = $Scale; label = $Label })
}

function Add-BusLine {
  param([string]$Id, [double]$X1, [double]$X2, [double]$Y, [string]$Label)
  if ($X2 -le $X1) { $X2 = $X1 + 200.0 }
  $points = New-Object System.Collections.Generic.List[object]
  $points.Add(@($X1, $Y))
  $points.Add(@($X2, $Y))
  Add-LineItem -Points $points -Weight 0.8
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
  $script:Terminals[$Id] = $terminalMap
  $script:BusAnchor[$Id] = @{ x1 = $X1; x2 = $X2; y = $Y }
  if ($Label) { Add-TextItem -X $X1 -Y ($Y - 2.5) -Text $Label -Size 3.0 -Anchor 'bottom-left' }
}

function Add-NodePlacement {
  param([string]$Id, $Node, [double]$CenterX, [double]$CenterY)
  $ref = Get-Str $Node 'symbolRef' ''
  $symbol = $script:Symbols[$ref]
  if ($null -eq $symbol) { throw "符号未注入：$ref（节点 $Id）" }
  $size = Get-InstanceSize $symbol 0 1.0
  Add-Placement -Id $Id -SymbolRef $ref -X ($CenterX - $size[0] / 2.0) -Y ($CenterY - $size[1] / 2.0) -Rot 0 -Scale 1.0 -Label (Get-Str $Node 'label' '')
  $script:Stats.symbols += 1
}

# ─────────────────────────── 布局①：电气主接线图（§4.4） ───────────────────────────
function Invoke-SingleLineLayout {
  $buses = @(Get-Items $script:Model 'buses')
  $bays = @(Get-Items $script:Model 'bays')
  $transformers = @(Get-Items $script:Model 'transformers')
  $busTies = @(Get-Items $script:Model 'busTies')

  # 母线排序：电压降序 → section → 模型顺序
  $ordered = New-Object System.Collections.Generic.List[object]
  $index = 0
  foreach ($bus in $buses) {
    $ordered.Add([pscustomobject]@{
        Bus = $bus
        Voltage = (-1.0 * (Get-Num $bus 'voltage' 0))
        Section = (Get-Str $bus 'section' '')
        Index = $index
      })
    $index += 1
  }
  $busRows = @($ordered | Sort-Object Voltage, Section, Index)

  $bayBusMap = @{}
  foreach ($bay in $bays) {
    $bayId = Get-Str $bay 'id' ''
    if ($bayId) { $bayBusMap[$bayId] = (Get-Str $bay 'bus' '') }
  }
  # 行高预留：该行起始的变压器高度
  $rowTransformerH = @{}
  foreach ($transformer in $transformers) {
    $fromBayId = Get-Str $transformer 'fromBay' ''
    $busId = ''
    if ($bayBusMap.ContainsKey($fromBayId)) { $busId = $bayBusMap[$fromBayId] }
    if (-not $busId) { continue }
    $symbol = $script:Symbols[(Get-Str $transformer 'symbolRef' '')]
    $height = 40.0
    if ($null -ne $symbol) { $height = (Get-InstanceSize $symbol 0 1.0)[1] + 12.0 }
    if (-not $rowTransformerH.ContainsKey($busId) -or $rowTransformerH[$busId] -lt $height) { $rowTransformerH[$busId] = $height }
  }

  $rowY = $script:MARGIN_MM
  $busInfo = @{}
  $bayInfo = @{}
  $maxRight = $script:MARGIN_MM + 240.0
  $currentBusY = $rowY
  foreach ($row in $busRows) {
    $bus = $row.Bus
    $busId = Get-Str $bus 'id' ''
    $rowBays = @($bays | Where-Object { (Get-Str $_ 'bus' '') -eq $busId })
    $cursorX = $script:MARGIN_MM
    $maxStack = 0.0
    foreach ($bay in $rowBays) {
      $bayId = Get-Str $bay 'id' ''
      $devices = @(Get-Items $bay 'devices')
      $colWidth = 40.0
      $deviceList = New-Object System.Collections.Generic.List[object]
      $stackHeight = $script:BAY_FIRST_MM
      foreach ($device in $devices) {
        $ref = Get-Str $device 'symbolRef' ''
        $symbol = $script:Symbols[$ref]
        if ($null -eq $symbol) { throw "符号未注入：$ref（设备 $((Get-Str $device 'id' ''))）" }
        $rot = Get-DeviceRot $symbol
        $size = Get-InstanceSize $symbol $rot 1.0
        $deviceList.Add([pscustomobject]@{ Device = $device; Ref = $ref; Rot = $rot; W = $size[0]; H = $size[1]; X = 0.0; Y = 0.0 })
        if (($size[0] + 40.0) -gt $colWidth) { $colWidth = $size[0] + 40.0 }
        $stackHeight += $size[1] + $script:BAY_GAP_MM
      }
      if ($deviceList.Count -eq 0) { $stackHeight = 0.0 }
      $centerX = $cursorX + ($colWidth / 2.0)
      $deviceY = $currentBusY + $script:BAY_FIRST_MM
      foreach ($entry in $deviceList) {
        $entry.X = $centerX - ($entry.W / 2.0)
        $entry.Y = $deviceY
        $deviceY += $entry.H + $script:BAY_GAP_MM
        Add-Placement -Id (Get-Str $entry.Device 'id' '') -SymbolRef $entry.Ref -X $entry.X -Y $entry.Y -Rot $entry.Rot -Scale 1.0 -Label (Get-Str $entry.Device 'label' '')
        $script:Stats.symbols += 1
      }
      $bottom = $currentBusY
      if ($deviceList.Count -gt 0) { $bottom = $deviceY - $script:BAY_GAP_MM }
      $bayLabel = Get-Str $bay 'label' ''
      if ($bayLabel) { Add-TextItem -X $centerX -Y ($currentBusY + 4.0) -Text $bayLabel -Size 2.4 -Anchor 'top-center' }
      $bayInfo[$bayId] = [pscustomobject]@{ Id = $bayId; BusId = $busId; CenterX = $centerX; ColWidth = $colWidth; Devices = $deviceList; Bottom = $bottom }
      $stackHeight += $script:OUT_LEAD_MM
      if ($stackHeight -gt $maxStack) { $maxStack = $stackHeight }
      $cursorX += $colWidth
      $script:Stats.bays += 1
    }
    if ($cursorX -gt $maxRight) { $maxRight = $cursorX }
    $busInfo[$busId] = [pscustomobject]@{ Bus = $bus; Y = $currentBusY; Left = $script:MARGIN_MM; Right = $cursorX; MaxStack = $maxStack }
    $extra = 0.0
    if ($rowTransformerH.ContainsKey($busId)) { $extra = $rowTransformerH[$busId] }
    $advance = [Math]::Max($script:BUS_ROW_GAP_MM, $maxStack + $extra + 20.0)
    $rowY = $currentBusY + $advance
    $currentBusY = $rowY
  }

  # 母联列占位：母线右端向右延展 70mm
  if ($busTies.Count -gt 0) { $maxRight += 70.0 }
  foreach ($busId in @($busInfo.Keys)) {
    $info = $busInfo[$busId]
    $info.Right = [Math]::Max($maxRight, $info.Right)
  }

  $script:Step = '绘制母线'
  foreach ($busId in @($busInfo.Keys)) {
    $info = $busInfo[$busId]
    Add-BusLine -Id $busId -X1 $info.Left -X2 $info.Right -Y $info.Y -Label (Get-Str $info.Bus 'label' '')
    $script:Stats.buses += 1
  }

  $script:Step = '放置变压器'
  foreach ($transformer in $transformers) {
    $id = Get-Str $transformer 'id' ''
    $ref = Get-Str $transformer 'symbolRef' ''
    $symbol = $script:Symbols[$ref]
    if ($null -eq $symbol) { throw "符号未注入：$ref（变压器 $id）" }
    $size = Get-InstanceSize $symbol 0 1.0
    $fromBayId = Get-Str $transformer 'fromBay' ''
    $toBayId = Get-Str $transformer 'toBay' ''
    $fromBay = $null
    $toBay = $null
    if ($bayInfo.ContainsKey($fromBayId)) { $fromBay = $bayInfo[$fromBayId] }
    if ($bayInfo.ContainsKey($toBayId)) { $toBay = $bayInfo[$toBayId] }
    $centerX = $script:MARGIN_MM + 120.0
    if ($null -ne $fromBay -and $null -ne $toBay) { $centerX = ($fromBay.CenterX + $toBay.CenterX) / 2.0 }
    elseif ($null -ne $fromBay) { $centerX = $fromBay.CenterX }
    elseif ($null -ne $toBay) { $centerX = $toBay.CenterX }
    $rowForFrom = $null
    if ($null -ne $fromBay -and $busInfo.ContainsKey($fromBay.BusId)) { $rowForFrom = $busInfo[$fromBay.BusId] }
    $rowForTo = $null
    if ($null -ne $toBay -and $busInfo.ContainsKey($toBay.BusId)) { $rowForTo = $busInfo[$toBay.BusId] }
    $y = $script:MARGIN_MM
    if ($null -ne $rowForFrom) { $y = $rowForFrom.Y + $rowForFrom.MaxStack + 6.0 }
    elseif ($null -ne $rowForTo) { $y = $rowForTo.Y + $rowForTo.MaxStack + 6.0 }
    if ($null -ne $rowForFrom -and $null -ne $rowForTo -and $rowForTo.Y -gt $rowForFrom.Y) {
      $limit = $rowForTo.Y - 20.0 - $size[1]
      if ($limit -gt $y) { $y = $limit }
    }
    Add-Placement -Id $id -SymbolRef $ref -X ($centerX - $size[0] / 2.0) -Y $y -Rot 0 -Scale 1.0 -Label (Get-Str $transformer 'label' '')
    $script:Stats.symbols += 1
  }

  $script:Step = '放置母联'
  foreach ($busTie in $busTies) {
    $between = @(Get-Items $busTie 'between')
    if ($between.Count -lt 2) { continue }
    $busA = $null
    $busB = $null
    if ($busInfo.ContainsKey([string]$between[0])) { $busA = $busInfo[[string]$between[0]] }
    if ($busInfo.ContainsKey([string]$between[1])) { $busB = $busInfo[[string]$between[1]] }
    if ($null -eq $busA -or $null -eq $busB) { continue }
    $upper = $busA
    if ($busB.Y -lt $upper.Y) { $upper = $busB }
    $cursorY = $upper.Y + $script:BAY_FIRST_MM + 4.0
    foreach ($device in @(Get-Items $busTie 'devices')) {
      $ref = Get-Str $device 'symbolRef' ''
      $symbol = $script:Symbols[$ref]
      if ($null -eq $symbol) { throw "符号未注入：$ref（母联设备）" }
      $rot = Get-DeviceRot $symbol
      $size = Get-InstanceSize $symbol $rot 1.0
      $x = $maxRight - 30.0 - ($size[0] / 2.0)
      Add-Placement -Id (Get-Str $device 'id' '') -SymbolRef $ref -X $x -Y $cursorY -Rot $rot -Scale 1.0 -Label (Get-Str $device 'label' '')
      $script:Stats.symbols += 1
      $cursorY += $size[1] + $script:BAY_GAP_MM
    }
  }
}

# ─────────────────────────── 布局②：配电系统图（分层 BFS） ───────────────────────────
function Add-Adjacency {
  param($Adj, [string]$A, [string]$B)
  if (-not $A -or -not $B -or $A -eq $B) { return }
  if (-not $Adj.ContainsKey($A) -or -not $Adj.ContainsKey($B)) { return }
  if (-not $Adj[$A].Contains($B)) { $Adj[$A].Add($B) }
  if (-not $Adj[$B].Contains($A)) { $Adj[$B].Add($A) }
}

function Get-BaseId {
  param([string]$Ref)
  $text = [string]$Ref
  $dot = $text.LastIndexOf('.')
  if ($dot -gt 0) { return $text.Substring(0, $dot) }
  return $text
}

function Invoke-DistributionLayout {
  $sources = @(Get-Items $script:Model 'sources')
  $nodes = @(Get-Items $script:Model 'nodes')
  $order = New-Object System.Collections.Generic.List[string]
  $instances = @{}
  foreach ($source in $sources) {
    $id = Get-Str $source 'id' ''
    if (-not $id -or $instances.ContainsKey($id)) { continue }
    $instances[$id] = @{ Id = $id; Node = $source; Ref = (Get-Str $source 'symbolRef' ''); IsSource = $true }
    $order.Add($id)
  }
  foreach ($node in $nodes) {
    $id = Get-Str $node 'id' ''
    if (-not $id -or $instances.ContainsKey($id)) { continue }
    $instances[$id] = @{ Id = $id; Node = $node; Ref = (Get-Str $node 'symbolRef' ''); IsSource = $false }
    $order.Add($id)
  }
  $adj = @{}
  foreach ($id in $order) { $adj[$id] = New-Object System.Collections.Generic.List[string] }
  foreach ($feeder in @(Get-Items $script:Model 'feeders')) {
    Add-Adjacency $adj (Get-BaseId (Get-Str $feeder 'from' '')) (Get-BaseId (Get-Str $feeder 'to' ''))
  }
  foreach ($link in @(Get-Items $script:Model 'links')) {
    Add-Adjacency $adj (Get-BaseId (Get-Str $link 'from' '')) (Get-BaseId (Get-Str $link 'to' ''))
  }

  $depth = @{}
  $queue = New-Object System.Collections.Generic.Queue[string]
  foreach ($id in $order) {
    if ($instances[$id].IsSource) {
      $depth[$id] = 0
      $queue.Enqueue($id)
    }
  }
  while ($queue.Count -gt 0) {
    $current = $queue.Dequeue()
    foreach ($next in $adj[$current]) {
      if (-not $depth.ContainsKey($next)) {
        $depth[$next] = $depth[$current] + 1
        $queue.Enqueue($next)
      }
    }
  }
  foreach ($id in $order) {
    if (-not $depth.ContainsKey($id)) { $depth[$id] = 0 }
  }

  $levelMap = @{}
  foreach ($id in $order) {
    $level = [int]$depth[$id]
    if (-not $levelMap.ContainsKey($level)) { $levelMap[$level] = New-Object System.Collections.Generic.List[string] }
    $levelMap[$level].Add($id)
  }
  $levelIds = @($levelMap.Keys | Sort-Object)
  $levelInfo = @{}
  $maxHeight = 0.0
  $cursorX = $script:MARGIN_MM
  foreach ($level in $levelIds) {
    $ids = @($levelMap[$level])
    $colWidth = 40.0
    $sizes = @{}
    $columnHeight = 0.0
    foreach ($id in $ids) {
      $symbol = $script:Symbols[$instances[$id].Ref]
      if ($null -eq $symbol) { throw "符号未注入：$($instances[$id].Ref)（节点 $id）" }
      $size = Get-InstanceSize $symbol 0 1.0
      $sizes[$id] = $size
      if ($size[0] -gt $colWidth) { $colWidth = $size[0] }
      $columnHeight += $size[1] + $script:ROW_GAP_MM
    }
    if ($ids.Count -gt 0) { $columnHeight -= $script:ROW_GAP_MM }
    if ($columnHeight -gt $maxHeight) { $maxHeight = $columnHeight }
    $levelInfo[$level] = [pscustomobject]@{ Ids = $ids; X = $cursorX; ColWidth = $colWidth; Sizes = $sizes; Height = $columnHeight }
    $cursorX += $colWidth + $script:COL_GAP_MM
  }
  foreach ($level in $levelIds) {
    $info = $levelInfo[$level]
    $y = $script:MARGIN_MM + (($maxHeight - $info.Height) / 2.0)
    foreach ($id in $info.Ids) {
      $size = $info.Sizes[$id]
      $nodeX = $info.X + (($info.ColWidth - $size[0]) / 2.0)
      Add-Placement -Id $id -SymbolRef $instances[$id].Ref -X $nodeX -Y $y -Rot 0 -Scale 1.0 -Label (Get-Str $instances[$id].Node 'label' '')
      $script:Stats.symbols += 1
      $y += $size[1] + $script:ROW_GAP_MM
    }
  }
}

# ─────────────────────────── 布局③：通信网络拓扑（环圆 + 径向树） ───────────────────────────
function Invoke-TopologyLayout {
  $rings = @(Get-Items $script:Model 'rings')
  $nodes = @(Get-Items $script:Model 'nodes')
  if ($rings.Count -eq 0) {
    Invoke-DistributionLayout
    return
  }
  $byId = @{}
  $order = New-Object System.Collections.Generic.List[string]
  foreach ($node in $nodes) {
    $id = Get-Str $node 'id' ''
    if (-not $id -or $byId.ContainsKey($id)) { continue }
    $byId[$id] = $node
    $order.Add($id)
  }
  $adj = @{}
  foreach ($id in $order) { $adj[$id] = New-Object System.Collections.Generic.List[string] }
  foreach ($link in @(Get-Items $script:Model 'links')) {
    Add-Adjacency $adj (Get-BaseId (Get-Str $link 'from' '')) (Get-BaseId (Get-Str $link 'to' ''))
  }
  foreach ($feeder in @(Get-Items $script:Model 'feeders')) {
    Add-Adjacency $adj (Get-BaseId (Get-Str $feeder 'from' '')) (Get-BaseId (Get-Str $feeder 'to' ''))
  }

  $ringRadii = New-Object System.Collections.Generic.List[double]
  $maxRadius = $script:RING_MIN_R_MM
  foreach ($ring in $rings) {
    $ringNodes = @(Get-Items $ring 'nodes')
    $radius = [Math]::Max($script:RING_MIN_R_MM, ($ringNodes.Count * 22.0) / (2.0 * [Math]::PI))
    $ringRadii.Add($radius)
    if ($radius -gt $maxRadius) { $maxRadius = $radius }
  }
  $cols = [Math]::Min(2, $rings.Count)
  if ($cols -lt 1) { $cols = 1 }
  $cellW = (2.0 * $maxRadius) + 110.0
  $cellH = (2.0 * $maxRadius) + 110.0
  $polar = @{}
  $ringNodeIds = @{}
  $ringIndex = 0
  foreach ($ring in $rings) {
    $ringId = Get-Str $ring 'id' ''
    $column = $ringIndex % $cols
    $row = [int][Math]::Floor($ringIndex / $cols)
    $cx = $script:MARGIN_MM + $maxRadius + $column * $cellW
    $cy = $script:MARGIN_MM + $maxRadius + $row * $cellH
    $radius = $ringRadii[$ringIndex]
    $ringNodeList = @(Get-Items $ring 'nodes')
    $thisRingId = $ringId
    if ($thisRingId) { $script:RingCenters[$thisRingId] = @($cx, $cy) }
    for ($index = 0; $index -lt $ringNodeList.Count; $index++) {
      $id = [string]$ringNodeList[$index]
      if (-not $byId.ContainsKey($id)) { continue }
      $angle = -90.0 + ($index * 360.0 / $ringNodeList.Count)
      $rad = $angle * [Math]::PI / 180.0
      $nx = $cx + $radius * [Math]::Cos($rad)
      $ny = $cy + $radius * [Math]::Sin($rad)
      $polar[$id] = @{ cx = $cx; cy = $cy; angle = $angle; radius = $radius }
      $ringNodeIds[$id] = $true
      Add-NodePlacement -Id $id -Node $byId[$id] -CenterX $nx -CenterY $ny
    }
    $ringIndex += 1
  }

  # 非环节点：从环节点 BFS 向外径向分层
  $level = @{}
  $parent = @{}
  $queue = New-Object System.Collections.Generic.Queue[string]
  foreach ($id in $order) {
    if ($ringNodeIds.ContainsKey($id)) {
      $level[$id] = 0
      $queue.Enqueue($id)
    }
  }
  $bfsOrder = New-Object System.Collections.Generic.List[string]
  while ($queue.Count -gt 0) {
    $current = $queue.Dequeue()
    foreach ($next in $adj[$current]) {
      if ($level.ContainsKey($next)) { continue }
      $level[$next] = $level[$current] + 1
      $parent[$next] = $current
      $bfsOrder.Add($next)
      $queue.Enqueue($next)
    }
  }
  foreach ($id in $bfsOrder) {
    if ($ringNodeIds.ContainsKey($id)) { continue }
    $anchorId = $parent[$id]
    $anchor = $polar[$anchorId]
    if ($null -eq $anchor) { continue }
    $radius = [double]$anchor.radius + $script:RING_LAYER_MM * [int]$level[$id]
    $rad = [double]$anchor.angle * [Math]::PI / 180.0
    $nx = [double]$anchor.cx + $radius * [Math]::Cos($rad)
    $ny = [double]$anchor.cy + $radius * [Math]::Sin($rad)
    $polar[$id] = @{ cx = $anchor.cx; cy = $anchor.cy; angle = $anchor.angle; radius = $radius }
    Add-NodePlacement -Id $id -Node $byId[$id] -CenterX $nx -CenterY $ny
  }

  # 完全孤立的节点排在所有环的下方
  $orphanIndex = 0
  $ringRows = [int][Math]::Ceiling($rings.Count / [double]$cols)
  $orphanY = $script:MARGIN_MM + $maxRadius + ($ringRows * $cellH) - 20.0
  foreach ($id in $order) {
    if ($polar.ContainsKey($id)) { continue }
    $x = $script:MARGIN_MM + 60.0 + $orphanIndex * 90.0
    Add-NodePlacement -Id $id -Node $byId[$id] -CenterX $x -CenterY $orphanY
    $polar[$id] = @{ cx = $x; cy = $orphanY; angle = 0.0; radius = 0.0 }
    $orphanIndex += 1
  }
}

# ─────────────────────────── 连线（§4.1 默认连线 + 显式 links/feeders） ───────────────────────────
function Add-MergedLink {
  param($List, $Seen, [string]$From, [string]$To, [string]$Kind, [string]$Label = '')
  if (-not $From -or -not $To) { return }
  $key = ''
  if ([string]::CompareOrdinal($From, $To) -le 0) { $key = "$From|$To" } else { $key = "$To|$From" }
  if ($Seen.ContainsKey($key)) { return }
  $Seen[$key] = $true
  $List.Add(@{ from = $From; to = $To; kind = $Kind; label = $Label })
}

function Get-LastDeviceId {
  param($Bay)
  if ($null -eq $Bay) { return '' }
  $devices = @(Get-Items $Bay 'devices')
  if ($devices.Count -eq 0) { return '' }
  return Get-Str $devices[$devices.Count - 1] 'id' ''
}

function Get-MergedLinks {
  $list = New-Object System.Collections.Generic.List[object]
  $seen = @{}
  foreach ($link in @(Get-Items $script:Model 'links')) {
    $kind = Get-Str $link 'kind' 'line'
    if (-not $kind) { $kind = 'line' }
    Add-MergedLink $list $seen (Get-Str $link 'from' '') (Get-Str $link 'to' '') $kind (Get-Str $link 'label' '')
  }
  foreach ($feeder in @(Get-Items $script:Model 'feeders')) {
    Add-MergedLink $list $seen (Get-Str $feeder 'from' '') (Get-Str $feeder 'to' '') 'feeder' (Get-Str $feeder 'label' '')
  }
  $bays = @(Get-Items $script:Model 'bays')
  $bayMap = @{}
  foreach ($bay in $bays) {
    $bayId = Get-Str $bay 'id' ''
    if ($bayId) { $bayMap[$bayId] = $bay }
  }
  foreach ($bay in $bays) {
    $busId = Get-Str $bay 'bus' ''
    $devices = @(Get-Items $bay 'devices')
    if ($busId -and $devices.Count -gt 0) {
      Add-MergedLink $list $seen $busId (Get-Str $devices[0] 'id' '') 'line'
    }
    for ($index = 0; $index -lt $devices.Count - 1; $index++) {
      Add-MergedLink $list $seen (Get-Str $devices[$index] 'id' '') (Get-Str $devices[$index + 1] 'id' '') 'line'
    }
  }
  foreach ($transformer in @(Get-Items $script:Model 'transformers')) {
    $transformerId = Get-Str $transformer 'id' ''
    if (-not $transformerId) { continue }
    $fromBay = $null
    $toBay = $null
    $fromBayId = Get-Str $transformer 'fromBay' ''
    $toBayId = Get-Str $transformer 'toBay' ''
    if ($bayMap.ContainsKey($fromBayId)) { $fromBay = $bayMap[$fromBayId] }
    if ($bayMap.ContainsKey($toBayId)) { $toBay = $bayMap[$toBayId] }
    $lastFrom = Get-LastDeviceId $fromBay
    if ($lastFrom) { Add-MergedLink $list $seen $lastFrom $transformerId 'line' }
    # 绘制口径：二次侧接 toBay 所在母线（间隔设备由该母线引出），避免连线穿越母线；
    # 校验口径（ops/grid-model.js）仍按 §4.1 规则接 toBay 末设备，电气连通性等价。
    $toBusId = ''
    if ($null -ne $toBay) { $toBusId = Get-Str $toBay 'bus' '' }
    if ($toBusId) { Add-MergedLink $list $seen $transformerId $toBusId 'line' }
    else {
      $lastTo = Get-LastDeviceId $toBay
      if ($lastTo) { Add-MergedLink $list $seen $transformerId $lastTo 'line' }
    }
  }
  foreach ($busTie in @(Get-Items $script:Model 'busTies')) {
    $between = @(Get-Items $busTie 'between')
    if ($between.Count -lt 2) { continue }
    $devices = @(Get-Items $busTie 'devices')
    if ($devices.Count -eq 0) {
      Add-MergedLink $list $seen ([string]$between[0]) ([string]$between[1]) 'line'
      continue
    }
    Add-MergedLink $list $seen ([string]$between[0]) (Get-Str $devices[0] 'id' '') 'line'
    for ($index = 0; $index -lt $devices.Count - 1; $index++) {
      Add-MergedLink $list $seen (Get-Str $devices[$index] 'id' '') (Get-Str $devices[$index + 1] 'id' '') 'line'
    }
    Add-MergedLink $list $seen (Get-Str $devices[$devices.Count - 1] 'id' '') ([string]$between[1]) 'line'
  }
  return $list
}

function Get-LinkPoint {
  param([string]$Ref, $OtherPoint)
  $text = [string]$Ref
  if (-not $text) { return $null }
  $base = $text
  $terminal = ''
  $dot = $text.LastIndexOf('.')
  if ($dot -gt 0) {
    $maybeBase = $text.Substring(0, $dot)
    if ($script:Terminals.ContainsKey($maybeBase) -or $script:Boxes.ContainsKey($maybeBase) -or $script:BusAnchor.ContainsKey($maybeBase)) {
      $base = $maybeBase
      $terminal = $text.Substring($dot + 1)
    }
  }
  if ($terminal -and $script:Terminals.ContainsKey($base)) {
    $terms = $script:Terminals[$base]
    if ($terms.ContainsKey($terminal)) { return $terms[$terminal] }
  }
  if ($script:BusAnchor.ContainsKey($base) -and -not $terminal) {
    $anchor = $script:BusAnchor[$base]
    $x = ($anchor.x1 + $anchor.x2) / 2.0
    if ($null -ne $OtherPoint) {
      $x = [Math]::Min([Math]::Max([double]$OtherPoint[0], $anchor.x1), $anchor.x2)
    }
    return @($x, [double]$anchor.y)
  }
  if ($script:Terminals.ContainsKey($base)) {
    $terms = $script:Terminals[$base]
    if ($terms.Count -gt 0) {
      if ($null -ne $OtherPoint -and $script:Boxes.ContainsKey($base) -and $terms.ContainsKey('T1') -and $terms.ContainsKey('T2')) {
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
  if ($script:BusAnchor.ContainsKey($base)) {
    $anchor = $script:BusAnchor[$base]
    $x = ($anchor.x1 + $anchor.x2) / 2.0
    if ($null -ne $OtherPoint) {
      $x = [Math]::Min([Math]::Max([double]$OtherPoint[0], $anchor.x1), $anchor.x2)
    }
    return @($x, [double]$anchor.y)
  }
  if ($script:Boxes.ContainsKey($base)) { return $script:Boxes[$base].center }
  return $null
}

function Draw-Links {
  $script:Step = '绘制连线'
  $merged = Get-MergedLinks
  foreach ($link in $merged) {
    $other = Get-LinkPoint $link.to $null
    $pointA = Get-LinkPoint $link.from $other
    $otherB = Get-LinkPoint $link.from $null
    $pointB = Get-LinkPoint $link.to $otherB
    if ($null -eq $pointA -or $null -eq $pointB) { continue }
    $points = New-Object System.Collections.Generic.List[object]
    $kind = [string]$link.kind
    if ($kind -eq 'feeder' -or $kind -eq 'fiber' -or $kind -eq 'cable') {
      $midX = ([double]$pointA[0] + [double]$pointB[0]) / 2.0
      $points.Add(@([double]$pointA[0], [double]$pointA[1]))
      $points.Add(@($midX, [double]$pointA[1]))
      $points.Add(@($midX, [double]$pointB[1]))
      $points.Add(@([double]$pointB[0], [double]$pointB[1]))
    }
    else {
      $points.Add(@([double]$pointA[0], [double]$pointA[1]))
      $points.Add(@([double]$pointB[0], [double]$pointB[1]))
    }
    Add-LineItem -Points $points -Weight 0.35 -Arrow $true
    $script:Stats.links += 1
    $label = [string]$link.label
    if ($label) {
      $p0 = $points[0]
      $p1 = $points[1]
      Add-TextItem -X (([double]$p0[0] + [double]$p1[0]) / 2.0) -Y (([double]$p0[1] + [double]$p1[1]) / 2.0 - 1.5) -Text $label -Size 2.4 -Anchor 'bottom-center'
    }
  }
}

function Draw-BayStubs {
  $bays = @(Get-Items $script:Model 'bays')
  if ($bays.Count -eq 0) { return }
  $usedBays = @{}
  foreach ($transformer in @(Get-Items $script:Model 'transformers')) {
    $fromBayId = Get-Str $transformer 'fromBay' ''
    $toBayId = Get-Str $transformer 'toBay' ''
    if ($fromBayId) { $usedBays[$fromBayId] = $true }
    if ($toBayId) { $usedBays[$toBayId] = $true }
  }
  foreach ($bay in $bays) {
    $bayId = Get-Str $bay 'id' ''
    if ($usedBays.ContainsKey($bayId)) { continue }
    $devices = @(Get-Items $bay 'devices')
    if ($devices.Count -eq 0) { continue }
    $lastId = Get-Str $devices[$devices.Count - 1] 'id' ''
    if (-not $script:Terminals.ContainsKey($lastId)) { continue }
    $terms = $script:Terminals[$lastId]
    $point = $null
    if ($terms.ContainsKey('T2')) { $point = $terms['T2'] }
    elseif ($terms.ContainsKey('T1')) { $point = $terms['T1'] }
    if ($null -eq $point) { continue }
    $points = New-Object System.Collections.Generic.List[object]
    $points.Add(@([double]$point[0], [double]$point[1]))
    $points.Add(@([double]$point[0], [double]$point[1] + $script:OUT_LEAD_MM))
    Add-LineItem -Points $points -Weight 0.35
  }
}

function Draw-RingPolygons {
  $rings = @(Get-Items $script:Model 'rings')
  foreach ($ring in $rings) {
    $points = New-Object System.Collections.Generic.List[object]
    foreach ($rawId in @(Get-Items $ring 'nodes')) {
      $id = [string]$rawId
      if (-not $script:Boxes.ContainsKey($id)) { continue }
      $center = $script:Boxes[$id].center
      $points.Add(@([double]$center[0], [double]$center[1]))
    }
    if ($points.Count -ge 2) { Add-LineItem -Points $points -Weight 0.5 -Closed $true }
    $ringId = Get-Str $ring 'id' ''
    $ringLabel = Get-Str $ring 'label' $ringId
    if ($ringId -and $script:RingCenters.ContainsKey($ringId) -and $ringLabel) {
      $center = $script:RingCenters[$ringId]
      Add-TextItem -X $center[0] -Y $center[1] -Text $ringLabel -Size 3.0 -Anchor 'center'
    }
  }
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
      if ([bool]$Item.arrow -and $index -eq ($segments - 1)) {
        try { $shape.Cells('EndArrow').FormulaU = '1' } catch { }
      }
      $script:Stats.shapes += 1
      $script:ShapeRefs.Add($shape)
    }
    return
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
    $script:ShapeRefs.Add($shape)
    return
  }
  if ($kind -eq 'text') {
    $text = [string]$Item.text
    if (-not $text) { return }
    $box = Get-TextBox -X ([double]$Item.x) -Y ([double]$Item.y) -Text $text -Size ([double]$Item.size) -Anchor ([string]$Item.anchor)
    $left = Map-X $box[0]
    $right = Map-X ($box[0] + $box[2])
    $top = Map-Y $box[1]
    $bottom = Map-Y ($box[1] + $box[3])
    $x1 = [Math]::Min($left, $right)
    $x2 = [Math]::Max($left, $right)
    $y1 = [Math]::Min($top, $bottom)
    $y2 = [Math]::Max($top, $bottom)
    if (($x2 - $x1) -lt 1e-4 -or ($y2 - $y1) -lt 1e-4) { return }
    $shape = $Page.DrawRectangle($x1, $y1, $x2, $y2)
    try { $shape.Cells('LinePattern').FormulaU = '0' } catch { }
    try { $shape.Cells('FillPattern').FormulaU = '0' } catch { }
    $shape.Text = $text
    $sizePt = [Math]::Max(4.0, [double]$Item.size * $script:Scale * $script:PT_PER_MM)
    try { $shape.Cells('Char.Size').FormulaU = ("{0:0.##} pt" -f $sizePt) } catch { }
    try { $shape.Cells('Para.HorzAlign').FormulaU = '1' } catch { }
    $script:Stats.shapes += 1
    $script:ShapeRefs.Add($shape)
  }
}

function Render-Items {
  param($Page, $Items)
  $script:ShapeRefs = New-Object System.Collections.Generic.List[object]
  foreach ($item in $Items) { Render-Item -Page $Page -Item $item }
  return $script:ShapeRefs
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

# ─────────────────────────── 主流程 ───────────────────────────
$script:Step = '解析参数'
try {
  $arg = $Json | ConvertFrom-Json
}
catch {
  Fail-Engine 'E2023' ("参数 JSON 解析失败：{0}" -f $_.Exception.Message)
}
if ($null -eq $arg) { Fail-Engine 'E2023' '缺少 -Json 参数内容' }

$model = Get-Prop $arg 'model'
if ($null -eq $model) { Fail-Engine 'E2021' '缺少 model 参数' }
$kind = Get-Str $arg 'kind' ''
if (-not $kind) { $kind = Get-Str $model 'kind' '' }
if (@('singleLine', 'distribution', 'networkTopology') -notcontains $kind) {
  Fail-Engine 'E2021' ("不支持的图类型 kind={0}（singleLine / distribution / networkTopology）" -f $kind)
}
$paperName = Get-Str (Get-Prop $model 'meta') 'paper' 'A3-landscape'
if (-not $script:PAPERS.ContainsKey($paperName)) { $paperName = 'A3-landscape' }
$title = Get-Str $arg 'title' ''
if (-not $title) { $title = Get-Str $model 'title' '' }
$outPath = Get-Str $arg 'outputPath' ''
if (-not $outPath) { Fail-Engine 'E2023' 'outputPath 不能为空' }
$validations = @(Get-Items $arg 'validations')

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

$script:Model = $model
$bays = @(Get-Items $model 'bays')
$busTies = @(Get-Items $model 'busTies')
$instanceTotal = (@(Get-Items $model 'nodes')).Count + (@(Get-Items $model 'sources')).Count + (@(Get-Items $model 'buses')).Count
$instanceTotal += (@(Get-Items $model 'transformers')).Count + $busTies.Count
foreach ($bay in $bays) { $instanceTotal += (@(Get-Items $bay 'devices')).Count }
foreach ($busTie in $busTies) { $instanceTotal += (@(Get-Items $busTie 'devices')).Count }
if ($instanceTotal -gt $script:MAX_SYMBOLS) {
  Fail-Engine 'E2021' ("符号实例数 {0} 超过上限 {1}" -f $instanceTotal, $script:MAX_SYMBOLS)
}

$outDir = Split-Path -Parent $outPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

$script:Step = '布局'
try {
  switch ($kind) {
    'singleLine' {
      if ((@(Get-Items $model 'buses')).Count -eq 0) { Fail-Engine 'E2021' 'singleLine 模型缺少 buses' }
      Invoke-SingleLineLayout
    }
    'distribution' {
      if ((@(Get-Items $model 'sources')).Count -eq 0) { Fail-Engine 'E2021' 'distribution 模型缺少 sources' }
      Invoke-DistributionLayout
    }
    'networkTopology' {
      Invoke-TopologyLayout
    }
  }
}
catch {
  Fail-Engine 'E2023' ("布局失败：{0}" -f $_.Exception.Message) '检查模型引用（节点/设备 symbolRef）是否都已在 symbols 中注入'
}

# 图元展开：逐实例计算端子绝对坐标（mm）
$script:Step = '展开符号图元'
$script:Global = $script:Items
try {
  foreach ($placement in $script:Placements) {
    $items = New-Object System.Collections.Generic.List[object]
    $script:Items = $items
    Add-SymbolInstance -Id $placement.id -SymbolRef $placement.symbolRef -X $placement.x -Y $placement.y -Rot $placement.rot -Scale $placement.scale -Label $placement.label
    $script:InstanceGroups.Add(@{ id = $placement.id; items = $items })
  }
  $script:Items = $script:Global
}
catch {
  Fail-Engine 'E2023' ("符号图元展开失败：{0}" -f $_.Exception.Message) '检查符号 JSON 的 primitives/view/terminals 字段'
}

$script:Step = '绘制母线'
try {
  if ($kind -eq 'singleLine') {
    foreach ($bus in @(Get-Items $model 'buses')) {
      $busId = Get-Str $bus 'id' ''
      if ($busId -and -not $script:BusAnchor.ContainsKey($busId)) {
        # 未在布局中登记的母线（理论上不会发生）：补一条默认宽度母线
        Add-BusLine -Id $busId -X1 $script:MARGIN_MM -X2 ($script:MARGIN_MM + 240.0) -Y $script:MARGIN_MM -Label (Get-Str $bus 'label' '')
      }
    }
  }
}
catch {
  Fail-Engine 'E2023' ("绘制母线失败：{0}" -f $_.Exception.Message)
}

$script:Step = '绘制连线'
try {
  Draw-Links
  Draw-BayStubs
  if ($kind -eq 'networkTopology') { Draw-RingPolygons }
}
catch {
  Fail-Engine 'E2023' ("绘制连线失败：{0}" -f $_.Exception.Message)
}

# 适配纸型：整体缩放 + 居中
$script:Step = '纸型适配'
$bounds = $script:Bounds
$bboxW = [Math]::Max(1.0, $bounds.maxX - $bounds.minX)
$bboxH = [Math]::Max(1.0, $bounds.maxY - $bounds.minY)
$paper = $script:PAPERS[$paperName]
$script:PaperW = [double]$paper[0]
$script:PaperH = [double]$paper[1]
$marginIn = 0.39
$availW = [Math]::Max(1.0, $script:PaperW - 2.0 * $marginIn)
$availH = [Math]::Max(1.0, $script:PaperH - 2.0 * $marginIn)
$scaleX = $availW / ($bboxW / $script:MM_PER_IN)
$scaleY = $availH / ($bboxH / $script:MM_PER_IN)
$script:Scale = [Math]::Min(1.0, [Math]::Min($scaleX, $scaleY))
$script:BBoxMinX = $bounds.minX
$script:BBoxMinY = $bounds.minY
$drawW = ($bboxW / $script:MM_PER_IN) * $script:Scale
$drawH = ($bboxH / $script:MM_PER_IN) * $script:Scale
$script:OffsetX = ($script:PaperW - $drawW) / 2.0
$script:OffsetY = ($script:PaperH - $drawH) / 2.0

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
  $page = $visio.ActivePage
  if ($null -eq $page) { throw '无法获取 Visio 活动页面' }
  try { $page.PageSheet.Cells('PageWidth').ResultIU = $script:PaperW } catch { }
  try { $page.PageSheet.Cells('PageHeight').ResultIU = $script:PaperH } catch { }
  if ($title) { try { $page.Name = $title } catch { } }

  $script:Step = '绘制图形'
  foreach ($group in $script:InstanceGroups) {
    $shapes = Render-Items -Page $page -Items $group.items
    Group-Shapes -Visio $visio -Shapes $shapes
  }
  Render-Items -Page $page -Items $script:Global | Out-Null

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
    kind = $kind
    paper = $paperName
    bytes = $fileSize
    stats = @{
      symbols = $script:Stats.symbols
      shapes = $script:Stats.shapes
      links = $script:Stats.links
      buses = $script:Stats.buses
      bays = $script:Stats.bays
      groups = $script:Stats.groups
    }
    validations = @($validations)
  }
} | ConvertTo-Json -Depth 10
