# grid-p3-smoke.ps1 —— visio-layouts.ps1 冒烟脚本（P3 五类图各一最小用例）
#
# 用法（Windows PowerShell 5.1）：
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\grid-p3-smoke.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\grid-p3-smoke.ps1 -OutDir runtime\tmp\diagrams
#
# 行为：
# 1. 从 suite\symbols 读取符号库（跳过 P1 示例 example-*.json），按用例注入所需符号；
# 2. 构造 protection / substationPlan / geoNetwork / powerFlow / terminalStrip 五个最小模型；
# 3. 以独立子进程调用 suite\engines\visio-layouts.ps1（-Json 载荷写临时文件，避免命令行长度限制）；
# 4. 断言返回 ok=true、输出 .vsdx 存在且大小 > 0、shapes > 0、terminalStrip 分页页数符合预期；
# 5. 打印统计表；任一用例失败则 exit 1。
#
# 产物位于 runtime\tmp\diagrams（git 忽略）；不写入仓库受版本控制文件。

[CmdletBinding()]
param(
  [string]$OutDir = '',
  [string]$SymbolsDir = '',
  [string]$EnginePath = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path $repoRoot 'runtime\tmp\diagrams' }
if (-not $SymbolsDir) { $SymbolsDir = Join-Path $repoRoot 'suite\symbols' }
if (-not $EnginePath) { $EnginePath = Join-Path $repoRoot 'suite\engines\visio-layouts.ps1' }
$payloadDir = Join-Path $repoRoot 'runtime\tmp\grid-p3-smoke'

if (-not (Test-Path -LiteralPath $EnginePath)) { throw "引擎脚本不存在：$EnginePath" }
if (-not (Test-Path -LiteralPath $SymbolsDir)) { throw "符号库目录不存在：$SymbolsDir" }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
if (-not (Test-Path -LiteralPath $payloadDir)) { New-Item -ItemType Directory -Force -Path $payloadDir | Out-Null }

function Read-SymbolLibrary {
  param([string]$Root)
  $library = @{}
  foreach ($directory in @(Get-ChildItem -LiteralPath $Root -Directory)) {
    if ($directory.Name -notmatch '^gb4728(\.\d+)?$') { continue }
    foreach ($file in @(Get-ChildItem -LiteralPath $directory.FullName -Filter '*.json' -File)) {
      if ($file.Name.ToLowerInvariant().StartsWith('example-')) { continue }
      $symbol = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
      $symbolId = [string]$symbol.id
      if ($symbolId) { $library[$symbolId] = $symbol }
    }
  }
  return $library
}

function Select-Symbols {
  param($Library, [string[]]$Ids)
  $selected = @{}
  foreach ($id in $Ids) {
    if (-not $Library.ContainsKey($id)) { throw "符号库缺少符号：$id" }
    $selected[$id] = $Library[$id]
  }
  return $selected
}

function Invoke-P3Case {
  param([string]$Name, [string]$Kind, $Model, [string[]]$SymbolIds, [int]$ExpectedMinPages = 1, [int]$ExpectedMaxPages = 999)

  $symbols = Select-Symbols -Library $script:Library -Ids $SymbolIds
  $outputPath = Join-Path $OutDir ("{0}-smoke.vsdx" -f $Name)
  if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }

  $payload = @{
    kind = $Kind
    title = "冒烟：$Name"
    model = $Model
    symbols = $symbols
    outputPath = $outputPath
    validations = @()
  }
  $payloadPath = Join-Path $payloadDir ("{0}.json" -f $Name)
  [System.IO.File]::WriteAllText($payloadPath, ($payload | ConvertTo-Json -Depth 20), (New-Object System.Text.UTF8Encoding($false)))

  $command = "& '$EnginePath' -Json (Get-Content -LiteralPath '$payloadPath' -Raw -Encoding UTF8)"
  $stdout = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $command 2>&1
  $text = ($stdout | Out-String)
  $parsed = $null
  try { $parsed = $text | ConvertFrom-Json } catch { $parsed = $null }

  $record = [pscustomobject]@{
    Case = $Name
    Kind = $Kind
    Ok = $false
    Pages = 0
    Symbols = 0
    Shapes = 0
    Links = 0
    Groups = 0
    Bytes = 0
    Output = $outputPath
    Error = ''
  }
  if ($null -eq $parsed) {
    $record.Error = "引擎输出非 JSON：$($text.Trim())"
    return $record
  }
  if (-not $parsed.ok) {
    $record.Error = ("{0}: {1}" -f $parsed.error.code, $parsed.error.message)
    return $record
  }
  $record.Ok = $true
  $record.Pages = [int]$parsed.result.pages
  $record.Symbols = [int]$parsed.result.stats.symbols
  $record.Shapes = [int]$parsed.result.stats.shapes
  $record.Links = [int]$parsed.result.stats.links
  $record.Groups = [int]$parsed.result.stats.groups
  if (Test-Path -LiteralPath $parsed.result.outputPath) {
    $record.Bytes = (Get-Item -LiteralPath $parsed.result.outputPath).Length
  }
  if (-not (Test-Path -LiteralPath $parsed.result.outputPath)) {
    $record.Ok = $false
    $record.Error = '输出文件不存在'
  }
  elseif ($record.Bytes -le 0) {
    $record.Ok = $false
    $record.Error = '输出文件为空'
  }
  elseif ($record.Shapes -le 0) {
    $record.Ok = $false
    $record.Error = '图形数为 0'
  }
  elseif ($record.Pages -lt $ExpectedMinPages -or $record.Pages -gt $ExpectedMaxPages) {
    $record.Ok = $false
    $record.Error = ("页数 {0} 不在预期 [{1},{2}]" -f $record.Pages, $ExpectedMinPages, $ExpectedMaxPages)
  }
  return $record
}

# ── 五类最小模型（docs/P3-设计.md §4.4） ──
$protectionModel = @{
  kind = 'protection'
  title = '冒烟：110kV 线路保护电流回路'
  meta = @{ paper = 'A3-landscape'; standard = 'GB/T 4728-2022' }
  circuits = @(
    @{ id = 'C1'; label = '电流回路'; voltage = 'DC110V'; devices = @(
        @{ id = 'TAa'; symbolRef = 'gb4728.6.TA'; label = '1TAa' },
        @{ id = 'KA1'; symbolRef = 'gb4728.7.QL'; label = 'KA1' },
        @{ id = 'X1T'; symbolRef = 'gb4728.3.J'; label = 'X1:1' }
      ) },
    @{ id = 'C2'; label = '电压回路'; voltage = 'AC100V'; devices = @(
        @{ id = 'TVa'; symbolRef = 'gb4728.6.TV'; label = '1TVa' },
        @{ id = 'KA2'; symbolRef = 'gb4728.4.C'; label = 'KA2' }
      ) }
  )
  loops = @(
    @{ id = 'L1'; from = 'KA1'; to = 'KA2'; kind = 'signal'; label = 'KA1-1' }
  )
  terminalStrips = @(
    @{ id = 'XT1'; label = '端子排 1X'; side = 'top'; terminals = @(
        @{ no = 1; ref = 'A411' }, @{ no = 2; ref = 'B411' }, @{ no = 3; ref = 'N411' }
      ) },
    @{ id = 'XT2'; label = '端子排 2X'; side = 'bottom'; terminals = @(
        @{ no = 1; ref = 'A421' }, @{ no = 2; ref = 'B421' }
      ) }
  )
}

$substationPlanModel = @{
  kind = 'substationPlan'
  title = '冒烟：110kV 白云变总平面布置'
  meta = @{ paper = 'A3-landscape'; compass = $true }
  plot = @{ widthMm = 800; heightMm = 600; gridMm = 50 }
  zones = @(
    @{ id = 'Z1'; label = '110kV 配电装置区'; xMm = 0; yMm = 0; widthMm = 300; heightMm = 200 },
    @{ id = 'Z2'; label = '主变区'; xMm = 300; yMm = 0; widthMm = 200; heightMm = 200 },
    @{ id = 'Z3'; label = '综合楼及生活区'; xMm = 0; yMm = 200; widthMm = 500; heightMm = 200 }
  )
  items = @(
    @{ id = 'T1'; label = '1# 主变'; symbolRef = 'gb4728.6.T2'; xMm = 350; yMm = 60; rot = 0 },
    @{ id = 'RD1'; label = '进站道路'; xMm = 0; yMm = 400; widthMm = 500; heightMm = 24; primitive = 'rect' },
    @{ id = 'F1'; label = '站区围墙'; xMm = 0; yMm = 0; widthMm = 800; heightMm = 600; primitive = 'rect'; dashed = $true }
  )
  links = @(
    @{ id = 'L1'; from = 'T1'; to = 'RD1'; label = '进站道路' }
  )
}

$geoNetworkModel = @{
  kind = 'geoNetwork'
  title = '冒烟：××供电区 35kV 及以上地理接线'
  meta = @{ paper = 'A3-landscape' }
  viewport = @{ west = 112.0; south = 30.0; east = 114.0; north = 32.0 }
  stations = @(
    @{ id = 'S1'; label = '白云变'; symbolRef = 'gb4728.9.NODE'; lng = 113.2; lat = 31.1; voltage = 110 },
    @{ id = 'S2'; label = '金亭变'; symbolRef = 'gb4728.9.NODE'; lng = 113.7; lat = 31.5; voltage = 110 },
    @{ id = 'S3'; label = '城东变'; symbolRef = 'gb4728.9.NODE'; lng = 112.4; lat = 30.6; voltage = 35 }
  )
  lines = @(
    @{ id = 'LN1'; from = 'S1'; to = 'S2'; voltage = 110; label = '白-金 I 回'; points = @(
        @('S1', $null), @(113.5, 31.3), @('S2', $null)
      ) },
    @{ id = 'LN2'; from = 'S1'; to = 'S2'; voltage = 110; label = '白-金 II 回' },
    @{ id = 'LN3'; from = 'S1'; to = 'S3'; voltage = 35; label = '白-城 35kV' }
  )
  labels = @(
    @{ id = 'LBL1'; text = '××工业园'; lng = 113.1; lat = 30.9 }
  )
}

$powerFlowModel = @{
  kind = 'powerFlow'
  title = '冒烟：××电网 110kV 系统接线潮流'
  meta = @{ paper = 'A3-landscape' }
  levels = @(
    @{ id = 'LV110'; voltage = 110; label = '110kV' },
    @{ id = 'LV35'; voltage = 35; label = '35kV' }
  )
  buses = @(
    @{ id = 'B1'; level = 'LV110'; label = '白云 110kV 母线' },
    @{ id = 'B2'; level = 'LV35'; label = '白云 35kV 母线' }
  )
  nodes = @(
    @{ id = 'G1'; symbolRef = 'gb4728.6.G'; label = '金亭电厂'; level = 'LV110'; tags = @{ source = $true } },
    @{ id = 'S3'; symbolRef = 'gb4728.9.NODE'; label = '外部系统'; level = 'LV110'; tags = @{ source = $true } }
  )
  transformers = @(
    @{ id = 'T1'; symbolRef = 'gb4728.6.T2'; label = '白 1# 主变'; from = 'B1'; to = 'B2' }
  )
  flows = @(
    @{ id = 'F1'; from = 'B1'; to = 'B2'; kind = 'flow'; direction = 'B1→B2'; label = 'P=85.4MW  Q=12.1Mvar'; style = 'solid' },
    @{ id = 'F2'; from = 'G1'; to = 'B1'; kind = 'flow'; direction = 'G1→B1'; label = 'P=120MW'; style = 'dashed' }
  )
  links = @(
    @{ id = 'L1'; from = 'B1'; to = 'S3' }
  )
}

$terminalTerms = @()
for ($index = 1; $index -le 25; $index++) {
  $terminalType = 'normal'
  if ($index -eq 4) { $terminalType = 'test' }
  if ($index -eq 8) { $terminalType = 'ground' }
  if ($index -eq 12) { $terminalType = 'shorting' }
  $linkValue = $null
  if ($index -eq 12) { $linkValue = 11 }
  $terminalTerms += @{
    no = $index
    ref = ('A{0}' -f (400 + $index))
    desc = ('保护回路 {0}' -f $index)
    type = $terminalType
    link = $linkValue
  }
}
$terminalStripModel = @{
  kind = 'terminalStrip'
  title = '冒烟：端子排图 1X'
  meta = @{ paper = 'A4-landscape'; terminalsPerPage = 20 }
  strips = @(
    @{ id = 'XT1'; label = '端子排 1X（保护屏）'; columns = 1; terminals = $terminalTerms }
  )
}

$script:Library = Read-SymbolLibrary -Root $SymbolsDir
Write-Host ("[grid-p3-smoke] 符号库：{0} 个符号（{1}）" -f $script:Library.Count, $SymbolsDir)
Write-Host ("[grid-p3-smoke] 引擎：{0}" -f $EnginePath)
Write-Host ("[grid-p3-smoke] 输出目录：{0}" -f $OutDir)

$results = @()
$results += Invoke-P3Case -Name 'protection' -Kind 'protection' -Model $protectionModel -SymbolIds @('gb4728.6.TA', 'gb4728.7.QL', 'gb4728.3.J', 'gb4728.6.TV', 'gb4728.4.C')
$results += Invoke-P3Case -Name 'substationPlan' -Kind 'substationPlan' -Model $substationPlanModel -SymbolIds @('gb4728.6.T2')
$results += Invoke-P3Case -Name 'geoNetwork' -Kind 'geoNetwork' -Model $geoNetworkModel -SymbolIds @('gb4728.9.NODE')
$results += Invoke-P3Case -Name 'powerFlow' -Kind 'powerFlow' -Model $powerFlowModel -SymbolIds @('gb4728.6.G', 'gb4728.9.NODE', 'gb4728.6.T2')
$results += Invoke-P3Case -Name 'terminalStrip' -Kind 'terminalStrip' -Model $terminalStripModel -SymbolIds @() -ExpectedMinPages 2 -ExpectedMaxPages 2

$results | Format-Table Case, Kind, Ok, Pages, Symbols, Shapes, Links, Groups, Bytes, Error -AutoSize | Out-String -Width 240 | Write-Host

$failed = @($results | Where-Object { -not $_.Ok })
if ($failed.Count -gt 0) {
  Write-Host ("[grid-p3-smoke] 失败 {0}/{1}" -f $failed.Count, $results.Count)
  exit 1
}
Write-Host ("[grid-p3-smoke] 全部通过（{0} 个用例，输出 .vsdx）" -f $results.Count)
exit 0
