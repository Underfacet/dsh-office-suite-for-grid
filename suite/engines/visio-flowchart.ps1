[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Json)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
  $arg = $Json | ConvertFrom-Json
  $nodes = @($arg.nodes)
  $edges = @($arg.edges)
  $outPath = [string]$arg.outputPath
  if (-not $outPath) { throw 'outputPath 不能为空' }
  $outDir = Split-Path -Parent $outPath
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

  # 按有向边求深度（拓扑层），每层一列，列内垂直居中排布
  $depth = @{}; $indeg = @{}; $adj = @{}
  foreach ($n in $nodes) { $depth[[string]$n.id] = 0; $indeg[[string]$n.id] = 0; $adj[[string]$n.id] = @() }
  foreach ($e in $edges) {
    $f = [string]$e.from; $t = [string]$e.to
    if ($adj.ContainsKey($f) -and $indeg.ContainsKey($t)) { $adj[$f] += $t; $indeg[$t]++ }
  }
  $queue = New-Object System.Collections.Generic.Queue[string]
  foreach ($n in $nodes) { if ($indeg[[string]$n.id] -eq 0) { $queue.Enqueue([string]$n.id) } }
  while ($queue.Count -gt 0) {
    $u = $queue.Dequeue()
    foreach ($v in $adj[$u]) {
      if ($depth[$v] -lt $depth[$u] + 1) { $depth[$v] = $depth[$u] + 1; $queue.Enqueue($v) }
    }
  }
  $cols = @{}
  foreach ($n in $nodes) {
    $d = $depth[[string]$n.id]
    if (-not $cols.ContainsKey($d)) { $cols[$d] = @() }
    $cols[$d] += $n
  }
  $maxD = 0
  foreach ($k in $cols.Keys) { if ($k -gt $maxD) { $maxD = $k } }

  $visio = New-Object -ComObject Visio.Application
  $visio.Visible = $false
  try {
    $doc = $visio.Documents.Add('')
    $page = $visio.ActivePage
    $shapes = @{}
    foreach ($d in 0..$maxD) {
      $colNodes = @($cols[$d])
      $y0 = -($colNodes.Count - 1) * 1.5 / 2
      $x = 1.0 + $d * 3.8
      for ($i = 0; $i -lt $colNodes.Count; $i++) {
        $n = $colNodes[$i]
        $y = $y0 + $i * 1.5
        if ($n.kind -eq 'start' -or $n.kind -eq 'end') {
          $sh = $page.DrawOval($x, $y, $x + 1.6, $y + 0.7)
        }
        elseif ($n.kind -eq 'decision') {
          $sh = $page.DrawRectangle($x, $y, $x + 2.2, $y + 0.8)
        }
        else {
          $sh = $page.DrawRectangle($x, $y, $x + 2.6, $y + 0.8)
        }
        $sh.Text = [string]$n.label
        $shapes[[string]$n.id] = $sh
      }
    }
    foreach ($e in $edges) {
      $a = $shapes[[string]$e.from]; $b = $shapes[[string]$e.to]
      if (-not $a -or -not $b) { continue }
      $ax = $a.Cells('PinX').ResultIU; $ay = $a.Cells('PinY').ResultIU
      $aw = $a.Cells('Width').ResultIU / 2
      $bx = $b.Cells('PinX').ResultIU; $by = $b.Cells('PinY').ResultIU
      $bw = $b.Cells('Width').ResultIU / 2
      $line = $page.DrawLine($ax + $aw, $ay, $bx - $bw, $by)
      try { $line.Cells('EndArrow').FormulaU = '1' } catch { }
    }
    if ($arg.title) { $page.Name = [string]$arg.title }
    $doc.SaveAs($outPath)
    try { $doc.Close() } catch { }
    @{ ok = $true; result = @{ outputPath = $outPath; nodes = $nodes.Count; edges = $edges.Count; format = 'vsdx' } } | ConvertTo-Json -Depth 4
  } finally {
    try { $visio.Quit() } catch { }
  }
}
catch {
  @{ ok = $false; error = @{ code = 'ENGINE_FAIL'; message = $_.Exception.Message } } | ConvertTo-Json -Depth 4
  exit 1
}
