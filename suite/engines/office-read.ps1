[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Json)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
  $arg = $Json | ConvertFrom-Json
  $path = [string]$arg.path
  if (-not $path -or -not (Test-Path -LiteralPath $path)) {
    @{ ok = $false; error = @{ code = 'NOT_FOUND'; message = "文件不存在: $path" } } | ConvertTo-Json -Depth 6
    exit 2
  }
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

  if ($ext -in '.doc', '.docx') {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($path, $false, $true)
    try {
      $paras = $doc.Paragraphs
      $headings = @()
      $preview = New-Object System.Collections.Generic.List[string]
      $n = [Math]::Min($paras.Count, 200)
      for ($i = 1; $i -le $n; $i++) {
        $p = $paras.Item($i)
        $style = [string]$p.Range.Style.NameLocal
        $text = ($p.Range.Text -replace "[\r\n\x07]", '').Trim()
        if ($style -match '标题|Heading') { $headings += [ordered]@{ style = $style; text = $text } }
        if ($preview.Count -lt 60 -and $text.Length -gt 0) { $preview.Add($text) }
      }
      $info = [ordered]@{
        kind       = 'docx'
        paragraphs = $paras.Count
        tables     = $doc.Tables.Count
        headings   = $headings
        preview    = @($preview)
      }
      @{ ok = $true; result = @{ file = $path; info = $info } } | ConvertTo-Json -Depth 8
    } finally {
      try { $doc.Close($false) | Out-Null } catch { }
      try { $word.Quit() } catch { }
    }
  }
  elseif ($ext -in '.xls', '.xlsx') {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($path, 0, $true)
    try {
      $sheets = @()
      foreach ($ws in $wb.Worksheets) {
        $rows = 0; $cols = 0
        $ur = $ws.UsedRange
        if ($ur) { $rows = $ur.Rows.Count; $cols = $ur.Columns.Count }
        $sample = @()
        $rMax = [Math]::Min(5, $rows); $cMax = [Math]::Min(6, $cols)
        for ($r = 1; $r -le $rMax; $r++) {
          $row = @()
          for ($c = 1; $c -le $cMax; $c++) { $row += [string]$ws.Cells.Item($r, $c).Text }
          $sample += ($row -join ' | ')
        }
        $sheets += [ordered]@{ name = $ws.Name; rows = $rows; cols = $cols; sample = $sample }
      }
      $info = [ordered]@{ kind = 'xlsx'; sheetCount = $sheets.Count; sheets = $sheets }
      @{ ok = $true; result = @{ file = $path; info = $info } } | ConvertTo-Json -Depth 8
    } finally {
      try { $wb.Close($false) | Out-Null } catch { }
      try { $excel.Quit() } catch { }
    }
  }
  else {
    @{ ok = $false; error = @{ code = 'UNSUPPORTED_TYPE'; message = "暂不支持的文件类型: $ext" } } | ConvertTo-Json -Depth 6
    exit 2
  }
}
catch {
  @{ ok = $false; error = @{ code = 'ENGINE_FAIL'; message = $_.Exception.Message } } | ConvertTo-Json -Depth 6
  exit 1
}
