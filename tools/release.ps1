#requires -Version 5.1
<#
.SYNOPSIS
  dsh-office-suite-for-grid 一键发布脚本（P4，八步幂等）。

.DESCRIPTION
  步骤：
    1 版本参数校验（package.json / SemVer / tag / 工作区）
    2 测试（node --test；-SkipTest 跳过）
    3 生成市场索引（tools/build-index.mjs → dist/index.json + docs-site/）
    4 npm pack（suite → dist/<name>-<version>.tgz）
    5 SHA256SUMS.txt
    6 打 tag（release-<version>）
    7 推送 tag
    8 GitHub Release（gh）+ Pages 提示

  幂等：tag 已存在且指向 HEAD → 跳过；Release 已存在 → 覆盖上传资产。
  DryRun：只做前置检查并打印每条命令，不执行测试 / 打包 / tag / push / gh。
  npm publish 不在本脚本内（需要 NPM_TOKEN，且发布不可回滚，由维护者确认后手工执行）。

.PARAMETER Version
  发布版本，必须与 suite/package.json 的 version 完全一致（本脚本不改 package.json）。

.PARAMETER Tag
  发布 tag，默认 release-<version>（release.yml 由 release-* 触发）。

.PARAMETER Message
  tag / Release 说明，默认 "Release <version>"。

.PARAMETER SkipTest
  跳过第 2 步测试（不建议用于正式发布）。

.PARAMETER SkipPages
  第 3 步不生成 docs-site/（仅构建 dist/index.json）。

.PARAMETER SkipAssets
  第 8 步不调用 gh，只打印手工上传指引。

.PARAMETER DryRun
  全链路只读演练：打印命令与前置检查结果。

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File tools\release.ps1 -DryRun

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File tools\release.ps1 -Version 1.0.0
#>
[CmdletBinding()]
param(
  [string]$Version,
  [string]$Tag,
  [string]$Message,
  [switch]$SkipTest,
  [switch]$SkipPages,
  [switch]$SkipAssets,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SuiteDir = Join-Path $RepoRoot 'suite'
$DistDir = Join-Path $RepoRoot 'dist'
$DocsSiteDir = Join-Path $RepoRoot 'docs-site'
$PackageJson = Join-Path $SuiteDir 'package.json'
$IndexJson = Join-Path $DistDir 'index.json'
$SumsFile = Join-Path $DistDir 'SHA256SUMS.txt'
$PythonZip = Join-Path $DistDir 'python-3.12.10-embed-amd64.zip'
$PagesIndexUrl = 'https://underfacet.github.io/dsh-office-suite-for-grid/index.json'
$SemverPattern = '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$'

function Write-Step {
  param([int]$Number, [string]$Text)
  Write-Host ''
  Write-Host ('[{0}/8] {1}' -f $Number, $Text) -ForegroundColor Cyan
}

function Write-DryRun {
  param([string]$Command)
  Write-Host ('  [DryRun] {0}' -f $Command) -ForegroundColor DarkYellow
}

function Find-CommandPath {
  param([string]$Name)
  $command = Get-Command $Name -ErrorAction SilentlyContinue
  if ($null -eq $command) { return $null }
  return $command.Source
}

function Invoke-Checked {
  param([string]$FilePath, [string[]]$Arguments)
  & $FilePath @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw ('命令失败（exit {0}）：{1} {2}' -f $LASTEXITCODE, $FilePath, ($Arguments -join ' '))
  }
}

try {
  Write-Host '==== dsh-office-suite-for-grid 发布（P4） ====' -ForegroundColor Green
  if ($DryRun) { Write-Host '模式：DryRun（只读演练，不执行任何发布动作）' -ForegroundColor Yellow }

  # ── 1. 版本参数校验与前置检查 ───────────────────────────────────────────
  Write-Step 1 '版本参数校验与前置检查'
  if (-not (Test-Path -LiteralPath $PackageJson)) { throw "找不到 $PackageJson" }
  $pkg = [System.IO.File]::ReadAllText($PackageJson, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
  $pkgName = [string]$pkg.name
  $pkgVersion = [string]$pkg.version
  if ([string]::IsNullOrWhiteSpace($pkgName)) { throw 'suite/package.json 缺少 name' }
  if ([string]::IsNullOrWhiteSpace($pkgVersion)) { throw 'suite/package.json 缺少 version' }

  if ([string]::IsNullOrWhiteSpace($Version)) { $Version = $pkgVersion }
  if ($Version -notmatch $SemverPattern) { throw "-Version 不是精确 SemVer：$Version" }
  if ($Version -ne $pkgVersion) {
    throw "suite/package.json 的 version 是 $pkgVersion，与 -Version $Version 不一致。请先手工升版（本脚本不改 suite\package.json），再重跑。"
  }

  if ([string]::IsNullOrWhiteSpace($Tag)) { $Tag = "release-$Version" }
  if ($Tag -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { throw "-Tag 含非法字符：$Tag" }
  if (-not $Tag.StartsWith('release-')) {
    Write-Warning "tag '$Tag' 不以 release- 开头；.github/workflows/release.yml 只监听 release-*，不会自动发版。"
  }
  if ([string]::IsNullOrWhiteSpace($Message)) { $Message = "Release $Version" }

  $git = Find-CommandPath 'git'
  $node = Find-CommandPath 'node'
  $npm = Find-CommandPath 'npm.cmd'
  if ($null -eq $npm) { $npm = Find-CommandPath 'npm' }
  $gh = Find-CommandPath 'gh'

  $missing = @()
  if ($null -eq $git) { $missing += 'git' }
  if ($null -eq $node) { $missing += 'node' }
  if ($null -eq $npm) { $missing += 'npm' }
  if ($missing.Count -gt 0) { throw ('缺少命令：{0}（请先安装并加入 PATH）' -f ($missing -join ', ')) }
  if ($null -eq $gh) { Write-Warning '未找到 gh：第 8 步会打印手工创建 Release 的指引（不阻塞打包与 Pages 提示）。' }

  $null = & $git -C $RepoRoot rev-parse --show-toplevel 2>&1
  if ($LASTEXITCODE -ne 0) { throw "$RepoRoot 不是 git 仓库" }

  $dirty = @(& $git -C $RepoRoot status --porcelain)
  if ($LASTEXITCODE -ne 0) { throw 'git status 失败' }
  if ($dirty.Count -gt 0) {
    $dirtyText = $dirty -join "`n"
    if ($DryRun) {
      Write-Warning "git 工作区不干净（DryRun 继续）：`n$dirtyText"
    } else {
      throw "git 工作区不干净，请先提交或暂存：`n$dirtyText"
    }
  }

  $epoch = (& $git -C $RepoRoot show -s --format=%ct HEAD).Trim()
  if ($LASTEXITCODE -eq 0 -and $epoch -match '^\d+$') {
    $env:SOURCE_DATE_EPOCH = $epoch
    Write-Host "SOURCE_DATE_EPOCH=$epoch（tag 提交时间，保证 Pages 与 Release 索引一致）"
  } else {
    Write-Warning '无法读取 HEAD 提交时间；build-index 将使用当前时间作为 publishTimestamp。'
  }

  $tgz = Join-Path $DistDir ("{0}-{1}.tgz" -f $pkgName, $Version)
  Write-Host ("包：{0}@{1}" -f $pkgName, $Version)
  Write-Host ("tag：{0}" -f $Tag)
  Write-Host ("输出：{0}" -f $DistDir)

  # ── 2. 测试 ─────────────────────────────────────────────────────────────
  Write-Step 2 '测试（node --test）'
  if ($SkipTest) {
    Write-Host '  跳过：-SkipTest'
  } elseif ($DryRun) {
    Write-DryRun 'node --test "tests/**/*.test.js"'
  } else {
    Push-Location $RepoRoot
    try {
      Invoke-Checked $node @('--test', 'tests/**/*.test.js')
    } finally {
      Pop-Location
    }
    Write-Host '  测试通过'
  }

  # ── 3. 生成市场索引 ─────────────────────────────────────────────────────
  Write-Step 3 '生成市场索引（tools/build-index.mjs）'
  $buildArgs = @('tools\build-index.mjs', '--out', 'dist\index.json')
  if (-not $SkipPages) { $buildArgs += @('--pages', 'docs-site') }
  if ($DryRun) {
    Write-DryRun ('$env:SOURCE_DATE_EPOCH={0}; node {1}' -f $env:SOURCE_DATE_EPOCH, ($buildArgs -join ' '))
    if (-not $SkipPages) {
      Write-DryRun "生成 $DocsSiteDir\index.json、icon.svg、index.html"
    }
  } else {
    Push-Location $RepoRoot
    try {
      Invoke-Checked $node $buildArgs
    } finally {
      Pop-Location
    }
    if (-not (Test-Path -LiteralPath $IndexJson)) { throw "build-index 未产出 $IndexJson" }
    Write-Host "  index.json：$IndexJson"
  }

  # ── 4. npm pack ─────────────────────────────────────────────────────────
  Write-Step 4 'npm pack（suite → dist）'
  if ($DryRun) {
    Write-DryRun ("npm pack --pack-destination `"{0}`"（工作目录 {1}）" -f $DistDir, $SuiteDir)
    Write-DryRun ("期望产物：{0}" -f $tgz)
  } else {
    New-Item -ItemType Directory -Force -Path $DistDir | Out-Null
    Push-Location $SuiteDir
    try {
      Invoke-Checked $npm @('pack', '--pack-destination', $DistDir)
    } finally {
      Pop-Location
    }
    if (-not (Test-Path -LiteralPath $tgz)) { throw "npm pack 未产出 $tgz" }
    $size = (Get-Item -LiteralPath $tgz).Length
    Write-Host ("  tgz：{0}（{1} bytes）" -f $tgz, $size)
  }

  # ── 5. SHA256SUMS ───────────────────────────────────────────────────────
  Write-Step 5 '生成 SHA256SUMS.txt'
  $assetPaths = @($tgz, $IndexJson)
  if (Test-Path -LiteralPath $PythonZip) {
    $assetPaths += $PythonZip
  } else {
    Write-Host '  [TODO] 便携 Python 资产未下载（P4 首版占位；见 market-entry.json assets.pythonEmbed）。'
  }
  if ($DryRun) {
    foreach ($path in $assetPaths) { Write-DryRun ("Get-FileHash -Algorithm SHA256 `"{0}`"" -f $path) }
    Write-DryRun ("写出 {0}" -f $SumsFile)
  } else {
    $lines = @()
    foreach ($path in $assetPaths) {
      if (-not (Test-Path -LiteralPath $path)) { throw "缺少资产：$path" }
      $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash.ToLowerInvariant()
      $lines += ('{0}  {1}' -f $hash, (Split-Path -Leaf $path))
    }
    Set-Content -LiteralPath $SumsFile -Value $lines -Encoding ASCII
    Write-Host ("  SHA256SUMS.txt：{0} 项" -f $lines.Count)
    $lines | ForEach-Object { Write-Host "    $_" }
  }

  # ── 6. 打 tag ───────────────────────────────────────────────────────────
  Write-Step 6 ("打 tag {0}" -f $Tag)
  if ($DryRun) {
    Write-DryRun ("git -C `"{0}`" tag -a `"{1}`" -m `"{2}`"" -f $RepoRoot, $Tag, $Message)
  } else {
    $null = & $git -C $RepoRoot rev-parse -q --verify "refs/tags/$Tag" 2>&1
    $tagExists = ($LASTEXITCODE -eq 0)
    if ($tagExists) {
      $tagCommit = (& $git -C $RepoRoot rev-list -n 1 $Tag).Trim()
      $headCommit = (& $git -C $RepoRoot rev-parse HEAD).Trim()
      if ($tagCommit -eq $headCommit) {
        Write-Host "  tag 已存在且指向 HEAD（$headCommit）：幂等跳过"
      } else {
        throw "tag $Tag 已存在但指向 $tagCommit（HEAD=$headCommit）；请改用新版本号或手工处理旧 tag。"
      }
    } else {
      Invoke-Checked $git @('-C', $RepoRoot, 'tag', '-a', $Tag, '-m', $Message)
      Write-Host "  已创建 tag $Tag"
    }
  }

  # ── 7. 推送 tag ─────────────────────────────────────────────────────────
  Write-Step 7 '推送 tag 到 origin'
  if ($DryRun) {
    Write-DryRun ("git -C `"{0}`" push origin `"{1}`"" -f $RepoRoot, $Tag)
  } else {
    Invoke-Checked $git @('-C', $RepoRoot, 'push', 'origin', $Tag)
    Write-Host '  已推送（若远端同 tag 已指向同一提交，git 会输出 up-to-date，幂等）'
  }

  # ── 8. GitHub Release + Pages 提示 ──────────────────────────────────────
  Write-Step 8 'GitHub Release 与 Pages'
  $releaseAssets = @($tgz, $IndexJson, $SumsFile)
  if (Test-Path -LiteralPath $PythonZip) { $releaseAssets += $PythonZip }
  if ($DryRun) {
    Write-DryRun ("gh release create {0} --title `"Release {1}`" --notes `"...`" {2}" -f $Tag, $Version, (($releaseAssets | ForEach-Object { '"{0}"' -f $_ }) -join ' '))
  } elseif ($SkipAssets) {
    Write-Host '  跳过 gh：-SkipAssets。可手工上传：'
    $releaseAssets | ForEach-Object { Write-Host "    - $_" }
  } elseif ($null -eq $gh) {
    Write-Warning 'gh 未安装：请在 GitHub Web 创建 Release 并上传以下资产：'
    $releaseAssets | ForEach-Object { Write-Host "    - $_" }
  } else {
    $null = & $gh release view $Tag 2>&1
    if ($LASTEXITCODE -eq 0) {
      Write-Host "  Release $Tag 已存在：覆盖上传资产（幂等）"
      Invoke-Checked $gh (@('release', 'upload', $Tag) + $releaseAssets + @('--clobber'))
    } else {
      Invoke-Checked $gh (@('release', 'create', $Tag, '--title', "Release $Version", '--notes', $Message) + $releaseAssets)
    }
    Write-Host "  Release：https://github.com/Underfacet/dsh-office-suite-for-grid/releases/tag/$Tag"
  }

  if (-not $SkipPages) {
    Write-Host '  Pages：.github/workflows/release.yml 的 pages job 将把 docs-site\ 发布到 gh-pages。'
    Write-Host '  首次使用需在仓库 Settings → Pages 选择 Deploy from a branch → gh-pages / (root)。'
    Write-Host "  索引 URL：$PagesIndexUrl"
  } else {
    Write-Host '  Pages：已跳过（-SkipPages）。'
  }

  # ── 发布报告（不计入八步） ─────────────────────────────────────────────
  Write-Host ''
  Write-Host '发布报告（JSON）：' -ForegroundColor Cyan
  $sha = [ordered]@{}
  if (-not $DryRun) {
    foreach ($path in $releaseAssets) {
      if (Test-Path -LiteralPath $path) {
        $sha[(Split-Path -Leaf $path)] = (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash.ToLowerInvariant()
      }
    }
  }
  $report = [ordered]@{
    ok          = $true
    dryRun      = [bool]$DryRun
    version     = $Version
    tag         = $Tag
    package     = $pkgName
    npm         = [ordered]@{
      published = $false
      note      = '未执行 npm publish：需要 NPM_TOKEN，且发布不可回滚；请在确认后手工执行：cd suite; npm publish --access public'
    }
    pages       = [ordered]@{
      indexUrl   = $PagesIndexUrl
      publishDir = $(if ($SkipPages) { $null } else { 'docs-site' })
      hint       = 'release.yml pages job 部署 gh-pages；首次需在 Settings → Pages 选择 gh-pages/(root)'
    }
    assets      = @($releaseAssets | ForEach-Object { Split-Path -Leaf $_ })
    sha256      = $sha
    generatedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
  }
  $report | ConvertTo-Json -Depth 6
} catch {
  Write-Host ''
  Write-Host ('发布失败：{0}' -f $_.Exception.Message) -ForegroundColor Red
  exit 1
}
