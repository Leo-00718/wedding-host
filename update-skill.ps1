# ============================================================
#  更新 Codex skill：wedding-host
#  规则：只增不删 —— 仓库里的文件会补上/更新；本地独有的文件绝不删、绝不改
#        覆盖前先全量备份；更新后列出“你本地独有的文件”和“被覆盖的文件”
# ============================================================
$ErrorActionPreference = "Stop"

# 规范化哈希：文本文件忽略换行符差异（避免把 CRLF/LF 当成本地改动）
function Get-NormHash([string]$path) {
  $ext = [System.IO.Path]::GetExtension($path).ToLower()
  if (@('.md','.txt','.py','.json','.ps1','.csv','.gitignore') -contains $ext) {
    $t = ([System.IO.File]::ReadAllText($path)) -replace "`r`n", "`n" -replace "`r", "`n"
    $b = [System.Text.Encoding]::UTF8.GetBytes($t)
  } else {
    $b = [System.IO.File]::ReadAllBytes($path)
  }
  $sha = [System.Security.Cryptography.SHA256]::Create()
  return [System.BitConverter]::ToString($sha.ComputeHash($b))
}
$dest  = "$env:USERPROFILE\.codex\skills\wedding-host"
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$tmp   = Join-Path $env:TEMP ("wh-update-" + $stamp)

Write-Host "[1/4] 下载最新 skill ..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$zip = Join-Path $tmp "wh.zip"
Invoke-WebRequest -UseBasicParsing -Uri "https://codeload.github.com/Leo-00718/wedding-host/zip/refs/heads/main" -OutFile $zip
Expand-Archive -LiteralPath $zip -DestinationPath $tmp -Force
$src = (Get-ChildItem -LiteralPath $tmp -Directory | Where-Object { $_.Name -like "wedding-host-*" } | Select-Object -First 1).FullName
if (-not $src) { Write-Host "★ 下载或解压失败，请检查网络" -ForegroundColor Red; return }

$isNew = -not (Test-Path -LiteralPath $dest)
$dest = if ($isNew) { $dest } else { (Resolve-Path -LiteralPath $dest).Path }   # 统一成长路径，避免 8.3 短名算错相对路径
$bak = $null
if ($isNew) {
  Write-Host "[2/4] 本地没有 skill -> 全新安装到 $dest" -ForegroundColor Yellow
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
} else {
  $bakRoot = Join-Path $env:USERPROFILE ".codex\skill-backups"
  New-Item -ItemType Directory -Force -Path $bakRoot | Out-Null
  $bak = Join-Path $bakRoot ("wedding-host." + $stamp)
  Write-Host "[2/4] 全量备份现有 skill -> $bak" -ForegroundColor Cyan
  Copy-Item -LiteralPath $dest -Destination $bak -Recurse -Force
}

Write-Host "[3/4] 合并（补上缺少的、更新已有的；本地独有的文件一个都不动）..." -ForegroundColor Cyan
$added = 0; $updated = 0
foreach ($f in (Get-ChildItem -LiteralPath $src -Recurse -File)) {
  $rel = $f.FullName.Substring($src.Length + 1)
  $dst = Join-Path $dest $rel
  $dir = Split-Path -Parent $dst
  if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
  if (Test-Path -LiteralPath $dst) { $updated++ } else { $added++ }
  Copy-Item -LiteralPath $f.FullName -Destination $dst -Force
}

Write-Host "[4/4] 完成" -ForegroundColor Green
Write-Host ("      新增 {0} 个文件，覆盖更新 {1} 个文件" -f $added, $updated)

if ($bak) {
  $relRepo = @{}
  $srcName = Split-Path -Leaf $src
  foreach ($f in (Get-ChildItem -LiteralPath $src -Recurse -File)) {
    $relRepo[$f.FullName.Substring($f.FullName.IndexOf($srcName) + $srcName.Length + 1)] = $true
  }
  $mine = @(); $changed = @()
  $bakName = Split-Path -Leaf $bak
  foreach ($f in (Get-ChildItem -LiteralPath $bak -Recurse -File)) {
    $rel = $f.FullName.Substring($f.FullName.IndexOf($bakName) + $bakName.Length + 1)
    $now = Join-Path $dest $rel
    if (-not $relRepo.ContainsKey($rel)) { $mine += $rel; continue }
    if ((Test-Path -LiteralPath $now) -and (Get-NormHash $f.FullName) -ne (Get-NormHash $now)) { $changed += $rel }
  }
  Write-Host ""
  Write-Host "  【你本地独有的文件】—— 已原样保留：" -ForegroundColor Green
  if ($mine.Count) { $mine | Select-Object -First 30 | ForEach-Object { Write-Host ("      + " + $_) } } else { Write-Host "      （无）" }
  Write-Host "  【本地有真改动、被仓库版本覆盖的文件】—— 原版都在备份里：" -ForegroundColor Yellow
  if ($changed.Count) { $changed | Select-Object -First 30 | ForEach-Object { Write-Host ("      ~ " + $_) } } else { Write-Host "      （无 —— 你本地没有文件内容被真正改动）" }
  Write-Host ("      回退方式：把备份 $bak 改名回 wedding-host 即可")
}

Write-Host ""
Write-Host "—— 换机自检 ——" -ForegroundColor Cyan
Write-Host ("  1) 装依赖：pip install -r `"$dest\requirements.txt`"")
$fnt = Get-ChildItem -LiteralPath "$env:USERPROFILE\Desktop" -Recurse -File -ErrorAction SilentlyContinue |
       Where-Object { $_.Name -like "*瘦金*" -and $_.Extension -match "ttf|otf" } | Select-Object -First 1
if ($fnt) { Write-Host ("  2) 瘦金字体：桌面已有 -> " + $fnt.FullName) -ForegroundColor Green }
else { Write-Host  "  2) 瘦金字体：桌面没找到！请把「瘦金简体.ttf」放到桌面（或放进 skill 的 assets\fonts\）" -ForegroundColor Yellow }
Write-Host  "  3) 首次出图需联网一次（自动下载抠图模型 ~179MB）"
Write-Host ""
Write-Host "  验证出图：" -ForegroundColor Cyan
Write-Host ("     python `"$dest\scripts\build_poster.py`" `"<你的照片.jpg>`" --name 测试 --cohort 第34期学员 --out `"$env:TEMP\测试.png`"")

