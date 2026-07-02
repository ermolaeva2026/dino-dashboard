param(
  [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)),
  [string]$RepoRoot = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))) 'Codex\dino-dashboard')
)

$ErrorActionPreference = 'Stop'

function Copy-ProjectFile([string]$SourceName, [string]$DestName) {
  $src = Join-Path $ProjectRoot $SourceName
  $dst = Join-Path $RepoRoot $DestName
  if (-not (Test-Path -LiteralPath $src)) { throw "Не найден файл для публикации: $src" }
  Copy-Item -LiteralPath $src -Destination $dst -Force
}

function Remove-LargeFiles([string]$Root, [int64]$MaxBytes) {
  Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Length -gt $MaxBytes } |
    ForEach-Object {
      Write-Host "Skip large file: $($_.FullName) ($([math]::Round($_.Length / 1MB, 1)) MB)"
      Remove-Item -LiteralPath $_.FullName -Force
    }
}

function Remove-LargeGalleryRefs([string]$HtmlPath, [string]$JsonPath) {
  if (Test-Path -LiteralPath $HtmlPath) {
    $html = Get-Content -LiteralPath $HtmlPath -Encoding UTF8 -Raw
    $m = [regex]::Match($html, '(?s)var GAL=(.*?);\s*var gC=')
    if ($m.Success) {
      $gal = $m.Groups[1].Value | ConvertFrom-Json
      foreach ($g in $gal) {
        $g.fls = @($g.fls | Where-Object { $_.t -ne 'vid' -and $_.url -notmatch '\.mp4($|[?#])' })
        $g.cnt = @($g.fls).Count
      }
      $newGal = $gal | ConvertTo-Json -Depth 40 -Compress
      $html = $html.Substring(0, $m.Groups[1].Index) + $newGal + $html.Substring($m.Groups[1].Index + $m.Groups[1].Length)
      [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $HtmlPath), $html, [System.Text.UTF8Encoding]::new($false))
    }
  }

  if (Test-Path -LiteralPath $JsonPath) {
    $data = Get-Content -LiteralPath $JsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
    if ($data.photos) {
      foreach ($g in $data.photos) {
        $g.fls = @($g.fls | Where-Object { $_.t -ne 'vid' -and $_.url -notmatch '\.mp4($|[?#])' })
        $g.cnt = @($g.fls).Count
      }
      [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $JsonPath), ($data | ConvertTo-Json -Depth 50 -Compress), [System.Text.UTF8Encoding]::new($false))
    }
  }
}

function ConvertTo-WebPhoto([string]$SourcePath, [string]$DestPath, [int]$MaxSide = 900, [int64]$Quality = 62) {
  Add-Type -AssemblyName System.Drawing
  $img = $null
  $bmp = $null
  $gfx = $null
  $encoderParams = $null
  try {
    $img = [System.Drawing.Image]::FromFile($SourcePath)
    $scale = [Math]::Min(1.0, $MaxSide / [double]([Math]::Max($img.Width, $img.Height)))
    $w = [Math]::Max(1, [int][Math]::Round($img.Width * $scale))
    $h = [Math]::Max(1, [int][Math]::Round($img.Height * $scale))
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $gfx.DrawImage($img, 0, 0, $w, $h)

    $dir = Split-Path -Parent $DestPath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, $Quality)
    $bmp.Save($DestPath, $codec, $encoderParams)
  }
  finally {
    if ($encoderParams) { $encoderParams.Dispose() }
    if ($gfx) { $gfx.Dispose() }
    if ($bmp) { $bmp.Dispose() }
    if ($img) { $img.Dispose() }
  }
}

function Publish-LightGallery([string]$HtmlPath, [string]$JsonPath, [string]$ProjectRoot, [string]$RepoRoot) {
  if (-not (Test-Path -LiteralPath $HtmlPath)) { return }
  $html = Get-Content -LiteralPath $HtmlPath -Encoding UTF8 -Raw
  $m = [regex]::Match($html, '(?s)var GAL=(.*?);\s*var gC=')
  if (-not $m.Success) { return }

  $lightRoot = Join-Path $RepoRoot 'photos-light'
  if (Test-Path -LiteralPath $lightRoot) { Remove-Item -LiteralPath $lightRoot -Recurse -Force }
  New-Item -ItemType Directory -Force -Path $lightRoot | Out-Null

  $gal = $m.Groups[1].Value | ConvertFrom-Json
  $converted = 0
  foreach ($g in $gal) {
    $newFiles = @()
    $fileIndex = 0
    foreach ($f in @($g.fls)) {
      if ($f.t -eq 'vid' -or $f.url -match '\.mp4($|[?#])') { continue }
      $relativeUrl = [System.Uri]::UnescapeDataString([string]$f.url).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
      $source = Join-Path $ProjectRoot $relativeUrl
      if (-not (Test-Path -LiteralPath $source)) { continue }
      if ([System.IO.Path]::GetExtension($source) -notmatch '^\.(jpg|jpeg|png)$') { continue }

      $groupDir = '{0:D2}' -f [int]$g.idx
      $destRel = ('photos-light/{0}/{1:D3}.jpg' -f $groupDir, $fileIndex)
      $dest = Join-Path $RepoRoot ($destRel.Replace('/', [System.IO.Path]::DirectorySeparatorChar))
      try {
        ConvertTo-WebPhoto -SourcePath $source -DestPath $dest
      }
      catch {
        Write-Host "Skip photo: $source ($($_.Exception.Message))"
        continue
      }
      $f.url = $destRel
      $f.t = 'img'
      $newFiles += $f
      $fileIndex++
      $converted++
    }
    $g.fls = @($newFiles)
    $g.cnt = @($newFiles).Count
  }

  $newGal = $gal | ConvertTo-Json -Depth 40 -Compress
  $html = $html.Substring(0, $m.Groups[1].Index) + $newGal + $html.Substring($m.Groups[1].Index + $m.Groups[1].Length)
  [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $HtmlPath), $html, [System.Text.UTF8Encoding]::new($false))

  if (Test-Path -LiteralPath $JsonPath) {
    $data = Get-Content -LiteralPath $JsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
    $data.photos = $gal
    [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $JsonPath), ($data | ConvertTo-Json -Depth 50 -Compress), [System.Text.UTF8Encoding]::new($false))
  }

  $bytes = (Get-ChildItem -LiteralPath $lightRoot -Recurse -File | Measure-Object Length -Sum).Sum
  Write-Host "Light gallery: $converted photo(s), $([math]::Round($bytes / 1MB, 1)) MB"
}

$updateScript = Join-Path $ProjectRoot 'Автообновление дашбордов\update-dashboards.ps1'
if (-not (Test-Path -LiteralPath $updateScript)) { throw "Не найден генератор дашбордов: $updateScript" }
if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot '.git'))) { throw "Не найден GitHub-репозиторий: $RepoRoot" }

Write-Host "Update local dashboards..."
& powershell -NoProfile -ExecutionPolicy Bypass -File $updateScript
if ($LASTEXITCODE -ne 0) { throw "Локальное обновление дашбордов завершилось с ошибкой $LASTEXITCODE" }

Push-Location $RepoRoot
try {
  Write-Host "Pull latest GitHub state..."
  & git pull --ff-only origin main
  if ($LASTEXITCODE -ne 0) { throw "git pull завершился с ошибкой $LASTEXITCODE" }

  Copy-ProjectFile 'PRK-2026-DS_Dashboard.html' 'index.html'
  Copy-ProjectFile 'PRK-2026-DS_Рабочий_дашборд.html' 'work.html'
  Copy-ProjectFile 'PRK-2026-DS_dashboard-data.json' 'PRK-2026-DS_dashboard-data.json'
  Set-Content -LiteralPath (Join-Path $RepoRoot '.nojekyll') -Value '' -Encoding ASCII

  $repoAutomation = Join-Path $RepoRoot 'Автообновление дашбордов'
  New-Item -ItemType Directory -Force -Path $repoAutomation | Out-Null
  Copy-Item -LiteralPath (Join-Path $ProjectRoot 'Автообновление дашбордов\README.md') -Destination (Join-Path $repoAutomation 'README.md') -Force
  Copy-Item -LiteralPath (Join-Path $ProjectRoot 'Автообновление дашбордов\update-dashboards.ps1') -Destination (Join-Path $repoAutomation 'update-dashboards.ps1') -Force
  Copy-Item -LiteralPath (Join-Path $ProjectRoot 'Автообновление дашбордов\update-dashboards.cmd') -Destination (Join-Path $repoAutomation 'update-dashboards.cmd') -Force
  Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $repoAutomation 'publish-to-github.ps1') -Force

  $repoPhotos = Join-Path $RepoRoot 'Фотофиксация работ'
  if (Test-Path -LiteralPath $repoPhotos) { Remove-Item -LiteralPath $repoPhotos -Recurse -Force }
  Remove-LargeGalleryRefs -HtmlPath (Join-Path $RepoRoot 'index.html') -JsonPath (Join-Path $RepoRoot 'PRK-2026-DS_dashboard-data.json')
  Publish-LightGallery -HtmlPath (Join-Path $RepoRoot 'index.html') -JsonPath (Join-Path $RepoRoot 'PRK-2026-DS_dashboard-data.json') -ProjectRoot $ProjectRoot -RepoRoot $RepoRoot

  & git add -A
  if ($LASTEXITCODE -ne 0) { throw "git add завершился с ошибкой $LASTEXITCODE" }

  $changes = & git status --porcelain
  if (-not $changes) {
    Write-Host "No GitHub changes to publish."
    exit 0
  }

  $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
  & git commit -m "Auto-update DinoPark dashboards $stamp"
  if ($LASTEXITCODE -ne 0) { throw "git commit завершился с ошибкой $LASTEXITCODE" }

  & git push origin main
  if ($LASTEXITCODE -ne 0) { throw "git push завершился с ошибкой $LASTEXITCODE" }
  & git push origin main:gh-pages
  if ($LASTEXITCODE -ne 0) { throw "git push gh-pages завершился с ошибкой $LASTEXITCODE" }

  Write-Host "Published to GitHub Pages: https://ermolaeva2026.github.io/dino-dashboard/"
}
finally {
  Pop-Location
}








