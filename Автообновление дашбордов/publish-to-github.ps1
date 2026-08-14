param(
  [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)),
  [string]$RepoRoot = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))) 'Codex\dino-dashboard'),
  [switch]$UseExistingKsgSnapshot
)

$ErrorActionPreference = 'Stop'

try {
  $publishMutex = New-Object System.Threading.Mutex($false, 'Global\PRK_2026_DS_DASHBOARD_PUBLISH')
} catch {
  $publishMutex = New-Object System.Threading.Mutex($false, 'PRK_2026_DS_DASHBOARD_PUBLISH')
}
$publishMutexHeld = $false
try { $publishMutexHeld = $publishMutex.WaitOne(0) }
catch [System.Threading.AbandonedMutexException] { $publishMutexHeld = $true }
if (-not $publishMutexHeld) {
  Write-Host 'Skip: another DinoPark dashboard publication is already running.'
  exit 0
}

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

function ConvertTo-WebPhoto([string]$SourcePath, [string]$DestPath, [int]$MaxSide = 700, [int64]$Quality = 50) {
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

function ConvertTo-WebPhotoDataUri([string]$SourcePath) {
  $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("dino-photo-" + [guid]::NewGuid().ToString("N") + ".jpg")
  try {
    ConvertTo-WebPhoto -SourcePath $SourcePath -DestPath $tmp
    $bytes = [System.IO.File]::ReadAllBytes($tmp)
    return 'data:image/jpeg;base64,' + [Convert]::ToBase64String($bytes)
  }
  finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
  }
}

function Publish-LightGallery([string]$HtmlPath, [string]$JsonPath, [string]$ProjectRoot, [string]$RepoRoot) {
  if (-not (Test-Path -LiteralPath $HtmlPath)) { return }
  $html = Get-Content -LiteralPath $HtmlPath -Encoding UTF8 -Raw
  $m = [regex]::Match($html, '(?s)var GAL=(.*?);\s*var gC=')
  if (-not $m.Success) { return }

  $lightRoot = Join-Path $RepoRoot 'photos-light'
  if (Test-Path -LiteralPath $lightRoot) { Remove-Item -LiteralPath $lightRoot -Recurse -Force }

  $gal = $m.Groups[1].Value | ConvertFrom-Json
  $converted = 0
  $totalBytes = 0
  foreach ($g in $gal) {
    $newFiles = @()
    foreach ($f in @($g.fls)) {
      if ($f.t -eq 'vid' -or $f.url -match '\.mp4($|[?#])') { continue }
      $relativeUrl = [System.Uri]::UnescapeDataString([string]$f.url).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
      $source = Join-Path $ProjectRoot $relativeUrl
      if (-not (Test-Path -LiteralPath $source)) { continue }
      if ([System.IO.Path]::GetExtension($source) -notmatch '^\.(jpg|jpeg|png)$') { continue }

      try {
        $dataUri = ConvertTo-WebPhotoDataUri -SourcePath $source
      }
      catch {
        Write-Host "Skip photo: $source ($($_.Exception.Message))"
        continue
      }
      $f.url = $dataUri
      $f.t = 'img'
      $newFiles += $f
      $converted++
      $totalBytes += [int64]([Math]::Floor(($dataUri.Length * 3) / 4))
    }
    $g.fls = @($newFiles)
    $g.cnt = @($newFiles).Count
  }

  $newGal = $gal | ConvertTo-Json -Depth 40 -Compress
  $html = $html.Substring(0, $m.Groups[1].Index) + $newGal + $html.Substring($m.Groups[1].Index + $m.Groups[1].Length)
  [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $HtmlPath), $html, [System.Text.UTF8Encoding]::new($false))

  if (Test-Path -LiteralPath $JsonPath) {
    $data = Get-Content -LiteralPath $JsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
    $jsonPhotos = @()
    foreach ($g in $gal) {
      $jsonPhotos += [ordered]@{ icon=$g.icon; cat=$g.cat; phase=$g.phase; cnt=$g.cnt; fls=@(); idx=$g.idx }
    }
    $data.photos = $jsonPhotos
    [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $JsonPath), ($data | ConvertTo-Json -Depth 50 -Compress), [System.Text.UTF8Encoding]::new($false))
  }

  Write-Host "Light gallery: $converted inline photo(s), about $([math]::Round($totalBytes / 1MB, 1)) MB"
}

function Get-PmoChipClass($Phase) {
  if ($Phase.pct -ge 80) { return 'green' }
  if ($Phase.pct -ge 50) { return 'amber' }
  if ($Phase.pct -gt 0) { return 'blue' }
  return ''
}

function Format-PmoPhaseName([string]$Name) {
  $n = $Name
  $n = $n -replace '^Ф\.(\d+)\s+', 'Ф$1. '
  $n = $n -replace 'Строительно-ремонтные работы', 'СМР'
  $n = $n -replace 'Ремонт и благоустройство детских зон и беседок', 'Детские зоны'
  return $n
}

function Format-PmoMillion($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '—' }
  $ru = [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')
  return (([double]$Value / 1000000).ToString('N2', $ru) + ' млн')
}

function Update-PmoDinoSection([string]$PmoPath, [string]$JsonPath) {
  if (-not (Test-Path -LiteralPath $PmoPath) -or -not (Test-Path -LiteralPath $JsonPath)) { return }

  $html = Get-Content -LiteralPath $PmoPath -Encoding UTF8 -Raw
  $data = Get-Content -LiteralPath $JsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
  if (-not $data.tasks -or -not $data.phases) { return }

  $tasks = @($data.tasks)
  $phases = @($data.phases)
  $done = @($tasks | Where-Object { [int]$_.pct -ge 100 }).Count
  $active = @($tasks | Where-Object { [int]$_.pct -gt 0 -and [int]$_.pct -lt 100 }).Count
  $total = $tasks.Count
  $remaining = [Math]::Max(0, $total - $done - $active)
  $progress = 0
  if ($data.metrics) {
    $metric = @($data.metrics | Where-Object { $_.lbl -eq 'Прогресс проекта' } | Select-Object -First 1)
    if ($metric -and ([string]$metric.val) -match '(\d+)') { $progress = [int]$matches[1] }
  }
  if ($progress -eq 0 -and $total -gt 0) {
    $progress = [math]::Round((($tasks | Measure-Object -Property pct -Average).Average), 0)
  }

  $generated = [datetime]$data.generated_at
  $statusDate = $generated.ToString('dd.MM.yyyy')
  $buildText = $generated.ToString('dd.MM.yyyy HH:mm')
  $ksgText = ([datetime]$data.ksg_updated).ToString('dd.MM.yyyy HH:mm')

  $activePhases = @(
    $phases |
      Where-Object { [int]$_.pct -gt 0 -and [int]$_.pct -lt 100 } |
      ForEach-Object {
        '<span class="chip ' + (Get-PmoChipClass $_) + '">' +
        [System.Net.WebUtility]::HtmlEncode((Format-PmoPhaseName $_.nm) + ' ' + $_.pct + '%') +
        '</span>'
      }
  ) -join "`r`n            "

  $workItems = @(
    $tasks |
      Where-Object { [int]$_.pct -gt 0 -and [int]$_.pct -lt 100 } |
      Sort-Object @{ Expression = { [int]$_.pct } } |
      Select-Object -First 3 |
      ForEach-Object {
        '<div class="work"><div>' + [System.Net.WebUtility]::HtmlEncode($_.nm) +
        '</div><div></div><div class="pct">' + [int]$_.pct + '%</div></div>'
      }
  ) -join "`r`n            "

  $needsCount = 0
  if ($data.needs_action) { $needsCount = @($data.needs_action).Count }

  $budgetRow = $null
  $budgetSource = $null
  if ($data.budget) {
    $budgetRow = '<div class="budget-row">' +
      '<div class="budget-cell"><strong>' + (Format-PmoMillion $data.budget.approved_with_reserve) + '</strong><span>Лимит с резервом</span></div>' +
      '<div class="budget-cell warn"><strong>' + (Format-PmoMillion $data.budget.confirmed) + '</strong><span>Подтверждено</span></div>' +
      '<div class="budget-cell good"><strong>' + (Format-PmoMillion $data.budget.paid) + '</strong><span>Оплачено</span></div>' +
      '</div>'
    $budgetUpdated = ([datetime]$data.budget.updated_at).ToString('dd.MM.yyyy HH:mm')
    $budgetSource = 'Источник: PRK-2026-DS_dashboard-data.json, КСГ ' + $ksgText + ' · автосборка ' + $buildText + '; бюджет: первая вкладка «Смета расходов Зерно» от ' + $budgetUpdated + '.'
  }

  $cardMatch = [regex]::Match($html, '(?s)<article class="card">\s*<div class="card-head">(?:(?!</article>).)*?PRK-2026-DS(?:(?!</article>).)*?</article>')
  if (-not $cardMatch.Success) {
    Write-Host "CEO PMO dashboard: DinoPark card not found, pmo.html was not changed."
    return
  }
  $card = $cardMatch.Value

  $html = [regex]::new('Статус на \d{2}\.\d{2}\.\d{4}').Replace($html, 'Статус на ' + $statusDate, 1)
  $card = [regex]::new('PRK-2026-DS · КСГ автосборка \d{2}\.\d{2}\.\d{4}').Replace($card, 'PRK-2026-DS · КСГ автосборка ' + $statusDate, 1)
  $card = [regex]::new('<div class="ring" style="--pct:\d+"><strong>\d+%</strong></div>').Replace($card, '<div class="ring" style="--pct:' + $progress + '"><strong>' + $progress + '%</strong></div>', 1)
  $card = [regex]::new('<div class="task"><strong>\d+</strong><span>Задач всего</span></div>').Replace($card, '<div class="task"><strong>' + $total + '</strong><span>Задач всего</span></div>', 1)
  $card = [regex]::new('<div class="task"><strong>\d+</strong><span>Выполнено</span></div>').Replace($card, '<div class="task"><strong>' + $done + '</strong><span>Выполнено</span></div>', 1)
  $card = [regex]::new('<div class="task"><strong>\d+</strong><span>В работе</span></div>').Replace($card, '<div class="task"><strong>' + $active + '</strong><span>В работе</span></div>', 1)
  $card = [regex]::new('<div class="task"><strong>\d+</strong><span>Остаток</span></div>').Replace($card, '<div class="task"><strong>' + $remaining + '</strong><span>Остаток</span></div>', 1)
  $card = [regex]::new('(?s)<div class="chips">\s*<span class="chip [^"]*">Ф3\..*?</div>').Replace($card, '<div class="chips">' + "`r`n            " + $activePhases + "`r`n          </div>", 1)
  $card = [regex]::new('<div class="work-count">· \d+ \+ \d+</div>').Replace($card, '<div class="work-count">· ' + $active + ' + ' + $needsCount + '</div>', 1)
  $card = [regex]::new('(?s)<div class="work-list">.*?</div>\s*</div>\s*<div class="source">').Replace($card, '<div class="work-list">' + "`r`n            " + $workItems + "`r`n          </div>`r`n        </div>`r`n        " + '<div class="source">', 1)
  if ($budgetRow) {
    $card = [regex]::new('(?s)<div class="budget-row">\s*(?:<div class="budget-cell[^"]*">.*?</div>\s*){3}</div>').Replace($card, $budgetRow, 1)
    $card = [regex]::new('(?s)<div class="source">Источник: PRK-2026-DS_dashboard-data\.json.*?</div>').Replace($card, '<div class="source">' + [System.Net.WebUtility]::HtmlEncode($budgetSource) + '</div>', 1)
  } else {
    $card = [regex]::new('Источник: PRK-2026-DS_dashboard-data\.json, [^;]+; бюджет').Replace($card, 'Источник: PRK-2026-DS_dashboard-data.json, КСГ ' + $ksgText + ' · автосборка ' + $buildText + '; бюджет', 1)
  }

  $html = $html.Substring(0, $cardMatch.Index) + $card + $html.Substring($cardMatch.Index + $cardMatch.Length)
  [System.IO.File]::WriteAllText((Resolve-Path -LiteralPath $PmoPath), $html, [System.Text.UTF8Encoding]::new($false))
  Write-Host "CEO PMO dashboard synced with KSG progress: $progress% ($done/$total done)"
}

$updateScript = Join-Path $ProjectRoot 'Автообновление дашбордов\update-dashboards.ps1'
if (-not (Test-Path -LiteralPath $updateScript)) { throw "Не найден генератор дашбордов: $updateScript" }
if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot '.git'))) { throw "Не найден GitHub-репозиторий: $RepoRoot" }

Write-Host "Update local dashboards..."
if ($UseExistingKsgSnapshot) {
  & powershell -NoProfile -ExecutionPolicy Bypass -File $updateScript -UseExistingKsgSnapshot
}
else {
  & powershell -NoProfile -ExecutionPolicy Bypass -File $updateScript
}
if ($LASTEXITCODE -ne 0) { throw "Локальное обновление дашбордов завершилось с ошибкой $LASTEXITCODE" }

$workspaceRoot = Split-Path -Parent $ProjectRoot
$localPmoPath = Join-Path $workspaceRoot 'PMO-05_04_CEO_Project_Blocks_Dashboard.html'
Update-PmoDinoSection -PmoPath $localPmoPath -JsonPath (Join-Path $ProjectRoot 'PRK-2026-DS_dashboard-data.json')

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
  Update-PmoDinoSection -PmoPath (Join-Path $RepoRoot 'pmo.html') -JsonPath (Join-Path $RepoRoot 'PRK-2026-DS_dashboard-data.json')

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









