param(
  [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'

function Join-ProjectPath([string]$Child) {
  Join-Path -Path $ProjectRoot -ChildPath $Child
}

function Format-DateIso($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
  try { return ([datetime]$Value).ToString('yyyy-MM-dd') } catch { return '' }
}

function Format-DateRu($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
  try { return ([datetime]$Value).ToString('dd.MM.yyyy') } catch { return '' }
}

function Format-DateShort($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
  try { return ([datetime]$Value).ToString('dd.MM') } catch { return '' }
}

function ConvertTo-JsLiteral($Object, [int]$Depth = 20) {
  ($Object | ConvertTo-Json -Depth $Depth -Compress).Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function Get-StatusClass([int]$Pct, [datetime]$Start, [datetime]$Finish, [bool]$Milestone, [datetime]$Ref) {
  if ($Pct -ge 100) { return 'done' }
  if ($Finish -lt $Ref) { return 'late' }
  if ($Milestone) { return 'milestone' }
  if ($Pct -gt 0) { return 'wip' }
  return 'open'
}

function Get-BadgeClass([string]$Status) {
  switch ($Status) {
    'done' { 'b-grn' }
    'late' { 'b-red' }
    'milestone' { 'b-pur' }
    'wip' { 'b-blu' }
    default { 'b-gray' }
  }
}

function Get-DashboardCategory([string]$Name, [bool]$Milestone, [string]$PhaseCode = '') {
  if ($Milestone -or $Name -match 'ВЕХА|Gate|Grand Opening|GRAND OPENING') { return 'milestone' }
  if ($PhaseCode -eq 'Ф.3') { return 'work' }
  if ($Name -match '^Закуп|Платеж|Договор|Спецификац|Подготовк|Финальн.*смет|инвойс|аванс') { return 'buy' }
  return 'work'
}

function Add-BigDashboardTask($PhaseMap, [string]$PhaseCode, [string]$Name, $Start, $Finish, [int]$Pct, [bool]$Milestone, [datetime]$Ref, [string]$Category) {
  if (-not $PhaseCode -or -not $PhaseMap.ContainsKey($PhaseCode)) { return }
  $status = Get-StatusClass $Pct $Start $Finish $Milestone $Ref
  $PhaseMap[$PhaseCode].tasks += [ordered]@{
    name = $Name
    start = (Format-DateShort $Start)
    finish = (Format-DateShort $Finish)
    pct = $Pct
    status = $status
    act_start = $(if ($Pct -gt 0) { Format-DateShort $Start } else { '' })
    act_finish = $(if ($Pct -ge 100) { Format-DateShort $Finish } else { '' })
    cat = $Category
    category = $Category
    slip = $(if ($status -eq 'late') { '+1 дн.' } elseif ($status -eq 'done') { 'в срок' } else { 'в графике' })
    slipover = ($status -eq 'late')
  }
}

function Get-DocStatus([string[]]$Patterns) {
  $files = @()
  foreach ($pattern in $Patterns) {
    $files += Get-ChildItem -LiteralPath $ProjectRoot -File -Filter $pattern -ErrorAction SilentlyContinue
  }
  $file = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($file) {
    return [ordered]@{
      status = 'Готово'
      badge = 'b-grn'
      file = $file.Name
      updated = $file.LastWriteTime.ToString('dd.MM.yyyy HH:mm')
    }
  }
  return [ordered]@{ status = 'Нет файла'; badge = 'b-gray'; file = $null; updated = '' }
}

$ksgPath = Join-ProjectPath 'PRK-2026-DS_KSG_DinoPark.mpp'
$workDashboardPath = Join-ProjectPath 'PRK-2026-DS_Рабочий_дашборд.html'
$mainDashboardPath = Join-ProjectPath 'PRK-2026-DS_Dashboard.html'
$photoRoot = Join-ProjectPath 'Фотофиксация работ'
$dataPath = Join-ProjectPath 'PRK-2026-DS_dashboard-data.json'

if (-not (Test-Path -LiteralPath $ksgPath)) { throw "Не найден КСГ: $ksgPath" }
if (-not (Test-Path -LiteralPath $workDashboardPath)) { throw "Не найден рабочий дашборд: $workDashboardPath" }
if (-not (Test-Path -LiteralPath $mainDashboardPath)) { throw "Не найден основной дашборд: $mainDashboardPath" }

$ref = Get-Date
$tasks = New-Object System.Collections.Generic.List[object]
$phaseSummaries = New-Object System.Collections.Generic.List[object]
$bigPhases = New-Object System.Collections.Generic.List[object]

$app = New-Object -ComObject MSProject.Application
$app.Visible = $false
try {
  $null = $app.FileOpen($ksgPath, $true)
  $project = $app.ActiveProject
  $currentPhase = $null
  $currentBlock = '—'
  $currentBlockCategory = 'work'
  $phaseMap = @{}

  foreach ($t in $project.Tasks) {
    if ($null -eq $t) { continue }
    $name = ([string]$t.Name).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { continue }

    $level = [int]$t.OutlineLevel
    $pct = [int]$t.PercentComplete
    $start = [datetime]$t.Start
    $finish = [datetime]$t.Finish
    $isSummary = [bool]$t.Summary
    $isMilestone = [bool]$t.Milestone

    if ($level -eq 2 -and $name -match 'ФАЗА\s*(\d+)') {
      $currentPhase = 'Ф.' + $Matches[1]
      $currentBlock = '—'
      $currentBlockCategory = 'work'
      $status = Get-StatusClass $pct $start $finish $false $ref
      $phaseObj = [ordered]@{
        nm = ($currentPhase + ' ' + ($name -replace '^ФАЗА\s*\d+\.\s*', ''))
        pct = $pct
        s = (Format-DateShort $start)
        e = (Format-DateShort $finish)
        cls = $(if ($pct -ge 100) { 'done' } elseif ($pct -gt 0) { 'act' } else { 'idle' })
      }
      $phaseSummaries.Add($phaseObj)

      $bigObj = [ordered]@{
        id = [string]$Matches[1]
        name = $name
        icon = '📌'
        plan_start = (Format-DateRu $start)
        plan_finish = (Format-DateRu $finish)
        act_start = $(if ($pct -gt 0) { Format-DateRu $start } else { '' })
        act_finish = $(if ($pct -ge 100) { Format-DateRu $finish } else { '' })
        pct = $pct
        budget = '—'
        status = $status
        slip = $(if ($status -eq 'late') { '+1 дн.' } elseif ($status -eq 'done') { 'в срок' } else { 'в графике' })
        slipover = ($status -eq 'late')
        tasks = @()
      }
      $phaseMap[$currentPhase] = $bigObj
      $bigPhases.Add($bigObj)
      continue
    }

    if ($level -eq 3 -and $isSummary) {
      $currentBlock = $name
      $currentBlockCategory = Get-DashboardCategory $name $isMilestone $currentPhase
      if ($currentPhase -and $phaseMap.ContainsKey($currentPhase)) {
        Add-BigDashboardTask $phaseMap $currentPhase $name $start $finish $pct $isMilestone $ref $currentBlockCategory
      }
      continue
    }

    if (-not $currentPhase) { continue }
    if ($isSummary) { continue }

    $leafCategory = Get-DashboardCategory $name $isMilestone $currentPhase
    if ($level -ge 4 -and -not $isMilestone -and ($name -match '^Закуп' -or $currentBlock -match '^Закуп')) {
      Add-BigDashboardTask $phaseMap $currentPhase $name $start $finish $pct $isMilestone $ref 'buy'
    }

    $taskStatus = Get-StatusClass $pct $start $finish $isMilestone $ref
    $tasks.Add([ordered]@{
      ph = $currentPhase
      blk = $currentBlock
      nm = $name
      s = (Format-DateIso $start)
      e = (Format-DateIso $finish)
      pct = $pct
      ms = $isMilestone
      status = $taskStatus
    })
  }
}
finally {
  try { $app.FileClose(0) | Out-Null } catch {}
  try { $app.Quit() | Out-Null } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}

$allTasks = @($tasks.ToArray() | ForEach-Object { [pscustomobject]$_ })
$active = @($allTasks | Where-Object { $_.pct -gt 0 -and $_.pct -lt 100 })
$doneCount = @($allTasks | Where-Object { $_.pct -ge 100 }).Count
$progress = if ($allTasks.Count -gt 0) { [math]::Round((($allTasks | Measure-Object -Property pct -Average).Average), 0) } else { 0 }

$needsAction = @(
  $allTasks |
    Where-Object {
      $_.pct -lt 100 -and
      (([datetime]$_.e -lt $ref.Date) -or (([datetime]$_.s -lt $ref.Date) -and $_.pct -eq 0))
    } |
    Sort-Object @{ Expression = { [datetime]$_.e } }, pct |
    Select-Object -First 20 |
    ForEach-Object {
      [ordered]@{
        nm = $_.nm
        ph = $_.ph
        blk = $_.blk
        dl = (Format-DateRu $_.e)
        pct = $_.pct
        act = $(if ($_.pct -eq 0) { 'Начать или перенести срок' } else { 'Завершить / закрыть в КСГ' })
        cls = $(if ([datetime]$_.e -lt $ref.Date) { 'red' } else { 'amb' })
      }
    }
)

$crit = @(
  $allTasks |
    Where-Object { $_.ms -or $_.nm -match 'GRAND|ВЕХА|готовност|открыт|при.мк|монтаж.*заверш' } |
    Sort-Object @{ Expression = { [datetime]$_.e } } |
    Select-Object -First 8 |
    ForEach-Object {
      $st = Get-StatusClass $_.pct ([datetime]$_.s) ([datetime]$_.e) $_.ms $ref
      [ordered]@{
        nm = $_.nm
        dt = (Format-DateRu $_.e)
        st = $st
        b = $(if ($_.pct -ge 100) { 'Готово' } elseif ([datetime]$_.e -lt $ref.Date) { 'Просрочено' } else { 'Ожидание' })
        bc = (Get-BadgeClass $st)
      }
    }
)

$docSpecs = @(
  @{ grp='init'; gl='Инициация проекта'; ico='🪪'; nm='Карточка инициативы'; code='PMO-01.01'; patterns=@('PMO-01_01_*PRK-2026-DS.*','PMO-01.01_*') },
  @{ grp='init'; ico='📄'; nm='Бриф проекта'; code='PMO-01.02'; patterns=@('PMO-01_02_*PRK-2026-DS.*','PMO-01.02_*') },
  @{ grp='init'; ico='📉'; nm='Экономическое обоснование / ФЭМ'; code='PMO-01.03'; patterns=@('PMO-01_03_*PRK-2026-DS.*','ФЭМ_PRK-2026-DS*.xlsx','PMO-01.03_*') },
  @{ grp='init'; ico='📋'; nm='Устав проекта'; code='PMO-01.04'; patterns=@('PMO-01_04_*PRK-2026-DS.*','PMO-01.04_*') },
  @{ grp='plan'; gl='Планирование'; ico='👥'; nm='Реестр стейкхолдеров'; code='PMO-02.01'; patterns=@('PMO-02_01_*PRK-2026-DS.*','PMO-02.01_*') },
  @{ grp='plan'; ico='📋'; nm='Реестр требований'; code='PMO-02.02'; patterns=@('PMO-02_02_*PRK-2026-DS.*','PMO-02.02_*') },
  @{ grp='plan'; ico='🔒'; nm='Реестр допущений и ограничений'; code='PMO-02.03'; patterns=@('PMO-02_03_*PRK-2026-DS.*') },
  @{ grp='plan'; ico='🌳'; nm='WBS'; code='PMO-02.04'; patterns=@('PMO-02_04_*PRK-2026-DS.*','PMO-02.04_*') },
  @{ grp='plan'; ico='⚡'; nm='Матрица RACI'; code='PMO-02.05'; patterns=@('PMO-02_05_*PRK-2026-DS.*','PMO-02.05_*') },
  @{ grp='plan'; ico='📡'; nm='План коммуникаций'; code='PMO-02.06'; patterns=@('PMO-02_06_*PRK-2026-DS.*') },
  @{ grp='plan'; ico='⚠️'; nm='Реестр рисков'; code='PMO-02.07'; patterns=@('PMO-02_07_*PRK-2026-DS*.xlsx','PMO-02.07_*') },
  @{ grp='exec'; gl='Реализация'; ico='📅'; nm='КСГ'; code='KSG'; patterns=@('PRK-2026-DS_KSG_DinoPark.mpp') },
  @{ grp='exec'; ico='💰'; nm='Смета'; code='Смета'; patterns=@('Смета_Динопарк_PRK-2026-DS*.xlsx') },
  @{ grp='exec'; ico='🐛'; nm='Журнал проблем'; code='PMO-03.02'; patterns=@('PMO-03_02_*PRK-2026-DS.*') },
  @{ grp='exec'; ico='📝'; nm='Реестр изменений'; code='PMO-03.03'; patterns=@('PMO-03_03_*PRK-2026-DS.*') },
  @{ grp='close'; gl='Завершение проекта'; ico='🏁'; nm='План приёмки'; code='PMO-04.01'; patterns=@('PMO-04_01_*PRK-2026-DS.*','PMO-04.01_*') }
)

$docs = @(
  foreach ($spec in $docSpecs) {
    $st = Get-DocStatus $spec.patterns
    $o = [ordered]@{
      grp = $spec.grp
      ico = $spec.ico
      nm = $spec.nm
      code = $spec.code
      bc = $st.badge
      st = $st.status
      file = $st.file
      updated = $st.updated
    }
    if ($spec.ContainsKey('gl')) { $o.gl = $spec.gl }
    $o
  }
)

$photoGroups = @()
if (Test-Path -LiteralPath $photoRoot) {
  $dirs = Get-ChildItem -LiteralPath $photoRoot -Directory | Sort-Object Name
  $idx = 0
  foreach ($dir in $dirs) {
    $files = Get-ChildItem -LiteralPath $dir.FullName -File -Include *.jpg,*.jpeg,*.png,*.webp,*.mp4,*.mov -Recurse |
      Sort-Object LastWriteTime -Descending
    if (-not $files) { continue }
    $fls = @(
      foreach ($f in $files) {
        $rel = $f.FullName.Substring($ProjectRoot.Length).TrimStart('\') -replace '\\','/'
        [ordered]@{
          url = [uri]::EscapeUriString($rel)
          n = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
          t = $(if ($f.Extension -match 'mp4|mov') { 'vid' } else { 'img' })
        }
      }
    )
    $phase = if ($dir.Name -match 'Производство') { 'Ф3' } elseif ($dir.Name -match 'Реставрация') { 'Ф4' } elseif ($dir.Name -match 'Настил|Забор') { 'Ф5' } else { 'Ф6' }
    $icon = if ($dir.Name -match 'Производство') { '🏭' } elseif ($dir.Name -match 'МАФ') { '🛝' } elseif ($dir.Name -match 'Настил|Забор') { '🔨' } else { '🦕' }
    $photoGroups += [ordered]@{ icon=$icon; cat=$dir.Name; phase=$phase; cnt=$files.Count; fls=$fls; idx=$idx }
    $idx++
  }
}

$grand = $allTasks | Where-Object { $_.nm -match 'Grand Opening|GRAND OPENING|Открытие' } | Sort-Object @{ Expression = { [datetime]$_.e } } | Select-Object -Last 1
$grandDate = if ($grand) { Format-DateRu $grand.e } else { '' }
$daysLeft = if ($grand) { [math]::Floor((([datetime]$grand.e).Date - $ref.Date).TotalDays) } else { $null }

$metrics = @(
  [ordered]@{ lbl='Grand Opening'; val=$grandDate; cls=''; sub=$(if ($daysLeft -ne $null) { "осталось $daysLeft дн." } else { 'дата не найдена в КСГ' }) },
  [ordered]@{ lbl='Прогресс проекта'; val=("$progress%"); cls=$(if ($progress -ge 80) { 'c-grn' } elseif ($progress -ge 50) { 'c-amb' } else { 'c-red' }); sub=('КСГ · ' + $phaseSummaries.Count + ' фаз · ' + $doneCount + '/' + $allTasks.Count + ' задач закрыто') }
)

$alerts = @()
if ($needsAction.Count -gt 0) {
  $alerts += [ordered]@{ cls='red'; icon='🔧'; html=('<b>Требуют актуализации в КСГ:</b> ' + $needsAction.Count + ' задач(и) по состоянию на ' + $ref.ToString('dd.MM.yyyy') + '.') }
}
$photoTotal = 0
foreach ($g in $photoGroups) { $photoTotal += [int]$g.cnt }
$alerts += [ordered]@{ cls='info'; icon='ℹ'; html=('КСГ ' + (Get-Item -LiteralPath $ksgPath).LastWriteTime.ToString('dd.MM.yyyy HH:mm') + ' · автосборка ' + $ref.ToString('dd.MM.yyyy HH:mm') + ' · фото: ' + $photoTotal + ' файл(ов)') }

$payload = [ordered]@{
  generated_at = $ref.ToString('s')
  ksg_file = (Get-Item -LiteralPath $ksgPath).Name
  ksg_updated = (Get-Item -LiteralPath $ksgPath).LastWriteTime.ToString('s')
  phases = $phaseSummaries
  tasks = $allTasks
  needs_action = $needsAction
  critical = $crit
  docs = $docs
  photos = $photoGroups
  metrics = $metrics
  alerts = $alerts
}
[System.IO.File]::WriteAllText($dataPath, (ConvertTo-JsLiteral $payload 30), [System.Text.UTF8Encoding]::new($false))

$ksgVer = 'v' + (Get-Item -LiteralPath $ksgPath).LastWriteTime.ToString('dd.MM')
$ksgDate = (Get-Item -LiteralPath $ksgPath).LastWriteTime.ToString('dd.MM.yyyy')
$refIso = $ref.ToString('yyyy-MM-dd')
$stamp = $ref.ToString('dd.MM.yyyy HH:mm')

$workDataBlock = @(
  "const KSG_VER  = '$(($ksgVer -replace "'", "\'"))';"
  "const KSG_DATE = '$(($ksgDate -replace "'", "\'"))';"
  "const REF_DATE = '$refIso';"
  "const BUILD_STAMP = '$stamp';"
  'const PHASES = ' + (ConvertTo-JsLiteral $phaseSummaries 20) + ';'
  'const CRIT = ' + (ConvertTo-JsLiteral $crit 20) + ';'
  'const ZI = [];'
  'const NEEDS_ACTION = ' + (ConvertTo-JsLiteral $needsAction 20) + ';'
  'const ALL_TASKS = ' + (ConvertTo-JsLiteral $allTasks 25) + ';'
  'const FILE_BLOBS = {};'
  'const DOCS = ' + (ConvertTo-JsLiteral $docs 20) + ';'
  'const METRICS = ' + (ConvertTo-JsLiteral $metrics 20) + ';'
  'const ALERTS = ' + (ConvertTo-JsLiteral $alerts 20) + ';'
) -join "`r`n"

$workHtml = [System.IO.File]::ReadAllText($workDashboardPath, [System.Text.Encoding]::UTF8)
$workPattern = '(?s)const KSG_VER\s*=.*?\r?\n\r?\nconst REF\s*='
$workUpdated = [regex]::Replace($workHtml, $workPattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $workDataBlock + "`r`n`r`nconst REF =" }, 1)
if ($workUpdated -eq $workHtml) { throw 'Не удалось обновить DATA-блок рабочего дашборда.' }
[System.IO.File]::WriteAllText($workDashboardPath, $workUpdated, [System.Text.UTF8Encoding]::new($false))

$galleryScript = @"
<script>
var GAL=$(ConvertTo-JsLiteral $photoGroups 30);
var gC=null,gI=0;
function glbOpen(ci){var g=GAL[ci];if(!g||!g.cnt)return;gC=ci;gI=0;
  document.getElementById("glb-title").textContent=g.icon+" "+g.cat;
  document.getElementById("glb-meta").textContent=g.phase+" · "+g.cnt+" файл(ов)";
  glbR();glbBuild();document.getElementById("glb").classList.add("open");document.body.style.overflow="hidden";}
function glbClose(){document.getElementById("glb").classList.remove("open");document.body.style.overflow="";gC=null;}
function glbMove(d){var g=GAL[gC];gI=(gI+d+g.fls.length)%g.fls.length;glbR();glbUpd();}
function glbGo(i){gI=i;glbR();glbUpd();}
function getSrc(f){return f.url||"";}
function glbR(){var g=GAL[gC],fi=g.fls[gI],src=getSrc(fi),med=document.getElementById("glb-media");
  if(fi.t==="vid"){med.innerHTML='<video controls autoplay playsinline style="max-width:92%;max-height:86%;border-radius:6px;box-shadow:0 8px 40px rgba(0,0,0,.6)"><source src="'+src+'" type="video/mp4"></video>';}
  else{med.innerHTML='<img src="'+src+'" onclick="glbMove(1)" style="cursor:e-resize;max-width:90%;max-height:86%;object-fit:contain;border-radius:6px;box-shadow:0 8px 40px rgba(0,0,0,.6)">';}
  document.getElementById("glb-ctr").textContent=(gI+1)+" / "+g.fls.length;
  var cap=document.getElementById("glb-caption");if(cap)cap.textContent=fi.n;
  var p=document.getElementById("glb-prev"),n=document.getElementById("glb-next");
  if(g.fls.length<=1){p.classList.add("hidden");n.classList.add("hidden");}else{p.classList.remove("hidden");n.classList.remove("hidden");}}
function glbBuild(){var g=GAL[gC],s=document.getElementById("glb-strip");s.innerHTML="";
  g.fls.forEach(function(fi,i){var th=document.createElement("div");th.className="glb-thumb"+(i===gI?" active":"");
    th.onclick=function(){glbGo(i);};th.innerHTML=(fi.t==="vid")?'<div class="glb-thumb-vid">🎬</div>':'<img src="'+getSrc(fi)+'" loading="lazy">';s.appendChild(th);});}
function glbUpd(){var ts=document.querySelectorAll(".glb-thumb");ts.forEach(function(t,i){t.classList.toggle("active",i===gI);});
  if(ts[gI])ts[gI].scrollIntoView({behavior:"smooth",inline:"nearest"});}
document.addEventListener("keydown",function(e){if(!document.getElementById("glb").classList.contains("open"))return;
  if(e.key==="Escape")glbClose();if(e.key==="ArrowLeft")glbMove(-1);if(e.key==="ArrowRight")glbMove(1);});
document.getElementById("glb").addEventListener("click",function(e){if(e.target===this)glbClose();});
document.addEventListener("DOMContentLoaded",function(){GAL.forEach(function(g,i){var el=document.getElementById("pci"+i);
  if(el&&g.fls.length&&getSrc(g.fls[0])){el.src=getSrc(g.fls[0]);}});});
</script>

<div class="photo-section">
"@

$mainHtml = [System.IO.File]::ReadAllText($mainDashboardPath, [System.Text.Encoding]::UTF8)
$mainHtml = [regex]::Replace($mainHtml, '(?s)<script>\s*var G0=.*?</script>\s*<div class="photo-section">', [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $galleryScript }, 1)
$mainHtml = [regex]::Replace($mainHtml, '(?s)const phases\s*=\s*\[.*?\];', [System.Text.RegularExpressions.MatchEvaluator]{ param($m) 'const phases = ' + (ConvertTo-JsLiteral $bigPhases 30) + ';' }, 1)

$photoCards = @()
foreach ($g in $photoGroups) {
  $photoCards += @"
    <div class="photo-card" onclick="glbOpen($($g.idx))">
      <div class="photo-cover">
        <img id="pci$($g.idx)" src="" alt="$($g.cat)" class="photo-img" loading="lazy" decoding="async">
        <span class="photo-count-badge">$($g.cnt)</span>
      </div>
      <div class="photo-caption">
        <div class="photo-phase">$($g.phase) · $($g.cat)</div>
        <div class="photo-date">обновлено автоматически</div>
      </div>
    </div>
"@
}
$photoSection = "<div class=`"photo-section`">`r`n  <div class=`"photo-section-title`">📸 Фотофиксация хода работ</div>`r`n  <div class=`"photo-grid`">`r`n" + ($photoCards -join "`r`n") + "`r`n  </div>`r`n</div>"
$mainHtml = [regex]::Replace(
  $mainHtml,
  '(?s)<div class="photo-section">.*?</div>\s*<script>\s*const phases',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $photoSection + "`r`n<script>`r`nconst phases" },
  1
)

[System.IO.File]::WriteAllText($mainDashboardPath, $mainHtml, [System.Text.UTF8Encoding]::new($false))

Write-Host "OK: dashboards updated"
Write-Host "KSG: $ksgDate"
Write-Host "Tasks: $($allTasks.Count), done: $doneCount, progress: $progress%"
Write-Host "Photos: $photoTotal"
Write-Host "Data: $dataPath"










