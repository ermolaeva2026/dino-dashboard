param(
  [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)),
  [switch]$UseExistingKsgSnapshot
)

$ErrorActionPreference = 'Stop'

function Join-ProjectPath([string]$Child) {
  Join-Path -Path $ProjectRoot -ChildPath $Child
}

function New-MsProjectApplication {
  for ($attempt = 1; $attempt -le 3; $attempt++) {
    try { return New-Object -ComObject MSProject.Application }
    catch {
      if ($attempt -ge 3) { throw }
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
      Start-Sleep -Seconds (2 * $attempt)
    }
  }
}

function Invoke-MsProjectWithRetry([scriptblock]$Action, [string]$Operation) {
  for ($attempt = 1; $attempt -le 8; $attempt++) {
    try { return & $Action }
    catch {
      $isBusy = $_.Exception.Message -match 'RPC_E_CALL_REJECTED|rejected by callee|занят|отклонен вызываемым'
      if (-not $isBusy -or $attempt -ge 8) { throw "${Operation}: $($_.Exception.Message)" }
      Start-Sleep -Seconds (1 + $attempt)
    }
  }
}

function Convert-ProjectDate($Value) {
  if ($null -eq $Value) { return $null }
  $text = [string]$Value
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  if ($text -match '^(NA|НД|Н/Д|Нет|None)$') { return $null }
  try { return [datetime]$Value } catch { return $null }
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

function Format-Rub($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '—' }
  try {
    $rounded = [math]::Round([double]$Value, 0)
    return ($rounded.ToString('N0', [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')).Replace(' ', [char]0x00A0) + ' ₽')
  } catch {
    return '—'
  }
}

function Format-Rub2($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '—' }
  try {
    return (([double]$Value).ToString('N2', [System.Globalization.CultureInfo]::GetCultureInfo('ru-RU')).Replace(' ', [char]0x00A0) + ' ₽')
  } catch {
    return '—'
  }
}

function ConvertTo-JsLiteral($Object, [int]$Depth = 20) {
  ($Object | ConvertTo-Json -Depth $Depth -Compress).Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function Read-XlsxCellNumber([string]$Path, [string]$SheetEntry, [string]$CellRef) {
  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = $null
  $stream = $null
  try {
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
    $entry = $zip.GetEntry($SheetEntry)
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }

    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $false
    $xml.LoadXml($xmlText)
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $node = $xml.SelectSingleNode("//x:c[@r='$CellRef']/x:v", $ns)
    if (-not $node) { return $null }
    return [double]::Parse($node.InnerText, [System.Globalization.CultureInfo]::InvariantCulture)
  }
  finally {
    if ($zip) { $zip.Dispose() }
    if ($stream) { $stream.Dispose() }
  }
}

function Read-XlsxSheetRows([string]$Path, [string]$SheetEntry = 'xl/worksheets/sheet1.xml') {
  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = $null
  $stream = $null
  try {
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)

    $sharedStrings = @()
    $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
    if ($sharedEntry) {
      $reader = New-Object System.IO.StreamReader($sharedEntry.Open())
      try { $sharedText = $reader.ReadToEnd() } finally { $reader.Dispose() }
      $sharedXml = New-Object System.Xml.XmlDocument
      $sharedXml.LoadXml($sharedText)
      $sharedNs = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
      $sharedNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
      $sharedStrings = @(
        $sharedXml.SelectNodes('//x:si', $sharedNs) | ForEach-Object {
          (@($_.SelectNodes('.//x:t', $sharedNs) | ForEach-Object { $_.InnerText }) -join '')
        }
      )
    }

    $entry = $zip.GetEntry($SheetEntry)
    if (-not $entry) { throw "В XLSX не найден лист: $SheetEntry" }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { $sheetText = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $sheetXml = New-Object System.Xml.XmlDocument
    $sheetXml.LoadXml($sheetText)
    $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $rows = @()
    foreach ($rowNode in $sheetXml.SelectNodes('//x:sheetData/x:row', $ns)) {
      $cells = @{}
      foreach ($cellNode in $rowNode.SelectNodes('x:c', $ns)) {
        $cellRef = $cellNode.GetAttribute('r')
        if ($cellRef -notmatch '^([A-Z]+)\d+$') { continue }
        $column = $Matches[1]
        $type = $cellNode.GetAttribute('t')
        $valueNode = $cellNode.SelectSingleNode('x:v', $ns)
        $value = if ($valueNode) { [string]$valueNode.InnerText } else { '' }
        if ($type -eq 's' -and $value -match '^\d+$') {
          $index = [int]$value
          if ($index -lt $sharedStrings.Count) { $value = [string]$sharedStrings[$index] }
        }
        elseif ($type -eq 'inlineStr') {
          $value = @($cellNode.SelectNodes('.//x:is//x:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
        }
        $cells[$column] = $value
      }
      $rows += [pscustomobject]@{ row = [int]$rowNode.GetAttribute('r'); cells = $cells }
    }
    return $rows
  }
  finally {
    if ($zip) { $zip.Dispose() }
    if ($stream) { $stream.Dispose() }
  }
}

function Get-RiskSnapshot([string]$RiskRoot) {
  $riskFile = Get-ChildItem -LiteralPath $RiskRoot -File -Filter 'PMO-02_07_Reestr_riskov_PRK-2026-DS*.xlsx' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if (-not $riskFile) { throw "Не найден актуальный реестр рисков в $RiskRoot" }

  $risks = @()
  foreach ($row in (Read-XlsxSheetRows $riskFile.FullName)) {
    $id = [string]$row.cells['A']
    if ($id -notmatch '^R-\d+$') { continue }
    $score = 0
    [void][int]::TryParse([string]$row.cells['G'], [ref]$score)
    $ownerStatus = ([string]$row.cells['N']).Trim()
    $comment = ([string]$row.cells['P']).Trim()
    $formalActive = $ownerStatus -notmatch '(?i)закрыт|реализован|исключ'
    $state = if (-not $formalActive) {
      'closed'
    }
    elseif ($comment -match '(?i)не реализовал|выполнен|выполнено|завершен|завершена|завершено|завершён|сдан|осуществлен|осуществлено|осуществлён') {
      'review'
    }
    else {
      'active'
    }
    $risks += [pscustomobject]@{
      id = $id
      description = ([string]$row.cells['D']).Trim()
      score = $score
      level = ([string]$row.cells['H']).Trim()
      trigger = ([string]$row.cells['L']).Trim()
      measures = ([string]$row.cells['M']).Trim()
      owner_status = $ownerStatus
      comment = $comment
      formal_active = $formalActive
      state = $state
    }
  }

  $active = @($risks | Where-Object { $_.formal_active })
  $critical = @($active | Where-Object { $_.score -ge 15 } | Sort-Object @{ Expression = { $_.score }; Descending = $true }, id)
  return [pscustomobject]@{
    source_file = $riskFile.Name
    source_updated = $riskFile.LastWriteTime.ToString('s')
    active_count = $active.Count
    critical_count = $critical.Count
    critical = $critical
    risks = $risks
  }
}

function Get-BudgetSnapshot([string]$SmetaPath) {
  if (-not (Test-Path -LiteralPath $SmetaPath)) { throw "Не найдена смета: $SmetaPath" }
  $sheetEntry = 'xl/worksheets/sheet1.xml'
  $snapshot = [pscustomobject]@{
    updated_at = (Get-Item -LiteralPath $SmetaPath).LastWriteTime.ToString('s')
    planned = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E25'
    approved = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E26'
    reserve = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E27'
    approved_with_reserve = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E28'
    remaining_approved = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E29'
    remaining_with_reserve = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E30'
    confirmed = Read-XlsxCellNumber $SmetaPath $sheetEntry 'D22'
    paid = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E22'
    outstanding = Read-XlsxCellNumber $SmetaPath $sheetEntry 'F22'
    phase_confirmed = @{
      '3' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'D16'
      '4' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'D7'
      '5' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'D21'
      '6' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'D12'
    }
    phase_paid = @{
      '3' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E16'
      '4' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E7'
      '5' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E21'
      '6' = Read-XlsxCellNumber $SmetaPath $sheetEntry 'E12'
    }
  }
  foreach ($name in @('planned','approved','reserve','approved_with_reserve','remaining_approved','remaining_with_reserve','confirmed','paid','outstanding')) {
    if ($null -eq $snapshot.$name) { throw "В первой вкладке сметы не рассчитан показатель: $name" }
  }
  return $snapshot
}

function Get-PhaseBudgetLabel($Budget, [string]$PhaseId) {
  if ($Budget.phase_confirmed.ContainsKey($PhaseId)) { return Format-Rub $Budget.phase_confirmed[$PhaseId] }
  return '—'
}

function Html-Escape([string]$Text) {
  return [System.Net.WebUtility]::HtmlEncode($Text)
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

function Get-RestorationPhotoSortKey($File) {
  $base = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
  $num = 9999
  $title = $base
  $state = 9

  if ($base -match '^(\d+)[_\-\s]+(.+)$') {
    $num = [int]$Matches[1]
    $title = $Matches[2]
  }

  if ($title -match '(?i)(?:^|[_\-\s])(после|стало|after)(?:\d+)?(?:$|[_\-\s])') {
    $state = 0
  }
  elseif ($title -match '(?i)(?:^|[_\-\s])(до|было|before)(?:\d+)?(?:$|[_\-\s])') {
    $state = 1
  }

  $cleanTitle = $title -replace '(?i)(?:^|[_\-\s])(после|стало|after|до|было|before)(?:\d+)?(?:$|[_\-\s])', ' '
  $cleanTitle = ($cleanTitle -replace '[_\-]+', ' ').Trim()

  '{0:D4}|{1}|{2:D2}|{3}' -f $num, $cleanTitle.ToLowerInvariant(), $state, $base.ToLowerInvariant()
}
function Get-DashboardStatusText([string]$Status, [int]$Pct) {
  if ($Pct -ge 100 -or $Status -eq 'done') { return @{ Text='завершено'; Class='green' } }
  if ($Status -eq 'late') { return @{ Text='просрочено'; Class='red' } }
  if ($Pct -gt 0 -or $Status -eq 'wip') { return @{ Text='в работе'; Class='amber' } }
  return @{ Text='—'; Class='' }
}

function Get-SlipLabel([string]$Status, [datetime]$Finish, [datetime]$Ref) {
  if ($Status -eq 'done') { return 'в срок' }
  if ($Status -eq 'late') {
    $days = [math]::Max(1, [int]([datetime]$Ref.Date - [datetime]$Finish.Date).TotalDays)
    return "+$days дн."
  }
  return 'в графике'
}

function Find-Task($Tasks, [string]$Pattern) {
  @($Tasks | Where-Object { $_.nm -match $Pattern } | Sort-Object @{ Expression = { [datetime]$_.e } } | Select-Object -First 1)
}

function Get-KeyDateRow($Tasks, [string]$Label, [string]$Pattern, [datetime]$Ref) {
  $t = Find-Task $Tasks $Pattern
  if (-not $t) {
    return '<tr><td style="font-size:11px;color:var(--text-2);padding:5px 0">' + (Html-Escape $Label) + '</td><td style="font-size:11px;font-weight:700;text-align:center;padding:5px 4px">—</td><td style="font-size:11px;text-align:center;color:#94a3b8;padding:5px 4px">—</td></tr>'
  }
  $status = Get-DashboardStatusText (Get-StatusClass $t.pct ([datetime]$t.s) ([datetime]$t.e) $t.ms $Ref) $t.pct
  $color = if ($status.Class -eq 'green') { 'var(--green)' } elseif ($status.Class -eq 'red') { 'var(--red)' } elseif ($status.Class -eq 'amber') { '#d97706' } else { '#94a3b8' }
  return '<tr><td style="font-size:11px;color:var(--text-2);padding:5px 0">' + (Html-Escape $Label) + '</td><td style="font-size:11px;font-weight:700;text-align:center;padding:5px 4px">' + (Format-DateRu $t.e) + '</td><td style="font-size:11px;text-align:center;color:' + $color + ';padding:5px 4px">' + $status.Text + '</td></tr>'
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
    slip = (Get-SlipLabel $status $Finish $Ref)
    slipover = ($status -eq 'late')
  }
}

function Get-DocStatus([string[]]$Patterns) {
  $files = @()
  $roots = @(
    $ProjectRoot,
    (Join-ProjectPath 'Документы PMO'),
    (Join-ProjectPath 'Прочие документы')
  ) | Where-Object { Test-Path -LiteralPath $_ }
  foreach ($pattern in $Patterns) {
    foreach ($root in $roots) {
      $files += Get-ChildItem -LiteralPath $root -File -Filter $pattern -ErrorAction SilentlyContinue
    }
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
$smetaPath = Join-ProjectPath 'Смета_Динопарк_PRK-2026-DS.xlsx'
$riskRoot = Join-ProjectPath 'Документы PMO'

if (-not (Test-Path -LiteralPath $ksgPath)) { throw "Не найден КСГ: $ksgPath" }
if (-not (Test-Path -LiteralPath $workDashboardPath)) { throw "Не найден рабочий дашборд: $workDashboardPath" }
if (-not (Test-Path -LiteralPath $mainDashboardPath)) { throw "Не найден основной дашборд: $mainDashboardPath" }

$ref = Get-Date
$budget = Get-BudgetSnapshot $smetaPath
$riskSnapshot = Get-RiskSnapshot $riskRoot
$tasks = New-Object System.Collections.Generic.List[object]
$phaseSummaries = New-Object System.Collections.Generic.List[object]
$bigPhases = New-Object System.Collections.Generic.List[object]

if ($UseExistingKsgSnapshot) {
  if (-not (Test-Path -LiteralPath $dataPath)) { throw "Не найден предыдущий снимок КСГ: $dataPath" }
  $existingSnapshot = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
  if (@($existingSnapshot.tasks).Count -eq 0 -or @($existingSnapshot.phases).Count -eq 0) {
    throw "Предыдущий снимок КСГ не содержит фаз или задач: $dataPath"
  }

  $phaseMap = @{}
  foreach ($phase in @($existingSnapshot.phases)) {
    if ([string]$phase.nm -notmatch '^Ф\.(\d+)\s+(.+)$') { continue }
    $phaseId = [string]$Matches[1]
    $phaseTitle = [string]$Matches[2]
    $phaseCode = 'Ф.' + $phaseId
    $phaseStart = [datetime]::ParseExact(([string]$phase.s + '.' + $ref.Year), 'dd.MM.yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    $phaseFinish = [datetime]::ParseExact(([string]$phase.e + '.' + $ref.Year), 'dd.MM.yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    $phasePct = [int]$phase.pct
    $phaseStatus = Get-StatusClass $phasePct $phaseStart $phaseFinish $false $ref
    $phaseSummaries.Add([ordered]@{ nm = [string]$phase.nm; pct = $phasePct; s = [string]$phase.s; e = [string]$phase.e; cls = [string]$phase.cls })
    $bigObj = [ordered]@{
      id = $phaseId
      name = ('ФАЗА ' + $phaseId + '. ' + $phaseTitle)
      icon = '📌'
      plan_start = (Format-DateRu $phaseStart)
      plan_finish = (Format-DateRu $phaseFinish)
      act_start = $(if ($phasePct -gt 0) { Format-DateRu $phaseStart } else { '' })
      act_finish = $(if ($phasePct -ge 100) { Format-DateRu $phaseFinish } else { '' })
      pct = $phasePct
      budget = (Get-PhaseBudgetLabel $budget $phaseId)
      status = $phaseStatus
      slip = (Get-SlipLabel $phaseStatus $phaseFinish $ref)
      slipover = ($phaseStatus -eq 'late')
      tasks = @()
    }
    $phaseMap[$phaseCode] = $bigObj
    $bigPhases.Add($bigObj)
  }

  foreach ($row in @($existingSnapshot.tasks)) {
    $taskStart = [datetime]::ParseExact([string]$row.s, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $taskFinish = [datetime]::ParseExact([string]$row.e, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $taskPct = [int]$row.pct
    $isMilestone = [bool]$row.ms
    $taskStatus = Get-StatusClass $taskPct $taskStart $taskFinish $isMilestone $ref
    $tasks.Add([ordered]@{
      ph = [string]$row.ph
      blk = [string]$row.blk
      nm = [string]$row.nm
      s = (Format-DateIso $taskStart)
      e = (Format-DateIso $taskFinish)
      pct = $taskPct
      ms = $isMilestone
      status = $taskStatus
    })
    $taskCategory = if ([string]$row.blk -match '^Закуп' -or [string]$row.nm -match '^Закуп') { 'buy' } else { Get-DashboardCategory ([string]$row.nm) $isMilestone ([string]$row.ph) }
    Add-BigDashboardTask $phaseMap ([string]$row.ph) ([string]$row.nm) $taskStart $taskFinish $taskPct $isMilestone $ref $taskCategory
  }
}
else {
  $app = New-MsProjectApplication
  $openedByScript = $false
  try {
  $project = $null
  try {
    $candidate = Invoke-MsProjectWithRetry { $app.ActiveProject } 'Не удалось прочитать открытый проект'
    if ($null -ne $candidate) {
      $candidatePath = Invoke-MsProjectWithRetry { [string]$candidate.FullName } 'Не удалось определить путь открытого проекта'
      if ([string]::Equals([System.IO.Path]::GetFullPath($candidatePath), [System.IO.Path]::GetFullPath($ksgPath), [System.StringComparison]::OrdinalIgnoreCase)) {
        $project = $candidate
      }
    }
  }
  catch {
    $project = $null
  }

  if ($null -eq $project) {
    Invoke-MsProjectWithRetry { $app.FileOpen($ksgPath, $true) } 'Не удалось открыть КСГ' | Out-Null
    $openedByScript = $true
    $project = Invoke-MsProjectWithRetry { $app.ActiveProject } 'Не удалось получить открытый КСГ'
  }
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
    $start = Convert-ProjectDate $t.Start
    $finish = Convert-ProjectDate $t.Finish
    if ($null -eq $start -and $null -ne $finish) { $start = $finish }
    if ($null -eq $finish -and $null -ne $start) { $finish = $start }
    if ($null -eq $start -or $null -eq $finish) { continue }
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
        budget = (Get-PhaseBudgetLabel $budget $Matches[1])
        status = $status
        slip = (Get-SlipLabel $status $finish $ref)
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

    $leafCategory = if ($name -match '^Закуп' -or $currentBlock -match '^Закуп') { 'buy' } else { Get-DashboardCategory $name $isMilestone $currentPhase }
    Add-BigDashboardTask $phaseMap $currentPhase $name $start $finish $pct $isMilestone $ref $leafCategory

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
  if ($openedByScript) {
    try { $app.FileClose(0) | Out-Null } catch {}
    try { $app.Quit() | Out-Null } catch {}
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
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
    $files = Get-ChildItem -LiteralPath $dir.FullName -File -Include *.jpg,*.jpeg,*.png,*.webp,*.mp4,*.mov -Recurse
    if ($dir.Name -eq 'Реставрация динозавров') {
      $files = $files | Sort-Object @{ Expression = { Get-RestorationPhotoSortKey $_ } }, LastWriteTime
    }
    else {
      $files = $files | Sort-Object LastWriteTime -Descending
    }
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
  risk_register = [ordered]@{
    source_file = $riskSnapshot.source_file
    source_updated = $riskSnapshot.source_updated
    active_count = $riskSnapshot.active_count
    critical_count = $riskSnapshot.critical_count
    critical = $riskSnapshot.critical
    risks = $riskSnapshot.risks
  }
  budget = [ordered]@{
    updated_at = $budget.updated_at
    planned = $budget.planned
    approved = $budget.approved
    reserve = $budget.reserve
    approved_with_reserve = $budget.approved_with_reserve
    confirmed = $budget.confirmed
    paid = $budget.paid
    outstanding = $budget.outstanding
    remaining_approved = $budget.remaining_approved
    remaining_with_reserve = $budget.remaining_with_reserve
    phase_confirmed = $budget.phase_confirmed
    phase_paid = $budget.phase_paid
  }
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
"@

$mainHtml = [System.IO.File]::ReadAllText($mainDashboardPath, [System.Text.Encoding]::UTF8)
$mainHtml = [regex]::Replace($mainHtml, 'Обновлено:\s*\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}', 'Обновлено: ' + $stamp)
$mainHtml = [regex]::Replace($mainHtml, '(?s)<script>\s*var G0=.*?</script>\s*<div class="photo-section">', [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $galleryScript }, 1)
$mainHtml = [regex]::Replace($mainHtml, '(?s)const phases\s*=\s*\[.*?\];', [System.Text.RegularExpressions.MatchEvaluator]{ param($m) 'const phases = ' + (ConvertTo-JsLiteral $bigPhases 30) + ';' }, 1)

$riskItemsHtml = @(
  foreach ($risk in @($riskSnapshot.critical)) {
    $ownerLine = $risk.owner_status
    if ($risk.trigger) { $ownerLine += ' · Триггер: ' + $risk.trigger }
    $commentText = if ($risk.comment) { $risk.comment } else { 'не заполнен' }
    $commentColor = if ($risk.state -eq 'review') { '#b45309' } elseif ($risk.comment) { '#15803d' } else { '#64748b' }
    $commentLabel = if ($risk.state -eq 'review') { 'Комментарий, статус требует финализации' } else { 'Комментарий' }
    '      <div class="risk-item"><div class="risk-id">' + (Html-Escape $risk.id) + '</div><div class="risk-body"><div class="risk-name">' + (Html-Escape $risk.description) + '</div><div class="risk-owner">' + (Html-Escape $ownerLine) + '</div><div style="font-size:10px;line-height:1.35;margin-top:4px;color:' + $commentColor + '"><b>' + $commentLabel + ':</b> ' + (Html-Escape $commentText) + '</div></div><div class="risk-badge rb-crit">🔴 ' + $risk.score + '</div></div>'
  }
) -join "`r`n"
if ([string]::IsNullOrWhiteSpace($riskItemsHtml)) {
  $riskItemsHtml = '      <div style="padding:12px 14px;font-size:11px;color:var(--green)">Критических рисков с формальным статусом «Активный» нет.</div>'
}
$riskSourceUpdated = ([datetime]$riskSnapshot.source_updated).ToString('dd.MM.yyyy HH:mm')
$riskCardHtml = @"
  <div class="card risks-card">
    <div class="card-head">
      <h2>🔴 Критические риски</h2>
      <span class="tag tag-red">$($riskSnapshot.critical_count) из $($riskSnapshot.active_count) формально активных</span>
    </div>
    <div class="risk-list">
${riskItemsHtml}
      <div style="padding:10px 14px;font-size:10px;color:var(--gray);border-top:1px solid var(--border)">Источник: $([System.Net.WebUtility]::HtmlEncode($riskSnapshot.source_file)) от $riskSourceUpdated. Формальный статус взят из колонки N, комментарий — из колонки P.</div>
    </div>
  </div>
"@
$mainHtml = [regex]::Replace(
  $mainHtml,
  '(<div class="header-kpi"><div class="val red">)\d+(</div><div class="lbl">Критич\. рисков)',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + $riskSnapshot.critical_count + $m.Groups[2].Value },
  1
)
$mainHtml = [regex]::Replace(
  $mainHtml,
  '(?s)(<div class="status-cell"><div class="icon">⚠️</div><div class="s-label">Риски</div>)<div class="s-val [^"]*">.*?</div></div>',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + '<div class="s-val ' + $(if ($riskSnapshot.critical_count -gt 0) { 's-amber' } else { 's-green' }) + '">' + $(if ($riskSnapshot.critical_count -gt 0) { '🟡 ' + $riskSnapshot.critical_count + ' крит.' } else { '🟢 OK' }) + '</div></div>' },
  1
)
$mainHtml = [regex]::Replace($mainHtml, '(?s)\s*<div class="card risks-card">.*?</div>\s*(?=</div>\s*<div style="padding: 0 32px 0; max-width:1600px; margin:0 auto;">)', "`r`n" + $riskCardHtml, 1)

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

$budgetPlanned = [double]$budget.planned
$budgetApproved = [double]$budget.approved
$budgetReserve = [double]$budget.reserve
$budgetApprovedWithReserve = [double]$budget.approved_with_reserve
$budgetConfirmed = [double]$budget.confirmed
$budgetPaid = [double]$budget.paid
$budgetOutstanding = [double]$budget.outstanding
$budgetRemainingApproved = [double]$budget.remaining_approved
$budgetRemainingWithReserve = [double]$budget.remaining_with_reserve
if ($budgetRemainingApproved -ge 0) {
  $budgetStatusClass = 's-green'; $budgetStatusText = '🟢 В лимите'
} elseif ($budgetRemainingWithReserve -ge 0) {
  $budgetStatusClass = 's-amber'; $budgetStatusText = '🟡 Резерв'
} else {
  $budgetStatusClass = 's-red'; $budgetStatusText = '🔴 Перерасход'
}

$vedBudget = Get-PhaseBudgetLabel $budget '3'
$vedAdvance = Find-Task $allTasks 'Внесение аванса ПИ'
$vedProduction = Find-Task $allTasks '^Производство 12 аниматроников'
$vedSettlement = Find-Task $allTasks 'Окончательный расч[её]т между ПИ'
$vedDelivery = Find-Task $allTasks 'Доставка из Китая'
$vedOnSite = Find-Task $allTasks 'Аниматроники на площадке парка'

function New-VedRow($Label, $Task, [datetime]$Ref) {
  if (-not $Task) {
    return '<tr><td style="font-size:11px;color:var(--text-2);padding:5px 0">' + (Html-Escape $Label) + '</td><td style="font-size:11px;font-weight:700;text-align:center;padding:5px 4px">—</td><td style="font-size:11px;text-align:center;color:#94a3b8;padding:5px 4px">—</td></tr>'
  }
  $status = Get-DashboardStatusText (Get-StatusClass $Task.pct ([datetime]$Task.s) ([datetime]$Task.e) $Task.ms $Ref) $Task.pct
  $color = if ($status.Class -eq 'green') { 'var(--green)' } elseif ($status.Class -eq 'red') { 'var(--red)' } elseif ($status.Class -eq 'amber') { '#d97706' } else { '#94a3b8' }
  $fact = if ($Task.pct -ge 100) { Format-DateRu $Task.e } elseif ($Task.pct -gt 0) { $status.Text } else { '—' }
  return '<tr><td style="font-size:11px;color:var(--text-2);padding:5px 0">' + (Html-Escape $Label) + '</td><td style="font-size:11px;font-weight:700;text-align:center;padding:5px 4px">' + (Format-DateRu $Task.e) + '</td><td style="font-size:11px;text-align:center;color:' + $color + ';padding:5px 4px">' + $fact + '</td></tr>'
}

$keyRows = @(
  Get-KeyDateRow $allTasks 'Реставрация завершена' 'ВЕХА: реставрационные работы завершены' $ref
  Get-KeyDateRow $allTasks 'Детские зоны' 'ВЕХА: Благоустройство детских зон завершено|Игровая зона .* принята' $ref
  Get-KeyDateRow $allTasks 'Электромонтаж' 'ВЕХА: Освещение введено в эксплуатацию 3 этап|Электросети введены в эксплуатацию' $ref
  Get-KeyDateRow $allTasks 'Аниматроники на площадке' 'Аниматроники на площадке парка' $ref
  Get-KeyDateRow $allTasks 'Монтаж завершён' 'Аниматроники смонтированы и работают штатно' $ref
  Get-KeyDateRow $allTasks '🏁 Grand Opening' 'Grand Opening|GRAND OPENING|Открытие' $ref
) -join "`r`n        "

$vedRows = @(
  New-VedRow 'Аванс ПИ' $vedAdvance $ref
  New-VedRow 'Производство' $vedProduction $ref
  New-VedRow 'Окончат. расчёт' $vedSettlement $ref
  New-VedRow 'Доставка' $vedDelivery $ref
  New-VedRow 'На площадке' $vedOnSite $ref
) -join "`r`n        "

$infoStrip = @"
<div class="info-strip">

  <div class="info-card">
    <div class="ic-label">💰 Финансы</div>
    <div class="info-row"><span class="ir-key">Плановый бюджет</span><span class="ir-val">$(Format-Rub2 $budgetPlanned)</span></div>
    <div class="info-row"><span class="ir-key">Утвержд. бюджет</span><span class="ir-val">$(Format-Rub2 $budgetApproved)</span></div>
    <div class="info-row"><span class="ir-key">Подтверждено</span><span class="ir-val amber">$(Format-Rub2 $budgetConfirmed)</span></div>
    <div class="info-row"><span class="ir-key">Оплачено</span><span class="ir-val green">$(Format-Rub2 $budgetPaid)</span></div>
    <div class="info-row"><span class="ir-key">К оплате</span><span class="ir-val amber">$(Format-Rub2 $budgetOutstanding)</span></div>
    <div class="info-row"><span class="ir-key">Резерв 10%</span><span class="ir-val">$(Format-Rub2 $budgetReserve)</span></div>
    <div class="info-row"><span class="ir-key">Остаток лимита</span><span class="ir-val $(if ($budgetRemainingApproved -lt 0) { 'red' } else { 'green' })">$(Format-Rub2 $budgetRemainingApproved)</span></div>
    <div class="info-row"><span class="ir-key">Остаток с резервом</span><span class="ir-val $(if ($budgetRemainingWithReserve -lt 0) { 'red' } else { 'green' })">$(Format-Rub2 $budgetRemainingWithReserve)</span></div>
  </div>

  <div class="info-card">
    <div class="ic-label">🚢 ВЭД-контур</div>
    <div class="info-row"><span class="ir-key">Аниматроники</span><span class="ir-val">12 шт. / $vedBudget</span></div>
    <table style="width:100%;border-collapse:collapse;margin-top:6px">
      <colgroup><col style="width:auto"><col style="width:96px"><col style="width:96px"></colgroup>
      <thead><tr><th style="font-size:9px;color:var(--text-2);text-align:left;padding:4px 0">Этап</th><th style="font-size:9px;color:var(--text-2);padding:4px">План</th><th style="font-size:9px;color:var(--text-2);padding:4px">Факт</th></tr></thead>
      <tbody>
        $vedRows
      </tbody>
    </table>
  </div>

  <div class="info-card">
    <div class="ic-label">📌 Ключевые даты</div>
    <table style="width:100%;border-collapse:collapse;margin-top:6px">
      <colgroup><col style="width:auto"><col style="width:90px"><col style="width:90px"></colgroup>
      <thead><tr><th style="font-size:9px;color:var(--text-2);text-align:left;padding:4px 0">Веха</th><th style="font-size:9px;color:var(--text-2);padding:4px">План</th><th style="font-size:9px;color:var(--text-2);padding:4px">Факт</th></tr></thead>
      <tbody>
        $keyRows
      </tbody>
    </table>
  </div>

  <div class="info-card" style="position:relative">
    <div class="ic-label">🏁 Сценарии Grand Opening</div>
    <div style="margin-top:10px;padding:8px 10px;border-radius:8px;background:#fefce8;border:1px solid #fde68a">
      <div style="font-size:9px;font-weight:700;letter-spacing:.6px;color:#92400e;margin-bottom:4px">ПЛАН КСГ — с резервами</div>
      <div style="display:grid;grid-template-columns:1fr auto;gap:2px 6px;align-items:center">
        <span style="font-size:10px;color:#78350f">Аниматроники на площадке</span><span style="font-size:11px;font-weight:700;color:#92400e;text-align:right">07.08.2026</span>
        <span style="font-size:10px;color:#78350f">Монтаж завершён</span><span style="font-size:11px;font-weight:700;color:#92400e;text-align:right">20.08.2026</span>
        <span style="font-size:10px;color:#78350f;font-weight:700">🏁 Grand Opening</span><span style="font-size:13px;font-weight:800;color:#92400e;text-align:right">24.08.2026</span>
      </div>
    </div>
    <div style="margin-top:8px;padding:6px 10px;border-radius:8px;background:var(--gray-light);border:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
      <span style="font-size:10px;color:var(--gray);font-weight:700">📋 Устав (крайний срок)</span><span style="font-size:12px;font-weight:800;color:var(--text)">01.09.2026</span>
    </div>
    <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap">
      <span style="font-size:9px;padding:2px 7px;border-radius:10px;background:#e0f2fe;color:#0369a1;font-weight:600">📦 Произв. 7 р.д.</span>
      <span style="font-size:9px;padding:2px 7px;border-radius:10px;background:#e0f2fe;color:#0369a1;font-weight:600">📦 Доставка 3 р.д.</span>
      <span style="font-size:9px;padding:2px 7px;border-radius:10px;background:#e0f2fe;color:#0369a1;font-weight:600">📦 Таможня 2 р.д.</span>
      <span style="font-size:9px;padding:2px 7px;border-radius:10px;background:#e0f2fe;color:#0369a1;font-weight:600">📦 Монтаж 4 р.д.</span>
    </div>
  </div>

  <div class="info-card">
    <div class="ic-label">👥 Команда</div>
    <div class="info-row"><span class="ir-key">PM</span><span class="ir-val">Ермолаева Ирина</span></div>
    <div class="info-row"><span class="ir-key">ГД / Заказчик</span><span class="ir-val">Горелов Виктор</span></div>
    <div class="info-row"><span class="ir-key">Куратор / Опер. дир.</span><span class="ir-val">Ксенофонтов Валерий</span></div>
    <div class="info-row"><span class="ir-key">Директор PMO</span><span class="ir-val">Иванов Сергей</span></div>
    <div class="info-row"><span class="ir-key">Директор ДЕЗ</span><span class="ir-val">Киселев Михаил</span></div>
    <div class="info-row"><span class="ir-key">Директор Сервис А</span><span class="ir-val">Саркитов Александр</span></div>
    <div class="info-row"><span class="ir-key">Реставратор</span><span class="ir-val">Березин</span></div>
  </div>
</div>
"@

$glbHtml = @"
<div class="glb" id="glb">
  <div class="glb-hdr">
    <div><div class="glb-title" id="glb-title"></div><div class="glb-meta" id="glb-meta"></div></div>
    <div style="display:flex;align-items:center;gap:10px"><span class="glb-counter" id="glb-ctr"></span><span class="glb-close" onclick="glbClose()">&#10005;</span></div>
  </div>
  <div class="glb-stage">
    <div class="glb-arrow left" id="glb-prev" onclick="glbMove(-1)">&#8249;</div>
    <div id="glb-media"></div>
    <div class="glb-arrow right" id="glb-next" onclick="glbMove(1)">&#8250;</div>
  </div>
  <div class="glb-caption-bar" id="glb-caption"></div>
  <div class="glb-strip" id="glb-strip"></div>
</div>
"@

$infoAndPhotos = "<div style=`"padding: 0 32px 0; max-width:1600px; margin:0 auto;`">`r`n" + $infoStrip + "`r`n</div>`r`n`r`n" + $glbHtml + "`r`n" + $galleryScript + "`r`n" + $photoSection

$mainHtml = [regex]::Replace($mainHtml, '(?s)<div class="header-kpi"><div class="val amber">.*?</div><div class="lbl">До Grand Opening</div></div>', '<div class="header-kpi"><div class="val amber">' + $daysLeft + ' дней</div><div class="lbl">До Grand Opening</div></div>')
$mainHtml = [regex]::Replace($mainHtml, '(?s)<div class="header-kpi"><div class="val green">\d+%</div><div class="lbl">Прогресс проекта</div></div>', '<div class="header-kpi"><div class="val green">' + $progress + '%</div><div class="lbl">Прогресс проекта</div></div>')
$mainHtml = [regex]::Replace($mainHtml, '(?s)<div class="header-kpi"><div class="val amber">.*?₽</div><div class="lbl">(?:Бюджет \(актуал\.\)|Фактические платежи)</div></div>', '<div class="header-kpi"><div class="val amber">' + (Format-Rub $budgetPaid) + '</div><div class="lbl">Фактические платежи</div></div>')
$mainHtml = [regex]::Replace(
  $mainHtml,
  '(?s)(<div class="status-cell"><div class="icon">💰</div><div class="s-label">Бюджет</div>)<div class="s-val [^"]*">.*?</div></div>',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + '<div class="s-val ' + $budgetStatusClass + '">' + $budgetStatusText + '</div></div>' },
  1
)
$mainHtml = [regex]::Replace($mainHtml, '<span class="ob-pct">\d+%</span>', '<span class="ob-pct">' + $progress + '%</span>')
$mainHtml = [regex]::Replace($mainHtml, 'style="width:\d+%"', 'style="width:' + $progress + '%"', 1)
$mainHtml = [regex]::Replace(
  $mainHtml,
  '(?s)(?:<div style="padding: 0 32px 0; max-width:1600px; margin:0 auto;">\s*)*<div class="info-strip">.*?<script>\s*const phases',
  [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $infoAndPhotos + "`r`n<script>`r`nconst phases" },
  1
)
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
















