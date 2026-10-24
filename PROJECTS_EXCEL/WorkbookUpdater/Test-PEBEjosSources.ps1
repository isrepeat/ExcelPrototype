$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = Split-Path $PSScriptRoot -Parent
$prefix = 'PersonalEventBuilder::validation.Ejos.'
$expectedColumns = '65|66|67|68|69|70|71 (А)|71 (Б)|72|73|74|75 (А)|75 (Б)|76|Стовпець1|Стовпець2'
$configs = @()
foreach ($name in @('wsConfig.txt', 'wsConfig.test.txt')) {
    $config = @{}
    foreach ($line in [IO.File]::ReadAllLines((Join-Path $projectRoot "config/PersonalEventBuilder/$name"))) {
        $parts = $line.Split([char]9)
        if ($parts.Count -ne 3) { throw "Invalid configuration row: $name" }
        if ($config.ContainsKey($parts[1])) { throw "Duplicate key: $($parts[1])" }
        $config.Add($parts[1], $parts[2])
    }
    if ($config["${prefix}columns"] -ne $expectedColumns) { throw 'EJOS source projection must retain the sixteen requested columns in order.' }
    if ($config["${prefix}text.NoteHeader"] -ne 'Примітка') { throw 'Missing configured EJOS notes header.' }
    if ($config["${prefix}fields.orderDate"] -ne '71 (А)' -or
        $config["${prefix}fields.rank"] -ne '65' -or
        $config["${prefix}fields.destination"] -ne '69') { throw 'Incorrect date/cancellation source fields.' }
    foreach ($field in @('name', 'code', 'kind', 'departure', 'return', 'orderNumber')) {
        if ($config["${prefix}fields.$field"] -notin $expectedColumns.Split('|')) { throw "Unknown source field: $field" }
    }
    $aliases = @{}
    for ($index = 1; $index -le [int]$config["${prefix}kinds.count"]; $index++) {
        if (-not $config["${prefix}kinds.$index.status"]) { throw "Missing event target status: $index" }
        foreach ($alias in $config["${prefix}kinds.$index.aliases"].Split('|')) {
            if (-not $alias.Trim() -or $aliases.ContainsKey($alias.Trim())) { throw "Empty or ambiguous kind alias: $alias" }
            $aliases.Add($alias.Trim(), $true)
        }
    }
    $configs += ,$config
}
foreach ($key in $configs[0].Keys) {
    if ($key.StartsWith($prefix) -and $configs[0][$key] -ne $configs[1][$key]) { throw "EJOS PROD/TEST settings differ: $key" }
}
[xml]$page = [IO.File]::ReadAllText((Join-Path $projectRoot 'ui/PersonalEventBuilder/MainPage.xaml'))
$source = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/ex_PEB_EjosValidation.vba'))
if ($source -notmatch '(?s)column\("name"\) = "ValidationNote"\s+column\("caption"\) = private_Message\("NoteHeader"\)\s+columns.Add column\s+For Each item In headers') {
    throw 'The diagnostic note must be the first display column, before all source columns.'
}
$controller = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/obj_PEB_PgMainController.cls.vba'))
$bindings = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/obj_PEB_PgMain.cls.vba'))
$table = $page.SelectSingleNode('//*[@name="EjosResults"]')
if (-not $table -or $table.GetAttribute('itemsSource') -ne '{Binding Path=Data.EjosResults}' -or
    $table.NextSibling.GetAttribute('name') -ne 'MedicalReports') { throw 'EJOS must precede medical reports in the workspace.' }
foreach ($command in @('CloseEjos')) {
    if (-not $bindings.Contains('"' + $command + '"') -or
        -not $controller.Contains("Public Function ${command}Handler(")) { throw "Missing command: $command" }
}
foreach ($tag in @('ejosRecorded', 'ejosReview', 'ejosCancelled')) {
    if ($page.SelectNodes("//*[@selector='type=table;part=row;tags=$tag']").Count -ne 1 -or
        -not $source.Contains('"' + $tag + '"')) { throw "Missing color route: $tag" }
}
if ($source -match 'MsgBox|fn_ShowMessage' -or
    $source -notmatch 'service.TryExecute\(source, query, data, diagnostic\)' -or
    $source -notmatch 'If configSnapshotActive Then ex_PEB_Export.fn_EndConfigSnapshot') { throw 'EJOS validation must use snapshots, balance configuration ownership and return errors through status.' }
if ($source -notmatch 'storedEvent\("Status"\) = "Active"' -or
    $source -notmatch 'storedEvent\("EventDate"\)\) = departure' -or
    $source -notmatch 'storedEvent\("AcceptedOrderDate"\)\) = orderDate' -or
    $source -notmatch 'plan\("targetStatus"\) = expected') { throw 'Recorded events must match person, order, departure date and event type.' }
if ($source -notmatch '(?s)If isCancelled Then.*?ejosCancelled.*?Else.*?ejosReview.*?ElseIf matched Then.*?ejosRecorded') { throw 'Cancellation and review colors must take precedence over green.' }
if ($table.HasAttribute('onSelect') -or $table.HasAttribute('selectedItem') -or
    $table.HasAttribute('selectionMode') -or $controller.Contains('SelectEjosHandler') -or
    $bindings.Contains('"SelectEjos"')) { throw 'EJOS navigation must not dispatch selection commands or update status.' }
$tableControl = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/common/[3] controls/obj_UiTableControl.cls.vba'))
if ($tableControl -notmatch '(?s)If Not m_selectCommand Is Nothing Or VBA.Len\(m_selectedSource\) > 0 Or VBA.Len\(m_columnPath\) > 0 Then\s+uiRenderContext.Router.RegisterSelection') {
    throw 'Display-only tables must not intercept selection or cancel native double-click editing.'
}
foreach ($stage in @('Source', 'Filter', 'Identity', 'History', 'Compare', 'Result', 'Load')) {
    if (-not $source.Contains('EjosValidation.' + $stage)) { throw "Missing EJOS stage timing: $stage" }
}
Write-Output 'PASS: EJOS columns, mappings, event matching, precedence, workspace order, status and timing contracts.'