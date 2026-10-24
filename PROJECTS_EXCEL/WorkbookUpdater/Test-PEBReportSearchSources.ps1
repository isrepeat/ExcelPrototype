$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$reports = Get-Content -LiteralPath (Join-Path $root 'vba/PersonalEventBuilder/ex_PEB_MedicalReports.vba') -Raw
$controller = Get-Content -LiteralPath (Join-Path $root 'vba/PersonalEventBuilder/obj_PEB_PgMainController.cls.vba') -Raw
$page = Get-Content -LiteralPath (Join-Path $root 'vba/PersonalEventBuilder/obj_PEB_PgMain.cls.vba') -Raw
[xml]$markup = Get-Content -LiteralPath (Join-Path $root 'ui/PersonalEventBuilder/MainPage.xaml') -Raw
$panels = $markup.SelectSingleNode("//*[local-name()='stackPanel' and @name='RecordPanels']")
$button = $markup.SelectSingleNode("//*[local-name()='button' and @name='OpenReportSearch']")
if ($button.GetAttribute('row') -ne '2' -or $button.GetAttribute('column') -ne '17' -or
    $button.GetAttribute('columnSpan') -ne '2') { throw 'Search button must be below New event.' }
if ($panels.LastChild.GetAttribute('name') -ne 'ReportSearchPanel') { throw 'Search must be the final workspace panel.' }
$tables = $panels.LastChild.SelectSingleNode("*[local-name()='tableList']")
if ($tables.GetAttribute('itemsSource') -ne '{Binding Path=Data.ReportSearchResults}') { throw 'Missing per-file table list.' }
if ($markup.SelectNodes("//*[local-name()='stackPanel' or local-name()='grid'][@visibility]").Count -gt 0) {
    throw 'Layout containers do not support visibility; bind visibility on controls.'
}
$form = $panels.LastChild.SelectSingleNode(".//*[local-name()='form']")
if ($form.GetAttribute('visibility') -ne '{Binding Path=Form.ReportSearchVisible}' -or
    $form.SelectNodes("*[local-name()='field']").Count -ne 4) { throw 'Expected a collapsible search form with four fields.' }
foreach ($field in $form.SelectNodes("*[local-name()='field']")) {
    if ($field.GetAttribute('columnSpan') -ne '2') { throw 'Search fields must each span two columns.' }
}
if ($tables.GetAttribute('titleColumnSpan') -ne '3') { throw 'Result titles must span the first three columns.' }
$run = $form.ParentNode.SelectSingleNode("*[local-name()='button' and @name='RunReportSearch']")
if ($null -eq $run -or $form.ParentNode.GetAttribute('orientation') -ne 'horizontal') { throw 'Search action must sit to the right of the form.' }
if ($form.GetAttribute('closeCommand') -ne '{Binding Path=Commands.CloseReportSearch}' -or
    $markup.SelectNodes("//*[local-name()='button' and @name='CloseReportSearch']").Count -gt 0) { throw 'Search must use the native left close button.' }
if ($form.GetAttribute('style') -ne 'candidateTable') { throw 'The native close button must inherit the standard table style.' }
$close = [regex]::Match($controller, '(?s)Public Function CloseReportSearchHandler\(.*?End Function').Value
if (-not $close.Contains('ReportResultsVisible')) { throw 'Closing search must hide its results.' }
foreach ($name in @('OpenReportSearch', 'RunReportSearch', 'CloseReportSearch')) {
    if ($page -notmatch ('"' + $name + '"') -or $controller -notmatch ('Public Function ' + $name + 'Handler')) { throw "Unwired search command: $name" }
}
$search = [regex]::Match($reports, '(?s)Public Function fn_TrySearch\(.*?End Function').Value
foreach ($contract in @('If day = dateTo Then Exit Do', 'If rows.Count > 0 Then', 'MedicalSearch.Discover', 'MedicalSearch.Read', 'MedicalSearch.Total')) {
    if (-not $search.Contains($contract)) { throw "Missing search contract: $contract" }
}
if ($search -match 'fn_ReadOrder|fn_TryApply|private_CheckPeople') { throw 'Report search must not depend on event history or personnel dependencies.' }
$reader = [regex]::Match($reports, '(?s)Private Function private_ReadFile\(.*?End Function').Value
if ($search -notmatch 'queryService.Backend = QueryAdo' -or $reader -match 'query.AddColumn private_Config' -or
    $reader -notmatch 'private_HeaderIndex\(data.Headers') { throw 'Search must use ADO and normalize provider headers before resolving columns.' }
if (-not $controller.Contains('MedicalSearch.Render | Success=')) { throw 'Search render timing is required.' }
$reset = [regex]::Match($controller, '(?s)Public Function ResetFormHandler\(.*?End Function').Value
if (-not $reset.Contains('private_ResetReportSearch(diagnostic)') -or -not $reset.Contains('ReportSearchVisible')) { throw 'New event must clear and hide report search.' }
foreach ($field in @('ReportSearchFrom', 'ReportSearchTo', 'ReportSearchName', 'ReportSearchDate')) {
    if (-not $page.Contains('SetValue("Form", "' + $field + '", VBA.vbNullString)')) { throw "Batch update target must be initialized: $field" }
}
$open = [regex]::Match($controller, '(?s)Public Function OpenReportSearchHandler\(.*?End Function').Value
if (-not $open.Contains('private_SetActionStatus(diagnostic, True)') -or -not $open.Contains('On Error GoTo EH')) { throw 'Opening failures must appear in status.' }
$defaults = [regex]::Match($controller, '(?s)Private Function private_ResetReportSearch\(.*?End Function').Value
if ($defaults -notmatch 'today = VBA.Date' -or $defaults -notmatch 'search.defaultDays') { throw 'Defaults must use the computer date and configured offset.' }
foreach ($file in @('wsConfig.txt', 'wsConfig.test.txt')) {
    $values = @{}
    foreach ($line in Get-Content -LiteralPath (Join-Path $root "config/PersonalEventBuilder/$file")) {
        if ($line -match 'PersonalEventBuilder::validation.MedicalReports.search\.') {
            $parts = $line.Split("`t")
            if ($parts.Count -ne 3 -or $values.ContainsKey($parts[1])) { throw "Invalid search configuration row: $file" }
            $values.Add($parts[1], $parts[2])
        }
    }
    if ($values['PersonalEventBuilder::validation.MedicalReports.search.defaultDays'] -ne '10') { throw 'Expected ten-day default.' }
    foreach ($key in @('To', 'Run', 'Date', 'From', 'Name', 'Open', 'Close', 'Loaded', 'OpenFailed', 'InvalidDate', 'InvalidRange', 'RangeTooWide')) {
        if (-not $values.ContainsKey('PersonalEventBuilder::validation.MedicalReports.search.text.' + $key)) { throw "Missing search text: $key" }
    }
}
Write-Output 'PASS: report search wiring, per-file results, range defaults, reset, independence and configuration.'