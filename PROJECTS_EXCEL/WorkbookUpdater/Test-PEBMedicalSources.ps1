$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = Split-Path $PSScriptRoot -Parent
$prefix = 'PersonalEventBuilder::validation.MedicalReports.'
$configs = @()
foreach ($name in @('wsConfig.txt', 'wsConfig.test.txt')) {
    $config = @{}
    foreach ($line in [IO.File]::ReadAllLines((Join-Path $projectRoot "config/PersonalEventBuilder/$name"))) {
        $parts = $line.Split([char]9)
        if ($parts.Count -ne 3) { throw "Invalid configuration row in $name" }
        if ($config.ContainsKey($parts[1])) { throw "Duplicate configuration key: $($parts[1])" }
        $config.Add($parts[1], $parts[2])
    }
    $fields = $config["${prefix}fields"].Split('|')
    if ($fields -notcontains 'Dates' -or $fields -notcontains 'PersonName' -or $fields -notcontains 'PersonId') { throw 'Missing mandatory medical report fields.' }
    foreach ($field in $fields) {
        if (-not $config["${prefix}headers.$field"]) { throw "Missing source header: $field" }
    }
    foreach ($field in $config["${prefix}display"].Split('|')) {
        if ($fields -notcontains $field -and $field -notin @('Sources', 'ReferenceName', 'IdentityStatus', 'Notes')) { throw "Unknown medical report display field: $field" }
        if (-not $config["${prefix}text.Columns.$field"]) { throw "Missing display caption: $field" }
    }
    if ([int]$config["${prefix}windowDays"] -ne 3) { throw 'Default medical report window must be three days.' }
    $patterns = $config["${prefix}patterns"].Split('|')
    foreach ($pattern in $patterns) {
        if ($pattern -notmatch '\{dd\.mm\.yyyy\}|\{yyyymmdd\}') { throw "Filename pattern needs a report-date token: $pattern" }
    }
    $configs += ,$config
    foreach ($status in @('Matched', 'MissingId', 'InvalidId', 'UnknownId', 'AmbiguousId', 'NameMismatch')) {
        if (-not $config["${prefix}text.Identity.$status"]) { throw "Missing identity diagnostic: $status" }
    }
}
foreach ($key in $configs[0].Keys) {
    if ($key.StartsWith($prefix) -and $configs[0][$key] -ne $configs[1][$key]) { throw "Medical report PROD/TEST settings differ: $key" }
}

[xml]$page = [IO.File]::ReadAllText((Join-Path $projectRoot 'ui/PersonalEventBuilder/MainPage.xaml'))
$bindings = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/obj_PEB_PgMain.cls.vba'))
$controller = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/obj_PEB_PgMainController.cls.vba'))
foreach ($command in @('ValidateMedical', 'CloseMedical', 'EmptyOrders')) {
    if ($bindings -notmatch ('"' + $command + '"')) { throw "Missing command registration: $command" }
    if ($controller -notmatch ("Public Function ${command}Handler\(")) { throw "Missing command handler: $command" }
    if ($page.OuterXml -notmatch ("Commands\.$command\}")) { throw "Missing command binding: $command" }
}
foreach ($name in @('ValidateMedical', 'MedicalReports')) {
    if ($page.SelectNodes("//*[@name='$name']").Count -ne 1) { throw "Missing or duplicate medical report control: $name" }
    if (('btn_' + $name).Length -gt 31) { throw "Excel shape name is too long: $name" }
}
$table = $page.SelectSingleNode("//*[@name='MedicalReports']")
if ($table.GetAttribute('itemsSource') -ne '{Binding Path=Data.MedicalReports}' -or
    $table.GetAttribute('visibility') -ne '{Binding Path=Form.MedicalVisibility}') { throw 'Incorrect report table bindings.' }
$lookupResult = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/common/[2] classes/obj_LookupResult.cls.vba'))
$tableControl = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/common/[3] controls/obj_UiTableControl.cls.vba'))
$reports = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/ex_PEB_MedicalReports.vba'))
if ($lookupResult -notmatch 'Implements obj_IUiRowTags' -or
    $lookupResult -notmatch 'Optional ByVal rowTagField As String = ""' -or
    $tableControl -notmatch 'rowTagSource.RowTagsAt\(rowIndex\)') { throw 'Missing optional row-tag pipeline.' }
if ($page.SelectNodes('//*[@selector="type=table;part=row;tags=medicalReview"]').Count -ne 1 -or
    $reports -notmatch 'result.Initialize\(records, columns, False, diagnostic, "RowTags"\)') { throw 'Missing identity warning presentation.' }
if ($reports -match 'private_Fail "MissingName"') { throw 'Missing report names must remain visible for identity review.' }
$validationHandler = [regex]::Match($controller, '(?s)Public Function ValidateMedicalHandler\(\).*?End Function').Value
if ($validationHandler -match 'fn_ShowMessage|MsgBox' -or
    $validationHandler -notmatch 'ValidateMedicalHandler = m_pageBase.FlushUpdates\(\)') { throw 'Validation must report handled failures through status only.' }
$orderSelect = $page.SelectSingleNode('//*[@name="OrderQuery"]')
if ($orderSelect.GetAttribute('onEmpty') -ne '{Binding Path=Commands.EmptyOrders}') { throw 'Empty order dropdown must explain its state.' }
$emptyOrdersHandler = [regex]::Match($controller, '(?s)Public Function EmptyOrdersHandler\(\).*?End Function').Value
if ($emptyOrdersHandler -match 'fn_Current\(' -or $emptyOrdersHandler -notmatch 'fn_Active\(') {
    throw 'Empty order dropdown must inspect snapshot availability without requiring a loaded snapshot.'
}
$controlFactory = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/common/[2] classes/obj_UiControlFactory.cls.vba'))
$selectSchema = [regex]::Match($controlFactory, '(?s)Case "input", "select"(.*?)Case "form"').Groups[1].Value
if ($selectSchema -notmatch 'schema.AddAttribute "onEmpty", "string", False, "", True, 0') { throw 'Select schema must accept the optional onEmpty command binding.' }
$selectAction = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/common/[2] classes/obj_UiSelectShapeAction.cls.vba'))
if ($selectAction -notmatch 'm_uiFieldControl.HandleEmptyDropdown\(\)') { throw 'Missing empty dropdown callback.' }
$medicalState = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/ex_PEB_MedicalState.vba'))
foreach ($tag in @('medicalRecorded', 'medicalConflict')) {
    if ($page.SelectNodes("//*[@selector='type=table;part=row;tags=$tag']").Count -ne 1 -or
        $medicalState -notmatch ('"' + $tag + '"')) { throw "Missing medical state presentation: $tag" }
}
if ($medicalState -notmatch 'fn_ReadOrder\(orderNumber, orderDate, events, diagnostic, False, people, False\)' -or
    $medicalState -notmatch 'storedEvent\("Status"\) = "Active"' -or
    $medicalState -notmatch 'storedEvent\("AcceptedOrderNumber"\) = orderNumber' -or
    $medicalState -notmatch 'storedEvent\("EventDate"\)' -or
    $medicalState -notmatch 'plan\("targetStatus"\)') { throw 'History matching must be live, allow empty results and require the active hospital event in the accepted order.' }
if ($medicalState -notmatch 'dateKey <= VBA.CDbl\(orderDate\)' -or
    $medicalState -notmatch 'state\("ArrivalDate"\)' -or
    $medicalState -notmatch 'status.HospitalTreatment' -or
    $medicalState -notmatch 'status.MedicalLeave') { throw 'Missing temporal or open-period guards for medical conflicts.' }
Write-Output 'PASS: medical report configuration, projections, UI commands and object names.'
$personIdCheck = [regex]::Match($reports, '(?s)Private Function private_PersonId\(.*?End Function').Value
if ($personIdCheck -match 'CreateObject|RegExp' -or $personIdCheck -notmatch 'text Like "##########"') {
    throw 'Tax ID validation must not allocate a regular expression per row.'
}
$firstDateCheck = [regex]::Match($reports, '(?s)Private Function private_FirstDateMatches\(.*?End Function').Value
if ($firstDateCheck -match 'CreateObject' -or $firstDateCheck -notmatch 'dateRx.Execute\(text\)' -or
    [regex]::Matches($reports, 'CreateObject\("VBScript.RegExp"\)').Count -ne 1) { throw 'Date matching must reuse one invocation-local regular expression.' }
$personnelCheck = [regex]::Match($reports, '(?s)Private Function private_CheckPeople\(.*?End Function').Value
if ($personnelCheck -notmatch '(?s)If neededIds.Exists\(personId\) Then.*?Set names = VBA.CreateObject' -or
    $personnelCheck -notmatch 'referenceValues = data.Values') { throw 'Personnel name sets must be limited to requested IDs.' }
if ($medicalState -notmatch 'If people.Exists\(personId\) Then' -or
    $medicalState -notmatch 'If Not completed.Exists\(key\) Then pending\(key\) = True' -or
    $medicalState -notmatch 'If people.Count = 0 Then') { throw 'Movement checks must be limited to unrecorded report people.' }
foreach ($stage in @('Config', 'Discover', 'Reports', 'Identity', 'Result', 'Load', 'File.Query', 'File.Filter', 'File.Total', 'Personnel.Query', 'Personnel.Index', 'Personnel.Compare')) {
    if (-not $reports.Contains('MedicalValidation.' + $stage)) { throw "Missing report timing stage: $stage" }
}
foreach ($stage in @('State.People', 'History.Read', 'History.Match', 'Movement.Query', 'Movement.Index', 'State.Notes', 'State.Total')) {
    if (-not $medicalState.Contains('MedicalValidation.' + $stage)) { throw "Missing state timing stage: $stage" }
}
if (-not $validationHandler.Contains('MedicalValidation.Render') -or
    -not $validationHandler.Contains('MedicalValidation.Total') -or
    $reports -notmatch 'If configSnapshotActive Then ex_PEB_Export.fn_EndConfigSnapshot') { throw 'Missing overall timing or configuration snapshot cleanup.' }
Write-Output 'PASS: medical validation allocation guards, restricted projections and stage timings.'