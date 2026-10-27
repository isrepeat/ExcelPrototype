$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = Split-Path $PSScriptRoot -Parent
[xml]$page = Get-Content -LiteralPath (Join-Path $projectRoot 'ui/PersonalEventBuilder/MainPage.xaml') -Raw
[xml]$routes = Get-Content -LiteralPath (Join-Path $projectRoot 'workbook/PersonalEventBuilder/PersonalEventBuilderEventRoutes.xml') -Raw
$catalog = $page.SelectSingleNode('//*[local-name()="fieldCatalog" and @name="EventFields"]')
$form = $page.SelectSingleNode('//*[local-name()="form" and @name="PersonnelForm"]')
if ($null -eq $catalog -or $null -eq $form) { throw 'Missing personnel form or field catalog.' }
if ($form.GetAttribute('fieldCatalog') -ne 'EventFields' -or
    $form.GetAttribute('fieldsSource') -ne '{Binding Path=Data.EventFields}') {
    throw 'Personnel form must consume the configured candidate field collection.'
}
if ($page.SelectNodes('//*[local-name()="formLayouts"]').Count -ne 0 -or
    $form.HasAttribute('layout') -or $form.HasAttribute('layoutSource')) {
    throw 'Personnel form must not duplicate configured fields in named XAML layouts.'
}
$catalogFields = @{}
foreach ($field in $catalog.SelectNodes('*')) {
    $name = $field.GetAttribute('name')
    if ($catalogFields.ContainsKey($name)) { throw "Duplicate catalog field: $name" }
    $catalogFields.Add($name, $true)
}
$profilesByEnvironment = @{}
foreach ($file in @('wsConfig.txt', 'wsConfig.test.txt')) {
    $config = @{}
    $path = Join-Path $projectRoot "config/PersonalEventBuilder/$file"
    foreach ($line in [IO.File]::ReadAllLines($path)) {
        $parts = $line.Split([char]9)
        if ($parts.Count -ne 3) { throw "Invalid configuration row: $file" }
        if ($config.ContainsKey($parts[1])) { throw "Duplicate configuration key: $($parts[1])" }
        $config.Add($parts[1], $parts[2])
        if ($parts[1] -match '\.arrivalLayout$') { throw 'Obsolete arrivalLayout setting.' }
    }
    $profiles = @{}
    foreach ($event in $routes.SelectNodes('//*[local-name()="event"]')) {
        $name = $event.GetAttribute('form')
        $prefix = "PersonalEventBuilder::forms.$name"
        if (-not $config.ContainsKey("$prefix.fields.count")) { throw "Missing form profile: $name" }
        $count = [int]$config["$prefix.fields.count"]
        if ($count -lt 1 -or $count -gt $catalogFields.Count) { throw "Invalid form field count: $name" }
        $fields = @()
        for ($index = 1; $index -le $count; $index++) {
            $fieldPrefix = "$prefix.fields.$index"
            $field = $config["$fieldPrefix.name"]
            $caption = $config["$fieldPrefix.captionKey"]
            if (-not $field -or -not $catalogFields.ContainsKey($field)) { throw "Unknown form field: $fieldPrefix" }
            if ($fields -contains $field) { throw "Duplicate configured field: $name/$field" }
            if (-not $caption -or -not $config.ContainsKey($caption)) { throw "Missing caption: $name/$field" }
            $fields += $field
        }
        if ($event.GetAttribute('direction') -eq 'Arrival' -and $fields -contains 'DurationDays') {
            throw "Arrival profile contains departure duration: $name"
        }
        $profiles[$name] = $fields -join '|'
    }
    foreach ($name in @('LeaveArrival', 'Leave')) {
        if (($profiles[$name] -split '\|') -notcontains 'EventDate') { throw "Missing leave event date: $name" }
    }
    $arrivalFields = @($profiles['LeaveArrival'] -split '\|')
    $arrivalDateIndex = [Array]::IndexOf($arrivalFields, 'EventDate') + 1
    if ($config.ContainsKey("PersonalEventBuilder::forms.LeaveArrival.fields.$arrivalDateIndex.record")) {
        throw 'Arrival date must not be filled from the previous departure.'
    }
    $medicalFields = @($profiles['MedicalCompany'] -split '\|')
    $expectedMedicalFields = @($profiles['Treatment'] -split '\|' | Where-Object { $_ -notin @('Hospital', 'HospitalShortName') })
    if (($medicalFields -join '|') -ne ($expectedMedicalFields -join '|')) {
        throw 'Medical company form must retain treatment fields except hospital names.'
    }
    foreach ($eventId in @('FromMedicalCompanyForTreatment', 'ToMedicalCompanyForTreatment')) {
        $event = $routes.SelectSingleNode("//*[local-name()='event' and @id='$eventId']")
        if ($event.GetAttribute('form') -ne 'MedicalCompany') { throw "Medical company profile missing: $eventId" }
    }
    foreach ($field in @('Hospital', 'HospitalShortName')) {
        if (($profiles['Treatment'] -split '\|') -notcontains $field) { throw "Hospital field removed from shared treatment form: $field" }
    }
    if (($config['PersonalEventBuilder::export.ticket.form'] -split '\|') -notcontains 'LeaveArrival') {
        throw 'Arrival profile must retain ticket preparation support.'
    }
    $profilesByEnvironment[$file] = $profiles
}
foreach ($name in $profilesByEnvironment['wsConfig.txt'].Keys) {
    if ($profilesByEnvironment['wsConfig.txt'][$name] -ne $profilesByEnvironment['wsConfig.test.txt'][$name]) {
        throw "PROD/TEST form fields differ: $name"
    }
}
Write-Output 'PASS: all event forms use configured catalog fields; arrival/departure profiles and PROD/TEST agree.'