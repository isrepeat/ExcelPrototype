$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = Split-Path $PSScriptRoot -Parent
$configKey = 'PersonalEventBuilder::export.Movement.destination.unitPattern'
$patterns = @()
foreach ($file in @('wsConfig.txt', 'wsConfig.test.txt')) {
    $rows = @([IO.File]::ReadAllLines((Join-Path $projectRoot "config/PersonalEventBuilder/$file")) | Where-Object { $_.Split([char]9)[1] -eq $configKey })
    if ($rows.Count -ne 1 -or $rows[0].Split([char]9).Count -ne 3) { throw "Missing or ambiguous destination pattern: $file" }
    $patterns += $rows[0].Split([char]9)[2]
}
if ($patterns[0] -ne $patterns[1]) { throw 'Destination patterns must agree between PROD and TEST.' }
$regex = New-Object -ComObject VBScript.RegExp
try {
    $regex.Global = $true
    $regex.IgnoreCase = $true
    $regex.Pattern = $patterns[0]
    function Format-TestDestination([string]$value) {
        $sections = $value.Trim().Split(';')
        $changed = $false
        for ($index = 0; $index -lt $sections.Length; $index++) {
            $matches = $regex.Execute($sections[$index].Replace([char]160, [char]32))
            $codes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($match in $matches) {
                if ($match.SubMatches.Count -ne 2) { throw 'Pattern must capture boundary and unit code.' }
                [void]$codes.Add(([string]$match.SubMatches.Item(1) -replace '[ \t\r\n]', '').ToUpperInvariant())
            }
            if ($codes.Count -eq 1) {
                $code = @($codes)[0]
                $changed = $true
                $sections[$index] = $code
            }
        }
        if ($changed) { return (($sections | ForEach-Object { $_.Trim() }) -join '; ') }
        return $value.Trim()
    }
    $cases = @(
        @('військова частина А7383, м. Київ', 'А7383'),
        @('військової частини A3990', 'A3990'),
        @('3019', '3019'),
        @('А7383;м. Київ', 'А7383; м. Київ'),
        @('в/ч3019, м. Київ', '3019'),
        @('частина Т0300, м. Львів', 'Т0300'),
        @('військова частина А7383, м. Київ; м. Полтава', 'А7383; м. Полтава'),
        @('м. Полтава; частина 3019, м. Київ', 'м. Полтава; 3019'),
        @('частина А7383, м. Київ; частина 3019, м. Одеса', 'А7383; 3019'),
        @('м. Київ;м. Полтава', 'м. Київ;м. Полтава'),
        @('м. Київ, 20.10.2026', 'м. Київ, 20.10.2026'),
        @('2026-10-20, м. Київ', '2026-10-20, м. Київ'),
        @('частина А73830', 'частина А73830'),
        @('частина 13019', 'частина 13019'),
        @('А7383 і А3990', 'А7383 і А3990'),
        @('частина А7383 (А7383), м. Київ', 'А7383'),
        @('Медична рота А7383; м. Полтава', 'А7383; м. Полтава'),
        @('  А 7383; м. Київ  ', 'А7383; м. Київ'),
        @('', '')
    )
    foreach ($case in $cases) {
        $actual = Format-TestDestination $case[0]
        if ($actual -cne $case[1]) { throw "Destination mismatch: '$($case[0])' -> '$actual'; expected '$($case[1])'." }
        if ((Format-TestDestination $actual) -cne $actual) { throw "Destination normalization must be idempotent: $actual" }
    }
    $registry = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/ex_PEB_MovementRegistry.vba'))
    $planner = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/ex_PEB_MovementPlanner.vba'))
    $validation = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/ex_PEB_MovementValidation.vba'))
    $exporter = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/obj_PEB_ExptrMovement.cls.vba'))
    if ($registry -notmatch 'private_ExportValue\(texts\(header\), header\)' -or
        $registry -notmatch 'private_ExportValue\(ex_PEB_Export.fn_Required\(values, field\), header\)' -or
        $registry -notmatch 'header = ex_PEB_Export.fn_Config\("export.Movement.headers.Destination"\)') { throw 'Both native and template values must normalize only the Movement destination column.' }
    if ($registry -notmatch 'Set m_destinationRegex = Nothing' -or
        $planner -notmatch 'fn_Destination\(mappings\(key\)\)' -or
        $exporter -notmatch 'fn_Destination\(expected\)' -or
        $validation -notmatch 'fn_Destination\(plan\("sourceDestination"\)\)') { throw 'Destination cache lifecycle and legacy source matching must remain consistent.' }
    Write-Output "PASS: $($cases.Count) destination fixtures, idempotence, VBScript regex compatibility and Movement integration contracts."
}
finally {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($regex)
}