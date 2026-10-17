param(
    [string]$UiRoot = (Join-Path $PSScriptRoot '..\..\PROJECTS_EXCEL\ui')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$UiRoot = (Resolve-Path -LiteralPath $UiRoot).Path
$profileNamespace = 'urn:excelprototype:profiles'

function Read-StyleXml([string]$Path) {
    $settings = [Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $reader = [Xml.XmlReader]::Create($Path, $settings)
    try {
        $document = [Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    } finally {
        $reader.Dispose()
    }
}

function Get-StyleSourcePaths([Xml.XmlDocument]$Page, [string]$PagePath) {
    $sources = $Page.DocumentElement.GetAttribute('styleSources').Trim()
    if ($sources.Length -eq 0) { return }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $sources.Split(';')) {
        $path = $entry.Trim().Replace('/', '\')
        if ($path.Length -eq 0) { throw 'Empty file path in page styleSources.' }
        if (-not [IO.Path]::IsPathRooted($path)) {
            $path = Join-Path ([IO.Path]::GetDirectoryName($PagePath)) $path
        }
        $path = [IO.Path]::GetFullPath($path)
        if (-not $seen.Add($path)) { throw "Duplicate page style source: $path" }
        if (-not [IO.File]::Exists($path)) { throw "Page style source file not found: $path" }
        $path
    }
}

function Get-PageStyleCatalog([Xml.XmlDocument]$Page, [string]$PagePath) {
    $catalog = @{}
    foreach ($path in @(Get-StyleSourcePaths $Page $PagePath)) {
        $document = Read-StyleXml $path
        $root = $document.DocumentElement
        if ($root.NamespaceURI -cne $profileNamespace -or $root.LocalName -cnotin @('styleCatalog', 'styles')) {
            throw "Invalid page style source root: $path"
        }
        if ($root.HasAttribute('version') -and $root.GetAttribute('version') -cne '1') {
            throw "Unsupported page style source version: $path"
        }
        foreach ($node in $document.SelectNodes("//*[local-name()='controlStyle']")) {
            $name = $node.GetAttribute('name')
            if ($name.Length -gt 0) { $catalog[$name] = $path }
        }
    }
    foreach ($node in $Page.SelectNodes("//*[local-name()='controlStyle']")) {
        $name = $node.GetAttribute('name')
        if ($name.Length -gt 0) { $catalog[$name] = $PagePath }
    }
    return $catalog
}

$pages = 0
$imports = 0
foreach ($file in Get-ChildItem -LiteralPath $UiRoot -Filter '*.xaml' -Recurse -File) {
    $page = Read-StyleXml $file.FullName
    if ($page.DocumentElement.LocalName -cne 'page') { continue }
    $paths = @(Get-StyleSourcePaths $page $file.FullName)
    $catalog = Get-PageStyleCatalog $page $file.FullName
    foreach ($node in $page.SelectNodes('//*[@style]')) {
        $name = $node.GetAttribute('style')
        if ($name.StartsWith('{Binding ')) { continue }
        if ($name.Length -gt 0 -and -not $catalog.ContainsKey($name)) {
            throw "Style is not declared: $($file.FullName) / $name"
        }
    }
    $pages++
    $imports += $paths.Count
}

$testPath = Join-Path $UiRoot 'PersonalEventBuilder\TestsPage.xaml'
$testPage = Read-StyleXml $testPath
$catalog = Get-PageStyleCatalog $testPage $testPath
if ($catalog['secondaryButton'] -ne $testPath) { throw 'Page-local style did not override the imported definition.' }
$mainPath = Join-Path $UiRoot 'PersonalEventBuilder\MainPage.xaml'
$mainPage = Read-StyleXml $mainPath
$catalog = Get-PageStyleCatalog $mainPage $mainPath
$commonPath = Join-Path ([IO.Path]::GetDirectoryName($mainPath)) 'CommonControlStyles.xaml'
if ($catalog['primaryButton'] -ne $commonPath) { throw 'Explicit shared style source was not loaded.' }

$negativeCases = @(
    @{ Value = ('missing-' + [Guid]::NewGuid().ToString() + '.xaml'); Error = 'Page style source file not found:' },
    @{ Value = 'CommonControlStyles.xaml;'; Error = 'Empty file path' },
    @{ Value = 'CommonControlStyles.xaml;./CommonControlStyles.xaml'; Error = 'Duplicate page style source:' }
)
foreach ($case in $negativeCases) {
    $page = [Xml.XmlDocument]::new()
    $page.LoadXml('<page xmlns="urn:excelprototype:profiles"/>')
    $page.DocumentElement.SetAttribute('styleSources', $case.Value)
    $rejected = $false
    try {
        $null = @(Get-StyleSourcePaths $page $mainPath)
    } catch {
        if (-not $_.Exception.Message.StartsWith($case.Error)) { throw }
        $rejected = $true
    }
    if (-not $rejected) { throw "Invalid source list was accepted: $($case.Value)" }
}

$noImports = [Xml.XmlDocument]::new()
$noImports.LoadXml('<page xmlns="urn:excelprototype:profiles"/>')
if (@(Get-StyleSourcePaths $noImports $mainPath).Count -ne 0) { throw 'Implicit style import is still active.' }
Write-Output "PASS: $pages pages, $imports explicit style imports, local overrides and 3 invalid source lists. Office/VBA were not executed."