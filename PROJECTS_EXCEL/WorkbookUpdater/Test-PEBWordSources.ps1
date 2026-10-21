Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$utf8 = [Text.UTF8Encoding]::new($false, $true)

function Read-Config([string]$Name) {
    $values = @{}
    $path = Join-Path $root "config/PersonalEventBuilder/$Name"
    foreach ($line in [IO.File]::ReadAllLines($path, $utf8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $fields = $line.Split([char]9)
        if ($fields.Count -ne 3) { throw "Invalid config row in ${Name}: $line" }
        if ($values.ContainsKey($fields[1])) { throw "Duplicate config key: $($fields[1])" }
        $values.Add($fields[1], $fields[2])
    }
    return $values
}

$test = Read-Config 'wsConfig.test.txt'
$production = Read-Config 'wsConfig.txt'
$prefix = 'PersonalEventBuilder::'
[xml]$routes = [IO.File]::ReadAllText((Join-Path $root 'workbook/PersonalEventBuilder/PersonalEventBuilderEventRoutes.xml'), $utf8)
$ns = [Xml.XmlNamespaceManager]::new($routes.NameTable)
$ns.AddNamespace('p', 'urn:excelprototype:profiles')
foreach ($event in $routes.SelectNodes('//p:event[p:word/@policy="Render"]', $ns)) {
    $key = $prefix + 'export.Word.textAnchors.' + $event.id
    if (-not $test.ContainsKey($key) -or -not $production.ContainsKey($key)) { throw "Unmapped event: $($event.id)" }
    if ($test[$key] -cne $production[$key]) { throw "TEST/PROD anchor mapping differs: $($event.id)" }
}
foreach ($field in @('WordCreateMs', 'WordInsertMs', 'WordVerifyMs', 'WordAppendMs', 'WordDocumentPath')) {
    foreach ($key in @('test.ui.resultFields')) {
        if ($field -notin ($test[$prefix + $key] -split '\|')) { throw "Missing $field in $key" }
    }
}

[xml]$wordXml = [IO.File]::ReadAllText((Join-Path $root 'workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml'), $utf8)
$wordNs = [Xml.XmlNamespaceManager]::new($wordXml.NameTable)
$wordNs.AddNamespace('p', 'urn:excelprototype:profiles')
foreach ($config in @($test, $production)) {
    $knownKeys = @{}
    foreach ($key in ($config[$prefix + 'export.Word.context.xmlKeys'] -split '\|')) {
        $knownKeys[$key] = $true
    }
    foreach ($node in $wordXml.SelectNodes('//p:derived | //p:aggregate', $wordNs)) {
        $knownKeys[$node.GetAttribute('name')] = $true
    }
    foreach ($node in $wordXml.SelectNodes('//p:text', $wordNs)) {
        foreach ($reference in [regex]::Matches($node.InnerText, '\[([A-Z][A-Za-z0-9_]*)\]')) {
            $key = $reference.Groups[1].Value
            if (-not $knownKeys.ContainsKey($key)) { throw "Undeclared Word context placeholder: $key" }
        }
    }
}

[xml]$page = [IO.File]::ReadAllText((Join-Path $root 'ui/PersonalEventBuilder/TestsPage.xaml'), $utf8)
$ns = [Xml.XmlNamespaceManager]::new($page.NameTable)
$ns.AddNamespace('c', 'urn:excelprototype:controls')
$checkbox = $page.SelectSingleNode('//c:checkbox[@name="TestWordReport"]', $ns)
if ($null -eq $checkbox -or $checkbox.GetAttribute('value') -cne '{Binding Path=TestSelection.WordReport}') {
    throw 'Missing Word report checkbox binding.'
}
foreach ($node in $page.SelectNodes('//*[@name]')) {
    if ($node.LocalName -in @('checkbox', 'button') -and ('chk_' + $node.GetAttribute('name')).Length -gt 31) {
        throw "Excel shape name exceeds 31 characters: $($node.GetAttribute('name'))"
    }
}

# Inspect the actual DOCX package without starting Word or executing VBA.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$templatePath = Join-Path (Join-Path $root 'workbook/PersonalEventBuilder') $test[$prefix + 'test.word.template']
$templateStream = [IO.File]::Open([IO.Path]::GetFullPath($templatePath), 'Open', 'Read', 'ReadWrite')
$zip = [IO.Compression.ZipArchive]::new($templateStream, [IO.Compression.ZipArchiveMode]::Read)
try {
    $reader = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
    try { [xml]$document = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $ns = [Xml.XmlNamespaceManager]::new($document.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $text = ($document.SelectNodes('//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
    $starts = @{}
    $ends = @{}
    foreach ($match in [regex]::Matches($text, '\{\s*\\\s*export\s*:[^{}]*\}')) {
        $marker = [regex]::Replace($match.Value, '[\s\u00a0]+', '')
        if ($marker -notmatch '^\{\\export:(.+)_(Begin|End)\}$') { throw "Invalid marker: $marker" }
        $identifier = $Matches[1]
        $kind = $Matches[2]
        $set = if ($kind -eq 'Begin') { $starts } else { $ends }
        if ($set.ContainsKey($identifier)) { throw "Duplicate $kind marker: $identifier" }
        $set.Add($identifier, $match.Index)
    }
    foreach ($identifier in $starts.Keys) {
        if (-not $ends.ContainsKey($identifier) -or $ends[$identifier] -le $starts[$identifier]) {
            throw "Unpaired or reversed anchor: $identifier"
        }
    }
    if ($starts.Count -eq 0 -or $starts.Count -ne $ends.Count) { throw 'Unpaired template markers.' }
    foreach ($key in $test.Keys | Sort-Object) {
        if ($key.StartsWith($prefix + 'export.Word.textAnchors.') -and -not $starts.ContainsKey($test[$key])) {
            Write-Warning "Template lacks anchor '$($test[$key])'; inserting this event must fail explicitly."
        }
    }
    Write-Output "DOCX inspected: $($starts.Count) paired anchors."
} finally {
    $zip.Dispose()
    $templateStream.Dispose()
}

$classes = @{}
foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'vba/PersonalEventBuilder') -Recurse -Filter '*.cls.vba') {
    $source = [IO.File]::ReadAllText($file.FullName, $utf8)
    if ($source -match 'Attribute VB_Name = "([^"]+)"') { $classes.Add($Matches[1], $source) }
}
foreach ($name in @('obj_PEB_WordAnchors', 'obj_PEB_TestWordSession')) {
    if (-not $classes.ContainsKey($name) -or $name.Length -gt 31) { throw "Invalid class name: $name" }
    $source = $classes[$name]
    $procedures = [regex]::Matches($source, '(?m)^(Public|Private) (Sub|Function) ').Count
    $endings = [regex]::Matches($source, '(?m)^End (Sub|Function)\r?$').Count
    if ($procedures -ne $endings) { throw "Unbalanced procedures in $name" }
    foreach ($procedure in [regex]::Matches($source, '(?ms)^(?:Public|Private) (?:Sub|Function|Property Get) (\w+)\b.*?^End (?:Sub|Function|Property)\r?$')) {
        $declared = @{}
        $declared[$procedure.Groups[1].Value] = $true
        foreach ($declaration in [regex]::Matches($source, '(?m)^Private (\w+) As ')) {
            $declared[$declaration.Groups[1].Value] = $true
        }
        foreach ($declaration in [regex]::Matches($procedure.Value, '(?:Dim |ByVal |ByRef )(\w+) As ')) {
            $declared[$declaration.Groups[1].Value] = $true
        }
        foreach ($assignment in [regex]::Matches($procedure.Value, '(?m)^\s*(?:Set )?(\w+)\s*=')) {
            if (-not $declared.ContainsKey($assignment.Groups[1].Value)) {
                throw "Undeclared assignment in ${name}: $($assignment.Groups[1].Value)"
            }
        }
    }
}
Write-Output 'Word source/configuration checks passed; Word/VBA runtime was not executed.'