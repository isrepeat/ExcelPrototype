param(
    [string]$ProjectRoot = (Join-Path $PSScriptRoot '..\..\PROJECTS_EXCEL')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$config = @{}
foreach ($line in [IO.File]::ReadAllLines((Join-Path $ProjectRoot 'config\PersonalEventBuilder\wsConfig.txt'))) {
    if ($line.Length -eq 0) { continue }
    $cells = $line.Split([char]9)
    if ($cells.Length -ne 3 -or $config.ContainsKey($cells[1])) { throw "Invalid or duplicate configuration row: $line" }
    $config[$cells[1]] = $cells[2]
}

function Read-Config([string]$Key) {
    $keyName = "PersonalEventBuilder::$Key"
    if (-not $config.ContainsKey($keyName) -or [string]::IsNullOrWhiteSpace($config[$keyName])) { throw "Missing configuration key: $keyName" }
    return $config[$keyName]
}

# Проверяем XML и контракт данных без открытия Excel и выполнения VBA.
$xmlSettings = [Xml.XmlReaderSettings]::new()
$xmlSettings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
$xmlSettings.XmlResolver = $null
$xmlPath = Join-Path (Join-Path $ProjectRoot 'workbook\PersonalEventBuilder') (Read-Config 'export.Word.templates.path')
$xmlReader = [Xml.XmlReader]::Create($xmlPath, $xmlSettings)
try {
    $document = [Xml.XmlDocument]::new()
    $document.XmlResolver = $null
    $document.Load($xmlReader)
}
finally { $xmlReader.Dispose() }
$namespaceManager = [Xml.XmlNamespaceManager]::new($document.NameTable)
$namespaceManager.AddNamespace('p', 'urn:excelprototype:profiles')
$root = $document.SelectSingleNode('/p:wordResultTemplates', $namespaceManager)
if ($null -eq $root -or $root.GetAttribute('version') -ne '1') { throw 'Unsupported XML root or version.' }
$shared = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
$templates = [Collections.Generic.Dictionary[string,Xml.XmlElement]]::new([StringComparer]::Ordinal)
foreach ($node in $root.SelectNodes('p:sharedTemplates/p:sharedTemplate', $namespaceManager)) {
    $name = $node.GetAttribute('id')
    $texts = $node.SelectNodes('p:text', $namespaceManager)
    if ($name.Length -eq 0 -or $shared.ContainsKey($name) -or $texts.Count -ne 1) { throw "Invalid shared template: $name" }
    $shared.Add($name, $texts[0].InnerText)
}
foreach ($node in $root.SelectNodes('p:template', $namespaceManager)) {
    $name = $node.GetAttribute('name')
    $texts = $node.SelectNodes('p:text', $namespaceManager)
    if ($name.Length -eq 0 -or $templates.ContainsKey($name) -or $texts.Count -ne 1) { throw "Invalid template: $name" }
    $templates.Add($name, $node)
}

function Expand-Includes([string]$Text, [string[]]$Stack = @()) {
    foreach ($match in [regex]::Matches($Text, '\{#include\s+([^}]+)\}')) {
        $name = $match.Groups[1].Value.Trim()
        if (-not $shared.ContainsKey($name)) { throw "Unknown XML include: $name" }
        if ($name -cin $Stack) { throw "Cyclic XML include: $($Stack -join ' -> ') -> $name" }
        if ($Stack.Count -ge [int](Read-Config 'export.dsl.maxDepth')) { throw 'XML include-depth limit exceeded.' }
        $body = Expand-Includes $shared[$name] @($Stack + $name)
        $Text = $Text.Replace($match.Value, $body)
    }
    return $Text
}

function Add-Declared([Collections.Generic.HashSet[string]]$Fields, [string]$Name) {
    if (-not $Fields.Add($Name)) { throw "Duplicate data field: $Name" }
}

foreach ($month in 1..12) { Read-Config "export.dsl.months.$month" | Out-Null }
$baseFields = (Read-Config 'export.fields').Split('|')
$modeKeys = @($config.Keys | Where-Object { $_ -match '^PersonalEventBuilder::export\.Word\.([^.]+)\.templateName$' } | Sort-Object)
if ($modeKeys.Count -ne 8) { throw "Expected eight mapped modes, found $($modeKeys.Count)." }
foreach ($modeKey in $modeKeys) {
    $mode = [regex]::Match($modeKey, 'export\.Word\.([^.]+)\.templateName$').Groups[1].Value
    $templateName = $config[$modeKey]
    if (-not $templates.ContainsKey($templateName)) { throw "Missing XML template for $($mode): $templateName" }
    $fields = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($field in $baseFields) { Add-Declared $fields $field }
    Add-Declared $fields 'EventName'
    foreach ($required in (Read-Config "export.Word.$mode.previewRequired").Split('|')) {
        if (-not $fields.Contains($required)) { throw "Unknown required field: $required" }
    }
    foreach ($lookupRoot in @('export.shared.lookups', "export.Word.$mode.lookups")) {
        $count = [int](Read-Config "$lookupRoot.count")
        for ($index = 1; $index -le $count; $index++) {
            $prefix = "$lookupRoot.$index"
            foreach ($key in @('input','when.field')) {
                $field = Read-Config "$prefix.$key"
                if (-not $fields.Contains($field)) { throw "Unknown lookup input: $prefix.$key=$field" }
            }
            [regex]::new((Read-Config "$prefix.when.pattern")) | Out-Null
            foreach ($key in @('path','sheet','range','keyHeader','valueHeader')) { Read-Config "$prefix.$key" | Out-Null }
            Add-Declared $fields (Read-Config "$prefix.output")
        }
    }
    for ($index = 1; $index -le [int](Read-Config 'export.shared.constants.count'); $index++) {
        $prefix = "export.shared.constants.$index"
        $kind = Read-Config "$prefix.kind"
        if ($kind -notin @('text','emptyText','boolean','emptyCollection')) { throw "Unknown constant kind: $kind" }
        if ($kind -in @('text','boolean')) { Read-Config "$prefix.value" | Out-Null }
        Add-Declared $fields (Read-Config "$prefix.name")
    }
    for ($index = 1; $index -le [int](Read-Config 'export.shared.aliases.count'); $index++) {
        $prefix = "export.shared.aliases.$index"
        if ((Read-Config "$prefix.kind") -notin @('value','shortDate','joinLowerFirst')) { throw "Unknown alias kind: $prefix" }
        foreach ($field in (Read-Config "$prefix.source").Split('|')) {
            if (-not $fields.Contains($field)) { throw "Unknown alias source: $field" }
        }
        Add-Declared $fields (Read-Config "$prefix.name")
    }
    $node = $templates[$templateName]
    $text = ''
    $headers = @{}
    foreach ($header in $node.SelectNodes('p:groupHeader', $namespaceManager)) {
        $by = $header.GetAttribute('by')
        if ($by.Length -eq 0 -or $headers.ContainsKey($by)) { throw "Invalid groupHeader: $templateName" }
        $headers[$by] = $header.InnerText
    }
    if ($node.HasAttribute('groupBy')) {
        foreach ($by in $node.GetAttribute('groupBy').Split(';')) {
            $by = $by.Trim()
            if (-not $headers.ContainsKey($by)) { throw "Missing groupHeader for $($templateName): $by" }
            $text += $headers[$by]
            $headers.Remove($by)
        }
    }
    if ($headers.Count -ne 0) { throw "Unmatched groupHeader: $templateName" }
    $text += $node.SelectSingleNode('p:text', $namespaceManager).InnerText
    $text = Expand-Includes $text
    $stack = [Collections.Generic.List[string]]::new()
    foreach ($match in [regex]::Matches($text, '\{#(if|else|endif|for|endfor)\b([^}]*)\}')) {
        $kind = $match.Groups[1].Value
        if ($kind -in @('if','for')) { $stack.Add($kind) }
        elseif ($kind -eq 'else') {
            if ($stack.Count -eq 0 -or $stack[$stack.Count - 1] -ne 'if') { throw "Misplaced #else: $templateName" }
        }
        else {
            $expected = if ($kind -eq 'endif') { 'if' } else { 'for' }
            if ($stack.Count -eq 0 -or $stack[$stack.Count - 1] -ne $expected) { throw "Unbalanced DSL blocks: $templateName" }
            $stack.RemoveAt($stack.Count - 1)
        }
    }
    if ($stack.Count -ne 0) { throw "Unclosed DSL blocks: $templateName" }
    $references = @([regex]::Matches($text, '\{\[([^\]]+)\]') | ForEach-Object { $_.Groups[1].Value })
    foreach ($condition in [regex]::Matches($text, '\{#if\s+([^}]+)\}')) {
        $references += @([regex]::Matches($condition.Groups[1].Value, '(?<![\w.])\[([^\]]+)\]') | ForEach-Object { $_.Groups[1].Value })
    }
    foreach ($field in $references) {
        if (-not $fields.Contains($field)) { throw "Undeclared DSL field in $($templateName): $field" }
    }
    foreach ($match in [regex]::Matches($text, '#let\s+(__\w+)\s*=\s*(\$[\w.]+)\(')) {
        if ($match.Groups[2].Value -notin @('$ex_Helpers.m_RegexIsMatch','$ex_Helpers.fn_IsDateOnOrAfter')) { throw "Unsupported #let function: $($match.Groups[2].Value)" }
    }
    Write-Output "XML preview contract OK: $mode -> $templateName"
}
Write-Output "Validated $($templates.Count) XML templates, $($shared.Count) shared blocks and eight preview contracts. Office/VBA were not executed."