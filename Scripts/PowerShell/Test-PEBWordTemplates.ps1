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
if ($null -eq $root -or $root.GetAttribute('version') -ne '2') { throw 'Unsupported XML root or version.' }
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
    $bodies = $node.SelectNodes('p:text | p:group', $namespaceManager)
    if ($name.Length -eq 0 -or $templates.ContainsKey($name) -or $bodies.Count -ne 1) { throw "Invalid template: $name" }
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
foreach ($templateName in $templates.Keys) {
    $fields = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($field in (Read-Config 'export.Word.context.xmlKeys').Split('|')) { Add-Declared $fields $field }
    $node = $templates[$templateName]
    foreach ($rule in $node.SelectNodes('p:reject', $namespaceManager)) {
        $field = $rule.GetAttribute('field')
        if (-not $fields.Contains($field) -and $field -ne 'VacationAdditionalDays') {
            throw "Unknown rejection field in ${templateName}: $field"
        }
        $messageKey = $rule.GetAttribute('messageKey')
        Read-Config $messageKey | Out-Null
        if ($rule.HasAttribute('equals') -eq $rule.HasAttribute('matches')) {
            throw "Rejection requires exactly one of equals or matches: $templateName"
        }
        if ($rule.HasAttribute('matches')) { [void][regex]::new($rule.GetAttribute('matches')) }
    }
    foreach ($computed in $node.SelectNodes('.//p:derived | .//p:aggregate', $namespaceManager)) {
        [void]$fields.Add($computed.GetAttribute('name'))
    }
    foreach ($textNode in $node.SelectNodes('.//p:text', $namespaceManager)) {
        $text = $textNode.InnerText
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
        foreach ($assignment in [regex]::Matches($text, '#let\s+(__\w+)\s*=')) {
            $fields.Add($assignment.Groups[1].Value) | Out-Null
        }
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
    }
    Write-Output "XML DSL contract OK: $templateName"
}
Write-Output "Validated $($templates.Count) XML templates, $($shared.Count) shared blocks and all template DSL blocks. Office/VBA were not executed."