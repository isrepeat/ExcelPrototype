Set-StrictMode -Version Latest

function Find-LoadedExcelWorkbook([string]$Name) {
    if (-not ('WorkbookModules.ExcelWindows' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace WorkbookModules {
    public static class ExcelWindows {
        private delegate bool WindowCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(WindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid interfaceId,
            [MarshalAs(UnmanagedType.Interface)] out object result);

        public static object[] GetNativeWindows() {
            var results = new List<object>();
            EnumWindows((parent, parameter) => {
                var name = new StringBuilder(256);
                GetClassName(parent, name, name.Capacity);
                if (name.ToString() != "XLMAIN") return true;
                EnumChildWindows(parent, (child, ignored) => {
                    var childName = new StringBuilder(256);
                    GetClassName(child, childName, childName.Capacity);
                    if (childName.ToString() != "EXCEL7") return true;
                    var interfaceId = new Guid("00020400-0000-0000-C000-000000000046");
                    object nativeWindow;
                    if (AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref interfaceId, out nativeWindow) == 0)
                        results.Add(nativeWindow);
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            return results.ToArray();
        }
    }
}
'@
    }
    $applications = New-Object 'System.Collections.Generic.List[object]'
    try { $applications.Add([Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')) }
    catch [Runtime.InteropServices.COMException] { }
    foreach ($window in [WorkbookModules.ExcelWindows]::GetNativeWindows()) {
        $applications.Add($window.Application)
    }
    $seenApplications = @{}
    $matches = New-Object 'System.Collections.Generic.List[object]'
    $availableBooks = New-Object 'System.Collections.Generic.List[string]'
    foreach ($application in $applications) {
        $identity = [string]$application.Hwnd
        if ($seenApplications.ContainsKey($identity)) { continue }
        $seenApplications[$identity] = $true
        foreach ($workbook in $application.Workbooks) {
            $availableBooks.Add($workbook.Name)
        }
        # Loaded add-ins are accessible by name but may be absent from the Workbooks enumeration.
        try {
            $namedWorkbook = $application.Workbooks.Item($Name)
            $matches.Add([pscustomobject]@{ Excel = $application; Workbook = $namedWorkbook })
        }
        catch [Runtime.InteropServices.COMException] { }
    }
    if ($matches.Count -gt 1) { throw "More than one Excel instance contains '$Name'. Close the duplicate target before updating." }
    if ($matches.Count -eq 0) {
        throw "Workbook is not loaded: $Name. Available workbooks: $($availableBooks -join ', ')"
    }
    return $matches[0]
}

function Read-WorkbookSourceText([string]$Path, [bool]$Native = $false) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($Native) {
        # Native VBE import depends on ANSI: accept only portable ASCII text.
        foreach ($value in $bytes) {
            if ($value -gt 127) { throw "Native source must be ASCII; convert code to UTF-8 .vba or remove non-ASCII designer text: $Path" }
        }
    }
    $encoding = New-Object Text.UTF8Encoding($false, $true)
    try { $text = $encoding.GetString($bytes) }
    catch { throw "Invalid UTF-8 source: $Path. $($_.Exception.Message)" }
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    if ($text.Contains([string][char]0xFFFD)) { throw "Source contains Unicode replacement character: $Path" }
    return $text
}

function Get-WorkbookProfile([object]$Workbook) {
    foreach ($sheet in $Workbook.Worksheets) {
        foreach ($table in $sheet.ListObjects) {
            if ($table.Name -ine 'tbConfig' -or $null -eq $table.DataBodyRange) { continue }
            $keyColumn = $table.ListColumns.Item('Key').Index
            if ($keyColumn -ge $table.ListColumns.Count) { throw 'tbConfig has no value column after Key.' }
            foreach ($row in $table.ListRows) {
                if ([string]$row.Range.Cells.Item(1, $keyColumn).Value2 -ieq 'ThisWorkbook::id') {
                    $value = ([string]$row.Range.Cells.Item(1, $keyColumn + 1).Value2).Trim()
                    if (-not $value) { throw 'ThisWorkbook::id is empty.' }
                    return $value
                }
            }
        }
    }
    throw 'Specify Profile or provide tbConfig with ThisWorkbook::id.'
}

function ConvertTo-WorkbookModuleCode([string]$Text) {
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $insideHeader = $false
    foreach ($line in ($Text -split '\r?\n')) {
        $trimmed = $line.Trim()
        if ($trimmed -ieq 'VERSION 1.0 CLASS') { $insideHeader = $true; continue }
        if ($insideHeader) {
            if ($trimmed -ieq 'END') { $insideHeader = $false }
            continue
        }
        if ($trimmed -match '^Attribute\s') { continue }
        # Encode Unicode string literals with ChrW$, independently of the VBE code page.
        $result = New-Object Text.StringBuilder
        $index = 0
        while ($index -lt $line.Length) {
            $character = $line[$index]
            if ($character -eq "'" -or ($line.Substring($index) -match '^Rem(?:\s|$)' -and ($index -eq 0 -or $line[$index - 1] -match '[\s:]'))) {
                [void]$result.Append($line.Substring($index)); break
            }
            if ($character -ne '"') {
                if ([int]$character -gt 127) { throw 'Non-ASCII VBA identifier is not portable. Use ASCII identifiers.' }
                [void]$result.Append($character); $index++; continue
            }
            $start = $index
            $index++
            $literal = New-Object Text.StringBuilder
            $closed = $false
            while ($index -lt $line.Length) {
                if ($line[$index] -eq '"') {
                    if ($index + 1 -lt $line.Length -and $line[$index + 1] -eq '"') {
                        [void]$literal.Append('"'); $index += 2; continue
                    }
                    $index++; $closed = $true; break
                }
                [void]$literal.Append($line[$index]); $index++
            }
            if (-not $closed) { throw 'Unterminated VBA string literal.' }
            if ($literal.ToString() -notmatch '[^\x00-\x7F]') {
                [void]$result.Append($line.Substring($start, $index - $start)); continue
            }
            if ($line -match '(?i)\bConst\s') { throw 'Unicode Const literals require runtime initialization; ChrW is not a constant expression.' }
            $parts = New-Object 'System.Collections.Generic.List[string]'
            $ascii = New-Object Text.StringBuilder
            foreach ($unit in $literal.ToString().ToCharArray()) {
                if ([int]$unit -le 127) { [void]$ascii.Append($unit); continue }
                if ($ascii.Length -gt 0) { $parts.Add('"' + $ascii.ToString().Replace('"', '""') + '"'); [void]$ascii.Clear() }
                $number = [int]$unit
                if ($number -gt 32767) { $number -= 65536 }
                $parts.Add("VBA.ChrW`$($number)")
            }
            if ($ascii.Length -gt 0) { $parts.Add('"' + $ascii.ToString().Replace('"', '""') + '"') }
            [void]$result.Append('(' + ($parts -join ' & ') + ')')
        }
        $lines.Add($result.ToString())
    }
    if ($insideHeader) { throw 'Unterminated class export header.' }
    return $lines -join "`r`n"
}

function Get-WorkbookModulePlan([string]$SourceFolder, [string]$Profile) {
    $root = (Resolve-Path -LiteralPath $SourceFolder -ErrorAction Stop).Path.TrimEnd('\', '/')
    $schemaPath = Join-Path $root 'modules.json'
    $schema = (Read-WorkbookSourceText $schemaPath) | ConvertFrom-Json
    $property = $schema.PSObject.Properties[$Profile]
    if ($null -eq $property -or $property.Value -isnot [Array] -or $property.Value.Count -eq 0) {
        throw "Profile '$Profile' must contain a nonempty array in $schemaPath"
    }
    $candidates = @(Get-ChildItem -LiteralPath $root -Recurse -File)
    $selected = New-Object 'System.Collections.Generic.List[object]'
    $seenPaths = @{}
    foreach ($pattern in $property.Value) {
        if ($pattern -isnot [string] -or [string]::IsNullOrWhiteSpace($pattern)) { throw 'Module patterns must be nonempty strings.' }
        $normalized = $pattern.Replace('/', '\')
        if ($normalized.StartsWith('\') -or $normalized.Contains(':') -or $normalized.Contains('..')) {
            throw "Invalid relative module path: $pattern"
        }
        if ([IO.Path]::GetExtension($normalized) -notin @('.vba', '.bas', '.cls', '.frm')) { throw "Unsupported module pattern: $pattern" }
        # Use the same wildcard syntax, including [[] for a literal opening bracket.
        $matcher = New-Object Management.Automation.WildcardPattern($normalized, [Management.Automation.WildcardOptions]::IgnoreCase)
        $matches = @($candidates | Where-Object {
            $matcher.IsMatch($_.FullName.Substring($root.Length + 1))
        } | Sort-Object FullName)
        if ($matches.Count -eq 0) { throw "No source files match: $pattern" }
        foreach ($file in $matches) {
            if (-not $seenPaths.ContainsKey($file.FullName)) { $selected.Add($file); $seenPaths[$file.FullName] = $true }
        }
    }
    $seenNames = @{}
    foreach ($file in $selected) {
        $native = $file.Extension -ine '.vba'
        $text = (Read-WorkbookSourceText $file.FullName $native)
        if ([string]::IsNullOrWhiteSpace($text)) { throw "Empty source: $($file.FullName)" }
        $stem = $file.Name -replace '(?i)\.(vba|bas|cls|frm)$', '' -replace '(?i)\.utf8$', '' -replace '(?i)\.(cls|bas|frm)$', ''
        $document = $stem -ieq 'ThisWorkbook' -or ($file.Extension -ieq '.vba' -and $stem.StartsWith('ws_', [StringComparison]::OrdinalIgnoreCase))
        $attribute = [regex]::Match($text, '(?im)^\s*Attribute VB_Name\s*=\s*"([^"]+)"')
        $name = if ($attribute.Success -and -not $document) { $attribute.Groups[1].Value } else { $stem }
        $type = if ($document) { 100 } elseif ($file.Name -match '(?i)\.cls(\.utf8)?(\.vba)?$') { 2 } elseif ($file.Name -match '(?i)\.frm(\.utf8)?(\.vba)?$' -or $text -match '(?im)^\s*BEGIN VB\.Form\b') { 3 } else { 1 }
        if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,39}$' -and -not ($document -and $stem -like 'ws_*')) { throw "Invalid component name: $name" }
        if ($seenNames.ContainsKey($name)) { throw "Duplicate component name: $name" }
        $seenNames[$name] = $true
        if ($type -eq 3 -and -not $native) { throw "UserForms require native .frm/.frx exports: $($file.FullName)" }
        $resource = [IO.Path]::ChangeExtension($file.FullName, '.frx')
        if ($type -eq 3 -and $text -match '(?i)\.frx' -and -not (Test-Path -LiteralPath $resource -PathType Leaf)) { throw "Missing form resource: $resource" }
        [pscustomobject]@{
            Name = $name; Type = $type; Document = $document; Native = $native
            Path = $file.FullName; SheetName = if ($stem -like 'ws_*') { $stem.Substring(3) } else { $null }
            Code = if (-not $native -or $document) { ConvertTo-WorkbookModuleCode $text } else { $null }
            ComponentName = $null
            NativeBytes = if ($native -and -not $document) { [IO.File]::ReadAllBytes($file.FullName) } else { $null }
            ResourceBytes = if ($type -eq 3 -and (Test-Path -LiteralPath $resource -PathType Leaf)) { [IO.File]::ReadAllBytes($resource) } else { $null }
        }
    }
}

function Assert-WorkbookModulePlan([object]$Workbook, [object[]]$Plan) {
    $names = @{}
    foreach ($item in $Plan) {
        $item.ComponentName = $item.Name
        if ($item.Document) {
            if ($null -eq $item.SheetName) { $item.ComponentName = $Workbook.CodeName }
            else { $item.ComponentName = $Workbook.Worksheets.Item($item.SheetName).CodeName }
            $component = $Workbook.VBProject.VBComponents.Item($item.ComponentName)
            if ($component.Type -ne 100) { throw "Not a document component: $($item.ComponentName)" }
        }
        elseif ($item.Native) {
            if (-not (Test-Path -LiteralPath $item.Path -PathType Leaf)) { throw "Source no longer exists: $($item.Path)" }
        }
        if ($names.ContainsKey($item.ComponentName)) { throw "Duplicate target component: $($item.ComponentName)" }
        $names[$item.ComponentName] = $true
    }
    foreach ($component in $Workbook.VBProject.VBComponents) {
        if ($component.Type -notin @(1, 2, 3, 100)) { throw "Unsupported component type: $($component.Type)" }
        if ($component.Type -eq 100 -and $names.ContainsKey($component.Name)) {
            $source = @($Plan | Where-Object { $_.ComponentName -ieq $component.Name })[0]
            if (-not $source.Document) { throw "Module name conflicts with document: $($component.Name)" }
        }
    }
}

function Set-WorkbookModulePlan([object]$Workbook, [object[]]$Plan, [string]$Mode) {
    $stagedFiles = New-Object 'System.Collections.Generic.List[string]'
    $stagedFolders = New-Object 'System.Collections.Generic.List[string]'
    try {
        # Native import uses a byte snapshot instead of the mutable source file.
        foreach ($item in $Plan) {
            if (-not $item.Native -or $item.Document) { continue }
            $folder = Join-Path ([IO.Path]::GetTempPath()) ('WorkbookModules-' + [Guid]::NewGuid().ToString('N'))
            [IO.Directory]::CreateDirectory($folder) | Out-Null
            $stagedFolders.Add($folder)
            $item.Path = Join-Path $folder ([IO.Path]::GetFileName($item.Path))
            [IO.File]::WriteAllBytes($item.Path, $item.NativeBytes)
            $stagedFiles.Add($item.Path)
            if ($null -ne $item.ResourceBytes) {
                $resourcePath = [IO.Path]::ChangeExtension($item.Path, '.frx')
                [IO.File]::WriteAllBytes($resourcePath, $item.ResourceBytes)
                $stagedFiles.Add($resourcePath)
            }
        }
        $project = $Workbook.VBProject
        for ($index = $project.VBComponents.Count; $index -ge 1; $index--) {
            $component = $project.VBComponents.Item($index)
            if ($component.Type -in @(1, 2, 3)) { $project.VBComponents.Remove($component) }
            elseif ($component.Type -eq 100 -and $component.CodeModule.CountOfLines -gt 0) {
                $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
            }
        }
        if ($Mode -eq 'Clear') { return }
        foreach ($item in $Plan) {
            if ($item.Document) { $component = $project.VBComponents.Item($item.ComponentName) }
            elseif ($item.Native) { $component = $project.VBComponents.Import($item.Path) }
            else { $component = $project.VBComponents.Add($item.Type); $component.Name = $item.Name }
            if (-not $item.Native -or $item.Document) { $component.CodeModule.AddFromString($item.Code) }
            if ($component.Name -cne $item.ComponentName -or $component.Type -ne $item.Type) { throw "Imported identity mismatch: $($item.Name)" }
            Write-Output "Imported: $($component.Name)"
        }
    }
    finally {
        foreach ($path in $stagedFiles) { [IO.File]::Delete($path) }
        foreach ($path in $stagedFolders) { [IO.Directory]::Delete($path, $false) }
    }
}