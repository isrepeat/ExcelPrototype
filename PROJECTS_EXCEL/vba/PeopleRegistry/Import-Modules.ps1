function Import-PeopleRegistryModules($Book, [string]$VbaRoot) {
    $utf8 = [Text.UTF8Encoding]::new($false, $true)
    $manifest = [IO.File]::ReadAllText((Join-Path $VbaRoot 'modules.json'), $utf8) | ConvertFrom-Json
    foreach ($pattern in $manifest.PeopleRegistry) {
        if ($pattern.Contains('*') -or $pattern.Contains('?')) {
            $files = @(Get-ChildItem -Path (Join-Path $VbaRoot $pattern) -File)
        } else {
            $files = @(Get-Item -LiteralPath (Join-Path $VbaRoot $pattern))
        }
        if ($files.Count -ne 1) {
            throw "Expected exactly one module for profile entry: $pattern"
        }
        $file = $files[0]
        $source = [IO.File]::ReadAllText($file.FullName, $utf8)
        $name = $file.Name -replace '(\.interface|\.base)?\.cls\.vba$|\.vba$', ''
        $componentType = 1
        if ($file.Name.EndsWith('.cls.vba')) {
            $componentType = 2
            $source = $source -replace '(?ms)\AVERSION[^\r\n]*\r?\nBEGIN\r?\n.*?^END\r?\n', ''
        }
        $source = $source -replace '(?m)^Attribute [^\r\n]*\r?\n', ''
        $source = $source.Replace("`r`n", "`n").Replace("`n", "`r`n")
        if ($name -eq 'ThisWorkbook') {
            $name = $Book.CodeName
        }
        $component = $null
        foreach ($candidate in $Book.VBProject.VBComponents) {
            if ($candidate.Name -eq $name) {
                $component = $candidate
                break
            }
        }
        if ($null -eq $component) {
            $component = $Book.VBProject.VBComponents.Add($componentType)
            $component.Name = $name
        }
        if ($component.CodeModule.CountOfLines -gt 0) {
            [void]$component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
        }
        [void]$component.CodeModule.AddFromString($source)
    }
}