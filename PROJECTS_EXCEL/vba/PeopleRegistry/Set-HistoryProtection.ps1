function Set-HistoryProtection($Book, [hashtable]$Config, [bool]$Enabled = $true) {
    foreach ($key in @('Sheet.History','Table.History')) {
        if (-not $Config.ContainsKey('PeopleRegistry::' + $key)) { throw "Missing configuration key: PeopleRegistry::$key" }
    }
    $sheet = $Book.Worksheets.Item([string]$Config['PeopleRegistry::Sheet.History'])
    $source = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ex_PeopleRegistry.vba'), [Text.Encoding]::UTF8)
    $passwordMatch = [regex]::Match($source, '(?m)^Private Const HISTORY_PASSWORD As String = "((?:[^"]|"")*)"\r?$')
    if (-not $passwordMatch.Success) { throw 'Missing HISTORY_PASSWORD constant in ex_PeopleRegistry.vba.' }
    $password = $passwordMatch.Groups[1].Value.Replace('""','"')
    if ($sheet.ProtectContents) { [void]$sheet.Unprotect($password) }
    if (-not $Enabled) { return }
    $sheet.Cells.Locked = $true
    $sheet.ListObjects.Item([string]$Config['PeopleRegistry::Table.History']).ShowAutoFilter = $true
    $button = $null
    foreach ($candidate in $sheet.Shapes) {
        if ($candidate.Name -eq 'btnHistoryLock') { $button = $candidate; break }
    }
    if ($null -eq $button) { $button = $sheet.Shapes.AddShape(1,0,0,1,1); $button.Name = 'btnHistoryLock' }
    $anchor = $sheet.Range([string]$Config['PeopleRegistry::Layout.HistoryButtonCell'])
    $button.Left = $anchor.Left + [double]$Config['PeopleRegistry::Layout.HistoryButtonOffsetX']
    $button.Top = $anchor.Top + [double]$Config['PeopleRegistry::Layout.HistoryButtonOffsetY']
    $button.Width = [double]$Config['PeopleRegistry::Layout.HistoryButtonWidth']
    $button.Height = [double]$Config['PeopleRegistry::Layout.HistoryButtonHeight']
    foreach ($pair in @(@('Theme.Alternate','Fill'), @('Theme.Border','Line'), @('Theme.Accent','Text'))) {
        $hex = [string]$Config['PeopleRegistry::' + $pair[0]]
        $color = [double]([Convert]::ToInt32($hex.Substring(1,2),16) + 256 * [Convert]::ToInt32($hex.Substring(3,2),16) + 65536 * [Convert]::ToInt32($hex.Substring(5,2),16))
        switch ($pair[1]) {
            'Fill' { $button.Fill.ForeColor.RGB = $color }
            'Line' { $button.Line.ForeColor.RGB = $color }
            'Text' { $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $color }
        }
    }
    $button.TextFrame2.TextRange.Text = [string]$Config['PeopleRegistry::Button.UnlockHistory']
    $button.TextFrame2.TextRange.Font.Name = [string]$Config['PeopleRegistry::Theme.FontName']
    $button.TextFrame2.TextRange.Font.Size = [double]$Config['PeopleRegistry::Theme.ButtonFontSize']
    $button.TextFrame2.TextRange.Font.Bold = -1
    $button.OnAction = "'" + $Book.Name.Replace("'","''") + "'!ex_PeopleRegistry.ToggleHistoryProtection"
    $button.Placement = 3
    $button.Locked = $true
    [void]$sheet.Protect($password, $true, $true, $true, $true,
        $false, $false, $false, $false, $false, $false, $false, $false, $false, $true, $false)
    $sheet.EnableSelection = 0
}