$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectRoot = Split-Path $PSScriptRoot -Parent
$controller = [IO.File]::ReadAllText((Join-Path $projectRoot 'vba/PersonalEventBuilder/obj_PEB_PgMainController.cls.vba'))
$export = [regex]::Match($controller, '(?s)Private Function private_Export\(.*?End Function').Value
$preview = [regex]::Match($controller, '(?s)Private Function private_PublishSavedPreview\(.*?End Function').Value
if ($export -notmatch '(?s)ElseIf exporter = "History" Then.*?success = ex_PEB_History.fn_Save.*?If success Then\s+If Not private_PublishSavedPreview\(values, previewDiagnostic\) Then') {
    throw 'Preview must be published only after a successful history save.'
}
if ($export -notmatch 'previewWarning = True' -or
    $export -notmatch 'private_SetActionStatus status, previewWarning' -or
    $export -notmatch 'private_Export = True') { throw 'A preview warning must preserve the successful save result.' }
if ($preview -notmatch 'If values\("WordPolicy"\) <> "Omit" Then' -or
    $preview -notmatch 'fn_Required\(values, "WordText"\)' -or
    $preview -notmatch 'SetValue\("Form", "WordPreviewText", VBA.vbNullString\)' -or
    $preview -notmatch 'SetValue\("Form", "WordPreviewText", text\)') { throw 'Preview must reuse prepared text and clear stale content for omitted templates.' }
if ($preview -match 'TryPrepare|TryExport|fn_Save|MsgBox|fn_ShowMessage') { throw 'Preview publication must not prepare, export or save another document/event.' }
foreach ($file in @('wsConfig.txt', 'wsConfig.test.txt')) {
    $rows = @([IO.File]::ReadAllLines((Join-Path $projectRoot "config/PersonalEventBuilder/$file")) | Where-Object { $_.Split([char]9)[1] -eq 'PersonalEventBuilder::text.HistoryPreviewFailed' })
    if ($rows.Count -ne 1 -or $rows[0].Split([char]9).Count -ne 3 -or -not $rows[0].Split([char]9)[2]) { throw "Missing preview-after-save warning: $file" }
}
Write-Output 'PASS: saved-event preview ordering, omitted templates, text reuse and save-success preservation.'