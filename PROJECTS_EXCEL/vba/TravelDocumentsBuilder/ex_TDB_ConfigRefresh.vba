Option Explicit

Private Const HOTKEY As String = "^%u"
Private m_isBound As Boolean

' namespace Lifecycle {
Public Sub fn_Module_Dispose()
    fn_Deactivate
End Sub
' } // namespace Lifecycle

' namespace API {
Public Sub fn_Activate()
    Application.OnKey HOTKEY, "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_TDB_ConfigRefresh.fn_Update"
    m_isBound = True
End Sub

Public Sub fn_Deactivate()
    If Not m_isBound Then Exit Sub
    Application.OnKey HOTKEY
    m_isBound = False
End Sub

Public Sub fn_Update()
    Dim runtimeContext As Object
    Dim fileSystem As Object
    Dim stream As Object
    Dim configPath As String
    Dim text As String
    Dim lines As Variant
    Dim fields As Variant
    Dim line As Variant
    Dim records As Collection
    Dim keys As Object
    Dim values As Variant
    Dim originalValues As Variant
    Dim originalRange As Range
    Dim resizedRange As Range
    Dim releasedRange As Range
    Dim table As ListObject
    Dim sheet As Worksheet
    Dim candidate As ListObject
    Dim index As Long
    Dim keyColumn As Long
    Dim kindColumn As Long
    Dim valueColumn As Long
    Dim previousEvents As Boolean
    Dim previousUpdating As Boolean
    Dim guarded As Boolean
    Dim changed As Boolean
    Dim diagnostic As String
    Dim failureTemplate As String
    Dim successTemplate As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    configPath = fileSystem.GetAbsolutePathName(fileSystem.BuildPath( _
        ThisWorkbook.Path, "../../config/TravelDocumentsBuilder/wsConfig.txt"))
    Set stream = VBA.CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile configPath
    text = stream.ReadText
    stream.Close
    text = VBA.Replace$(text, VBA.vbCrLf, VBA.vbLf)
    text = VBA.Replace$(text, VBA.vbCr, VBA.vbLf)
    If VBA.Left$(text, 1) = VBA.ChrW$(-257) Then text = VBA.Mid$(text, 2)
    lines = VBA.Split(text, VBA.vbLf)
    Set records = New Collection
    Set keys = VBA.CreateObject("Scripting.Dictionary")
    keys.CompareMode = VBA.vbTextCompare
    For Each line In lines
        If VBA.Len(VBA.Trim$(VBA.CStr(line))) > 0 Then
            fields = VBA.Split(VBA.CStr(line), VBA.vbTab)
            If UBound(fields) <> 2 Then VBA.Err.Raise 5, , "Invalid configuration row: " & VBA.CStr(line)
            If VBA.Len(VBA.Trim$(fields(0))) = 0 Or VBA.Len(VBA.Trim$(fields(1))) = 0 Then
                VBA.Err.Raise 5, , "Empty configuration kind or key."
            End If
            If keys.Exists(VBA.Trim$(fields(1))) Then VBA.Err.Raise 5, , "Duplicate configuration key: " & fields(1)
            keys.Add VBA.Trim$(fields(1)), fields(2)
            records.Add fields
        End If
    Next line
    If records.Count = 0 Then VBA.Err.Raise 5, , "Configuration file is empty: " & configPath
    If Not keys.Exists("ThisWorkbook::id") Then VBA.Err.Raise 5, , "Required configuration key not found: ThisWorkbook::id"
    If keys("ThisWorkbook::id") <> "TravelDocumentsBuilder" Then VBA.Err.Raise 5, , "Configuration profile mismatch."
    successTemplate = keys("TravelDocumentsBuilder::configRefresh.text.Success")
    failureTemplate = keys("TravelDocumentsBuilder::configRefresh.text.Failed")
    If VBA.Len(successTemplate) = 0 Or VBA.Len(failureTemplate) = 0 Then
        VBA.Err.Raise 5, , "Required configuration key unavailable: configRefresh.text.Success / configRefresh.text.Failed"
    End If
    For Each sheet In ThisWorkbook.Worksheets
        For Each candidate In sheet.ListObjects
            If candidate.Name = "tbConfig" Then
                If Not table Is Nothing Then VBA.Err.Raise 5, , "Duplicate table: tbConfig"
                Set table = candidate
            End If
        Next candidate
    Next sheet
    If table Is Nothing Then VBA.Err.Raise 5, , "Required table not found: tbConfig"
    kindColumn = table.ListColumns("Kind").Index
    keyColumn = table.ListColumns("Key").Index
    valueColumn = keyColumn + 1
    If valueColumn > table.ListColumns.Count Then VBA.Err.Raise 5, , "Required value column not found: tbConfig"
    Set originalRange = table.Range
    originalValues = originalRange.Value2
    ReDim values(1 To records.Count, 1 To table.ListColumns.Count)
    For index = 1 To records.Count
        fields = records(index)
        values(index, kindColumn) = fields(0)
        values(index, keyColumn) = fields(1)
        values(index, valueColumn) = fields(2)
    Next index
    previousEvents = Application.EnableEvents
    previousUpdating = Application.ScreenUpdating
    guarded = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    changed = True
    table.Resize table.HeaderRowRange.Resize(records.Count + 1, table.ListColumns.Count)
    Set resizedRange = table.Range
    If resizedRange.Rows.Count > originalRange.Rows.Count And originalRange.Rows.Count > 1 Then
        ' Extend the existing column formatting to newly added table rows.
        originalRange.Rows(originalRange.Rows.Count).Copy
        resizedRange.Rows(originalRange.Rows.Count + 1).Resize( _
            resizedRange.Rows.Count - originalRange.Rows.Count).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    If originalRange.Rows.Count > 1 Then
        originalRange.Offset(1).Resize(originalRange.Rows.Count - 1).ClearContents
    End If
    table.DataBodyRange.NumberFormat = "@"
    table.DataBodyRange.Value2 = values
    If resizedRange.Rows.Count < originalRange.Rows.Count Then
        ' Restore the surrounding sheet formatting only in rows released by the table.
        Set releasedRange = originalRange.Rows(resizedRange.Rows.Count + 1).Resize( _
            originalRange.Rows.Count - resizedRange.Rows.Count)
        originalRange.Rows(originalRange.Rows.Count).Offset(1).Copy
        releasedRange.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    changed = False
    ex_Core.fn_Diagnostic_WriteLog "TDB_CONFIG_REFRESHED | Path=" & configPath & " | Rows=" & records.Count
    ex_Core.fn_StatusBar_Show VBA.Replace$(successTemplate, "{path}", configPath)
Cleanup:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not stream Is Nothing Then stream.Close
    If guarded Then
        Application.EnableEvents = previousEvents
        Application.ScreenUpdating = previousUpdating
    End If
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    diagnostic = VBA.Err.Description
    On Error Resume Next
    If changed Then
        VBA.Err.Clear
        If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents
        table.Resize originalRange
        originalRange.Value2 = originalValues
        If VBA.Err.Number <> 0 Then diagnostic = diagnostic & " | Rollback: " & VBA.Err.Description
    End If
    ex_Core.fn_Diagnostic_WriteLog "TDB_CONFIG_REFRESH_FAILED | " & diagnostic
    If VBA.Len(failureTemplate) = 0 Then
        ex_Core.fn_TryGetWorkbookConfigValue "TravelDocumentsBuilder::configRefresh.text.Failed", failureTemplate
    End If
    If VBA.Len(failureTemplate) > 0 Then
        ex_WindowsUi.fn_ShowMessage VBA.Replace$(failureTemplate, "{error}", diagnostic), VBA.vbExclamation
    Else
        ex_WindowsUi.fn_ShowMessage "Required configuration key unavailable: TravelDocumentsBuilder::configRefresh.text.Failed | " & diagnostic, VBA.vbExclamation
    End If
    Resume Cleanup
End Sub
' } // namespace API