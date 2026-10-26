Attribute VB_Name = "ex_PEB_MovementDiff"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Build( _
    ByVal plan As Object, _
    ByVal target As ListObject, _
    ByVal entries As Collection _
) As Object
    Dim report As Object
    Dim rows As New Collection
    Dim links As Object
    Dim internalHeaders As Object
    Dim changedEvents As Object
    Dim action As Object
    Dim entry As Object
    Dim row As Object
    Dim key As Variant
    Dim added As Long
    Dim updated As Long
    Dim unchanged As Long
    Dim summary As String
    Dim originalRows As Long

    Set report = VBA.CreateObject("Scripting.Dictionary")
    originalRows = target.ListRows.Count
    Set links = ex_PEB_MovementLinks.fn_Headers()
    Set internalHeaders = VBA.CreateObject("Scripting.Dictionary")
    For Each key In links.Keys
        internalHeaders(VBA.CStr(links(key))) = True
    Next key
    Set changedEvents = VBA.CreateObject("Scripting.Dictionary")
    For Each entry In entries
        If Not internalHeaders.Exists(entry("Header")) Then
            If ex_PEB_MovementLinks.fn_ValueToken(entry("Value")) <> _
                    ex_PEB_MovementLinks.fn_ValueToken(entry("Expected")) Then
                Set action = entry("Action")
                Set row = VBA.CreateObject("Scripting.Dictionary")
                row("PersonName") = action("PersonName")
                row("EventRowId") = action("RowId")
                row("EventName") = action("EventName")
                row("Worksheet") = target.Parent.Name
                row("TargetRow") = VBA.CStr(entry("Row"))
                row("Field") = entry("Header")
                row("Before") = private_Text(entry("Value"), VBA.CStr(entry("Format")))
                row("After") = private_Text(entry("Expected"))
                row("Action") = ex_PEB_Export.fn_Config("movementDiff.updated")
                If entry("TableRow") > originalRows Then
                    row("Action") = ex_PEB_Export.fn_Config("movementDiff.added")
                End If
                rows.Add row
                changedEvents(action("RowId")) = True
            End If
        End If
    Next entry
    For Each action In plan("Actions")
        If Not changedEvents.Exists(action("RowId")) Then
            unchanged = unchanged + 1
        ElseIf action("IsUpdate") Or action("Operation") = "close" Then
            updated = updated + 1
        Else
            added = added + 1
        End If
    Next action
    summary = ex_PEB_Export.fn_Config("movementDiff.summary")
    summary = VBA.Replace$(summary, "{added}", VBA.CStr(added))
    summary = VBA.Replace$(summary, "{updated}", VBA.CStr(updated))
    summary = VBA.Replace$(summary, "{unchanged}", VBA.CStr(unchanged))
    summary = VBA.Replace$(summary, "{time}", VBA.Format$(VBA.Now, "dd.mm.yyyy hh:nn:ss"))
    report.Add "Summary", summary
    report.Add "Added", added
    report.Add "Updated", updated
    report.Add "Unchanged", unchanged
    report.Add "Target", target.Parent.Parent.FullName & " / " & target.Parent.Name & " / " & target.Name
    report.Add "AbsenceSheet", target.Parent.Name
    report.Add "TvoSheet", ex_PEB_Export.fn_Config("reporter.tvo.source.sheet")
    report.Add "Rows", rows
    Set fn_Build = report
End Function

Public Function fn_Result( _
    ByVal report As Object, _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String, _
    Optional ByVal worksheetName As String = "" _
) As Boolean
    Dim records As New Collection
    Dim columns As New Collection
    Dim row As Object
    Dim record As obj_LookupRecord
    Dim column As Object
    Dim key As Variant
    Dim index As Long

    For Each row In report("Rows")
        If Len(worksheetName) > 0 Then
            If row("Worksheet") <> worksheetName Then
                GoTo NextDiffRow
            End If
        End If
        index = index + 1
        Set record = New obj_LookupRecord
        If Not record.Initialize(VBA.CStr(index), row) Then
            Exit Function
        End If
        records.Add record
NextDiffRow:
    Next row
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("movementDiff.fields"), "|")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column.Add "name", key
        column.Add "caption", ex_PEB_Export.fn_Config("movementDiff.columns." & key)
        columns.Add column
    Next key
    Set result = New obj_LookupResult
    fn_Result = result.Initialize(records, columns, False, diagnostic)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Text( _
    ByVal value As Variant, _
    Optional ByVal numberFormat As String = "" _
) As String
    If VBA.IsError(value) Then
        private_Text = ex_PEB_Export.fn_Config("movementDiff.error")
    ElseIf VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        private_Text = ex_PEB_Export.fn_Config("movementDiff.empty")
    ElseIf VBA.Len(VBA.CStr(value)) = 0 Then
        private_Text = ex_PEB_Export.fn_Config("movementDiff.empty")
    ElseIf VBA.VarType(value) = VBA.vbDate Then
        private_Text = VBA.Format$(value, ex_PEB_Export.fn_Config("export.Movement.dateFormat"))
    ElseIf VBA.IsNumeric(value) And (VBA.InStr(1, numberFormat, "yy", VBA.vbTextCompare) > 0 Or _
            VBA.InStr(1, numberFormat, "dd", VBA.vbTextCompare) > 0) Then
        private_Text = VBA.Format$(VBA.CDate(value), ex_PEB_Export.fn_Config("export.Movement.dateFormat"))
    Else
        private_Text = VBA.CStr(value)
    End If
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------