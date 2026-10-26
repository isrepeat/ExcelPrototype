Attribute VB_Name = "ex_PEB_MovementSource"
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
Public Function fn_ReadLatest( _
    ByVal personId As String, _
    Optional ByVal includeTvo As Boolean = False, _
    Optional ByVal useSnapshot As Boolean = False _
) As Object
    Dim result As Object
    Dim fields As Object
    Dim field As Variant
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim table As obj_DataTable
    Dim diagnostic As String
    Dim row As Long
    Dim dateKey As Double
    Dim arrivalKey As Double
    Dim latestKey As Double
    Dim latestArrivalKey As Double
    Dim startedAt As Double
    Dim failure As String

    On Error GoTo EH
    startedAt = VBA.Timer
    If VBA.Len(VBA.Trim$(personId)) = 0 Then
        ex_PEB_MovementLinks.fn_Fail "MissingPersonId", ""
    End If
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Array("EventName", "DepartureDate", "ArrivalDate")
        fields.Add field, ex_PEB_Export.fn_Config("export.Movement.headers." & field)
    Next field
    If includeTvo Then
        fields.Add "PositionCode", ex_PEB_Export.fn_Config("export.Movement.headers.PositionCode")
        fields.Add "ArrivalOrder", ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalOrder")
        For Each field In VBA.Array("TvoName", "TvoPersonId", "TvoPositionCode")
            fields.Add field, ex_PEB_Export.fn_Config("returnTvo.headers." & field)
        Next field
    End If
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        Err.Raise 5, , "Cannot initialize Movement snapshot query."
    End If
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("validation.Movement.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("validation.Movement.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("validation.Movement.source.range")
    For Each field In fields.Keys
        query.AddColumn fields(field)
    Next field
    query.AddCondition ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), QueryEquals, personId, QueryText, True
    If Not service.TryExecute(source, query, table, diagnostic, Not includeTvo Or useSnapshot) Then
        Err.Raise 5, , diagnostic
    End If
    ' Preserve the latest-period ordering and all subsequent business validation.
    For row = 1 To table.RowCount
        dateKey = fn_OrderDate(table.ValueAt(row, ex_TableQuery.fn_HeaderIndex(table.Headers, fields("DepartureDate"))))
        arrivalKey = fn_OrderDate(table.ValueAt(row, ex_TableQuery.fn_HeaderIndex(table.Headers, fields("ArrivalDate"))))
        If result Is Nothing Or dateKey > latestKey Or (dateKey = latestKey And arrivalKey <= latestArrivalKey) Then
            latestKey = dateKey
            latestArrivalKey = arrivalKey
            Set result = VBA.CreateObject("Scripting.Dictionary")
            For Each field In fields.Keys
                result(field) = table.ValueAt(row, ex_TableQuery.fn_HeaderIndex(table.Headers, fields(field)))
            Next field
        End If
    Next row
    Set fn_ReadLatest = result
Cleanup:
    If Not table Is Nothing Then
        table.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not service Is Nothing Then
        service.Dispose
    End If
    ex_Core.fn_Diagnostic_WritePerf "MovementSource.ReadLatest", startedAt
    If VBA.Len(failure) > 0 Then
        On Error GoTo 0
        Err.Raise 5, , failure
    End If
    Exit Function
EH:
    failure = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_OrderDate(ByVal value As Variant) As Double
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(VBA.CStr(value))) = 0 Then
        Exit Function
    End If
    If VBA.IsNumeric(value) Then
        fn_OrderDate = VBA.CDbl(value)
    ElseIf VBA.VarType(value) = VBA.vbDate Then
        fn_OrderDate = VBA.CDbl(value)
    Else
        fn_OrderDate = VBA.CDbl(ex_PEB_History.fn_Date(value))
    End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        ex_PEB_MovementLinks.fn_Fail "InvalidSourceCell", ""
    End If
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
        private_Text = VBA.Trim$(VBA.CStr(value))
    End If
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------