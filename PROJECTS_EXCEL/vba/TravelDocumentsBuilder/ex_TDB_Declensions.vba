Attribute VB_Name = "ex_TDB_Declensions"
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
Public Sub fn_Enrich(ByVal values As Object)
    Dim index As Long
    Dim prefix As String
    Dim output As String
    Dim detail As String

    On Error GoTo Failed
    For index = 1 To VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("declensions.count"))
        prefix = "declensions.lookups." & VBA.CStr(index)
        output = ex_TravelDocumentsBuilder.fn_Config(prefix & ".output")
        values(output) = private_Lookup(prefix, values)
    Next index
    Exit Sub
Failed:
    detail = VBA.Err.Description
    VBA.Err.Raise 5, "Tickets.Declensions", VBA.Replace$(VBA.Replace$( _
        ex_TravelDocumentsBuilder.fn_Config("declensions.error"), "{field}", output), "{detail}", detail)
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Lookup( _
    ByVal prefix As String, _
    ByVal values As Object _
) As String
    Dim snapshot As obj_TableSnapshotSet
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim dataTable As obj_DataTable
    Dim fileSystem As Object
    Dim sourcePrefix As String
    Dim path As String
    Dim keyHeader As String
    Dim valueHeader As String
    Dim lookupInput As String
    Dim value As Variant
    Dim diagnostic As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim specialPosition As String

    On Error GoTo Failed
    Set snapshot = ex_TableSnapshots.fn_Active()
    If snapshot Is Nothing Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    sourcePrefix = "declensions.sources." & ex_TravelDocumentsBuilder.fn_Config(prefix & ".source")
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    If Not tableSource.Initialize() Or Not tableQuery.Initialize() Then VBA.Err.Raise 5, , sourcePrefix
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_TravelDocumentsBuilder.fn_Config(sourcePrefix & ".path")
    If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
        path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    tableSource.WorkbookPath = fileSystem.GetAbsolutePathName(path)
    tableSource.SheetName = ex_TravelDocumentsBuilder.fn_Config(sourcePrefix & ".sheet")
    tableSource.RangeAddress = ex_TravelDocumentsBuilder.fn_Config(sourcePrefix & ".range")
    keyHeader = ex_TravelDocumentsBuilder.fn_Config(prefix & ".keyHeader")
    valueHeader = ex_TravelDocumentsBuilder.fn_Config(prefix & ".valueHeader")
    lookupInput = VBA.CStr(values(ex_TravelDocumentsBuilder.fn_Config(prefix & ".input")))
    If VBA.Len(lookupInput) = 0 Then VBA.Err.Raise 5, , keyHeader
    If ex_TravelDocumentsBuilder.fn_Config(prefix & ".output") = "PositionGenitive" Then
        lookupInput = private_PositionKey(lookupInput)
        specialPosition = private_SpecialPosition(lookupInput, VBA.CStr(values("PersonRank")))
        If VBA.Len(specialPosition) > 0 Then
            private_Lookup = specialPosition
            tableQuery.Dispose
            tableSource.Dispose
            Exit Function
        End If
    End If
    tableQuery.AddColumn keyHeader
    tableQuery.AddColumn valueHeader
    tableQuery.AddCondition keyHeader, QueryEquals, lookupInput, QueryText, True
    tableQuery.Limit = 2
    If Not snapshot.TryQuery(tableSource, tableQuery, dataTable, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If dataTable.RowCount = 0 Then
        GoTo Cleanup
    End If
    If dataTable.RowCount > 1 Then
        VBA.Err.Raise 5, , tableSource.SheetName & ": " & keyHeader & "=" & lookupInput
    End If
    value = dataTable.ValueAt(1, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, valueHeader))
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        GoTo Cleanup
    End If
    private_Lookup = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
Cleanup:
    dataTable.Dispose
    tableQuery.Dispose
    tableSource.Dispose
    Exit Function
Failed:
    failureNumber = VBA.Err.Number
    failureDescription = VBA.Err.Description
    On Error Resume Next
    If Not dataTable Is Nothing Then dataTable.Dispose
    If Not tableQuery Is Nothing Then tableQuery.Dispose
    If Not tableSource Is Nothing Then tableSource.Dispose
    On Error GoTo 0
    VBA.Err.Raise failureNumber, "Tickets.Declensions.Lookup", failureDescription
End Function

Private Function private_PositionKey(ByVal value As String) As String
    Dim normalized As String
    Dim mapping As Variant
    Dim pair As Variant
    Dim kind As Variant
    Dim prefix As String
    Dim code As String

    private_PositionKey = VBA.UCase$(ex_TableQuery.fn_NormalizeText(value))
    normalized = VBA.Replace$(private_PositionKey, " ", VBA.vbNullString)
    For Each mapping In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("declensions.positions.characters"), "|")
        pair = VBA.Split(mapping, "=")
        normalized = VBA.Replace$(normalized, pair(0), pair(1))
    Next mapping
    For Each kind In VBA.Array("tp", "rozp", "spis")
        prefix = ex_TravelDocumentsBuilder.fn_Config("declensions.positions." & kind & ".prefix")
        code = ex_TravelDocumentsBuilder.fn_Config("declensions.positions." & kind & ".code")
        If private_PositionKey = code Or VBA.Left$(normalized, VBA.Len(prefix)) = prefix Then
            private_PositionKey = code
            Exit Function
        End If
    Next kind
End Function

Private Function private_SpecialPosition( _
    ByVal positionCode As String, _
    ByVal personRank As String _
) As String
    Dim pattern As Object
    Dim unit As String

    If positionCode = ex_TravelDocumentsBuilder.fn_Config("declensions.positions.tp.code") Then
        private_SpecialPosition = ex_TravelDocumentsBuilder.fn_Config("declensions.positions.tp.text")
    ElseIf positionCode = ex_TravelDocumentsBuilder.fn_Config("declensions.positions.rozp.code") Then
        Set pattern = VBA.CreateObject("VBScript.RegExp")
        pattern.Pattern = ex_TravelDocumentsBuilder.fn_Config("declensions.positions.officerPattern")
        pattern.IgnoreCase = True
        If pattern.Test(personRank) Then
            unit = ex_TravelDocumentsBuilder.fn_Config("declensions.positions.rozp.officerUnit")
        Else
            unit = ex_TravelDocumentsBuilder.fn_Config("declensions.positions.rozp.otherUnit")
        End If
        private_SpecialPosition = VBA.Replace$( _
            ex_TravelDocumentsBuilder.fn_Config("declensions.positions.rozp.text"), "{unit}", unit)
    End If
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------