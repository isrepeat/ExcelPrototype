Attribute VB_Name = "ex_TDB_HistoryRecovery"
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
Public Function fn_Prepare( _
    ByVal table As ListObject, _
    ByVal versions As ListObject, _
    ByVal identifier As String, _
    ByVal transaction As Object _
) As String
    Dim document As Object
    Dim root As Object
    Dim rows As Object
    Dim group As Collection
    Dim previous As Collection
    Dim row As Object
    Dim token As String
    Dim index As Long

    token = ex_TDB_History.fn_NewId()
    Set rows = ex_TDB_History.fn_Read(table)
    Set group = ex_TDB_History.fn_Group(rows, identifier)
    Set previous = transaction("Previous")
    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    Set root = document.createElement("Recovery")
    document.appendChild root
    root.setAttribute "version", "1"
    root.setAttribute "operation", token
    root.setAttribute "state", "PREPARE"
    root.setAttribute "workbook", table.Parent.Parent.FullName
    root.setAttribute "time", VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    root.setAttribute "user", Application.UserName & " / " & VBA.Environ$("USERNAME")
    For Each row In previous
        private_RecordRow document, root, table, row, "Before"
    Next row
    For Each row In group
        private_RecordRow document, root, table, row, "After"
    Next row
    For index = VBA.CLng(transaction("RevisionsCount")) + 1 To versions.ListRows.Count
        private_Row document, root, versions, versions.ListRows(index).Range.Value2, "Revision"
    Next index
    private_Write table.Parent.Parent, document
    fn_Prepare = token
End Function

Public Sub fn_Mark( _
    ByVal book As Workbook, _
    ByVal token As String, _
    ByVal state As String _
)
    Dim document As Object
    Dim root As Object

    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    Set root = document.createElement("Recovery")
    document.appendChild root
    root.setAttribute "version", "1"
    root.setAttribute "operation", token
    root.setAttribute "state", state
    root.setAttribute "time", VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    private_Write book, document
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_RecordRow( _
    ByVal document As Object, _
    ByVal root As Object, _
    ByVal table As ListObject, _
    ByVal record As Object, _
    ByVal role As String _
)
    Dim data As Variant
    Dim index As Long

    ReDim data(1 To 1, 1 To table.ListColumns.Count)
    For index = 1 To table.ListColumns.Count
        data(1, index) = record(table.ListColumns(index).Name)
    Next index
    private_Row document, root, table, data, role
End Sub

Private Sub private_Row( _
    ByVal document As Object, _
    ByVal root As Object, _
    ByVal table As ListObject, _
    ByVal data As Variant, _
    ByVal role As String _
)
    Dim row As Object
    Dim field As Object
    Dim index As Long
    Dim value As Variant

    Set row = document.createElement("Row")
    row.setAttribute "role", role
    row.setAttribute "sheet", table.Parent.Name
    row.setAttribute "table", table.Name
    root.appendChild row
    For index = 1 To table.ListColumns.Count
        Set field = document.createElement("Field")
        field.setAttribute "name", table.ListColumns(index).Name
        value = data(1, index)
        field.setAttribute "type", VBA.CStr(VBA.VarType(value))
        If Not IsEmpty(value) Then
            field.Text = VBA.CStr(value)
        End If
        row.appendChild field
    Next index
End Sub

Private Sub private_Write( _
    ByVal book As Workbook, _
    ByVal document As Object _
)
    Dim files As Object
    Dim stream As Object
    Dim path As String
    Dim payload As String
    Dim failure As String
    Dim startedAt As Double

    On Error GoTo EH
    startedAt = VBA.Timer
    Set files = VBA.CreateObject("Scripting.FileSystemObject")
    path = files.BuildPath(book.Path, book.Name & ex_TDB_History.fn_Config("recovery.suffix"))
    payload = document.XML
    Do While VBA.Right$(payload, 1) = VBA.vbCr Or VBA.Right$(payload, 1) = VBA.vbLf
        payload = VBA.Left$(payload, VBA.Len(payload) - 1)
    Loop
    ' Preserve embedded line breaks while keeping each XML record on one physical line.
    payload = VBA.Replace$(payload, VBA.vbCr, "&#13;")
    payload = VBA.Replace$(payload, VBA.vbLf, "&#10;")
    Set stream = files.OpenTextFile(path, 8, True, -1)
    ' A leading separator isolates a truncated record left by a previous process crash.
    stream.WriteLine VBA.vbCrLf & payload
    stream.Close
    Set stream = Nothing
    On Error Resume Next
    ex_Core.fn_Diagnostic_WritePerf "History.Recovery.Append", startedAt
    On Error GoTo 0
    Exit Sub
EH:
    failure = VBA.Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
    End If
    On Error GoTo 0
    Err.Raise 5, , VBA.Replace$(ex_TDB_History.fn_Config("text.RecoveryFailed"), "{details}", failure)
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------