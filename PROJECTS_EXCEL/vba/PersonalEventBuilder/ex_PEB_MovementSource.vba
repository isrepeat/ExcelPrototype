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
Public Function fn_ReadLatest(ByVal personId As String) As Object
    Dim result As Object
    Dim book As Workbook
    Dim table As ListObject
    Dim fields As Object
    Dim names As Object
    Dim ids As Variant
    Dim person As String
    Dim path As String
    Dim field As Variant
    Dim index As Long
    Dim column As Long
    Dim dateKey As Double
    Dim latestKey As Double
    Dim arrivalKey As Double
    Dim latestArrivalKey As Double
    Dim conn As Object
    Dim command As Object
    Dim records As Object
    Dim parameter As Object
    Dim sql As String
    Dim props As String
    Dim extension As String
    Dim failure As String

    On Error GoTo EH
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("validation.Movement.source.path"))
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Array("EventName", "DepartureDate", "ArrivalDate")
        fields.Add field, ex_PEB_Export.fn_Config("export.Movement.headers." & field)
    Next field
    If VBA.Len(VBA.Trim$(personId)) = 0 Then
        ex_PEB_MovementLinks.fn_Fail "MissingPersonId", ""
    End If
    For Each book In Application.Workbooks
        If VBA.StrComp(book.FullName, path, VBA.vbTextCompare) = 0 Then
            Set table = book.Worksheets(ex_PEB_Export.fn_Config("validation.Movement.source.sheet")).ListObjects( _
                ex_PEB_Export.fn_Config("validation.Movement.source.table"))
            Set names = VBA.CreateObject("Scripting.Dictionary")
            For Each field In fields.Keys
                names(field) = table.ListColumns(fields(field)).Index
            Next field
            If Not table.DataBodyRange Is Nothing Then
                ids = table.ListColumns(ex_PEB_Export.fn_Config("export.Movement.headers.PersonId")).DataBodyRange.Value2
                For index = 1 To table.ListRows.Count
                    If table.ListRows.Count = 1 Then
                        person = private_Text(ids)
                    Else
                        person = private_Text(ids(index, 1))
                    End If
                    If person = personId Then
                        dateKey = fn_OrderDate(table.DataBodyRange.Cells(index, names("DepartureDate")).Value2)
                        arrivalKey = fn_OrderDate(table.DataBodyRange.Cells(index, names("ArrivalDate")).Value2)
                        If result Is Nothing Or dateKey > latestKey Or _
                                (dateKey = latestKey And arrivalKey <= latestArrivalKey) Then
                            latestKey = dateKey
                            latestArrivalKey = arrivalKey
                            Set result = VBA.CreateObject("Scripting.Dictionary")
                            For Each field In fields.Keys
                                result(field) = table.DataBodyRange.Cells(index, names(field)).Value2
                            Next field
                        End If
                    End If
                Next index
            End If
            Set fn_ReadLatest = result
            Exit Function
        End If
    Next book

    ' Read one period; dates establish order and are not business constraints.
    For Each field In fields.Keys
        If VBA.Len(sql) > 0 Then
            sql = sql & ", "
        End If
        sql = sql & private_Identifier(fields(field))
    Next field
    sql = "SELECT TOP 1 " & sql & " FROM " & private_Identifier( _
        ex_PEB_Export.fn_Config("validation.Movement.source.sheet") & "$" & _
        ex_PEB_Export.fn_Config("validation.Movement.source.range"))
    sql = sql & " WHERE Trim(" & private_Identifier(ex_PEB_Export.fn_Config("export.Movement.headers.PersonId")) & " & '') = ?"
    sql = sql & " ORDER BY " & private_Identifier(fields("DepartureDate")) & " DESC, " & _
        private_Identifier(fields("ArrivalDate")) & " ASC"
    extension = VBA.LCase$(VBA.Mid$(path, VBA.InStrRev(path, ".") + 1))
    Select Case extension
        Case "xlsx", "xlsm"
            props = "Excel 12.0 Xml"
        Case "xlsb"
            props = "Excel 12.0"
        Case "xls"
            props = "Excel 8.0"
        Case Else
            ex_PEB_MovementLinks.fn_Fail "UnsupportedSourceFormat", path
    End Select
    Set conn = VBA.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & VBA.Replace$(path, """", """""") & _
        """;Extended Properties=""" & props & ";HDR=YES;IMEX=1;ReadOnly=True"";"
    Set command = VBA.CreateObject("ADODB.Command")
    command.CommandType = 1
    Set parameter = command.CreateParameter("p1", 202, 1, VBA.Len(personId), personId)
    command.Parameters.Append parameter
    Set command.ActiveConnection = conn
    command.CommandText = sql
    Set records = command.Execute
    If Not records.EOF Then
        Set result = VBA.CreateObject("Scripting.Dictionary")
        column = 0
        For Each field In fields.Keys
            result(field) = records.Fields(column).Value
            column = column + 1
        Next field
    End If
    Set fn_ReadLatest = result
Cleanup:
    On Error Resume Next
    If Not records Is Nothing Then
        records.Close
    End If
    If Not conn Is Nothing Then
        conn.Close
    End If
    On Error GoTo 0
    If VBA.Len(failure) > 0 Then
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
Private Function private_Identifier(ByVal value As String) As String
    value = VBA.Replace$(value, ".", "#")
    private_Identifier = "[" & VBA.Replace$(value, "]", "]]") & "]"
End Function

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