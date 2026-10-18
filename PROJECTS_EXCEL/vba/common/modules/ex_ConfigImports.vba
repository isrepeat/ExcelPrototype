Option Explicit

#Const CONFIG_QUERY_AVAILABLE = True

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_TryExpand( _
    ByVal records As Collection, _
    ByVal baseFolder As String, _
    ByRef output As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim entries As Object
    Dim resolved As Object
    Dim visiting As Object
    Dim imported As Object
    Dim localKeys As Object
    Dim row As Variant
    Dim key As Variant
    Dim path As String

    On Error GoTo EH
    Set output = Nothing
    Set entries = private_Dictionary()
    Set resolved = private_Dictionary()
    Set visiting = private_Dictionary()
    Set imported = private_Dictionary()
    Set localKeys = private_Dictionary()

    ' Load shared keys first; explicit local overrides do not depend on row order.
    For Each row In records
        If VBA.StrComp(row(0), "importCfg", VBA.vbTextCompare) = 0 Then
            path = private_Path(baseFolder, VBA.CStr(row(2)))
            private_LoadShared path, entries, imported
        End If
    Next row
    For Each row In records
        key = VBA.Trim$(VBA.CStr(row(1)))
        If localKeys.Exists(key) Then
            Err.Raise 5, , "Duplicate local configuration key: " & key
        End If
        localKeys.Add key, True
        entries(key) = VBA.Array(VBA.Trim$(VBA.CStr(row(0))), key, VBA.CStr(row(2)), baseFolder)
    Next row

    Set output = New Collection
    For Each key In entries.Keys
        row = entries(key)
        If VBA.StrComp(row(0), "importCfg", VBA.vbTextCompare) = 0 Then
            output.Add VBA.Array("importCfg", key, private_Path(baseFolder, VBA.CStr(row(2))))
        Else
            output.Add VBA.Array("value", key, private_Resolve(VBA.CStr(key), entries, resolved, visiting))
        End If
    Next key
    fn_TryExpand = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Set output = Nothing
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Sub private_LoadShared( _
    ByVal path As String, _
    ByVal entries As Object, _
    ByVal imported As Object _
)
#If CONFIG_QUERY_AVAILABLE Then
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim table As obj_DataTable
    Dim fileSystem As Object
    Dim diagnostic As String
    Dim kindColumn As Long
    Dim keyColumn As Long
    Dim valueColumn As Long
    Dim row As Long
    Dim kind As String
    Dim key As String
    Dim value As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(path) Then
        Err.Raise 5, , "Shared configuration workbook not found: " & path
    End If
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        Err.Raise 5, , "Cannot initialize shared configuration reader: " & path
    End If
    source.WorkbookPath = path
    source.TableName = "tbConfig"
    service.Backend = QueryAuto
    If Not service.TryExecute(source, query, table, diagnostic) Then
        Err.Raise 5, , "Cannot read shared configuration '" & path & "': " & diagnostic
    End If
    kindColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, "Kind")
    keyColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, "Key")
    valueColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, "Value")
    For row = 1 To table.RowCount
        kind = private_Text(table.ValueAt(row, kindColumn))
        key = private_Text(table.ValueAt(row, keyColumn))
        value = private_Text(table.ValueAt(row, valueColumn))
        If VBA.Len(kind) = 0 Or VBA.Len(key) = 0 Or VBA.Len(value) = 0 Then
            Err.Raise 5, , "Empty shared configuration kind, key or value: " & path & " / row " & VBA.CStr(row)
        End If
        If VBA.StrComp(kind, "importCfg", VBA.vbTextCompare) = 0 Then
            Err.Raise 5, , "Nested shared configuration imports are not allowed: " & path
        End If
        If imported.Exists(key) Then
            Err.Raise 5, , "Duplicate shared configuration key: " & key & " / " & path
        End If
        imported.Add key, True
        entries.Add key, VBA.Array(kind, key, value, fileSystem.GetParentFolderName(path))
    Next row
Cleanup:
    If Not table Is Nothing Then
        table.Dispose
    End If
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, , errorText
    End If
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
#Else
    private_LoadSharedWorkbook path, entries, imported
#End If
End Sub

#If Not CONFIG_QUERY_AVAILABLE Then
Private Sub private_LoadSharedWorkbook( _
    ByVal path As String, _
    ByVal entries As Object, _
    ByVal imported As Object _
)
    Dim book As Workbook
    Dim candidate As Workbook
    Dim table As ListObject
    Dim fileSystem As Object
    Dim opened As Boolean
    Dim previousEvents As Boolean
    Dim row As Long
    Dim values As Variant
    Dim kind As String
    Dim key As String
    Dim value As String
    Dim kindColumn As Long
    Dim keyColumn As Long
    Dim valueColumn As Long
    Dim errorNumber As Long
    Dim errorText As String

    previousEvents = Application.EnableEvents
    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(path) Then Err.Raise 5, , "Shared configuration workbook not found: " & path
    For Each candidate In Application.Workbooks
        If VBA.StrComp(candidate.FullName, path, VBA.vbTextCompare) = 0 Then
            Set book = candidate
            Exit For
        End If
    Next candidate
    Application.EnableEvents = False
    If book Is Nothing Then
        Set book = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        opened = True
    End If
    Set table = book.Worksheets("wsConfig").ListObjects("tbConfig")
    If table.DataBodyRange Is Nothing Then Err.Raise 5, , "Empty shared configuration: " & path
    kindColumn = table.ListColumns("Kind").Index
    keyColumn = table.ListColumns("Key").Index
    valueColumn = table.ListColumns("Value").Index
    values = table.DataBodyRange.Value2
    For row = 1 To UBound(values, 1)
        kind = private_Text(values(row, kindColumn))
        key = private_Text(values(row, keyColumn))
        value = private_Text(values(row, valueColumn))
        If VBA.Len(kind) = 0 Or VBA.Len(key) = 0 Or VBA.Len(value) = 0 Then
            Err.Raise 5, , "Empty shared configuration kind, key or value: " & path & " / row " & VBA.CStr(row)
        End If
        If VBA.StrComp(kind, "importCfg", VBA.vbTextCompare) = 0 Then
            Err.Raise 5, , "Nested shared configuration imports are not allowed: " & path
        End If
        If imported.Exists(key) Then Err.Raise 5, , "Duplicate shared configuration key: " & key & " / " & path
        imported.Add key, True
        entries.Add key, VBA.Array(kind, key, value, fileSystem.GetParentFolderName(path))
    Next row
Cleanup:
    On Error Resume Next
    If opened Then book.Close SaveChanges:=False
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, , errorText
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Sub
#End If

Private Function private_Resolve( _
    ByVal key As String, _
    ByVal entries As Object, _
    ByVal resolved As Object, _
    ByVal visiting As Object _
) As String
    Dim row As Variant
    Dim value As String

    If resolved.Exists(key) Then
        private_Resolve = resolved(key)
        Exit Function
    End If
    If Not entries.Exists(key) Then
        Err.Raise 5, , "Referenced configuration key not found: " & key
    End If
    If visiting.Exists(key) Then
        Err.Raise 5, , "Configuration reference cycle: " & key
    End If
    visiting.Add key, True
    row = entries(key)
    value = VBA.Trim$(VBA.CStr(row(2)))
    Select Case VBA.LCase$(row(0))
        Case "ref"
            value = private_Resolve(value, entries, resolved, visiting)
            If VBA.Len(value) = 0 Then
                Err.Raise 5, , "Empty referenced configuration value: " & key
            End If
        Case "file"
            value = private_Path(VBA.CStr(row(3)), value)
        Case "importcfg"
            Err.Raise 5, , "Cannot reference an import directive: " & key
    End Select
    visiting.Remove key
    resolved.Add key, value
    private_Resolve = value
End Function

Private Function private_Path( _
    ByVal baseFolder As String, _
    ByVal value As String _
) As String
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    value = VBA.Replace$(VBA.Trim$(value), "/", "\")
    If VBA.Len(value) = 0 Then
        Err.Raise 5, , "Empty configuration file path."
    End If
    If fileSystem.GetDriveName(value) = VBA.vbNullString Then
        value = fileSystem.BuildPath(baseFolder, value)
    End If
    private_Path = fileSystem.GetAbsolutePathName(value)
End Function

Private Function private_Dictionary() As Object
    Set private_Dictionary = VBA.CreateObject("Scripting.Dictionary")
    private_Dictionary.CompareMode = VBA.vbTextCompare
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then
        Err.Raise 5, , "Invalid shared configuration cell value."
    End If
    private_Text = VBA.Trim$(VBA.CStr(value))
End Function