Attribute VB_Name = "ex_PR_Source"
Option Explicit

Private m_config As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Load( _
    ByRef identifiers As Object, _
    ByRef names As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim executor As obj_AdoQueryExecutor
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim data As obj_DataTable
    Dim kind As Variant
    Dim values As Variant
    Dim record As Variant
    Dim row As Long
    Dim nameColumn As Long
    Dim idColumn As Long
    Dim originColumn As Long
    Dim fullName As String
    Dim identifier As String
    Dim origin As String
    Dim sourcePath As String
    Dim candidates As Object
    Dim fileSystem As Object
    Dim errorText As String

    On Error GoTo Failed
    Set m_config = Nothing
    diagnostic = vbNullString
    Set identifiers = private_Dictionary()
    Set names = private_Dictionary()
    sourcePath = private_Cfg("Source.Path")
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(sourcePath) Then
        Err.Raise vbObjectError + 2100, , Replace(private_Cfg("Message.SourceMissing"), "{path}", sourcePath)
    End If

    ' Explicit ranges bypass the shared executor's workbook/table metadata path.
    Set executor = New obj_AdoQueryExecutor
    If Not executor.Initialize() Then
        Err.Raise vbObjectError + 2106, , private_Cfg("Message.InitializeFailed")
    End If
    For Each kind In Array("List", "Arrivals")
        Set source = New obj_TableSource
        Set query = New obj_TableQuery
        If Not source.Initialize() Or Not query.Initialize() Then
            Err.Raise vbObjectError + 2106, , private_Cfg("Message.InitializeFailed")
        End If
        source.WorkbookPath = sourcePath
        source.SheetName = private_Cfg("Source." & kind & ".Sheet")
        source.RangeAddress = private_Cfg("Source." & kind & ".Range")
        query.AddColumn private_Cfg("Source." & kind & ".NameColumn")
        query.AddColumn private_Cfg("Source." & kind & ".IDColumn")
        If CStr(kind) = "Arrivals" Then
            query.AddColumn private_Cfg("Source.Arrivals.OriginColumn")
        End If
        If Not executor.TryExecute(source, query, data, errorText, CLng(private_Cfg("Source.FetchSize"))) Then
            Err.Raise vbObjectError + 2107, , errorText
        End If
        nameColumn = ex_TableQuery.fn_HeaderIndex(data.Headers, private_Cfg("Source." & kind & ".NameColumn"))
        idColumn = ex_TableQuery.fn_HeaderIndex(data.Headers, private_Cfg("Source." & kind & ".IDColumn"))
        originColumn = 0
        If CStr(kind) = "Arrivals" Then
            originColumn = ex_TableQuery.fn_HeaderIndex(data.Headers, private_Cfg("Source.Arrivals.OriginColumn"))
        End If
        If data.RowCount > 0 Then
            values = data.Values
            For row = 1 To data.RowCount
                fullName = private_Text(values(row, nameColumn))
                identifier = private_Text(values(row, idColumn))
                origin = vbNullString
                If originColumn > 0 Then
                    origin = private_Text(values(row, originColumn))
                End If
                If Len(identifier) > 0 Then
                    If identifiers.Exists(identifier) Then
                        record = identifiers(identifier)
                        If StrComp(CStr(record(0)), fullName, vbTextCompare) <> 0 Then
                            record(2) = True
                        End If
                        If Len(CStr(record(1))) = 0 And Len(origin) > 0 Then
                            record(1) = origin
                        End If
                        identifiers(identifier) = record
                    Else
                        identifiers.Add identifier, Array(fullName, origin, Len(fullName) = 0)
                    End If
                End If
                If Len(fullName) > 0 Then
                    If Not names.Exists(fullName) Then
                        Set candidates = private_Dictionary()
                        names.Add fullName, candidates
                    End If
                    Set candidates = names(fullName)
                    If Len(identifier) > 0 Then
                        If Not candidates.Exists(identifier) Then
                            candidates.Add identifier, origin
                        ElseIf Len(origin) > 0 Then
                            candidates(identifier) = origin
                        End If
                    End If
                End If
            Next row
        End If
        data.Dispose
        query.Dispose
        source.Dispose
        Set data = Nothing
        Set query = Nothing
        Set source = Nothing
    Next kind
    If names.Count = 0 Then
        Err.Raise vbObjectError + 2101, , private_Cfg("Message.SourceEmpty")
    End If
    fn_Load = True
Cleanup:
    On Error Resume Next
    If Not data Is Nothing Then data.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    If Not executor Is Nothing Then executor.Dispose
    Set data = Nothing
    Set query = Nothing
    Set source = Nothing
    Set executor = Nothing
    On Error GoTo 0
    Exit Function
Failed:
    errorText = Err.Description
    diagnostic = Replace(private_Cfg("Message.SourceFailed"), "{error}", errorText)
    Set identifiers = Nothing
    Set names = Nothing
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Text(ByVal value As Variant) As String
    Dim result As String

    If IsError(value) Then
        Err.Raise vbObjectError + 2105, , private_Cfg("Message.SourceCellError")
    End If
    If IsNull(value) Or IsEmpty(value) Then
        Exit Function
    End If
    result = CStr(value)
    result = Replace(result, ChrW(160), " ")
    result = Replace(result, ChrW(&H202F), " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    private_Text = Trim$(result)
End Function

Private Function private_Dictionary() As Object
    Set private_Dictionary = CreateObject("Scripting.Dictionary")
    private_Dictionary.CompareMode = vbTextCompare
End Function

Private Function private_Cfg(ByVal key As String) As String
    Dim values As Variant
    Dim row As Long
    Dim fullKey As String

    If m_config Is Nothing Then
        Set m_config = private_Dictionary()
        values = ThisWorkbook.Worksheets("wsConfig").ListObjects("tbConfig").DataBodyRange.Value2
        For row = 1 To UBound(values, 1)
            If IsError(values(row, 2)) Or IsError(values(row, 3)) Then
                Err.Raise vbObjectError + 2050, , "Invalid configuration cell."
            End If
            m_config.Add CStr(values(row, 2)), CStr(values(row, 3))
        Next row
    End If
    fullKey = "PeopleRegistry::" & key
    If m_config.Exists(fullKey) Then
        private_Cfg = CStr(m_config(fullKey))
        If Len(private_Cfg) > 0 Then Exit Function
    End If
    Err.Raise vbObjectError + 2050, , "Missing configuration key: PeopleRegistry::" & key
End Function