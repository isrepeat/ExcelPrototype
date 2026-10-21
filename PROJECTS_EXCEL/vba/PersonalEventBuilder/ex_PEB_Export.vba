Attribute VB_Name = "ex_PEB_Export"
Option Explicit

Private m_configSnapshot As Object
Private m_configSnapshotDepth As Long

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_configSnapshot = Nothing
    m_configSnapshotDepth = 0
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Config(ByVal key As String) As String
    Dim value As String

    If Not m_configSnapshot Is Nothing Then
        If Not m_configSnapshot.Exists("PersonalEventBuilder::" & key) Then
            Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::" & key
        End If
        value = m_configSnapshot("PersonalEventBuilder::" & key)
    Else
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::" & key, value) Then
            Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::" & key
        End If
    End If
    If VBA.Len(VBA.Trim$(value)) = 0 Then
        Err.Raise 5, , "Empty configuration key: PersonalEventBuilder::" & key
    End If
    fn_Config = value
End Function

Public Function fn_TryApplyReporterPositions( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim names As New Collection
    Dim positions As Object
    Dim fields As Object
    Dim name As Variant
    Dim key As String

    On Error GoTo EH
    For Each record In records
        If Not record.TryGetValue("Personnel.FullName", name) Then Err.Raise 5, , "Reporter name not found."
        names.Add VBA.CStr(name)
    Next record
    If Not fn_TryReporterTvoPositions(names, positions, diagnostic) Then Exit Function
    For Each record In records
        If Not record.TryGetValue("Personnel.FullName", name) Then Err.Raise 5, , "Reporter name not found."
        key = ex_TableQuery.fn_NormalizeText(VBA.CStr(name))
        If positions.Exists(key) Then
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            fields("Personnel.PositionCode") = positions(key)
            Set supplement = New obj_LookupRecord
            If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize reporter TVO position."
            If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot merge reporter TVO position."
        End If
    Next record
    fn_TryApplyReporterPositions = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function fn_TryReporterTvoPosition( _
    ByVal reporterName As String, _
    ByRef positionCode As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim names As New Collection
    Dim positions As Object
    Dim key As String

    positionCode = VBA.vbNullString
    names.Add reporterName
    If Not fn_TryReporterTvoPositions(names, positions, diagnostic) Then Exit Function
    key = ex_TableQuery.fn_NormalizeText(reporterName)
    If positions.Exists(key) Then positionCode = positions(key)
    fn_TryReporterTvoPosition = True
End Function

Public Function fn_TryReporterTvoPositions( _
    ByVal reporterNames As Collection, _
    ByRef positions As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim result As obj_DataTable
    Dim requested As Object
    Dim available As Boolean
    Dim nameHeader As String
    Dim codeHeader As String
    Dim reporterName As Variant
    Dim row As Long
    Dim key As String
    Dim positionCode As String

    On Error GoTo EH
    Set positions = VBA.CreateObject("Scripting.Dictionary")
    positions.CompareMode = VBA.vbTextCompare
    Set requested = VBA.CreateObject("Scripting.Dictionary")
    requested.CompareMode = VBA.vbTextCompare
    If reporterNames.Count = 0 Then
        fn_TryReporterTvoPositions = True
        Exit Function
    End If
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        Err.Raise 5, , "Cannot initialize reporter TVO query."
    End If
    source.WorkbookPath = fn_Path(fn_Config("lookup.Personnel.source.path"))
    source.SheetName = fn_Config("lookup.Personnel.source.sheet")
    source.RangeAddress = fn_Config("lookup.Personnel.source.range")
    nameHeader = fn_Config("reporter.tvo.nameHeader")
    codeHeader = fn_Config("reporter.tvo.codeHeader")
    query.AddColumn nameHeader
    query.AddColumn codeHeader
    For Each reporterName In reporterNames
        requested(ex_TableQuery.fn_NormalizeText(VBA.CStr(reporterName))) = True
    Next reporterName
    query.AddCondition nameHeader, QueryIsNotEmpty, Empty, QueryText, True
    If Not ex_TableDataRuntime.fn_TryQueryLoaded(source, query, available, result, diagnostic) Then GoTo Cleanup
    If Not available Then
        service.Backend = QueryAuto
        If Not service.TryExecute(source, query, result, diagnostic) Then GoTo Cleanup
    End If
    For row = 1 To result.RowCount
        reporterName = fn_Text(result.ValueAt(row, ex_TableQuery.fn_HeaderIndex(result.Headers, nameHeader)))
        key = ex_TableQuery.fn_NormalizeText(VBA.CStr(reporterName))
        If requested.Exists(key) Then
            If positions.Exists(key) Then
                Err.Raise 5, , VBA.Replace$(fn_Config("reporter.tvo.ambiguous"), "{name}", VBA.CStr(reporterName))
            End If
            positionCode = fn_Text(result.ValueAt(row, ex_TableQuery.fn_HeaderIndex(result.Headers, codeHeader)))
            If VBA.Len(positionCode) = 0 Then
                Err.Raise 5, , VBA.Replace$(fn_Config("reporter.tvo.missingCode"), "{name}", VBA.CStr(reporterName))
            End If
            positions.Add key, positionCode
        End If
    Next row
    fn_TryReporterTvoPositions = True
Cleanup:
    If Not result Is Nothing Then result.Dispose
    If Not service Is Nothing Then service.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Sub fn_BeginConfigSnapshot()
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim data As Variant
    Dim keyColumn As Long
    Dim row As Long
    Dim key As String
    Dim snapshot As Object

    If m_configSnapshotDepth > 0 Then
        m_configSnapshotDepth = m_configSnapshotDepth + 1
        Exit Sub
    End If
    For Each sheet In ThisWorkbook.Worksheets
        For Each table In sheet.ListObjects
            If VBA.StrComp(table.Name, "tbConfig", VBA.vbTextCompare) = 0 Then
                keyColumn = table.ListColumns("Key").Index
                If table.DataBodyRange Is Nothing Or keyColumn >= table.ListColumns.Count Then
                    Err.Raise 5, , "Export configuration table is empty or invalid."
                End If
                data = table.DataBodyRange.Value2
                Set snapshot = VBA.CreateObject("Scripting.Dictionary")
                snapshot.CompareMode = VBA.vbTextCompare
                For row = 1 To UBound(data, 1)
                    key = VBA.Trim$(VBA.CStr(data(row, keyColumn)))
                    If snapshot.Exists(key) Then
                        Err.Raise 5, , "Duplicate configuration key: " & key
                    End If
                    snapshot.Add key, VBA.Trim$(VBA.CStr(data(row, keyColumn + 1)))
                Next row
                Set m_configSnapshot = snapshot
                m_configSnapshotDepth = 1
                Exit Sub
            End If
        Next table
    Next sheet
    Err.Raise 5, , "Export configuration table not found: tbConfig"
End Sub

Public Sub fn_EndConfigSnapshot()
    If m_configSnapshotDepth > 0 Then m_configSnapshotDepth = m_configSnapshotDepth - 1
    If m_configSnapshotDepth = 0 Then Set m_configSnapshot = Nothing
End Sub

Public Function fn_Count( _
    ByVal key As String, _
    Optional ByVal maximum As Long = 1000 _
) As Long
    Dim value As String
    Dim number As Double

    value = fn_Config(key)
    If Not VBA.IsNumeric(value) Then
        Err.Raise 5, , "Invalid configuration count: " & key
    End If
    number = VBA.CDbl(value)
    If number < 0 Or number > maximum Or number <> VBA.Fix(number) Then
        Err.Raise 5, , "Invalid configuration count: " & key
    End If
    fn_Count = VBA.CLng(number)
End Function

Public Function fn_Path(ByVal value As String) As String
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    value = VBA.Replace$(value, "/", "\")
    If VBA.Left$(value, 2) <> "\\" And VBA.Mid$(value, 2, 1) <> ":" Then
        value = fileSystem.BuildPath(ThisWorkbook.Path, value)
    End If
    fn_Path = fileSystem.GetAbsolutePathName(value)
End Function

Public Function fn_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        Err.Raise 5, , "An export value contains an Excel error."
    End If
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        Exit Function
    End If
    fn_Text = VBA.Trim$(VBA.CStr(value))
End Function

Public Function fn_Required( _
    ByVal values As Object, _
    ByVal key As String _
) As Variant
    If Not values.Exists(key) Then
        Err.Raise 5, , "Unknown export field: " & key
    End If
    If VBA.IsObject(values(key)) Then
        Err.Raise 5, , "Scalar export field required: " & key
    End If
    fn_Required = values(key)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------