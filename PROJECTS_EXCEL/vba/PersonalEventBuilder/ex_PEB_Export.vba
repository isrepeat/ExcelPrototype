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