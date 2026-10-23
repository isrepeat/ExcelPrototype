Attribute VB_Name = "ex_TableSnapshots"
Option Explicit

Private m_loader As obj_ITableSnapshotLoader
Private m_active As obj_TableSnapshotSet
Private m_building As Boolean
Private m_generation As Long
Private m_notLoadedMessage As String

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_active = Nothing
    Set m_loader = Nothing
    m_notLoadedMessage = VBA.vbNullString
    m_building = False
    m_generation = m_generation + 1
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Register(ByVal loader As obj_ITableSnapshotLoader, ByVal notLoadedMessage As String)
    fn_Module_Dispose
    Set m_loader = loader
    m_notLoadedMessage = notLoadedMessage
End Sub

Public Function fn_Generation() As Long
    fn_Generation = m_generation
End Function

Public Sub fn_Invalidate()
    Set m_active = Nothing
    m_generation = m_generation + 1
End Sub

Public Function fn_Current() As obj_TableSnapshotSet
    If m_active Is Nothing Then
        If Not m_loader Is Nothing Then
            Err.Raise 5, , m_notLoadedMessage
        End If
    End If
    Set fn_Current = m_active
End Function

Public Function fn_Active() As obj_TableSnapshotSet
    Set fn_Active = m_active
End Function

Public Function fn_Refresh(ByRef diagnostic As String) As Boolean
    Dim staged As obj_TableSnapshotSet
    Dim startedAt As Double

    If m_building Then
        diagnostic = "Snapshot loading is already in progress."
        Exit Function
    End If
    On Error GoTo EH
    startedAt = VBA.Timer
    If m_loader Is Nothing Then
        Err.Raise 5, , "Snapshot loader is not registered."
    End If
    m_building = True
    Set staged = New obj_TableSnapshotSet
    If Not staged.Initialize() Then
        Err.Raise 5, , "Cannot initialize snapshot."
    End If
    If Not m_loader.TryBuild(staged, diagnostic) Then
        GoTo Cleanup
    End If
    ' Publish the complete generation only after all preparation succeeds.
    Set m_active = staged
    m_generation = m_generation + 1
    fn_Refresh = True
Cleanup:
    m_building = False
    ex_Core.fn_Diagnostic_WritePerf "Dependencies.Refresh | Success=" & VBA.CStr(fn_Refresh), startedAt
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_TryQuery( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByRef available As Boolean, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim snapshot As obj_TableSnapshotSet

    On Error GoTo EH
    available = False
    fn_TryQuery = True
    If m_loader Is Nothing Then
        Exit Function
    End If
    Set snapshot = fn_Current()
    If Not snapshot.Contains(source) Then
        If snapshot.OwnsWorkbook(source) Then
            Err.Raise 5, , "Dependency source is not registered in the snapshot: " & ex_TableSnapshots.fn_SourceKey(source)
        End If
        Exit Function
    End If
    available = True
    fn_TryQuery = snapshot.TryQuery(source, query, output, diagnostic)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    fn_TryQuery = False
End Function

Public Function fn_SourceKey(ByVal source As obj_TableSource) As String
    fn_SourceKey = VBA.LCase$(VBA.Replace$(source.WorkbookPath, "/", "\")) & VBA.vbNullChar & _
        VBA.LCase$(source.SheetName) & VBA.vbNullChar & VBA.LCase$(source.RangeAddress) & VBA.vbNullChar & VBA.LCase$(source.TableName)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------