VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TDB_DependencyLoader"
Option Explicit

Implements obj_ITableSnapshotLoader

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
End Sub

Public Function TryBuild( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByRef diagnostic As String _
) As Boolean
    Dim prefix As Variant
    Dim source As obj_TableSource
    Dim profile As obj_LookupProfile
    Dim lookup As obj_LookupService
    Dim orders As obj_OrderResolver
    Dim fileSystem As Object
    Dim signatures As Object
    Dim path As Variant
    Dim busyEntered As Boolean
    Dim sourceDefinitions As Variant
    Dim progressStep As Long
    Dim progressTotal As Long

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Dependency loader is unavailable."
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set signatures = VBA.CreateObject("Scripting.Dictionary")
    signatures.CompareMode = VBA.vbTextCompare
    busyEntered = ex_UiBusy.fn_Begin(ex_TravelDocumentsBuilder.fn_Config("text.DependenciesLoading"))
    sourceDefinitions = VBA.Split(ex_TravelDocumentsBuilder.fn_Config("snapshots.sources"), "|")
    progressTotal = UBound(sourceDefinitions) + 1 + 3

    ' Load every configured source before preparing projections.
    For Each prefix In sourceDefinitions
        Set source = New obj_TableSource
        If Not source.Initialize() Then Err.Raise 5, , "Cannot initialize dependency source."
        path = ex_TravelDocumentsBuilder.fn_Config(VBA.CStr(prefix) & ".path")
        If fileSystem.GetDriveName(VBA.CStr(path)) = VBA.vbNullString And VBA.Left$(VBA.CStr(path), 2) <> "\\" Then
            path = fileSystem.BuildPath(ThisWorkbook.Path, VBA.CStr(path))
        End If
        source.WorkbookPath = fileSystem.GetAbsolutePathName(VBA.CStr(path))
        source.SheetName = ex_TravelDocumentsBuilder.fn_Config(VBA.CStr(prefix) & ".sheet")
        source.RangeAddress = ex_TravelDocumentsBuilder.fn_Config(VBA.CStr(prefix) & ".range")
        If Not source.TryValidate(diagnostic) Then GoTo Cleanup
        If Not signatures.Exists(source.WorkbookPath) Then
            signatures.Add source.WorkbookPath, private_FileStamp(source.WorkbookPath)
        End If
        private_Progress progressStep, progressTotal, "Table", source.SheetName
        If Not snapshot.TryLoad(source, diagnostic) Then GoTo Cleanup
        source.Dispose
        Set source = Nothing
    Next prefix

    ' Prepare lookup and orders against the staged generation only.
    Set profile = New obj_LookupProfile
    Set lookup = New obj_LookupService
    If Not profile.Initialize("TravelDocumentsBuilder::lookup.Personnel", diagnostic) Then GoTo Cleanup
    If Not lookup.Initialize(profile) Then Err.Raise 5, , "Cannot initialize dependency lookup."
    private_Progress progressStep, progressTotal, "Profile", profile.Source.SheetName
    If Not lookup.PrepareSnapshot(snapshot, diagnostic) Then GoTo Cleanup
    Set orders = New obj_OrderResolver
    If Not orders.Initialize("TravelDocumentsBuilder::orders", "TravelDocumentsBuilder::text", diagnostic) Then GoTo Cleanup
    private_Progress progressStep, progressTotal, "Orders"
    If Not orders.PrepareSnapshot(snapshot, diagnostic) Then GoTo Cleanup

    ' Publish only a consistent set of saved files.
    private_Progress progressStep, progressTotal, "Verify"
    For Each path In signatures.Keys
        If signatures(path) <> private_FileStamp(VBA.CStr(path)) Then
            Err.Raise 5, , VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("text.DependenciesSourceChanged"), "{path}", VBA.CStr(path))
        End If
    Next path
    TryBuild = True
Cleanup:
    If Not source Is Nothing Then source.Dispose
    If Not lookup Is Nothing Then lookup.Dispose
    If Not profile Is Nothing Then profile.Dispose
    If Not orders Is Nothing Then orders.Dispose
    If busyEntered Then ex_UiBusy.fn_End
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

' //
' // Interface
' //
Private Function obj_ITableSnapshotLoader_TryBuild( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByRef diagnostic As String _
) As Boolean
    obj_ITableSnapshotLoader_TryBuild = Me.TryBuild(snapshot, diagnostic)
End Function

' //
' // Private
' //
Private Function private_FileStamp(ByVal path As String) As String
    Dim fileSystem As Object
    Dim file As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set file = fileSystem.GetFile(path)
    private_FileStamp = VBA.CStr(file.Size) & "|" & VBA.CStr(VBA.CDbl(file.DateLastModified))
End Function

Private Sub private_Progress( _
    ByRef currentStep As Long, _
    ByVal total As Long, _
    ByVal stage As String, _
    Optional ByVal source As String = VBA.vbNullString _
)
    Dim text As String
    Dim stageText As String

    currentStep = currentStep + 1
    stageText = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("snapshots.progress." & stage), "{source}", source)
    text = ex_TravelDocumentsBuilder.fn_Config("text.DependenciesProgress")
    text = VBA.Replace$(text, "{step}", VBA.CStr(currentStep))
    text = VBA.Replace$(text, "{total}", VBA.CStr(total))
    text = VBA.Replace$(text, "{stage}", stageText)
    ex_Core.fn_StatusBar_Show text
End Sub