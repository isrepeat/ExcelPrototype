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
Public Function fn_TryUpdate( _
    ByRef diagnostic As String, _
    Optional ByVal sourcePath As String = "" _
) As Boolean
    Dim updater As Workbook
    Dim fileSystem As Object
    Dim vbaPath As String
    Dim updaterPath As String

    On Error GoTo EH
    If Not ex_Core.fn_TryGetWorkbookConfigValue("ThisWorkbook::vbaPath", vbaPath) Then
        Err.Raise 5, , "Required configuration key unavailable: ThisWorkbook::vbaPath"
    End If
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If fileSystem.GetDriveName(vbaPath) = "" And VBA.Left$(vbaPath, 2) <> "\\" Then
        vbaPath = fileSystem.GetAbsolutePathName(ThisWorkbook.Path & "\" & vbaPath)
    End If
    updaterPath = fileSystem.GetParentFolderName(vbaPath) & "\WorkbookUpdater\WorkbookUpdater.xlam"
    If Not fileSystem.FileExists(updaterPath) Then Err.Raise 5, , "Build WorkbookUpdater.xlam first: " & updaterPath
    On Error Resume Next
    Set updater = Application.Workbooks("WorkbookUpdater.xlam")
    On Error GoTo EH
    If updater Is Nothing Then Set updater = Application.Workbooks.Open(updaterPath)
    If VBA.StrComp(updater.FullName, updaterPath, VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , "A different WorkbookUpdater.xlam is already loaded."
    End If
    fn_TryUpdate = VBA.CBool(Application.Run( _
        "'WorkbookUpdater.xlam'!ex_ConfigRefresh.fn_TryUpdateWorkbook", ThisWorkbook, diagnostic, sourcePath))
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------