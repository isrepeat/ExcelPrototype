Option Explicit

Private m_uiFolderPath As String

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    m_uiFolderPath = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_SetUiFolder(ByVal uiFolderPath As String)
    m_uiFolderPath = VBA.Trim$(uiFolderPath)
End Sub

Public Function fn_TryGetUiFolder(ByRef outUiFolderPath As String) As Boolean
    outUiFolderPath = m_uiFolderPath
    If VBA.Len(outUiFolderPath) > 0 Then
        fn_TryGetUiFolder = True
        Exit Function
    End If
    fn_TryGetUiFolder = ex_Core.fn_TryGetWorkbookConfigFolder( _
        "ThisWorkbook::uiPath", outUiFolderPath)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------