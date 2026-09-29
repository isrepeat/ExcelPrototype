Attribute VB_Name = "ex_UiRenderer"
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
Public Sub fn_RenderPages(ByVal uiFolderRelativePath As String)
    ex_UiRuntime.fn_RenderPages uiFolderRelativePath
End Sub

Public Sub fn_RenderActivePage(ByVal uiFolderRelativePath As String)
    ex_UiRuntime.fn_RenderActivePage uiFolderRelativePath
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------