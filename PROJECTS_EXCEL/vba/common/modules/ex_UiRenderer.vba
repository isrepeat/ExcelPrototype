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
Public Function fn_RenderPages( _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    fn_RenderPages = ex_UiRuntime.fn_RenderPages(uiFolderRelativePath, uiBindingContext)
End Function

Public Function fn_RenderActivePage( _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    fn_RenderActivePage = ex_UiRuntime.fn_RenderActivePage(uiFolderRelativePath, uiBindingContext)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------