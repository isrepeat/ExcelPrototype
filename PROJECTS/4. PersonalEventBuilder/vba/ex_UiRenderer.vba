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
Public Sub fn_RenderPages()
    ex_UiRuntime.fn_RenderPages
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------