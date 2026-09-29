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
Public Sub fn_Initialize()
    ex_UiRenderer.fn_RenderPages
End Sub

Public Sub fn_HelloWorld()
    VBA.MsgBox "Hello World from PersonalEventBuilder.", VBA.vbInformation, _
        "PersonalEventBuilder"
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------