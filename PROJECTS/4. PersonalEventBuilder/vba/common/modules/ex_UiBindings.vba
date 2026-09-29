Option Explicit

Private shapeCommands As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set shapeCommands = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Reset()
    Set shapeCommands = VBA.CreateObject("Scripting.Dictionary")
    shapeCommands.CompareMode = VBA.vbTextCompare
End Sub

Public Sub fn_Register(ByVal shapeName As String, ByVal uiCommand As obj_UiCommand)
    If uiCommand Is Nothing Then Exit Sub
    If shapeCommands Is Nothing Then fn_Reset
    Set shapeCommands(shapeName) = uiCommand
End Sub

Public Function fn_TryGetCommand(ByVal shapeName As String, ByRef outUiCommand As obj_UiCommand) As Boolean
    Set outUiCommand = Nothing
    If shapeCommands Is Nothing Then Exit Function
    If Not shapeCommands.Exists(shapeName) Then Exit Function
    Set outUiCommand = shapeCommands(shapeName)
    fn_TryGetCommand = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------