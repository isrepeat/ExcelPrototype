Attribute VB_Name = "ex_UiPageLoader"
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
Public Function fn_TryLoad( _
    ByVal xamlPath As String, _
    ByRef outUiPageDefinition As obj_UiPageDefinition _
) As Boolean
    Dim document As Object
    Dim uiPageDefinition As obj_UiPageDefinition

    Set outUiPageDefinition = Nothing
    If Not ex_UiParser.fn_TryLoadPage(xamlPath, document) Then Exit Function

    Set uiPageDefinition = New obj_UiPageDefinition
    If Not uiPageDefinition.fn_Initialize(document, xamlPath) Then
        VBA.MsgBox "The XAML page definition is invalid: " & xamlPath, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set outUiPageDefinition = uiPageDefinition
    fn_TryLoad = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------