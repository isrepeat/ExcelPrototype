Attribute VB_Name = "ex_UiPageFactory"
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
Public Function fn_Create(ByVal pageId As String) As obj_IPage
    Dim factoryName As String

    On Error GoTo Failed
    factoryName = ex_WorkbookCallbacks.fn_Resolve("ThisWorkbook::pageFactory")
    Set fn_Create = Application.Run(factoryName, pageId)
    If fn_Create Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 2211, "ex_UiPageFactory.fn_Create", _
            "Configured page factory returned no page: " & pageId
    End If
    Exit Function
Failed:
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, "Configuration"
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------