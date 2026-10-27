Option Explicit

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub Workbook_Open()
    Dim context As Object
    Dim initialized As Boolean
    Dim errorText As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(context) Then
        Exit Sub
    End If
    On Error GoTo Failed
    initialized = ex_PeopleRegistry.fn_Initialize()
    ex_RuntimeLifecycle.fn_Leave context
    If initialized Then
        ex_PeopleRegistry.ClearResults
    End If
    Exit Sub
Failed:
    errorText = Err.Description
    ex_RuntimeLifecycle.fn_Leave context
    ex_WindowsUi.fn_ShowMessage errorText, vbCritical
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------