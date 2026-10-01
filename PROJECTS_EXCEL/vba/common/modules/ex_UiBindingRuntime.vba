Attribute VB_Name = "ex_UiBindingRuntime"
Option Explicit

Private Const BINDING_PREFIX As String = "{Binding "
Private Const BINDING_SUFFIX As String = "}"

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
Public Function fn_TryResolveText( _
    ByVal rawText As String, _
    ByVal uiBindingContext As obj_UiBindingContext, _
    ByRef outText As String _
) As Boolean
    Dim resolvedValue As Variant
    Dim resolvedObject As Object
    Dim isObject As Boolean

    If Not fn_TryResolveValue(rawText, uiBindingContext, resolvedValue, resolvedObject, isObject) Then Exit Function
    If isObject Then
        VBA.MsgBox "A text binding must resolve to a scalar value: " & rawText, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    outText = VBA.CStr(resolvedValue)
    fn_TryResolveText = True
End Function

Public Function fn_TryResolveCommand( _
    ByVal rawText As String, _
    ByVal uiBindingContext As obj_UiBindingContext, _
    ByRef outUiCommand As obj_UiCommand _
) As Boolean
    Dim resolvedValue As Variant
    Dim resolvedObject As Object
    Dim isObject As Boolean

    Set outUiCommand = Nothing
    If Not fn_TryResolveValue(rawText, uiBindingContext, resolvedValue, resolvedObject, isObject) Then Exit Function
    If Not isObject Or Not TypeOf resolvedObject Is obj_UiCommand Then
        VBA.MsgBox "A command binding must resolve to obj_UiCommand: " & rawText, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set outUiCommand = resolvedObject
    fn_TryResolveCommand = True
End Function

Public Function fn_TryResolveValue( _
    ByVal rawText As String, _
    ByVal uiBindingContext As obj_UiBindingContext, _
    ByRef outValue As Variant, _
    ByRef outObject As Object, _
    ByRef outIsObject As Boolean _
) As Boolean
    Dim bindingBody As String
    Dim sourceName As String
    Dim bindingPath As String

    Set outObject = Nothing
    outIsObject = False
    rawText = VBA.Trim$(rawText)
    If Not private_TryExtractBindingBody(rawText, bindingBody) Then
        outValue = rawText
        fn_TryResolveValue = True
        Exit Function
    End If
    If uiBindingContext Is Nothing Then
        VBA.MsgBox "A binding context is required for: " & rawText, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    If Not private_TryReadArgument(bindingBody, "Source", sourceName) Then sourceName = "Text"
    If Not private_TryReadArgument(bindingBody, "Path", bindingPath) Then
        VBA.MsgBox "Binding Path is required: " & rawText, VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    If Not uiBindingContext.TryGetValue(sourceName, bindingPath, outValue, outObject, outIsObject) Then
        VBA.MsgBox "Binding was not found: Source=" & sourceName & "; Path=" & bindingPath, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_TryResolveValue = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryExtractBindingBody(ByVal rawText As String, ByRef outBody As String) As Boolean
    If VBA.Len(rawText) <= VBA.Len(BINDING_PREFIX) Then Exit Function
    If VBA.StrComp(VBA.Left$(rawText, VBA.Len(BINDING_PREFIX)), BINDING_PREFIX, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(rawText, VBA.Len(BINDING_SUFFIX)) <> BINDING_SUFFIX Then Exit Function
    outBody = VBA.Trim$(VBA.Mid$(rawText, VBA.Len(BINDING_PREFIX) + 1, VBA.Len(rawText) - VBA.Len(BINDING_PREFIX) - 1))
    private_TryExtractBindingBody = True
End Function

Private Function private_TryReadArgument(ByVal bindingBody As String, ByVal argumentName As String, ByRef outValue As String) As Boolean
    Dim arguments As Variant
    Dim argumentText As Variant
    Dim separatorPosition As Long
    Dim currentName As String

    arguments = VBA.Split(bindingBody, ";")
    For Each argumentText In arguments
        separatorPosition = VBA.InStr(1, VBA.CStr(argumentText), "=", VBA.vbBinaryCompare)
        If separatorPosition > 0 Then
            currentName = VBA.Trim$(VBA.Left$(VBA.CStr(argumentText), separatorPosition - 1))
            If VBA.StrComp(currentName, argumentName, VBA.vbTextCompare) = 0 Then
                outValue = VBA.Trim$(VBA.Mid$(VBA.CStr(argumentText), separatorPosition + 1))
                private_TryReadArgument = (VBA.Len(outValue) > 0)
                Exit Function
            End If
        End If
    Next argumentText
End Function