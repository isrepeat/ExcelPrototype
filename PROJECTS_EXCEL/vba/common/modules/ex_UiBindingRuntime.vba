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
        ex_WindowsUi.fn_ShowMessage "A text binding must resolve to a scalar value: " & rawText, _
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
        ex_WindowsUi.fn_ShowMessage "A command binding must resolve to obj_UiCommand: " & rawText, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set outUiCommand = resolvedObject
    fn_TryResolveCommand = True
End Function

Public Function fn_TryParseBinding( _
    ByVal rawBinding As String, _
    ByVal defaultSource As String, _
    ByRef outSourceName As String, _
    ByRef outBindingPath As String, _
    Optional ByVal context As obj_UiBindingContext _
) As Boolean
    Dim bindingBody As String
    Dim separatorPosition As Long
    Dim rootSource As Boolean
    Dim qualified As Boolean
    Dim contextPosition As Long
    Dim contextPath As String

    outSourceName = VBA.vbNullString
    outBindingPath = VBA.vbNullString
    If Not private_TryExtractBindingBody(VBA.Trim$(rawBinding), bindingBody) Then Exit Function
    If Not private_TryReadArgument(bindingBody, "Path", outBindingPath) Then Exit Function
    If Not context Is Nothing Then rootSource = context.HasSource(outBindingPath)
    If Not private_TryReadArgument(bindingBody, "Source", outSourceName) Then
        separatorPosition = VBA.InStr(1, outBindingPath, ".", VBA.vbBinaryCompare)
        qualified = (separatorPosition > 0)
        If qualified And Not context Is Nothing And VBA.Len(defaultSource) > 0 Then
            qualified = context.HasSource(VBA.Trim$(VBA.Left$(outBindingPath, separatorPosition - 1)))
        End If
        If qualified Then
            outSourceName = VBA.Trim$(VBA.Left$(outBindingPath, separatorPosition - 1))
            outBindingPath = VBA.Trim$(VBA.Mid$(outBindingPath, separatorPosition + 1))
        ElseIf rootSource Then
            outSourceName = outBindingPath
            outBindingPath = VBA.vbNullString
        Else
            contextPosition = VBA.InStr(1, defaultSource, ".", VBA.vbBinaryCompare)
            If contextPosition > 0 Then
                outSourceName = VBA.Left$(defaultSource, contextPosition - 1)
                contextPath = VBA.Mid$(defaultSource, contextPosition + 1)
                outBindingPath = contextPath & "." & outBindingPath
            Else
                outSourceName = defaultSource
            End If
        End If
    End If
    fn_TryParseBinding = (VBA.Len(outSourceName) > 0)
End Function

Public Function fn_TryDataContext( _
    ByVal raw As String, _
    ByVal inheritedContext As String, _
    ByVal context As obj_UiBindingContext, _
    ByRef resolvedContext As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As String, path As String
    Dim value As Variant, sourceObject As Object, isObject As Boolean

    resolvedContext = inheritedContext
    If VBA.Len(raw) = 0 Then
        fn_TryDataContext = True
        Exit Function
    End If
    If Not fn_TryParseBinding(raw, inheritedContext, source, path, context) Then
        diagnostic = "dataContext requires a Binding expression: " & raw
        Exit Function
    End If
    If Not context.TryGetValue(source, path, value, sourceObject, isObject) Then
        diagnostic = "dataContext was not found: " & raw
        Exit Function
    End If
    If Not isObject Then
        diagnostic = "dataContext must resolve to an object: " & raw
        Exit Function
    End If
    resolvedContext = source
    If VBA.Len(path) > 0 Then resolvedContext = source & "." & path
    fn_TryDataContext = True
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
        ex_WindowsUi.fn_ShowMessage "A binding context is required for: " & rawText, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    If Not fn_TryParseBinding(rawText, VBA.vbNullString, sourceName, bindingPath, uiBindingContext) Then
        ex_WindowsUi.fn_ShowMessage "Binding Path is required: " & rawText, VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    If Not uiBindingContext.TryGetValue(sourceName, bindingPath, outValue, outObject, outIsObject) Then
        ex_WindowsUi.fn_ShowMessage "Binding was not found: Source=" & sourceName & "; Path=" & bindingPath, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_TryResolveValue = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_TryExtractBindingBody( _
    ByVal rawText As String, _
    ByRef outBody As String _
) As Boolean
    If VBA.Len(rawText) <= VBA.Len(BINDING_PREFIX) Then Exit Function
    If VBA.StrComp(VBA.Left$(rawText, VBA.Len(BINDING_PREFIX)), BINDING_PREFIX, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(rawText, VBA.Len(BINDING_SUFFIX)) <> BINDING_SUFFIX Then Exit Function
    outBody = VBA.Trim$(VBA.Mid$(rawText, VBA.Len(BINDING_PREFIX) + 1, VBA.Len(rawText) - VBA.Len(BINDING_PREFIX) - 1))
    private_TryExtractBindingBody = True
End Function

Private Function private_TryReadArgument( _
    ByVal bindingBody As String, _
    ByVal argumentName As String, _
    ByRef outValue As String _
) As Boolean
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
' --------------------------------------
' } // namespace Private
' --------------------------------------