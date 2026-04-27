Attribute VB_Name = "ex_BindingRuntime"
Option Explicit

Private Const BINDING_PREFIX As String = "{Binding "
Private Const BINDING_SUFFIX As String = "}"

' //
' // API
' //
' Callstack[1]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_ButtonControlVM.obj_IControl_Configure -> ex_BindingRuntime.m_TryResolveTextBinding
' Callstack[2]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_LabelControlVM.obj_IControl_Configure -> ex_BindingRuntime.m_TryResolveTextBinding
Public Function m_TryResolveTextBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outText As String _
) As Boolean
    Dim resolvedValue As Variant

    If Not private_TryResolveBindingValue(rawText, sourceObject, resolvedValue) Then Exit Function

    If VBA.IsObject(resolvedValue) Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: text binding must resolve to scalar value."
        Exit Function
    End If

    outText = VBA.CStr(resolvedValue)
    m_TryResolveTextBinding = True
End Function


' Callstack[1]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_ButtonControlVM.obj_IControl_Configure -> ex_BindingRuntime.m_TryResolveMacroBinding
' Callstack[2]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_SelectControlVM.obj_IControl_Configure -> ex_BindingRuntime.m_TryResolveMacroBinding
Public Function m_TryResolveMacroBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outMacroRef As String _
) As Boolean
    Dim resolvedValue As Variant
    Dim macroName As String

    If Not private_TryResolveBindingValue(rawText, sourceObject, resolvedValue) Then Exit Function

    If VBA.IsObject(resolvedValue) Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: macro binding must resolve to text value."
        Exit Function
    End If

    macroName = VBA.Trim$(VBA.CStr(resolvedValue))
    If VBA.Len(macroName) = 0 Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: macro binding resolved to an empty value."
        Exit Function
    End If

    outMacroRef = private_QualifyMacroName(macroName)
    m_TryResolveMacroBinding = True
End Function


' // Helper for Visibility attribute.
' Callstack[1]: ex_XmlLayoutEngine.private_TryIsNodeVisible -> ex_BindingRuntime.m_TryResolveVisibilityBinding
' Callstack[2]: ex_LayoutListRenderer.private_ApplyNodeBindingsRecursive -> ex_BindingRuntime.m_TryResolveVisibilityBinding
' Callstack[3]: ex_LayoutItemControlRenderer.private_ApplyNodeBindingsRecursive -> ex_BindingRuntime.m_TryResolveVisibilityBinding
' Callstack[4]: obj_TableListControlVM.private_TryApplyItemVisibilityFilter -> ex_BindingRuntime.m_TryResolveVisibilityBinding
Public Function m_TryResolveVisibilityBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outVisible As Boolean _
) As Boolean
    Dim resolvedValue As Variant

    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then
        outVisible = True
        m_TryResolveVisibilityBinding = True
        Exit Function
    End If

    If Not m_TryResolveValueBinding(rawText, sourceObject, resolvedValue) Then Exit Function
    If Not private_TryParseBooleanVariant(resolvedValue, outVisible) Then
        If VBA.IsObject(resolvedValue) Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: visibility value resolved to object '" & VBA.TypeName(resolvedValue) & "'. Expected boolean-compatible value."
        Else
            ex_Core.m_Diagnostic_LogError "PrototypeNew: visibility value '" & VBA.CStr(resolvedValue) & "' is invalid. Supported values: true/false/visible/collapsed."
        End If
        Exit Function
    End If

    m_TryResolveVisibilityBinding = True
End Function


' Callstack[1]: ex_BindingRuntime.m_TryResolveVisibilityBinding -> ex_BindingRuntime.m_TryResolveValueBinding
' Callstack[2]: ex_RuntimeSourceResolver.m_TryResolveItemsSource -> ex_BindingRuntime.m_TryResolveValueBinding
' Callstack[3]: ex_RuntimeSourceResolver.m_TryResolveObjectSource -> ex_BindingRuntime.m_TryResolveValueBinding
' Callstack[4]: ex_LayoutListRenderer.private_ApplyNodeBindingsRecursive -> ex_BindingRuntime.m_TryResolveValueBinding
' Callstack[5]: ex_LayoutListRenderer.private_TryResolveItemsSourceForMeasure -> ex_BindingRuntime.m_TryResolveValueBinding
' Callstack[6]: ex_LayoutItemControlRenderer.private_ApplyNodeBindingsRecursive -> ex_BindingRuntime.m_TryResolveValueBinding
' Callstack[7]: ex_LayoutItemControlRenderer.private_TryResolveObjectSourceForMeasure -> ex_BindingRuntime.m_TryResolveValueBinding
Public Function m_TryResolveValueBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outValue As Variant _
) As Boolean
    If Not private_TryResolveBindingValue(rawText, sourceObject, outValue) Then Exit Function
    m_TryResolveValueBinding = True
End Function

' //
' // Internal
' //
Private Function private_TryResolveBindingValue( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outValue As Variant _
) As Boolean
    Dim bindingBody As String
    Dim methodName As String
    Dim moduleName As String
    Dim mappedText As String
    Dim usedConditionalBranch As Boolean
    Dim bindingPath As String

    If Not private_TryExtractBindingBody(rawText, bindingBody) Then
        outValue = rawText
        private_TryResolveBindingValue = True
        Exit Function
    End If

    If private_TryExtractNamedArg(bindingBody, "Method", methodName) Then
        If private_TryExtractNamedArg(bindingBody, "Module", moduleName) Then
            If VBA.InStr(1, methodName, ".", VBA.vbBinaryCompare) = 0 Then
                methodName = moduleName & "." & methodName
            End If
        End If

        methodName = VBA.Trim$(methodName)
        If VBA.Len(methodName) = 0 Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: binding expression contains empty Method value."
            Exit Function
        End If

        outValue = methodName
        private_TryResolveBindingValue = True
        Exit Function
    End If

    If Not private_TryResolveConditionalBindingAsText(bindingBody, sourceObject, mappedText, usedConditionalBranch) Then Exit Function
    If usedConditionalBranch Then
        outValue = mappedText
        private_TryResolveBindingValue = True
        Exit Function
    End If

    If private_TryExtractNamedArg(bindingBody, "Path", bindingPath) Then
    Else
        bindingPath = VBA.Trim$(bindingBody)
    End If

    If VBA.Len(bindingPath) = 0 Then bindingPath = "."
    If Not private_TryReadBindingPathValue(sourceObject, bindingPath, outValue) Then Exit Function

    private_TryResolveBindingValue = True
End Function


Private Function private_TryResolveConditionalBindingAsText( _
    ByVal bindingBody As String, _
    ByVal sourceObject As Object, _
    ByRef outText As String, _
    ByRef outUsedConditionalBranch As Boolean _
) As Boolean
    Dim bindingPath As String
    Dim opText As String
    Dim valueText As String
    Dim trueAsText As String
    Dim falseAsText As String
    Dim resolvedValue As Variant
    Dim hasOp As Boolean
    Dim hasValue As Boolean
    Dim hasTrueAs As Boolean
    Dim hasFalseAs As Boolean
    Dim conditionResult As Boolean

    outUsedConditionalBranch = False

    hasOp = private_TryExtractNamedArg(bindingBody, "Op", opText)
    hasValue = private_TryExtractNamedArg(bindingBody, "Value", valueText)
    hasTrueAs = private_TryExtractNamedArg(bindingBody, "TrueAs", trueAsText)
    hasFalseAs = private_TryExtractNamedArg(bindingBody, "FalseAs", falseAsText)

    If Not hasOp And Not hasValue And Not hasTrueAs And Not hasFalseAs Then
        private_TryResolveConditionalBindingAsText = True
        Exit Function
    End If

    outUsedConditionalBranch = True

    If sourceObject Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional binding requires source object."
        Exit Function
    End If

    If private_TryExtractNamedArg(bindingBody, "Path", bindingPath) Then
    Else
        bindingPath = VBA.Trim$(bindingBody)
    End If

    If VBA.Len(bindingPath) = 0 Then bindingPath = "."
    If Not private_TryReadBindingPathValue(sourceObject, bindingPath, resolvedValue) Then Exit Function

    If hasValue And Not hasOp Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional binding argument 'Value' requires 'Op'."
        Exit Function
    End If

    If hasOp Then
        opText = VBA.LCase$(VBA.Trim$(opText))
        If Not private_TryEvaluateConditionalOperation(resolvedValue, opText, valueText, conditionResult) Then Exit Function
    Else
        If Not private_TryParseBooleanVariant(resolvedValue, conditionResult) Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional binding path '" & bindingPath & "' must resolve to boolean-compatible value when Op is omitted."
            Exit Function
        End If
    End If

    If Not private_TryMapConditionalResultToText( _
        conditionResult, _
        hasTrueAs, trueAsText, _
        hasFalseAs, falseAsText, _
        outText) Then Exit Function

    private_TryResolveConditionalBindingAsText = True
End Function


Private Function private_TryMapConditionalResultToText( _
    ByVal conditionResult As Boolean, _
    ByVal hasTrueAs As Boolean, _
    ByVal trueAsText As String, _
    ByVal hasFalseAs As Boolean, _
    ByVal falseAsText As String, _
    ByRef outText As String _
) As Boolean
    If conditionResult Then
        If hasTrueAs Then
            outText = trueAsText
        Else
            outText = "True"
        End If
    Else
        If hasFalseAs Then
            outText = falseAsText
        Else
            outText = "False"
        End If
    End If

    private_TryMapConditionalResultToText = True
End Function


Private Function private_TryExtractBindingBody(ByVal rawText As String, ByRef outBody As String) As Boolean
    Dim normalized As String
    Dim prefixLen As Long

    normalized = VBA.Trim$(rawText)
    prefixLen = VBA.Len(BINDING_PREFIX)

    If VBA.Len(normalized) < prefixLen + 1 Then Exit Function
    If VBA.StrComp(VBA.Left$(normalized, prefixLen), BINDING_PREFIX, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(normalized, VBA.Len(BINDING_SUFFIX)) <> BINDING_SUFFIX Then Exit Function

    outBody = VBA.Trim$(VBA.Mid$(normalized, prefixLen + 1, VBA.Len(normalized) - prefixLen - 1))
    private_TryExtractBindingBody = True
End Function


Private Function private_TryExtractNamedArg( _
    ByVal bindingBody As String, _
    ByVal argName As String, _
    ByRef outValue As String _
) As Boolean
    Dim argPos As Long
    Dim valueStart As Long
    Dim valueText As String
    Dim sepPos As Long

    argPos = VBA.InStr(1, bindingBody, argName & "=", VBA.vbTextCompare)
    If argPos = 0 Then Exit Function

    valueStart = argPos + VBA.Len(argName) + 1
    valueText = VBA.Mid$(bindingBody, valueStart)

    sepPos = VBA.InStr(1, valueText, ";", VBA.vbBinaryCompare)
    If sepPos = 0 Then sepPos = VBA.InStr(1, valueText, ",", VBA.vbBinaryCompare)

    If sepPos > 0 Then
        outValue = VBA.Trim$(VBA.Left$(valueText, sepPos - 1))
    Else
        outValue = VBA.Trim$(valueText)
    End If

    outValue = private_Unquote(outValue)
    private_TryExtractNamedArg = (VBA.Len(outValue) > 0)
End Function


Private Function private_Unquote(ByVal textValue As String) As String
    textValue = VBA.Trim$(textValue)
    If VBA.Len(textValue) < 2 Then
        private_Unquote = textValue
        Exit Function
    End If

    If (VBA.Left$(textValue, 1) = VBA.Chr$(34) And VBA.Right$(textValue, 1) = VBA.Chr$(34)) Or _
       (VBA.Left$(textValue, 1) = "'" And VBA.Right$(textValue, 1) = "'") Then
        private_Unquote = VBA.Mid$(textValue, 2, VBA.Len(textValue) - 2)
    Else
        private_Unquote = textValue
    End If
End Function


Private Function private_TryReadBindingPathValue( _
    ByVal sourceObject As Object, _
    ByVal bindingPath As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim segments As Variant
    Dim segmentIndex As Long
    Dim segmentName As String
    Dim currentObject As Object
    Dim memberIsObject As Boolean
    Dim memberObject As Object
    Dim memberScalar As Variant

    If sourceObject Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: binding source object is not specified."
        Exit Function
    End If

    bindingPath = VBA.Trim$(bindingPath)
    If VBA.Len(bindingPath) = 0 Or bindingPath = "." Then
        Set outValue = sourceObject
        private_TryReadBindingPathValue = True
        Exit Function
    End If

    Set currentObject = sourceObject
    segments = VBA.Split(bindingPath, ".")

    For segmentIndex = LBound(segments) To UBound(segments)
        segmentName = VBA.Trim$(VBA.CStr(segments(segmentIndex)))
        If VBA.Len(segmentName) = 0 Then GoTo ContinueLoop

        If currentObject Is Nothing Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: binding path '" & bindingPath & "' reached Nothing before segment '" & segmentName & "'."
            Exit Function
        End If

        If Not private_TryReadMemberValue(currentObject, segmentName, memberIsObject, memberObject, memberScalar) Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: member '" & segmentName & "' was not found on object '" & VBA.TypeName(currentObject) & "'."
            Exit Function
        End If

        If segmentIndex < UBound(segments) Then
            If Not memberIsObject Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: member '" & segmentName & "' in binding path '" & bindingPath & "' is not an object."
                Exit Function
            End If
            Set currentObject = memberObject
        Else
            If memberIsObject Then
                Set outValue = memberObject
            Else
                outValue = memberScalar
            End If
        End If

ContinueLoop:
    Next segmentIndex

    private_TryReadBindingPathValue = True
End Function


Private Function private_TryReadMemberValue( _
    ByVal sourceObject As Object, _
    ByVal memberName As String, _
    ByRef outIsObject As Boolean, _
    ByRef outObject As Object, _
    ByRef outScalar As Variant _
) As Boolean
    Dim dictObj As Object
    Dim collObj As Collection
    Dim itemIndex As Long

    outIsObject = False
    Set outObject = Nothing

    Set dictObj = private_AsDictionary(sourceObject)
    If Not dictObj Is Nothing Then
        If Not dictObj.Exists(memberName) Then Exit Function

        On Error Resume Next
        Set outObject = dictObj.Item(memberName)
        If Err.Number = 0 Then
            outIsObject = True
            private_TryReadMemberValue = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear

        outScalar = dictObj.Item(memberName)
        If Err.Number = 0 Then
            private_TryReadMemberValue = True
        End If
        On Error GoTo 0

        Exit Function
    End If

    Set collObj = private_AsCollection(sourceObject)
    If Not collObj Is Nothing Then
        Select Case VBA.LCase$(VBA.Trim$(memberName))
            Case "count"
                outScalar = VBA.CLng(collObj.Count)
                private_TryReadMemberValue = True
                Exit Function
        End Select

        If VBA.IsNumeric(memberName) Then
            itemIndex = VBA.CLng(memberName)
            If itemIndex <= 0 Or itemIndex > collObj.Count Then Exit Function

            On Error Resume Next
            Set outObject = collObj.Item(itemIndex)
            If Err.Number = 0 Then
                outIsObject = True
                private_TryReadMemberValue = True
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear

            outScalar = collObj.Item(itemIndex)
            If Err.Number = 0 Then
                private_TryReadMemberValue = True
            End If
            On Error GoTo 0

            Exit Function
        End If
    End If

    On Error Resume Next
    Set outObject = VBA.CallByName(sourceObject, memberName, VbGet)
    If Err.Number = 0 Then
        outIsObject = True
        private_TryReadMemberValue = True
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    outScalar = VBA.CallByName(sourceObject, memberName, VbGet)
    If Err.Number = 0 Then
        private_TryReadMemberValue = True
    End If
    On Error GoTo 0
End Function


Private Function private_AsDictionary(ByVal sourceObject As Object) As Object
    Dim typeNameText As String

    If sourceObject Is Nothing Then Exit Function

    typeNameText = VBA.TypeName(sourceObject)
    If VBA.StrComp(typeNameText, "Dictionary", VBA.vbTextCompare) = 0 Or _
       VBA.StrComp(typeNameText, "Scripting.Dictionary", VBA.vbTextCompare) = 0 Then
        Set private_AsDictionary = sourceObject
    End If
End Function


Private Function private_AsCollection(ByVal sourceObject As Object) As Collection
    If sourceObject Is Nothing Then Exit Function
    If VBA.StrComp(VBA.TypeName(sourceObject), "Collection", VBA.vbTextCompare) <> 0 Then Exit Function

    Set private_AsCollection = sourceObject
End Function


Private Function private_QualifyMacroName(ByVal macroName As String) As String
    macroName = VBA.Trim$(macroName)
    If VBA.InStr(1, macroName, "!", VBA.vbBinaryCompare) > 0 Then
        private_QualifyMacroName = macroName
    Else
        private_QualifyMacroName = "'" & ThisWorkbook.Name & "'!" & macroName
    End If
End Function


Private Function private_TryEvaluateConditionalOperation( _
    ByVal actualValue As Variant, _
    ByVal opText As String, _
    ByVal expectedText As String, _
    ByRef outResult As Boolean _
) As Boolean
    Dim actualNumber As Double
    Dim expectedNumber As Double
    Dim actualBool As Boolean
    Dim expectedBool As Boolean
    Dim actualText As String

    Select Case opText
        Case "eq", "ne"
            If private_TryParseBooleanVariant(actualValue, actualBool) And private_TryParseBooleanValue(expectedText, expectedBool) Then
                outResult = (actualBool = expectedBool)
            ElseIf private_TryParseNumberVariant(actualValue, actualNumber) And private_TryParseNumberText(expectedText, expectedNumber) Then
                outResult = (actualNumber = expectedNumber)
            Else
                actualText = VBA.LCase$(VBA.Trim$(VBA.CStr(actualValue)))
                outResult = (VBA.StrComp(actualText, VBA.LCase$(VBA.Trim$(expectedText)), VBA.vbBinaryCompare) = 0)
            End If

            If VBA.StrComp(opText, "ne", VBA.vbBinaryCompare) = 0 Then
                outResult = Not outResult
            End If

        Case "gt", "ge", "lt", "le"
            If Not private_TryParseNumberVariant(actualValue, actualNumber) Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional operator '" & opText & "' requires numeric binding value."
                Exit Function
            End If
            If Not private_TryParseNumberText(expectedText, expectedNumber) Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional operator '" & opText & "' requires numeric Value."
                Exit Function
            End If

            Select Case opText
                Case "gt": outResult = (actualNumber > expectedNumber)
                Case "ge": outResult = (actualNumber >= expectedNumber)
                Case "lt": outResult = (actualNumber < expectedNumber)
                Case "le": outResult = (actualNumber <= expectedNumber)
            End Select

        Case "istrue"
            If Not private_TryParseBooleanVariant(actualValue, outResult) Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional operator 'isTrue' requires boolean-compatible value."
                Exit Function
            End If

        Case "isfalse"
            If Not private_TryParseBooleanVariant(actualValue, outResult) Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: conditional operator 'isFalse' requires boolean-compatible value."
                Exit Function
            End If
            outResult = Not outResult

        Case Else
            ex_Core.m_Diagnostic_LogError "PrototypeNew: unsupported conditional Op '" & opText & "'."
            Exit Function
    End Select

    private_TryEvaluateConditionalOperation = True
End Function


Private Function private_TryParseNumberVariant(ByVal rawValue As Variant, ByRef outNumber As Double) As Boolean
    If VBA.IsObject(rawValue) Then Exit Function
    If Not VBA.IsNumeric(rawValue) Then Exit Function

    outNumber = VBA.CDbl(rawValue)
    private_TryParseNumberVariant = True
End Function


Private Function private_TryParseNumberText(ByVal rawText As String, ByRef outNumber As Double) As Boolean
    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then Exit Function
    If Not VBA.IsNumeric(rawText) Then Exit Function

    outNumber = VBA.CDbl(rawText)
    private_TryParseNumberText = True
End Function


Private Function private_TryParseBooleanVariant(ByVal rawValue As Variant, ByRef outBoolean As Boolean) As Boolean
    Dim typeCode As VbVarType

    If VBA.IsObject(rawValue) Then Exit Function

    typeCode = VBA.VarType(rawValue)
    If typeCode = vbBoolean Then
        outBoolean = VBA.CBool(rawValue)
        private_TryParseBooleanVariant = True
        Exit Function
    End If

    If VBA.IsNumeric(rawValue) Then
        outBoolean = (VBA.CDbl(rawValue) <> 0#)
        private_TryParseBooleanVariant = True
        Exit Function
    End If

    private_TryParseBooleanVariant = private_TryParseBooleanValue(VBA.CStr(rawValue), outBoolean)
End Function


Private Function private_TryParseBooleanValue(ByVal rawText As String, ByRef outBoolean As Boolean) As Boolean
    rawText = VBA.LCase$(VBA.Trim$(rawText))

    Select Case rawText
        Case "1", "true", "yes", "y", "on", "visible", "show", "shown"
            outBoolean = True
            private_TryParseBooleanValue = True

        Case "0", "false", "no", "n", "off", VBA.vbNullString, "collapsed", "collapse", "hidden", "hide", "none"
            outBoolean = False
            private_TryParseBooleanValue = True
    End Select
End Function
