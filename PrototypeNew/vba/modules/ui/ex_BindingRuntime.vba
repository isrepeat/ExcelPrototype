Attribute VB_Name = "ex_BindingRuntime"
Option Explicit

Private Const BINDING_PREFIX As String = "{Binding "
Private Const BINDING_SUFFIX As String = "}"

Public Function m_TryResolveTextBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outText As String _
) As Boolean
    Dim resolvedValue As Variant

    If Not mp_TryResolveBindingValue(rawText, sourceObject, resolvedValue) Then Exit Function

    If IsObject(resolvedValue) Then
        MsgBox "PrototypeNew: text binding must resolve to scalar value.", vbExclamation
        Exit Function
    End If

    outText = CStr(resolvedValue)
    m_TryResolveTextBinding = True
End Function

Public Function m_TryResolveMacroBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outMacroRef As String _
) As Boolean
    Dim resolvedValue As Variant
    Dim macroName As String

    If Not mp_TryResolveBindingValue(rawText, sourceObject, resolvedValue) Then Exit Function

    If IsObject(resolvedValue) Then
        MsgBox "PrototypeNew: macro binding must resolve to text value.", vbExclamation
        Exit Function
    End If

    macroName = Trim$(CStr(resolvedValue))
    If Len(macroName) = 0 Then
        MsgBox "PrototypeNew: macro binding resolved to an empty value.", vbExclamation
        Exit Function
    End If

    outMacroRef = mp_QualifyMacroName(macroName)
    m_TryResolveMacroBinding = True
End Function

' // Helper for Visibility attribute.
Public Function m_TryResolveVisibilityBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outVisible As Boolean _
) As Boolean
    Dim resolvedValue As Variant

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then
        outVisible = True
        m_TryResolveVisibilityBinding = True
        Exit Function
    End If

    If Not m_TryResolveValueBinding(rawText, sourceObject, resolvedValue) Then Exit Function
    If Not mp_TryParseBooleanVariant(resolvedValue, outVisible) Then
        If IsObject(resolvedValue) Then
            MsgBox "PrototypeNew: visibility value resolved to object '" & TypeName(resolvedValue) & "'. Expected boolean-compatible value.", vbExclamation
        Else
            MsgBox "PrototypeNew: visibility value '" & CStr(resolvedValue) & "' is invalid. Supported values: true/false/visible/collapsed.", vbExclamation
        End If
        Exit Function
    End If

    m_TryResolveVisibilityBinding = True
End Function

Public Function m_TryResolveValueBinding( _
    ByVal rawText As String, _
    ByVal sourceObject As Object, _
    ByRef outValue As Variant _
) As Boolean
    If Not mp_TryResolveBindingValue(rawText, sourceObject, outValue) Then Exit Function
    m_TryResolveValueBinding = True
End Function

Private Function mp_TryResolveBindingValue( _
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

    If Not mp_TryExtractBindingBody(rawText, bindingBody) Then
        outValue = rawText
        mp_TryResolveBindingValue = True
        Exit Function
    End If

    If mp_TryExtractNamedArg(bindingBody, "Method", methodName) Then
        If mp_TryExtractNamedArg(bindingBody, "Module", moduleName) Then
            If InStr(1, methodName, ".", vbBinaryCompare) = 0 Then
                methodName = moduleName & "." & methodName
            End If
        End If

        methodName = Trim$(methodName)
        If Len(methodName) = 0 Then
            MsgBox "PrototypeNew: binding expression contains empty Method value.", vbExclamation
            Exit Function
        End If

        outValue = methodName
        mp_TryResolveBindingValue = True
        Exit Function
    End If

    If Not mp_TryResolveConditionalBindingAsText(bindingBody, sourceObject, mappedText, usedConditionalBranch) Then Exit Function
    If usedConditionalBranch Then
        outValue = mappedText
        mp_TryResolveBindingValue = True
        Exit Function
    End If

    If mp_TryExtractNamedArg(bindingBody, "Path", bindingPath) Then
    Else
        bindingPath = Trim$(bindingBody)
    End If

    If Len(bindingPath) = 0 Then bindingPath = "."
    If Not mp_TryReadBindingPathValue(sourceObject, bindingPath, outValue) Then Exit Function

    mp_TryResolveBindingValue = True
End Function

Private Function mp_TryResolveConditionalBindingAsText( _
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

    hasOp = mp_TryExtractNamedArg(bindingBody, "Op", opText)
    hasValue = mp_TryExtractNamedArg(bindingBody, "Value", valueText)
    hasTrueAs = mp_TryExtractNamedArg(bindingBody, "TrueAs", trueAsText)
    hasFalseAs = mp_TryExtractNamedArg(bindingBody, "FalseAs", falseAsText)

    If Not hasOp And Not hasValue And Not hasTrueAs And Not hasFalseAs Then
        mp_TryResolveConditionalBindingAsText = True
        Exit Function
    End If

    outUsedConditionalBranch = True

    If sourceObject Is Nothing Then
        MsgBox "PrototypeNew: conditional binding requires source object.", vbExclamation
        Exit Function
    End If

    If mp_TryExtractNamedArg(bindingBody, "Path", bindingPath) Then
    Else
        bindingPath = Trim$(bindingBody)
    End If

    If Len(bindingPath) = 0 Then bindingPath = "."
    If Not mp_TryReadBindingPathValue(sourceObject, bindingPath, resolvedValue) Then Exit Function

    If hasValue And Not hasOp Then
        MsgBox "PrototypeNew: conditional binding argument 'Value' requires 'Op'.", vbExclamation
        Exit Function
    End If

    If hasOp Then
        opText = LCase$(Trim$(opText))
        If Not mp_TryEvaluateConditionalOperation(resolvedValue, opText, valueText, conditionResult) Then Exit Function
    Else
        If Not mp_TryParseBooleanVariant(resolvedValue, conditionResult) Then
            MsgBox "PrototypeNew: conditional binding path '" & bindingPath & "' must resolve to boolean-compatible value when Op is omitted.", vbExclamation
            Exit Function
        End If
    End If

    If Not mp_TryMapConditionalResultToText( _
        conditionResult, _
        hasTrueAs, trueAsText, _
        hasFalseAs, falseAsText, _
        outText) Then Exit Function

    mp_TryResolveConditionalBindingAsText = True
End Function

Private Function mp_TryMapConditionalResultToText( _
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

    mp_TryMapConditionalResultToText = True
End Function

Private Function mp_TryExtractBindingBody(ByVal rawText As String, ByRef outBody As String) As Boolean
    Dim normalized As String
    Dim prefixLen As Long

    normalized = Trim$(rawText)
    prefixLen = Len(BINDING_PREFIX)

    If Len(normalized) < prefixLen + 1 Then Exit Function
    If StrComp(Left$(normalized, prefixLen), BINDING_PREFIX, vbTextCompare) <> 0 Then Exit Function
    If Right$(normalized, Len(BINDING_SUFFIX)) <> BINDING_SUFFIX Then Exit Function

    outBody = Trim$(Mid$(normalized, prefixLen + 1, Len(normalized) - prefixLen - 1))
    mp_TryExtractBindingBody = True
End Function

Private Function mp_TryExtractNamedArg( _
    ByVal bindingBody As String, _
    ByVal argName As String, _
    ByRef outValue As String _
) As Boolean
    Dim argPos As Long
    Dim valueStart As Long
    Dim valueText As String
    Dim sepPos As Long

    argPos = InStr(1, bindingBody, argName & "=", vbTextCompare)
    If argPos = 0 Then Exit Function

    valueStart = argPos + Len(argName) + 1
    valueText = Mid$(bindingBody, valueStart)

    sepPos = InStr(1, valueText, ";", vbBinaryCompare)
    If sepPos = 0 Then sepPos = InStr(1, valueText, ",", vbBinaryCompare)

    If sepPos > 0 Then
        outValue = Trim$(Left$(valueText, sepPos - 1))
    Else
        outValue = Trim$(valueText)
    End If

    outValue = mp_Unquote(outValue)
    mp_TryExtractNamedArg = (Len(outValue) > 0)
End Function

Private Function mp_Unquote(ByVal textValue As String) As String
    textValue = Trim$(textValue)
    If Len(textValue) < 2 Then
        mp_Unquote = textValue
        Exit Function
    End If

    If (Left$(textValue, 1) = Chr$(34) And Right$(textValue, 1) = Chr$(34)) Or _
       (Left$(textValue, 1) = "'" And Right$(textValue, 1) = "'") Then
        mp_Unquote = Mid$(textValue, 2, Len(textValue) - 2)
    Else
        mp_Unquote = textValue
    End If
End Function

Private Function mp_TryReadBindingPathValue( _
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
        MsgBox "PrototypeNew: binding source object is not specified.", vbExclamation
        Exit Function
    End If

    bindingPath = Trim$(bindingPath)
    If Len(bindingPath) = 0 Or bindingPath = "." Then
        Set outValue = sourceObject
        mp_TryReadBindingPathValue = True
        Exit Function
    End If

    Set currentObject = sourceObject
    segments = Split(bindingPath, ".")

    For segmentIndex = LBound(segments) To UBound(segments)
        segmentName = Trim$(CStr(segments(segmentIndex)))
        If Len(segmentName) = 0 Then GoTo ContinueLoop

        If currentObject Is Nothing Then
            MsgBox "PrototypeNew: binding path '" & bindingPath & "' reached Nothing before segment '" & segmentName & "'.", vbExclamation
            Exit Function
        End If

        If Not mp_TryReadMemberValue(currentObject, segmentName, memberIsObject, memberObject, memberScalar) Then
            MsgBox "PrototypeNew: member '" & segmentName & "' was not found on object '" & TypeName(currentObject) & "'.", vbExclamation
            Exit Function
        End If

        If segmentIndex < UBound(segments) Then
            If Not memberIsObject Then
                MsgBox "PrototypeNew: member '" & segmentName & "' in binding path '" & bindingPath & "' is not an object.", vbExclamation
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

    mp_TryReadBindingPathValue = True
End Function

Private Function mp_TryReadMemberValue( _
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

    Set dictObj = mp_AsDictionary(sourceObject)
    If Not dictObj Is Nothing Then
        If Not dictObj.Exists(memberName) Then Exit Function

        On Error Resume Next
        Set outObject = dictObj.Item(memberName)
        If Err.Number = 0 Then
            outIsObject = True
            mp_TryReadMemberValue = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear

        outScalar = dictObj.Item(memberName)
        If Err.Number = 0 Then
            mp_TryReadMemberValue = True
        End If
        On Error GoTo 0

        Exit Function
    End If

    Set collObj = mp_AsCollection(sourceObject)
    If Not collObj Is Nothing Then
        Select Case LCase$(Trim$(memberName))
            Case "count"
                outScalar = CLng(collObj.Count)
                mp_TryReadMemberValue = True
                Exit Function
        End Select

        If IsNumeric(memberName) Then
            itemIndex = CLng(memberName)
            If itemIndex <= 0 Or itemIndex > collObj.Count Then Exit Function

            On Error Resume Next
            Set outObject = collObj.Item(itemIndex)
            If Err.Number = 0 Then
                outIsObject = True
                mp_TryReadMemberValue = True
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear

            outScalar = collObj.Item(itemIndex)
            If Err.Number = 0 Then
                mp_TryReadMemberValue = True
            End If
            On Error GoTo 0

            Exit Function
        End If
    End If

    On Error Resume Next
    Set outObject = CallByName(sourceObject, memberName, VbGet)
    If Err.Number = 0 Then
        outIsObject = True
        mp_TryReadMemberValue = True
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    outScalar = CallByName(sourceObject, memberName, VbGet)
    If Err.Number = 0 Then
        mp_TryReadMemberValue = True
    End If
    On Error GoTo 0
End Function

Private Function mp_AsDictionary(ByVal sourceObject As Object) As Object
    Dim typeNameText As String

    If sourceObject Is Nothing Then Exit Function

    typeNameText = TypeName(sourceObject)
    If StrComp(typeNameText, "Dictionary", vbTextCompare) = 0 Or _
       StrComp(typeNameText, "Scripting.Dictionary", vbTextCompare) = 0 Then
        Set mp_AsDictionary = sourceObject
    End If
End Function

Private Function mp_AsCollection(ByVal sourceObject As Object) As Collection
    If sourceObject Is Nothing Then Exit Function
    If StrComp(TypeName(sourceObject), "Collection", vbTextCompare) <> 0 Then Exit Function

    Set mp_AsCollection = sourceObject
End Function

Private Function mp_QualifyMacroName(ByVal macroName As String) As String
    macroName = Trim$(macroName)
    If InStr(1, macroName, "!", vbBinaryCompare) > 0 Then
        mp_QualifyMacroName = macroName
    Else
        mp_QualifyMacroName = "'" & ThisWorkbook.Name & "'!" & macroName
    End If
End Function

Private Function mp_TryEvaluateConditionalOperation( _
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
            If mp_TryParseBooleanVariant(actualValue, actualBool) And mp_TryParseBooleanValue(expectedText, expectedBool) Then
                outResult = (actualBool = expectedBool)
            ElseIf mp_TryParseNumberVariant(actualValue, actualNumber) And mp_TryParseNumberText(expectedText, expectedNumber) Then
                outResult = (actualNumber = expectedNumber)
            Else
                actualText = LCase$(Trim$(CStr(actualValue)))
                outResult = (StrComp(actualText, LCase$(Trim$(expectedText)), vbBinaryCompare) = 0)
            End If

            If StrComp(opText, "ne", vbBinaryCompare) = 0 Then
                outResult = Not outResult
            End If

        Case "gt", "ge", "lt", "le"
            If Not mp_TryParseNumberVariant(actualValue, actualNumber) Then
                MsgBox "PrototypeNew: conditional operator '" & opText & "' requires numeric binding value.", vbExclamation
                Exit Function
            End If
            If Not mp_TryParseNumberText(expectedText, expectedNumber) Then
                MsgBox "PrototypeNew: conditional operator '" & opText & "' requires numeric Value.", vbExclamation
                Exit Function
            End If

            Select Case opText
                Case "gt": outResult = (actualNumber > expectedNumber)
                Case "ge": outResult = (actualNumber >= expectedNumber)
                Case "lt": outResult = (actualNumber < expectedNumber)
                Case "le": outResult = (actualNumber <= expectedNumber)
            End Select

        Case "istrue"
            If Not mp_TryParseBooleanVariant(actualValue, outResult) Then
                MsgBox "PrototypeNew: conditional operator 'isTrue' requires boolean-compatible value.", vbExclamation
                Exit Function
            End If

        Case "isfalse"
            If Not mp_TryParseBooleanVariant(actualValue, outResult) Then
                MsgBox "PrototypeNew: conditional operator 'isFalse' requires boolean-compatible value.", vbExclamation
                Exit Function
            End If
            outResult = Not outResult

        Case Else
            MsgBox "PrototypeNew: unsupported conditional Op '" & opText & "'.", vbExclamation
            Exit Function
    End Select

    mp_TryEvaluateConditionalOperation = True
End Function

Private Function mp_TryParseNumberVariant(ByVal rawValue As Variant, ByRef outNumber As Double) As Boolean
    If IsObject(rawValue) Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    outNumber = CDbl(rawValue)
    mp_TryParseNumberVariant = True
End Function

Private Function mp_TryParseNumberText(ByVal rawText As String, ByRef outNumber As Double) As Boolean
    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then Exit Function
    If Not IsNumeric(rawText) Then Exit Function

    outNumber = CDbl(rawText)
    mp_TryParseNumberText = True
End Function

Private Function mp_TryParseBooleanVariant(ByVal rawValue As Variant, ByRef outBoolean As Boolean) As Boolean
    Dim typeCode As VbVarType

    If IsObject(rawValue) Then Exit Function

    typeCode = VarType(rawValue)
    If typeCode = vbBoolean Then
        outBoolean = CBool(rawValue)
        mp_TryParseBooleanVariant = True
        Exit Function
    End If

    If IsNumeric(rawValue) Then
        outBoolean = (CDbl(rawValue) <> 0#)
        mp_TryParseBooleanVariant = True
        Exit Function
    End If

    mp_TryParseBooleanVariant = mp_TryParseBooleanValue(CStr(rawValue), outBoolean)
End Function

Private Function mp_TryParseBooleanValue(ByVal rawText As String, ByRef outBoolean As Boolean) As Boolean
    rawText = LCase$(Trim$(rawText))

    Select Case rawText
        Case "1", "true", "yes", "y", "on", "visible", "show", "shown"
            outBoolean = True
            mp_TryParseBooleanValue = True

        Case "0", "false", "no", "n", "off", vbNullString, "collapsed", "collapse", "hidden", "hide", "none"
            outBoolean = False
            mp_TryParseBooleanValue = True
    End Select
End Function
