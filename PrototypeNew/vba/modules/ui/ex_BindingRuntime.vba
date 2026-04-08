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

    If mp_TryExtractNamedArg(bindingBody, "Path", bindingPath) Then
    Else
        bindingPath = Trim$(bindingBody)
    End If

    If Len(bindingPath) = 0 Then bindingPath = "."
    If Not mp_TryReadBindingPathValue(sourceObject, bindingPath, outValue) Then Exit Function

    mp_TryResolveBindingValue = True
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

Private Function mp_QualifyMacroName(ByVal macroName As String) As String
    macroName = Trim$(macroName)
    If InStr(1, macroName, "!", vbBinaryCompare) > 0 Then
        mp_QualifyMacroName = macroName
    Else
        mp_QualifyMacroName = "'" & ThisWorkbook.Name & "'!" & macroName
    End If
End Function
