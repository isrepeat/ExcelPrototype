Attribute VB_Name = "ex_XmlCore"
Option Explicit

Public Function m_CombineBasePath(ByVal wb As Workbook, ByVal relPath As String) As String
    Dim basePath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    basePath = wb.Path
    If Len(basePath) = 0 Then basePath = CurDir$

    m_CombineBasePath = basePath & "\" & relPath
End Function

Public Function m_CreateDom(Optional ByVal nsUri As String = vbNullString) As Object
    Dim doc As Object

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If Len(Trim$(nsUri)) > 0 Then
        doc.setProperty "SelectionNamespaces", "xmlns:p='" & nsUri & "'"
    End If

    Set m_CreateDom = doc
End Function

Public Function m_LoadDomByRelativePath(ByVal wb As Workbook, ByVal relPath As String, ByVal nsUri As String, ByVal missingPrefix As String, ByVal parsePrefix As String) As Object
    Dim filePath As String

    filePath = m_CombineBasePath(wb, relPath)
    If Len(filePath) = 0 Then Exit Function

    Set m_LoadDomByRelativePath = m_LoadDomByFilePath(filePath, nsUri, missingPrefix, parsePrefix)
End Function

Public Function m_LoadDomByFilePath(ByVal filePath As String, ByVal nsUri As String, ByVal missingPrefix As String, ByVal parsePrefix As String) As Object
    Dim doc As Object

    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then Exit Function

    If Len(Dir(filePath)) = 0 Then
        MsgBox missingPrefix & filePath, vbExclamation
        Exit Function
    End If

    Set doc = m_CreateDom(nsUri)
    If Not doc.Load(filePath) Then
        MsgBox parsePrefix & filePath, vbExclamation
        Exit Function
    End If

    Set m_LoadDomByFilePath = doc
End Function

Public Function m_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    m_NodeAttrText = CStr(node.Attributes.getNamedItem(attrName).Text)
    If Err.Number <> 0 Then
        Err.Clear
        m_NodeAttrText = vbNullString
    End If
    On Error GoTo 0

    If Len(m_NodeAttrText) > 0 Then Exit Function

    On Error Resume Next
    m_NodeAttrText = CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        m_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

Public Function m_XPathLiteral(ByVal value As String) As String
    Dim parts() As String
    Dim i As Long

    If InStr(value, "'") = 0 Then
        m_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If InStr(value, """") = 0 Then
        m_XPathLiteral = """" & value & """"
        Exit Function
    End If

    parts = Split(value, "'")
    m_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            m_XPathLiteral = m_XPathLiteral & ", ""'"" , "
        End If
        m_XPathLiteral = m_XPathLiteral & "'" & parts(i) & "'"
    Next i
    m_XPathLiteral = m_XPathLiteral & ")"
End Function

Public Function m_TryParseBoolean(ByVal valueText As String, ByRef result As Boolean) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "1", "true", "yes"
            result = True
            m_TryParseBoolean = True
        Case "0", "false", "no"
            result = False
            m_TryParseBoolean = True
    End Select
End Function

Public Function m_TryParseDouble(ByVal valueText As String, ByRef result As Double, Optional ByVal localeAware As Boolean = False) As Boolean
    Dim normalized As String
    Dim decSep As String
    Dim altSep As String

    On Error GoTo EH

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    If localeAware Then
        decSep = CStr(Application.International(xlDecimalSeparator))
        If decSep = "." Then
            altSep = ","
        Else
            altSep = "."
        End If
        normalized = Replace(normalized, altSep, decSep)
    End If

    If Not IsNumeric(normalized) Then Exit Function

    result = CDbl(normalized)
    m_TryParseDouble = True
    Exit Function
EH:
    m_TryParseDouble = False
End Function

Public Function m_TryParseLong(ByVal valueText As String, ByRef result As Long) As Boolean
    If Not IsNumeric(valueText) Then Exit Function
    result = CLng(valueText)
    m_TryParseLong = True
End Function

Public Function m_TryParseHexColor(ByVal textValue As String, ByRef outValue As Long) As Boolean
    Dim r As Long
    Dim g As Long
    Dim b As Long

    textValue = Trim$(textValue)
    If Len(textValue) <> 7 Then Exit Function
    If Left$(textValue, 1) <> "#" Then Exit Function
    If Not mp_IsHexPair(Mid$(textValue, 2, 2)) Then Exit Function
    If Not mp_IsHexPair(Mid$(textValue, 4, 2)) Then Exit Function
    If Not mp_IsHexPair(Mid$(textValue, 6, 2)) Then Exit Function

    r = CLng("&H" & Mid$(textValue, 2, 2))
    g = CLng("&H" & Mid$(textValue, 4, 2))
    b = CLng("&H" & Mid$(textValue, 6, 2))
    outValue = RGB(r, g, b)
    m_TryParseHexColor = True
End Function

Public Function m_TryParseColor(ByVal valueText As String, ByRef colorValue As Long) As Boolean
    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    If Left$(valueText, 1) = "#" Then
        m_TryParseColor = m_TryParseHexColor(valueText, colorValue)
        Exit Function
    End If

    If IsNumeric(valueText) Then
        colorValue = CLng(valueText)
        m_TryParseColor = True
    End If
End Function

Public Function m_ReadRequiredAttrText(ByVal node As Object, ByVal attrName As String, ByVal fieldName As String, ByVal entityLabel As String) As String
    If node Is Nothing Then
        MsgBox "Missing required node for " & entityLabel & ": " & fieldName, vbExclamation
        Exit Function
    End If

    m_ReadRequiredAttrText = Trim$(m_NodeAttrText(node, attrName))
    If Len(m_ReadRequiredAttrText) = 0 Then
        MsgBox "Missing required " & entityLabel & " attribute: " & fieldName, vbExclamation
    End If
End Function

Public Function m_ReadRequiredAttrDouble(ByVal node As Object, ByVal attrName As String, ByRef outValue As Double, ByVal fieldName As String, ByVal entityLabel As String) As Boolean
    Dim textValue As String

    textValue = m_ReadRequiredAttrText(node, attrName, fieldName, entityLabel)
    If Len(textValue) = 0 Then Exit Function

    If Not m_TryParseDouble(textValue, outValue) Then
        MsgBox "Invalid numeric " & entityLabel & " attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    m_ReadRequiredAttrDouble = True
End Function

Public Function m_ReadRequiredAttrLong(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long, ByVal fieldName As String, ByVal entityLabel As String) As Boolean
    Dim textValue As String

    textValue = m_ReadRequiredAttrText(node, attrName, fieldName, entityLabel)
    If Len(textValue) = 0 Then Exit Function

    If Not m_TryParseLong(textValue, outValue) Then
        MsgBox "Invalid integer " & entityLabel & " attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    m_ReadRequiredAttrLong = True
End Function

Public Function m_ReadRequiredAttrBoolean(ByVal node As Object, ByVal attrName As String, ByRef outValue As Boolean, ByVal fieldName As String, ByVal entityLabel As String) As Boolean
    Dim textValue As String

    textValue = LCase$(m_ReadRequiredAttrText(node, attrName, fieldName, entityLabel))
    If Len(textValue) = 0 Then Exit Function

    If Not m_TryParseBoolean(textValue, outValue) Then
        MsgBox "Invalid boolean " & entityLabel & " attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    m_ReadRequiredAttrBoolean = True
End Function

Public Function m_ReadRequiredAttrHexColor(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long, ByVal fieldName As String, ByVal entityLabel As String) As Boolean
    Dim textValue As String

    textValue = m_ReadRequiredAttrText(node, attrName, fieldName, entityLabel)
    If Len(textValue) = 0 Then Exit Function

    If Not m_TryParseHexColor(textValue, outValue) Then
        MsgBox "Invalid color " & entityLabel & " attribute '" & fieldName & "': expected #RRGGBB, got " & textValue, vbExclamation
        Exit Function
    End If

    m_ReadRequiredAttrHexColor = True
End Function

Private Function mp_IsHexPair(ByVal value As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(value) <> 2 Then Exit Function

    For i = 1 To 2
        ch = Mid$(value, i, 1)
        If InStr(1, "0123456789ABCDEFabcdef", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i

    mp_IsHexPair = True
End Function
