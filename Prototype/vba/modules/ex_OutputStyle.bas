Attribute VB_Name = "ex_OutputStyle"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const OUTPUT_STYLE_REL_PATH As String = "config\OutputStyle.xml"

Public Type t_OutputStyle
    FontName As String
    FontSize As Double
    RowHeight As Double

    ContentColor As Long
    HeaderColor As Long
    HeaderBold As Boolean
    SectionColor As Long
    SectionBold As Boolean
    SectionMergeColumns As Long

    HorizontalAlignment As Long
    VerticalAlignment As Long
End Type

Public Function m_LoadOutputStyle(ByRef style As t_OutputStyle, Optional ByVal wb As Workbook) As Boolean
    Dim doc As Object
    Dim rootNode As Object
    Dim nodeFont As Object
    Dim nodeRows As Object
    Dim nodeContent As Object
    Dim nodeHeader As Object
    Dim nodeSection As Object
    Dim nodeAlignment As Object
    Dim filePath As String

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then
        MsgBox "Failed to load output style: workbook is not specified.", vbExclamation
        Exit Function
    End If

    filePath = mp_GetOutputStyleFilePath(wb)
    If Len(Dir(filePath)) = 0 Then
        MsgBox "Output style config file was not found: " & filePath, vbExclamation
        Exit Function
    End If

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    If Not doc.Load(filePath) Then
        MsgBox "Failed to parse output style config file: " & filePath, vbExclamation
        Exit Function
    End If
    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"

    Set rootNode = doc.selectSingleNode("/p:outputStyle")
    If rootNode Is Nothing Then
        MsgBox "Invalid output style config format. Expected root '/outputStyle'.", vbExclamation
        Exit Function
    End If

    Set nodeFont = rootNode.selectSingleNode("p:font")
    Set nodeRows = rootNode.selectSingleNode("p:rows")
    Set nodeContent = rootNode.selectSingleNode("p:content")
    Set nodeHeader = rootNode.selectSingleNode("p:header")
    Set nodeSection = rootNode.selectSingleNode("p:section")
    Set nodeAlignment = rootNode.selectSingleNode("p:alignment")

    If nodeFont Is Nothing Or nodeRows Is Nothing Or nodeContent Is Nothing Or _
       nodeHeader Is Nothing Or nodeSection Is Nothing Or nodeAlignment Is Nothing Then
        MsgBox "Output style config must contain nodes: font, rows, content, header, section, alignment.", vbExclamation
        Exit Function
    End If

    style.FontName = mp_ReadRequiredAttrText(nodeFont, "name", "font@name")
    If Len(style.FontName) = 0 Then Exit Function
    If Not mp_ReadRequiredAttrDouble(nodeFont, "size", style.FontSize, "font@size") Then Exit Function
    If Not mp_ReadRequiredAttrDouble(nodeRows, "height", style.RowHeight, "rows@height") Then Exit Function

    If Not mp_ReadRequiredAttrColor(nodeContent, "color", style.ContentColor, "content@color") Then Exit Function
    If Not mp_ReadRequiredAttrColor(nodeHeader, "color", style.HeaderColor, "header@color") Then Exit Function
    If Not mp_ReadRequiredAttrBoolean(nodeHeader, "bold", style.HeaderBold, "header@bold") Then Exit Function
    If Not mp_ReadRequiredAttrColor(nodeSection, "color", style.SectionColor, "section@color") Then Exit Function
    If Not mp_ReadRequiredAttrBoolean(nodeSection, "bold", style.SectionBold, "section@bold") Then Exit Function
    If Not mp_ReadRequiredAttrLong(nodeSection, "mergeColumns", style.SectionMergeColumns, "section@mergeColumns") Then Exit Function
    If style.SectionMergeColumns < 1 Then
        MsgBox "Invalid value for output style 'section@mergeColumns': must be >= 1.", vbExclamation
        Exit Function
    End If

    If Not mp_ReadRequiredAttrHorizontalAlignment(nodeAlignment, "horizontal", style.HorizontalAlignment) Then Exit Function
    If Not mp_ReadRequiredAttrVerticalAlignment(nodeAlignment, "vertical", style.VerticalAlignment) Then Exit Function

    m_LoadOutputStyle = True
End Function

Private Function mp_GetOutputStyleFilePath(ByVal wb As Workbook) As String
    Dim basePath As String

    basePath = wb.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    mp_GetOutputStyleFilePath = basePath & "\" & OUTPUT_STYLE_REL_PATH
End Function

Private Function mp_ReadRequiredAttrText(ByVal node As Object, ByVal attrName As String, ByVal fieldName As String) As String
    mp_ReadRequiredAttrText = Trim$(CStr(node.getAttribute(attrName)))
    If Len(mp_ReadRequiredAttrText) = 0 Then
        MsgBox "Missing required output style attribute: " & fieldName, vbExclamation
    End If
End Function

Private Function mp_ReadRequiredAttrDouble(ByVal node As Object, ByVal attrName As String, ByRef outValue As Double, ByVal fieldName As String) As Boolean
    Dim textValue As String
    textValue = mp_ReadRequiredAttrText(node, attrName, fieldName)
    If Len(textValue) = 0 Then Exit Function

    If Not IsNumeric(textValue) Then
        MsgBox "Invalid numeric output style attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    outValue = CDbl(textValue)
    mp_ReadRequiredAttrDouble = True
End Function

Private Function mp_ReadRequiredAttrLong(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long, ByVal fieldName As String) As Boolean
    Dim textValue As String
    textValue = mp_ReadRequiredAttrText(node, attrName, fieldName)
    If Len(textValue) = 0 Then Exit Function

    If Not IsNumeric(textValue) Then
        MsgBox "Invalid integer output style attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    outValue = CLng(textValue)
    mp_ReadRequiredAttrLong = True
End Function

Private Function mp_ReadRequiredAttrBoolean(ByVal node As Object, ByVal attrName As String, ByRef outValue As Boolean, ByVal fieldName As String) As Boolean
    Dim textValue As String
    textValue = LCase$(mp_ReadRequiredAttrText(node, attrName, fieldName))
    If Len(textValue) = 0 Then Exit Function

    Select Case textValue
        Case "true", "1", "yes"
            outValue = True
        Case "false", "0", "no"
            outValue = False
        Case Else
            MsgBox "Invalid boolean output style attribute '" & fieldName & "': " & textValue, vbExclamation
            Exit Function
    End Select

    mp_ReadRequiredAttrBoolean = True
End Function

Private Function mp_ReadRequiredAttrColor(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long, ByVal fieldName As String) As Boolean
    Dim textValue As String
    textValue = mp_ReadRequiredAttrText(node, attrName, fieldName)
    If Len(textValue) = 0 Then Exit Function

    If Not mp_TryParseHexColor(textValue, outValue) Then
        MsgBox "Invalid color output style attribute '" & fieldName & "': expected #RRGGBB, got " & textValue, vbExclamation
        Exit Function
    End If

    mp_ReadRequiredAttrColor = True
End Function

Private Function mp_TryParseHexColor(ByVal textValue As String, ByRef outValue As Long) As Boolean
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
    mp_TryParseHexColor = True
End Function

Private Function mp_IsHexPair(ByVal value As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(value) <> 2 Then Exit Function

    For i = 1 To 2
        ch = Mid$(value, i, 1)
        If InStr(1, "0123456789ABCDEFabcdef", ch, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i

    mp_IsHexPair = True
End Function

Private Function mp_ReadRequiredAttrHorizontalAlignment(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long) As Boolean
    Dim textValue As String

    textValue = LCase$(mp_ReadRequiredAttrText(node, attrName, "alignment@" & attrName))
    If Len(textValue) = 0 Then Exit Function

    Select Case textValue
        Case "center"
            outValue = xlCenter
        Case "left"
            outValue = xlLeft
        Case "right"
            outValue = xlRight
        Case Else
            MsgBox "Invalid output alignment value for '" & attrName & "': " & textValue & ". Allowed: left, center, right.", vbExclamation
            Exit Function
    End Select

    mp_ReadRequiredAttrHorizontalAlignment = True
End Function

Private Function mp_ReadRequiredAttrVerticalAlignment(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long) As Boolean
    Dim textValue As String

    textValue = LCase$(mp_ReadRequiredAttrText(node, attrName, "alignment@" & attrName))
    If Len(textValue) = 0 Then Exit Function

    Select Case textValue
        Case "center"
            outValue = xlCenter
        Case "top"
            outValue = xlTop
        Case "bottom"
            outValue = xlBottom
        Case Else
            MsgBox "Invalid output alignment value for '" & attrName & "': " & textValue & ". Allowed: top, center, bottom.", vbExclamation
            Exit Function
    End Select

    mp_ReadRequiredAttrVerticalAlignment = True
End Function
