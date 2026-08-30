VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WordResultStylesParser"
Option Explicit

' Общий parser DSL-атрибутов WORD. Он отвечает за нормализацию и проверку
' значений, но намеренно не знает о mode-specific экспортёрах и Word COM API.
' Экспортёр получает отсюда типизированные данные и только применяет их.

Private m_FontName As String
Private m_FontSize As Single
Private m_ParagraphSpacingBefore As Single
Private m_ParagraphSpacingAfter As Single
Private m_LineSpacing As String
Private m_ShowTableGridlines As Boolean
Private m_PaperSize As String
Private m_Orientation As String
Private m_MarginTopPoints As Single
Private m_MarginRightPoints As Single
Private m_MarginBottomPoints As Single
Private m_MarginLeftPoints As Single

Public Function Initialize( _
    ByVal documentStylesMarkup As String, _
    ByVal sectionStylesMarkup As String _
) As Boolean
    Dim documentNode As Object
    Dim sectionNode As Object
    If Not private_TryLoadRoot( _
        documentStylesMarkup, "documentStyles", documentNode) Then Exit Function
    If Not private_TryLoadRoot( _
        sectionStylesMarkup, "sectionStyles", sectionNode) Then Exit Function

    m_FontName = HtmlAttributeText(documentNode, "fontName")
    If VBA.Len(VBA.Trim$(m_FontName)) = 0 Then GoTo InvalidDocumentStyles
    If Not private_TryParsePositiveSingle( _
        HtmlAttributeText(documentNode, "fontSize"), m_FontSize) Then _
        GoTo InvalidDocumentStyles
    If Not private_TryParseNonNegativeSingle( _
        HtmlAttributeText(documentNode, "paragraphSpacingBefore"), _
        m_ParagraphSpacingBefore) Then GoTo InvalidDocumentStyles
    If Not private_TryParseNonNegativeSingle( _
        HtmlAttributeText(documentNode, "paragraphSpacingAfter"), _
        m_ParagraphSpacingAfter) Then GoTo InvalidDocumentStyles
    m_LineSpacing = VBA.LCase$(VBA.Trim$( _
        HtmlAttributeText(documentNode, "lineSpacing")))
    If m_LineSpacing <> "single" Then GoTo InvalidDocumentStyles
    If Not TryParseBoolean( _
        HtmlAttributeText(documentNode, "showTableGridlines"), _
        m_ShowTableGridlines) Then GoTo InvalidDocumentStyles

    m_PaperSize = VBA.UCase$(VBA.Trim$( _
        HtmlAttributeText(sectionNode, "paperSize")))
    If m_PaperSize <> "A4" Then GoTo InvalidSectionStyles
    m_Orientation = VBA.LCase$(VBA.Trim$( _
        HtmlAttributeText(sectionNode, "orientation")))
    If m_Orientation <> "portrait" And m_Orientation <> "landscape" Then _
        GoTo InvalidSectionStyles
    If Not private_TryParseMillimeters( _
        HtmlAttributeText(sectionNode, "marginTop"), m_MarginTopPoints) Then _
        GoTo InvalidSectionStyles
    If Not private_TryParseMillimeters( _
        HtmlAttributeText(sectionNode, "marginRight"), m_MarginRightPoints) Then _
        GoTo InvalidSectionStyles
    If Not private_TryParseMillimeters( _
        HtmlAttributeText(sectionNode, "marginBottom"), m_MarginBottomPoints) Then _
        GoTo InvalidSectionStyles
    If Not private_TryParseMillimeters( _
        HtmlAttributeText(sectionNode, "marginLeft"), m_MarginLeftPoints) Then _
        GoTo InvalidSectionStyles

    Initialize = True
    Exit Function

InvalidDocumentStyles:
    VBA.MsgBox "PrototypeNew: invalid or incomplete documentStyles markup.", _
        VBA.vbExclamation, "PrototypeNew / WORD styles"
    Exit Function
InvalidSectionStyles:
    VBA.MsgBox "PrototypeNew: invalid or incomplete sectionStyles markup.", _
        VBA.vbExclamation, "PrototypeNew / WORD styles"
End Function

Public Function HtmlAttributeText( _
    ByVal htmlNode As Object, _
    ByVal attributeName As String _
) As String
    Dim rawValue As Variant

    If htmlNode Is Nothing Then Exit Function
    On Error GoTo EH
    rawValue = htmlNode.getAttribute(attributeName)
    HtmlAttributeText = HtmlValueText(rawValue)
    Exit Function
EH:
    VBA.MsgBox "PrototypeNew: failed to read WORD style attribute '" & _
        attributeName & "': [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "PrototypeNew / WORD styles"
End Function

Public Function HtmlValueText(ByVal rawValue As Variant) As String
    ' MSHTML возвращает Null для отсутствующего необязательного атрибута.
    If VBA.IsNull(rawValue) Or VBA.IsEmpty(rawValue) Or VBA.IsError(rawValue) Then _
        Exit Function
    HtmlValueText = VBA.CStr(rawValue)
End Function

Public Function TryParseColumnWidths( _
    ByVal widthsText As String, _
    ByVal expectedCount As Long, _
    ByRef outWidths As Collection _
) As Boolean
    Dim widthParts As Variant
    Dim widthText As String
    Dim widthValue As Double
    Dim totalWidth As Double
    Dim widthIndex As Long

    Set outWidths = New Collection
    widthsText = VBA.Trim$(widthsText)
    If VBA.Len(widthsText) = 0 Then
        TryParseColumnWidths = True
        Exit Function
    End If
    widthParts = VBA.Split(widthsText, ";")
    If UBound(widthParts) - LBound(widthParts) + 1 <> expectedCount Then
        VBA.MsgBox "PrototypeNew: columnWidths must contain " & _
            VBA.CStr(expectedCount) & " values, configured: " & widthsText, _
            VBA.vbExclamation, "PrototypeNew / WORD styles"
        Exit Function
    End If
    For widthIndex = LBound(widthParts) To UBound(widthParts)
        widthText = VBA.Trim$(VBA.CStr(widthParts(widthIndex)))
        If VBA.Right$(widthText, 1) <> "%" Then GoTo InvalidWidth
        widthText = VBA.Trim$(VBA.Left$(widthText, VBA.Len(widthText) - 1))
        If Not VBA.IsNumeric(widthText) Then GoTo InvalidWidth
        widthValue = VBA.CDbl(widthText)
        If widthValue <= 0 Then GoTo InvalidWidth
        outWidths.Add VBA.CSng(widthValue)
        totalWidth = totalWidth + widthValue
    Next widthIndex
    If VBA.Abs(totalWidth - 100#) > 0.0001 Then
        VBA.MsgBox "PrototypeNew: columnWidths must total 100%, configured " & _
            "total: " & VBA.CStr(totalWidth) & "%.", VBA.vbExclamation, _
            "PrototypeNew / WORD styles"
        Exit Function
    End If
    TryParseColumnWidths = True
    Exit Function
InvalidWidth:
    VBA.MsgBox "PrototypeNew: invalid column width '" & widthText & _
        "'. Expected a positive percentage, for example 25%.", _
        VBA.vbExclamation, "PrototypeNew / WORD styles"
End Function

Public Function TryParseAlignment( _
    ByVal valueText As String, _
    ByRef outAlignment As String _
) As Boolean
    outAlignment = ex_Helpers.fn_NormalizeText(valueText)
    If VBA.Len(outAlignment) = 0 Then outAlignment = "left"
    If outAlignment <> "left" And outAlignment <> "center" Then
        VBA.MsgBox "PrototypeNew: unsupported table cell alignment '" & _
            outAlignment & "'.", VBA.vbExclamation, _
            "PrototypeNew / WORD styles"
        Exit Function
    End If
    TryParseAlignment = True
End Function

Public Function TryParseOptionalBoolean( _
    ByVal valueText As String, _
    ByRef outHasValue As Boolean, _
    ByRef outValue As Boolean _
) As Boolean
    valueText = VBA.Trim$(valueText)
    outHasValue = (VBA.Len(valueText) > 0)
    If Not outHasValue Then
        TryParseOptionalBoolean = True
        Exit Function
    End If
    If Not TryParseBoolean(valueText, outValue) Then
        VBA.MsgBox "PrototypeNew: invalid boolean WORD style value '" & _
            valueText & "'. Expected true or false.", VBA.vbExclamation, _
            "PrototypeNew / WORD styles"
        Exit Function
    End If
    TryParseOptionalBoolean = True
End Function

Public Function TryParseOptionalPositiveSingle( _
    ByVal valueText As String, _
    ByRef outHasValue As Boolean, _
    ByRef outValue As Single _
) As Boolean
    valueText = VBA.Trim$(valueText)
    outHasValue = (VBA.Len(valueText) > 0)
    If Not outHasValue Then
        TryParseOptionalPositiveSingle = True
        Exit Function
    End If
    If Not private_TryParsePositiveSingle(valueText, outValue) Then
        VBA.MsgBox "PrototypeNew: invalid positive WORD style number '" & _
            valueText & "'.", VBA.vbExclamation, _
            "PrototypeNew / WORD styles"
        Exit Function
    End If
    TryParseOptionalPositiveSingle = True
End Function

Public Function TryParseBorders( _
    ByVal valueText As String, _
    ByRef outBorders As Collection _
) As Boolean
    Dim borderParts As Variant
    Dim borderPart As Variant
    Dim borderName As String

    Set outBorders = New Collection
    valueText = ex_Helpers.fn_NormalizeText(valueText)
    If VBA.Len(valueText) = 0 Then
        TryParseBorders = True
        Exit Function
    End If
    borderParts = VBA.Split(valueText, ";")
    For Each borderPart In borderParts
        borderName = VBA.Trim$(VBA.CStr(borderPart))
        Select Case borderName
            Case "bottom", "top", "left", "right", "none"
                outBorders.Add borderName
            Case Else
                VBA.MsgBox "PrototypeNew: unsupported table cell border '" & _
                    borderName & "'.", VBA.vbExclamation, _
                    "PrototypeNew / WORD styles"
                Exit Function
        End Select
    Next borderPart
    TryParseBorders = True
End Function

Public Function TryParseTableBorders( _
    ByVal valueText As String, _
    ByRef outBordersDisabled As Boolean _
) As Boolean
    valueText = ex_Helpers.fn_NormalizeText(valueText)
    Select Case valueText
        Case VBA.vbNullString
            outBordersDisabled = False
        Case "none"
            outBordersDisabled = True
        Case Else
            VBA.MsgBox "PrototypeNew: unsupported table borders value '" & _
                valueText & "'.", VBA.vbExclamation, _
                "PrototypeNew / WORD styles"
            Exit Function
    End Select
    TryParseTableBorders = True
End Function

Public Function TryParseTableStyle( _
    ByVal valueText As String, _
    ByRef outStyleName As String _
) As Boolean
    outStyleName = VBA.LCase$(VBA.Trim$(valueText))
    Select Case outStyleName
        Case VBA.vbNullString, "builtin:tablegrid"
            TryParseTableStyle = True
        Case Else
            VBA.MsgBox "PrototypeNew: unsupported Word table style '" & _
                valueText & "'.", VBA.vbExclamation, _
                "PrototypeNew / WORD styles"
    End Select
End Function

Public Function TryParseOptionalPositiveLong( _
    ByVal valueText As String, _
    ByRef outValue As Long _
) As Boolean
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then
        TryParseOptionalPositiveLong = True
        Exit Function
    End If
    If Not VBA.IsNumeric(valueText) Then GoTo InvalidValue
    If VBA.CDbl(valueText) <= 0 Or VBA.CDbl(valueText) <> VBA.Fix(VBA.CDbl(valueText)) Then _
        GoTo InvalidValue
    outValue = VBA.CLng(valueText)
    TryParseOptionalPositiveLong = True
    Exit Function
InvalidValue:
    VBA.MsgBox "PrototypeNew: invalid positive integer WORD style value '" & _
        valueText & "'.", VBA.vbExclamation, "PrototypeNew / WORD styles"
End Function

Public Function TryParseBoolean( _
    ByVal valueText As String, _
    ByRef outValue As Boolean _
) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(valueText))
        Case "true": outValue = True
        Case "false": outValue = False
        Case Else: Exit Function
    End Select
    TryParseBoolean = True
End Function

Public Function TryParseHtmlColor( _
    ByVal colorText As String, _
    ByRef outColor As Long _
) As Boolean
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    colorText = VBA.Trim$(colorText)
    If VBA.Left$(colorText, 1) = "#" Then colorText = VBA.Mid$(colorText, 2)
    If VBA.Len(colorText) <> 6 Or Not colorText Like _
        "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" & _
        "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
        VBA.MsgBox "PrototypeNew: invalid HTML color '#" & colorText & _
            "'. Expected #RRGGBB.", VBA.vbExclamation, _
            "PrototypeNew / WORD styles"
        Exit Function
    End If
    redValue = VBA.CLng("&H" & VBA.Left$(colorText, 2))
    greenValue = VBA.CLng("&H" & VBA.Mid$(colorText, 3, 2))
    blueValue = VBA.CLng("&H" & VBA.Right$(colorText, 2))
    outColor = VBA.RGB(redValue, greenValue, blueValue)
    TryParseHtmlColor = True
End Function

Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Get ParagraphSpacingBefore() As Single
    ParagraphSpacingBefore = m_ParagraphSpacingBefore
End Property

Public Property Get ParagraphSpacingAfter() As Single
    ParagraphSpacingAfter = m_ParagraphSpacingAfter
End Property

Public Property Get LineSpacing() As String
    LineSpacing = m_LineSpacing
End Property

Public Property Get ShowTableGridlines() As Boolean
    ShowTableGridlines = m_ShowTableGridlines
End Property

Public Property Get PaperSize() As String
    PaperSize = m_PaperSize
End Property

Public Property Get Orientation() As String
    Orientation = m_Orientation
End Property

Public Property Get MarginTopPoints() As Single
    MarginTopPoints = m_MarginTopPoints
End Property

Public Property Get MarginRightPoints() As Single
    MarginRightPoints = m_MarginRightPoints
End Property

Public Property Get MarginBottomPoints() As Single
    MarginBottomPoints = m_MarginBottomPoints
End Property

Public Property Get MarginLeftPoints() As Single
    MarginLeftPoints = m_MarginLeftPoints
End Property

Private Function private_TryLoadRoot( _
    ByVal markupText As String, _
    ByVal expectedName As String, _
    ByRef outRoot As Object _
) As Boolean
    Dim xmlDoc As Object

    Set outRoot = Nothing
    Set xmlDoc = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    If Not xmlDoc.LoadXML(markupText) Then
        VBA.MsgBox "PrototypeNew: invalid " & expectedName & " markup: " & _
            xmlDoc.parseError.reason, VBA.vbExclamation, _
            "PrototypeNew / WORD styles"
        Exit Function
    End If
    Set outRoot = xmlDoc.DocumentElement
    If outRoot Is Nothing Then Exit Function
    If VBA.StrComp(VBA.CStr(outRoot.baseName), expectedName, _
        VBA.vbTextCompare) <> 0 Then Exit Function
    private_TryLoadRoot = True
End Function

Private Function private_TryParsePositiveSingle( _
    ByVal valueText As String, _
    ByRef outValue As Single _
) As Boolean
    If Not VBA.IsNumeric(valueText) Then Exit Function
    If VBA.CDbl(valueText) <= 0 Then Exit Function
    outValue = VBA.CSng(valueText)
    private_TryParsePositiveSingle = True
End Function

Private Function private_TryParseNonNegativeSingle( _
    ByVal valueText As String, _
    ByRef outValue As Single _
) As Boolean
    If Not VBA.IsNumeric(valueText) Then Exit Function
    If VBA.CDbl(valueText) < 0 Then Exit Function
    outValue = VBA.CSng(valueText)
    private_TryParseNonNegativeSingle = True
End Function

Private Function private_TryParseMillimeters( _
    ByVal valueText As String, _
    ByRef outPoints As Single _
) As Boolean
    Dim millimeters As Double

    valueText = VBA.LCase$(VBA.Trim$(valueText))
    If VBA.Right$(valueText, 2) <> "mm" Then Exit Function
    valueText = VBA.Trim$(VBA.Left$(valueText, VBA.Len(valueText) - 2))
    If Not VBA.IsNumeric(valueText) Then Exit Function
    millimeters = VBA.CDbl(valueText)
    If millimeters < 0 Then Exit Function
    outPoints = VBA.CSng(millimeters * 2.834645669)
    private_TryParseMillimeters = True
End Function
