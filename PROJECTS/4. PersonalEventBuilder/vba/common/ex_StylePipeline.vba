Option Explicit

Private Const COMMON_STYLE_CATALOG_FILE_NAME As String = "CommonControlStyles.xaml"
Private Const SHEET_SCOPE_MIN_COLUMN As Long = 40
Private Const SHEET_SCOPE_MIN_ROW As Long = 100
Private m_stylesByName As Object
Private m_pageDocument As Object
Private m_targetWorksheet As Worksheet

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_stylesByName = Nothing
    Set m_pageDocument = Nothing
    Set m_targetWorksheet = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_BeginPage( _
    ByVal targetWorksheet As Worksheet, _
    ByVal pageDocument As Object, _
    ByVal uiFolderPath As String _
)
    Dim fileSystem As Object
    Dim commonStyleDocument As Object
    Dim commonStylePath As String

    Set m_stylesByName = VBA.CreateObject("Scripting.Dictionary")
    m_stylesByName.CompareMode = VBA.vbTextCompare
    Set m_targetWorksheet = targetWorksheet
    Set m_pageDocument = pageDocument
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    commonStylePath = uiFolderPath & "\" & COMMON_STYLE_CATALOG_FILE_NAME
    If fileSystem.FileExists(commonStylePath) Then
        ex_Core.fn_Diagnostic_WriteLog "STYLE_CATALOG_LOAD_STARTED | Path=" & commonStylePath
        Set commonStyleDocument = VBA.CreateObject("Msxml2.DOMDocument.6.0")
        commonStyleDocument.async = False
        If commonStyleDocument.Load(commonStylePath) Then
            private_RegisterStyleNodes commonStyleDocument
            ex_Core.fn_Diagnostic_WriteLog "STYLE_CATALOG_LOAD_COMPLETED | Path=" & commonStylePath
        Else
            ex_Core.fn_Diagnostic_WriteLog "STYLE_CATALOG_LOAD_ERROR | Path=" & _
                commonStylePath & " | Description=" & commonStyleDocument.parseError.reason
        End If
    End If
    private_RegisterStyleNodes pageDocument
    ex_Core.fn_Diagnostic_WriteLog "STYLE_PAGE_READY | Sheet=" & targetWorksheet.Name & _
        " | StyleCount=" & VBA.CStr(m_stylesByName.Count)
End Sub

Public Sub fn_ApplyControlStyle( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal controlNode As Object _
)
    Dim styleName As String
    Dim styleProperties As Object
    Dim directProperties As Object

    If m_stylesByName Is Nothing Then Exit Sub
    styleName = private_ReadAttribute(controlNode, "style")
    If VBA.Len(styleName) > 0 Then
        If m_stylesByName.Exists(styleName) Then
            Set styleProperties = m_stylesByName(styleName)
            private_ApplyProperties targetRange, targetShape, styleProperties
        Else
            ex_Core.fn_Diagnostic_WriteLog "STYLE_NOT_FOUND | Sheet=" & _
                m_targetWorksheet.Name & " | Style=" & styleName
            VBA.MsgBox "Control style is not declared: " & styleName, _
                VBA.vbExclamation, "PersonalEventBuilder / Styles"
            VBA.Err.Raise 5, "ex_StylePipeline.fn_ApplyControlStyle", _
                "Control style is not declared: " & styleName
        End If
    End If

    Set directProperties = private_ReadVisualAttributes(controlNode)
    private_ApplyProperties targetRange, targetShape, directProperties
    private_ApplyPipelineRules targetRange, targetShape, controlNode
End Sub

Public Sub fn_ApplyPagePipeline(ByVal targetWorksheet As Worksheet)
    Dim stageNode As Object
    Dim layerNode As Object
    Dim ruleNode As Object
    Dim targetName As String
    Dim properties As Object
    Dim sheetScope As Range

    If m_pageDocument Is Nothing Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "STYLE_PIPELINE_STARTED | Sheet=" & targetWorksheet.Name
    For Each stageNode In m_pageDocument.SelectNodes( _
            "//*[local-name()='stylePipelineStage']")
        If private_IsNodeEnabled(stageNode) Then
            ex_Core.fn_Diagnostic_WriteLog "STYLE_PIPELINE_STAGE | Name=" & _
                private_ReadAttribute(stageNode, "name")
            For Each layerNode In stageNode.ChildNodes
                If layerNode.NodeType = 1 Then
                    For Each ruleNode In layerNode.ChildNodes
                        If ruleNode.NodeType = 1 And private_IsNodeEnabled(ruleNode) Then
                            targetName = VBA.LCase$(private_ReadAttribute(ruleNode, "target"))
                            Set properties = private_ParseStyleDeclarations( _
                                private_ReadAttribute(ruleNode, "styles"))
                            Select Case targetName
                                Case "sheet"
                                    Set sheetScope = private_GetSheetScope(targetWorksheet)
                                    ex_Core.fn_Diagnostic_WriteLog "STYLE_PIPELINE_RULE | Target=sheet | Scope=" & _
                                        sheetScope.Address(False, False)
                                    private_ApplyProperties sheetScope, Nothing, properties
                                    private_ApplyWorksheetProperties targetWorksheet, properties
                                Case "column"
                                    ex_Core.fn_Diagnostic_WriteLog "STYLE_PIPELINE_RULE | Target=column | Selector=" & _
                                        private_ReadAttribute(ruleNode, "selector")
                                    private_ApplyColumnRule targetWorksheet, ruleNode, properties
                            End Select
                        End If
                    Next ruleNode
                End If
            Next layerNode
        End If
    Next stageNode
    ex_Core.fn_Diagnostic_WriteLog "STYLE_PIPELINE_COMPLETED | Sheet=" & targetWorksheet.Name
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Sub private_RegisterStyleNodes(ByVal styleDocument As Object)
    Dim styleNode As Object
    Dim styleName As String
    Dim properties As Object

    For Each styleNode In styleDocument.SelectNodes("//*[local-name()='controlStyle']")
        styleName = private_ReadAttribute(styleNode, "name")
        If VBA.Len(styleName) > 0 Then
            Set properties = private_ReadVisualAttributes(styleNode)
            If m_stylesByName.Exists(styleName) Then m_stylesByName.Remove styleName
            m_stylesByName.Add styleName, properties
            ex_Core.fn_Diagnostic_WriteLog "STYLE_REGISTERED | Name=" & styleName
        End If
    Next styleNode
End Sub

Private Sub private_ApplyPipelineRules( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal controlNode As Object _
)
    Dim ruleNode As Object
    Dim targetName As String
    Dim properties As Object

    If m_pageDocument Is Nothing Then Exit Sub
    For Each ruleNode In m_pageDocument.SelectNodes( _
            "//*[local-name()='stylePipelineStage']//*[local-name()='rule']")
        targetName = VBA.LCase$(private_ReadAttribute(ruleNode, "target"))
        If private_IsNodeEnabled(ruleNode) Then
            If (targetName = "control" Or VBA.Len(targetName) = 0) And _
               private_MatchesSelector(controlNode, private_ReadAttribute(ruleNode, "selector")) Then
                Set properties = private_ParseStyleDeclarations( _
                    private_ReadAttribute(ruleNode, "styles"))
                private_ApplyProperties targetRange, targetShape, properties
            End If
            If targetName = "controlpart" Then _
                private_ApplyControlPartRule targetRange, targetShape, controlNode, ruleNode
        End If
    Next ruleNode
End Sub

Private Sub private_ApplyControlPartRule( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal controlNode As Object, _
    ByVal ruleNode As Object _
)
    Dim properties As Object
    Dim selectorText As String

    selectorText = private_ReadAttribute(ruleNode, "selector")
    Set properties = private_ParseStyleDeclarations(private_ReadAttribute(ruleNode, "styles"))
    If private_MatchesSelector(controlNode, selectorText, "control") Then _
        private_ApplyProperties targetRange, targetShape, properties
    If private_MatchesSelector(controlNode, selectorText, "cell") Then _
        private_ApplyProperties targetRange, Nothing, properties
    If Not targetShape Is Nothing Then
        If private_MatchesSelector(controlNode, selectorText, "shape") Then _
            private_ApplyProperties targetRange, targetShape, properties
    End If
End Sub

Private Function private_MatchesSelector( _
    ByVal controlNode As Object, _
    ByVal selectorText As String, _
    Optional ByVal partName As String = VBA.vbNullString _
) As Boolean
    Dim selectorParts As Variant
    Dim selectorPart As Variant
    Dim separatorPosition As Long
    Dim keyName As String
    Dim expectedValue As String
    Dim actualValue As String

    selectorText = VBA.Trim$(selectorText)
    If VBA.Len(selectorText) = 0 Then
        private_MatchesSelector = True
        Exit Function
    End If
    selectorParts = VBA.Split(selectorText, ";")
    For Each selectorPart In selectorParts
        separatorPosition = VBA.InStr(1, VBA.CStr(selectorPart), "=")
        If separatorPosition = 0 Then Exit Function
        keyName = VBA.LCase$(VBA.Trim$(VBA.Left$(selectorPart, separatorPosition - 1)))
        expectedValue = VBA.Trim$(VBA.Mid$(selectorPart, separatorPosition + 1))
        Select Case keyName
            Case "sheet"
                actualValue = m_targetWorksheet.Name
            Case "type", "name", "tags"
                actualValue = private_ReadAttribute(controlNode, keyName)
            Case "part"
                actualValue = partName
            Case Else
                Exit Function
        End Select
        If VBA.StrComp(actualValue, expectedValue, VBA.vbTextCompare) <> 0 Then Exit Function
    Next selectorPart
    private_MatchesSelector = True
End Function

Private Function private_ReadVisualAttributes(ByVal node As Object) As Object
    Dim result As Object
    Dim attributeNode As Object
    Dim propertyName As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    For Each attributeNode In node.Attributes
        propertyName = VBA.LCase$(VBA.CStr(attributeNode.nodeName))
        Select Case propertyName
            Case "backcolor", "textcolor", "fontcolor", "bordercolor", "borderweight", _
                 "borderlinestyle", "fontbold", "fontitalic", "fontsize", _
                 "fontname", "horizontal", "vertical", "columnwidth", "width", "rowheight", _
                 "overflow", "zoom", "gridlines"
                result(propertyName) = VBA.CStr(attributeNode.Text)
        End Select
    Next attributeNode
    Set private_ReadVisualAttributes = result
End Function

Private Function private_ParseStyleDeclarations(ByVal declarationText As String) As Object
    Dim result As Object
    Dim declarationParts As Variant
    Dim declarationPart As Variant
    Dim separatorPosition As Long
    Dim propertyName As String
    Dim propertyValue As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    declarationText = VBA.Trim$(declarationText)
    If VBA.Left$(declarationText, 1) = "{" Then declarationText = VBA.Mid$(declarationText, 2)
    If VBA.Right$(declarationText, 1) = "}" Then _
        declarationText = VBA.Left$(declarationText, VBA.Len(declarationText) - 1)
    declarationParts = VBA.Split(declarationText, ";")
    For Each declarationPart In declarationParts
        separatorPosition = VBA.InStr(1, VBA.CStr(declarationPart), ":")
        If separatorPosition > 0 Then
            propertyName = VBA.LCase$(private_TrimXmlWhitespace( _
                VBA.Left$(declarationPart, separatorPosition - 1)))
            propertyValue = private_TrimXmlWhitespace( _
                VBA.Mid$(declarationPart, separatorPosition + 1))
            If private_IsVisualProperty(propertyName) Then
                result(propertyName) = propertyValue
            End If
        End If
    Next declarationPart
    Set private_ParseStyleDeclarations = result
End Function

Private Function private_IsVisualProperty(ByVal propertyName As String) As Boolean
    Select Case propertyName
            Case "backcolor", "textcolor", "fontcolor", "bordercolor", "borderweight", _
             "borderlinestyle", "fontbold", "fontitalic", "fontsize", _
                 "fontname", "horizontal", "vertical", "columnwidth", "width", "rowheight", _
                  "overflow", "zoom", "gridlines"
            private_IsVisualProperty = True
    End Select
End Function

Private Function private_TrimXmlWhitespace(ByVal valueText As String) As String
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    private_TrimXmlWhitespace = VBA.Trim$(valueText)
End Function

Private Sub private_ApplyProperties( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal properties As Object _
)
    Dim colorValue As Long

    If properties Is Nothing Then Exit Sub
    If properties.Exists("backcolor") Then
        If private_TryParseColor(properties("backcolor"), colorValue) Then
            If targetShape Is Nothing Then
                targetRange.Interior.Pattern = xlSolid
                targetRange.Interior.Color = colorValue
            Else
                targetShape.Fill.ForeColor.RGB = colorValue
            End If
        End If
    End If
    If properties.Exists("textcolor") Or properties.Exists("fontcolor") Then
        If properties.Exists("textcolor") Then
            If Not private_TryParseColor(properties("textcolor"), colorValue) Then Exit Sub
        ElseIf Not private_TryParseColor(properties("fontcolor"), colorValue) Then
            Exit Sub
        End If
        targetRange.Font.Color = colorValue
        If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorValue
    End If
    If properties.Exists("fontbold") Then
        targetRange.Font.Bold = private_ReadBoolean(properties("fontbold"))
        If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.Font.Bold = _
            IIf(private_ReadBoolean(properties("fontbold")), -1, 0)
    End If
    If properties.Exists("fontitalic") Then
        targetRange.Font.Italic = private_ReadBoolean(properties("fontitalic"))
        If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.Font.Italic = _
            IIf(private_ReadBoolean(properties("fontitalic")), -1, 0)
    End If
    If properties.Exists("fontsize") And VBA.IsNumeric(properties("fontsize")) Then
        targetRange.Font.Size = VBA.CDbl(properties("fontsize"))
        If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.Font.Size = _
            VBA.CDbl(properties("fontsize"))
    End If
    If properties.Exists("fontname") Then
        targetRange.Font.Name = properties("fontname")
        If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.Font.Name = properties("fontname")
    End If
    If properties.Exists("horizontal") Then private_ApplyHorizontalAlignment _
        targetRange, targetShape, properties("horizontal")
    If properties.Exists("vertical") Then private_ApplyVerticalAlignment _
        targetRange, targetShape, properties("vertical")
    If properties.Exists("columnwidth") And VBA.IsNumeric(properties("columnwidth")) Then _
        targetRange.EntireColumn.ColumnWidth = VBA.CDbl(properties("columnwidth"))
    If properties.Exists("width") And VBA.IsNumeric(properties("width")) Then _
        targetRange.EntireColumn.ColumnWidth = VBA.CDbl(properties("width"))
    If properties.Exists("rowheight") And VBA.IsNumeric(properties("rowheight")) Then _
        targetRange.EntireRow.RowHeight = VBA.CDbl(properties("rowheight"))
    If properties.Exists("overflow") Then private_ApplyOverflow targetRange, properties("overflow")
    private_ApplyBorders targetRange, targetShape, properties
End Sub

Private Sub private_ApplyBorders( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal properties As Object _
)
    Dim colorValue As Long
    Dim borderIndex As Variant
    Dim lineStyle As Long
    Dim lineWeight As Long
    Dim shapeLineWeight As Single

    If Not properties.Exists("bordercolor") And Not properties.Exists("borderweight") And _
       Not properties.Exists("borderlinestyle") Then Exit Sub
    lineStyle = private_ReadBorderLineStyle(properties)
    lineWeight = private_ReadBorderWeight(properties)
    shapeLineWeight = private_ReadShapeBorderWeight(properties)
    If targetShape Is Nothing Then
        For Each borderIndex In Array( _
                xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                xlInsideHorizontal, xlInsideVertical)
            targetRange.Borders(borderIndex).LineStyle = lineStyle
            targetRange.Borders(borderIndex).Weight = lineWeight
            If properties.Exists("bordercolor") Then
                If private_TryParseColor(properties("bordercolor"), colorValue) Then _
                    targetRange.Borders(borderIndex).Color = colorValue
            End If
        Next borderIndex
    End If
    If Not targetShape Is Nothing Then
        targetShape.Line.Visible = -1
        targetShape.Line.Weight = shapeLineWeight
        If properties.Exists("bordercolor") Then
            If private_TryParseColor(properties("bordercolor"), colorValue) Then _
                targetShape.Line.ForeColor.RGB = colorValue
        End If
    End If
End Sub

Private Sub private_ApplyWorksheetProperties( _
    ByVal targetWorksheet As Worksheet, _
    ByVal properties As Object _
)
    Dim zoomText As String

    If properties.Exists("gridlines") Then _
        ActiveWindow.DisplayGridlines = private_ReadBoolean(properties("gridlines"))
    If properties.Exists("zoom") Then
        zoomText = VBA.Replace(VBA.Trim$(properties("zoom")), "%", VBA.vbNullString)
        If VBA.IsNumeric(zoomText) Then ActiveWindow.Zoom = VBA.CLng(zoomText)
    End If
End Sub

Private Function private_GetSheetScope(ByVal targetWorksheet As Worksheet) As Range
    Dim lastRow As Long
    Dim lastColumn As Long

    lastRow = Application.Max(SHEET_SCOPE_MIN_ROW, targetWorksheet.UsedRange.Row + _
        targetWorksheet.UsedRange.Rows.Count - 1)
    lastColumn = Application.Max(SHEET_SCOPE_MIN_COLUMN, targetWorksheet.UsedRange.Column + _
        targetWorksheet.UsedRange.Columns.Count - 1)
    Set private_GetSheetScope = targetWorksheet.Range( _
        targetWorksheet.Cells(1, 1), targetWorksheet.Cells(lastRow, lastColumn))
End Function

Private Sub private_ApplyColumnRule( _
    ByVal targetWorksheet As Worksheet, _
    ByVal ruleNode As Object, _
    ByVal properties As Object _
)
    Dim selectorText As String
    Dim addressText As String
    Dim targetRange As Range

    selectorText = private_ReadAttribute(ruleNode, "selector")
    addressText = private_ReadSelectorValue(selectorText, "address")
    If VBA.Len(addressText) = 0 Then
        VBA.MsgBox "Column style rule requires selector address=... .", _
            VBA.vbExclamation, "PersonalEventBuilder / Styles"
        Exit Sub
    End If
    On Error GoTo EH
    Set targetRange = targetWorksheet.Range(addressText)
    private_ApplyProperties targetRange, Nothing, properties
    Exit Sub
EH:
    VBA.MsgBox "Column style rule address is invalid: " & addressText, _
        VBA.vbExclamation, "PersonalEventBuilder / Styles"
End Sub

Private Sub private_ApplyOverflow(ByVal targetRange As Range, ByVal overflowText As String)
    Select Case VBA.LCase$(VBA.Trim$(overflowText))
        Case "wrap"
            targetRange.WrapText = True
        Case "clip"
            targetRange.WrapText = False
        Case Else
            VBA.MsgBox "Unsupported overflow value: " & overflowText, _
                VBA.vbExclamation, "PersonalEventBuilder / Styles"
    End Select
End Sub

Private Sub private_ApplyHorizontalAlignment( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal alignmentText As String _
)
    Select Case VBA.LCase$(alignmentText)
        Case "left"
            targetRange.HorizontalAlignment = xlLeft
            If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        Case "right"
            targetRange.HorizontalAlignment = xlRight
            If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        Case Else
            targetRange.HorizontalAlignment = xlCenter
            If Not targetShape Is Nothing Then targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End Select
End Sub

Private Sub private_ApplyVerticalAlignment( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal alignmentText As String _
)
    Select Case VBA.LCase$(alignmentText)
        Case "top"
            targetRange.VerticalAlignment = xlTop
            If Not targetShape Is Nothing Then targetShape.TextFrame2.VerticalAnchor = msoAnchorTop
        Case "bottom"
            targetRange.VerticalAlignment = xlBottom
            If Not targetShape Is Nothing Then targetShape.TextFrame2.VerticalAnchor = msoAnchorBottom
        Case Else
            targetRange.VerticalAlignment = xlCenter
            If Not targetShape Is Nothing Then targetShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    End Select
End Sub

Private Function private_TryParseColor(ByVal colorText As String, ByRef outColor As Long) As Boolean
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    colorText = VBA.Trim$(colorText)
    If VBA.Left$(colorText, 1) = "#" Then colorText = VBA.Mid$(colorText, 2)
    If VBA.Len(colorText) <> 6 Then Exit Function
    On Error GoTo EH
    redValue = VBA.CLng("&H" & VBA.Mid$(colorText, 1, 2))
    greenValue = VBA.CLng("&H" & VBA.Mid$(colorText, 3, 2))
    blueValue = VBA.CLng("&H" & VBA.Mid$(colorText, 5, 2))
    outColor = VBA.RGB(redValue, greenValue, blueValue)
    private_TryParseColor = True
EH:
End Function

Private Function private_ReadBoolean(ByVal valueText As String) As Boolean
    private_ReadBoolean = VBA.StrComp(VBA.Trim$(valueText), "true", VBA.vbTextCompare) = 0 Or _
        VBA.StrComp(VBA.Trim$(valueText), "yes", VBA.vbTextCompare) = 0 Or _
        VBA.Trim$(valueText) = "1"
End Function

Private Function private_ReadBorderLineStyle(ByVal properties As Object) As Long
    private_ReadBorderLineStyle = xlContinuous
    If Not properties.Exists("borderlinestyle") Then Exit Function
    Select Case VBA.LCase$(properties("borderlinestyle"))
        Case "dash": private_ReadBorderLineStyle = xlDash
        Case "dot": private_ReadBorderLineStyle = xlDot
        Case "none": private_ReadBorderLineStyle = xlLineStyleNone
    End Select
End Function

Private Function private_ReadBorderWeight(ByVal properties As Object) As Long
    private_ReadBorderWeight = xlThin
    If Not properties.Exists("borderweight") Then Exit Function
    If VBA.IsNumeric(properties("borderweight")) Then
        private_ReadBorderWeight = VBA.CLng(VBA.CDbl(properties("borderweight")))
        Exit Function
    End If
    Select Case VBA.LCase$(properties("borderweight"))
        Case "hairline": private_ReadBorderWeight = xlHairline
        Case "medium": private_ReadBorderWeight = xlMedium
        Case "thick": private_ReadBorderWeight = xlThick
    End Select
End Function

Private Function private_ReadShapeBorderWeight(ByVal properties As Object) As Single
    private_ReadShapeBorderWeight = 0.75
    If Not properties.Exists("borderweight") Then Exit Function
    If VBA.IsNumeric(properties("borderweight")) Then
        private_ReadShapeBorderWeight = VBA.CSng(properties("borderweight"))
        Exit Function
    End If
    Select Case VBA.LCase$(properties("borderweight"))
        Case "hairline": private_ReadShapeBorderWeight = 0.25
        Case "medium": private_ReadShapeBorderWeight = 1.5
        Case "thick": private_ReadShapeBorderWeight = 2.25
    End Select
End Function

Private Function private_IsNodeEnabled(ByVal node As Object) As Boolean
    Dim enabledText As String

    enabledText = private_ReadAttribute(node, "enabled")
    If VBA.Len(enabledText) = 0 Then
        private_IsNodeEnabled = True
    Else
        private_IsNodeEnabled = private_ReadBoolean(enabledText)
    End If
End Function

Private Function private_ReadSelectorValue( _
    ByVal selectorText As String, _
    ByVal requestedKey As String _
) As String
    Dim selectorParts As Variant
    Dim selectorPart As Variant
    Dim separatorPosition As Long
    Dim keyName As String

    selectorParts = VBA.Split(selectorText, ";")
    For Each selectorPart In selectorParts
        separatorPosition = VBA.InStr(1, VBA.CStr(selectorPart), "=")
        If separatorPosition > 0 Then
            keyName = VBA.LCase$(VBA.Trim$(VBA.Left$(selectorPart, separatorPosition - 1)))
            If VBA.StrComp(keyName, requestedKey, VBA.vbTextCompare) = 0 Then
                private_ReadSelectorValue = VBA.Trim$( _
                    VBA.Mid$(selectorPart, separatorPosition + 1))
                Exit Function
            End If
        End If
    Next selectorPart
End Function

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant

    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function