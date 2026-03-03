Attribute VB_Name = "ex_ConfigStylesParser"
Option Explicit

Private Enum e_ConfigStylePropertyId
    cfgStylePropUnknown = 0
    cfgStylePropWidth = 1
    cfgStylePropMinWidth = 2
    cfgStylePropMaxWidth = 3
    cfgStylePropAutoFitColumns = 4
    cfgStylePropOverflow = 5
    cfgStylePropAutoHeight = 6
    cfgStylePropRowHeight = 7
    cfgStylePropMergeColumns = 8
    cfgStylePropFontName = 9
    cfgStylePropFontSize = 10
    cfgStylePropFontBold = 11
    cfgStylePropBackColor = 12
    cfgStylePropFontColor = 13
    cfgStylePropHorizontal = 14
    cfgStylePropVertical = 15
End Enum

' Supported style properties (declarations):
' width, minWidth, maxWidth, autoFitColumns
' overflow, autoHeight, rowHeight, mergeColumns
' fontName, fontSize, fontBold
' backColor, fontColor
' horizontal, vertical
Private Const STYLE_PROPERTY_WIDTH As String = "width"
Private Const STYLE_PROPERTY_MIN_WIDTH As String = "minWidth"
Private Const STYLE_PROPERTY_MAX_WIDTH As String = "maxWidth"
Private Const STYLE_PROPERTY_AUTO_FIT_COLUMNS As String = "autoFitColumns"
Private Const STYLE_PROPERTY_OVERFLOW As String = "overflow"
Private Const STYLE_PROPERTY_AUTO_HEIGHT As String = "autoHeight"
Private Const STYLE_PROPERTY_ROW_HEIGHT As String = "rowHeight"
Private Const STYLE_PROPERTY_MERGE_COLUMNS As String = "mergeColumns"
Private Const STYLE_PROPERTY_FONT_NAME As String = "fontName"
Private Const STYLE_PROPERTY_FONT_SIZE As String = "fontSize"
Private Const STYLE_PROPERTY_FONT_BOLD As String = "fontBold"
Private Const STYLE_PROPERTY_BACK_COLOR As String = "backColor"
Private Const STYLE_PROPERTY_FONT_COLOR As String = "fontColor"
Private Const STYLE_PROPERTY_HORIZONTAL As String = "horizontal"
Private Const STYLE_PROPERTY_VERTICAL As String = "vertical"

Public Sub m_ApplyColumnStylesByMapKeys( _
    ByVal ws As Worksheet, _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgNotes As Object _
)
    Dim i As Long
    Dim target As Object
    Dim mapKey As String
    Dim noteText As String
    Dim parsedStyles As Object
    Dim hasStyleBlock As Boolean
    Dim parseError As String
    Dim outCol As Long
    Dim rowStart As Long
    Dim rowEnd As Long

    If ws Is Nothing Then Exit Sub
    If resultFieldRanges Is Nothing Then Exit Sub
    If cfgNotes Is Nothing Then Exit Sub
    If resultFieldRanges.Count = 0 Then Exit Sub

    For i = 1 To resultFieldRanges.Count
        Set target = resultFieldRanges(i)
        If target Is Nothing Then GoTo ContinueTarget

        mapKey = Trim$(CStr(target("MapKey")))
        If Len(mapKey) = 0 Then GoTo ContinueTarget
        If Not cfgNotes.Exists(mapKey) Then GoTo ContinueTarget

        noteText = Trim$(CStr(cfgNotes(mapKey)))
        If Len(noteText) = 0 Then GoTo ContinueTarget

        hasStyleBlock = False
        parseError = vbNullString
        If Not mp_TryParseStyleMap(noteText, parsedStyles, hasStyleBlock, parseError) Then
            Err.Raise vbObjectError + 1491, "ex_ConfigStylesParser", _
                "Invalid styles definition for key '" & mapKey & "': " & parseError & ". Source: '" & noteText & "'."
        End If
        If Not hasStyleBlock Then GoTo ContinueTarget

        outCol = CLng(target("ColumnIndex"))
        If outCol <= 0 Then GoTo ContinueTarget

        If target.Exists("RowStart") Then rowStart = CLng(target("RowStart")) Else rowStart = 1
        If target.Exists("RowEnd") Then rowEnd = CLng(target("RowEnd")) Else rowEnd = rowStart
        If rowStart <= 0 Then rowStart = 1
        If rowEnd < rowStart Then rowEnd = rowStart

        mp_ApplyParsedStylesToColumn ws, outCol, rowStart, rowEnd, parsedStyles, mapKey

ContinueTarget:
    Next i
End Sub

Public Function m_TryParseStyleDeclarations( _
    ByVal styleText As String, _
    ByRef outDeclarations As Object, _
    ByRef outHasDeclarationBlock As Boolean, _
    ByRef outErrorText As String _
) As Boolean
    Dim normalized As String

    normalized = Trim$(styleText)
    outErrorText = vbNullString
    outHasDeclarationBlock = False

    If Not mp_TryParseStyleMap(normalized, outDeclarations, outHasDeclarationBlock, outErrorText) Then
        Exit Function
    End If

    If outHasDeclarationBlock Then
        m_TryParseStyleDeclarations = True
        Exit Function
    End If

    ' Allow compact declarations syntax without braces for style catalogs:
    ' width:40;overflow:wrap;autoHeight:true
    If InStr(1, normalized, ":", vbBinaryCompare) > 0 Then
        If Not mp_TryParseStyleMap("{" & normalized & "}", outDeclarations, outHasDeclarationBlock, outErrorText) Then
            Exit Function
        End If
    End If

    m_TryParseStyleDeclarations = True
End Function

Public Function m_ValidateColumnStylesByMapKeys( _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgNotes As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim target As Object
    Dim mapKey As String
    Dim noteText As String
    Dim parsedStyles As Object
    Dim hasStyleBlock As Boolean
    Dim parseError As String
    Dim stepName As String

    On Error GoTo EH

    If resultFieldRanges Is Nothing Then
        m_ValidateColumnStylesByMapKeys = True
        Exit Function
    End If
    If cfgNotes Is Nothing Then
        m_ValidateColumnStylesByMapKeys = True
        Exit Function
    End If
    If resultFieldRanges.Count = 0 Then
        m_ValidateColumnStylesByMapKeys = True
        Exit Function
    End If

    For i = 1 To resultFieldRanges.Count
        stepName = "read-target"
        Set target = resultFieldRanges(i)
        If target Is Nothing Then GoTo ContinueTarget

        stepName = "read-map-key"
        mapKey = Trim$(CStr(target("MapKey")))
        If Len(mapKey) = 0 Then GoTo ContinueTarget
        If Not cfgNotes.Exists(mapKey) Then GoTo ContinueTarget

        stepName = "read-note"
        noteText = Trim$(CStr(cfgNotes(mapKey)))
        If Len(noteText) = 0 Then GoTo ContinueTarget

        stepName = "parse-style-map"
        parseError = vbNullString
        hasStyleBlock = False
        If Not mp_TryParseStyleMap(noteText, parsedStyles, hasStyleBlock, parseError) Then
            outErrorText = "Invalid styles definition for key '" & mapKey & "': " & parseError & ". Source: '" & noteText & "'."
            Exit Function
        End If

ContinueTarget:
    Next i

    m_ValidateColumnStylesByMapKeys = True
    Exit Function

EH:
    outErrorText = "Style validation runtime error"
    If Len(mapKey) > 0 Then
        outErrorText = outErrorText & " for key '" & mapKey & "'"
    End If
    If Len(stepName) > 0 Then
        outErrorText = outErrorText & " at step '" & stepName & "'"
    End If
    outErrorText = outErrorText & ": [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_TryParseStyleMap( _
    ByVal noteText As String, _
    ByRef outStyles As Object, _
    ByRef hasStyleBlock As Boolean, _
    ByRef outErrorText As String _
) As Boolean
    Dim sourceText As String
    Dim blockText As String
    Dim openPos As Long
    Dim closePos As Long
    Dim pairs As Variant
    Dim pairText As String
    Dim colonPos As Long
    Dim propertyName As String
    Dim propertyValue As String
    Dim propertyId As e_ConfigStylePropertyId
    Dim propertyIds As Object
    Dim discreteValidators As Object
    Dim i As Long

    Set outStyles = CreateObject("Scripting.Dictionary")
    outStyles.CompareMode = 1

    sourceText = mp_NormalizeStyleToken(noteText)
    If Len(sourceText) = 0 Then
        mp_TryParseStyleMap = True
        Exit Function
    End If

    openPos = InStr(1, sourceText, "{", vbBinaryCompare)
    closePos = InStrRev(sourceText, "}", -1, vbBinaryCompare)

    If openPos = 0 And closePos = 0 Then
        mp_TryParseStyleMap = True
        Exit Function
    End If

    hasStyleBlock = True

    If openPos <= 0 Or closePos <= 0 Or closePos <= openPos Then
        outErrorText = "style block must match pattern '{prop:value;...}'"
        Exit Function
    End If

    blockText = mp_NormalizeStyleToken(Mid$(sourceText, openPos + 1, closePos - openPos - 1))
    If Len(blockText) = 0 Then
        outErrorText = "style block is empty"
        Exit Function
    End If

    Set propertyIds = mp_BuildSupportedPropertyIds()
    Set discreteValidators = mp_BuildDiscreteValidators()

    pairs = Split(blockText, ";")
    For i = LBound(pairs) To UBound(pairs)
        pairText = mp_NormalizeStyleToken(CStr(pairs(i)))
        If Len(pairText) = 0 Then GoTo ContinuePair

        colonPos = InStr(1, pairText, ":", vbBinaryCompare)
        If colonPos <= 1 Then
            outErrorText = "invalid style token '" & pairText & "'"
            Exit Function
        End If

        propertyName = LCase$(mp_NormalizeStyleToken(Left$(pairText, colonPos - 1)))
        propertyValue = mp_UnquoteText(mp_NormalizeStyleToken(Mid$(pairText, colonPos + 1)))

        If Len(propertyName) = 0 Then
            outErrorText = "property name is empty in token '" & pairText & "'"
            Exit Function
        End If
        If Len(propertyValue) = 0 Then
            outErrorText = "value is empty for property '" & propertyName & "'"
            Exit Function
        End If

        If Not propertyIds.Exists(propertyName) Then
            outErrorText = "unsupported style property '" & propertyName & "'"
            Exit Function
        End If

        propertyId = propertyIds(propertyName)
        If Not mp_ValidatePropertyValue(propertyId, propertyValue, discreteValidators, outErrorText) Then
            Exit Function
        End If

        outStyles(propertyName) = propertyValue

ContinuePair:
    Next i

    mp_TryParseStyleMap = True
End Function

Private Function mp_BuildSupportedPropertyIds() As Object
    Dim propertyIds As Object
    Set propertyIds = CreateObject("Scripting.Dictionary")
    propertyIds.CompareMode = 1

    propertyIds(STYLE_PROPERTY_WIDTH) = cfgStylePropWidth
    propertyIds(LCase$(STYLE_PROPERTY_MIN_WIDTH)) = cfgStylePropMinWidth
    propertyIds(LCase$(STYLE_PROPERTY_MAX_WIDTH)) = cfgStylePropMaxWidth
    propertyIds(LCase$(STYLE_PROPERTY_AUTO_FIT_COLUMNS)) = cfgStylePropAutoFitColumns
    propertyIds(STYLE_PROPERTY_OVERFLOW) = cfgStylePropOverflow
    propertyIds(STYLE_PROPERTY_AUTO_HEIGHT) = cfgStylePropAutoHeight
    propertyIds(LCase$(STYLE_PROPERTY_ROW_HEIGHT)) = cfgStylePropRowHeight
    propertyIds(LCase$(STYLE_PROPERTY_MERGE_COLUMNS)) = cfgStylePropMergeColumns
    propertyIds(LCase$(STYLE_PROPERTY_FONT_NAME)) = cfgStylePropFontName
    propertyIds(LCase$(STYLE_PROPERTY_FONT_SIZE)) = cfgStylePropFontSize
    propertyIds(LCase$(STYLE_PROPERTY_FONT_BOLD)) = cfgStylePropFontBold
    propertyIds(LCase$(STYLE_PROPERTY_BACK_COLOR)) = cfgStylePropBackColor
    propertyIds(LCase$(STYLE_PROPERTY_FONT_COLOR)) = cfgStylePropFontColor
    propertyIds(LCase$(STYLE_PROPERTY_HORIZONTAL)) = cfgStylePropHorizontal
    propertyIds(LCase$(STYLE_PROPERTY_VERTICAL)) = cfgStylePropVertical

    Set mp_BuildSupportedPropertyIds = propertyIds
End Function

Private Function mp_BuildDiscreteValidators() As Object
    Dim validators As Object
    Dim overflowAllowed As Object
    Dim horizontalAllowed As Object
    Dim verticalAllowed As Object

    Set validators = CreateObject("Scripting.Dictionary")
    validators.CompareMode = 1

    Set overflowAllowed = CreateObject("Scripting.Dictionary")
    overflowAllowed.CompareMode = 1
    overflowAllowed("wrap") = True
    overflowAllowed("clip") = True
    overflowAllowed("shrink") = True

    Dim boolAllowed As Object
    Set boolAllowed = CreateObject("Scripting.Dictionary")
    boolAllowed.CompareMode = 1
    boolAllowed("true") = True
    boolAllowed("false") = True

    validators.Add CStr(cfgStylePropOverflow), overflowAllowed
    validators.Add CStr(cfgStylePropAutoHeight), boolAllowed
    validators.Add CStr(cfgStylePropAutoFitColumns), boolAllowed
    validators.Add CStr(cfgStylePropFontBold), boolAllowed

    Set horizontalAllowed = CreateObject("Scripting.Dictionary")
    horizontalAllowed.CompareMode = 1
    horizontalAllowed("left") = True
    horizontalAllowed("center") = True
    horizontalAllowed("right") = True
    horizontalAllowed("fill") = True
    horizontalAllowed("justify") = True
    horizontalAllowed("distributed") = True
    horizontalAllowed("general") = True
    validators.Add CStr(cfgStylePropHorizontal), horizontalAllowed

    Set verticalAllowed = CreateObject("Scripting.Dictionary")
    verticalAllowed.CompareMode = 1
    verticalAllowed("top") = True
    verticalAllowed("center") = True
    verticalAllowed("bottom") = True
    verticalAllowed("justify") = True
    verticalAllowed("distributed") = True
    validators.Add CStr(cfgStylePropVertical), verticalAllowed

    Set mp_BuildDiscreteValidators = validators
End Function

Private Function mp_ValidatePropertyValue( _
    ByVal propertyId As e_ConfigStylePropertyId, _
    ByVal propertyValue As String, _
    ByVal discreteValidators As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim widthValue As Double
    Dim normalizedValue As String
    Dim allowedValues As Object
    Dim boolValue As Boolean
    Dim colorValue As Long
    Dim longValue As Long

    Select Case propertyId
        Case cfgStylePropWidth, cfgStylePropMinWidth, cfgStylePropMaxWidth
            If Not mp_TryParseWidth(propertyValue, widthValue) Then
                outErrorText = "invalid numeric width value '" & propertyValue & "' (expected positive number)"
                Exit Function
            End If

        Case cfgStylePropAutoFitColumns
            If Not mp_TryParseBoolean(propertyValue, boolValue) Then
                outErrorText = "invalid autoFitColumns value '" & propertyValue & "' (expected true/false)"
                Exit Function
            End If

        Case cfgStylePropOverflow
            normalizedValue = LCase$(Trim$(propertyValue))
            If Not discreteValidators.Exists(CStr(propertyId)) Then
                outErrorText = "overflow validator is not configured"
                Exit Function
            End If

            Set allowedValues = discreteValidators(CStr(propertyId))
            If Not allowedValues.Exists(normalizedValue) Then
                outErrorText = "unsupported overflow value '" & propertyValue & "'"
                Exit Function
            End If

        Case cfgStylePropAutoHeight
            If Not mp_TryParseBoolean(propertyValue, boolValue) Then
                outErrorText = "invalid autoHeight value '" & propertyValue & "' (expected true/false)"
                Exit Function
            End If

        Case cfgStylePropRowHeight
            If Not mp_TryParsePositiveDouble(propertyValue, widthValue) Then
                outErrorText = "invalid rowHeight value '" & propertyValue & "' (expected positive number)"
                Exit Function
            End If

        Case cfgStylePropMergeColumns
            If Not ex_XmlCore.m_TryParseLong(propertyValue, longValue) Then
                outErrorText = "invalid mergeColumns value '" & propertyValue & "' (expected integer >= 1)"
                Exit Function
            End If
            If longValue < 1 Then
                outErrorText = "invalid mergeColumns value '" & propertyValue & "' (expected integer >= 1)"
                Exit Function
            End If

        Case cfgStylePropFontName
            If Len(Trim$(propertyValue)) = 0 Then
                outErrorText = "invalid fontName value (expected non-empty text)"
                Exit Function
            End If

        Case cfgStylePropFontSize
            If Not mp_TryParsePositiveDouble(propertyValue, widthValue) Then
                outErrorText = "invalid fontSize value '" & propertyValue & "' (expected positive number)"
                Exit Function
            End If

        Case cfgStylePropFontBold
            If Not mp_TryParseBoolean(propertyValue, boolValue) Then
                outErrorText = "invalid fontBold value '" & propertyValue & "' (expected true/false)"
                Exit Function
            End If

        Case cfgStylePropBackColor, cfgStylePropFontColor
            If Not ex_XmlCore.m_TryParseColor(propertyValue, colorValue) Then
                outErrorText = "invalid color value '" & propertyValue & "'"
                Exit Function
            End If

        Case cfgStylePropHorizontal, cfgStylePropVertical
            normalizedValue = LCase$(Trim$(propertyValue))
            If Not discreteValidators.Exists(CStr(propertyId)) Then
                outErrorText = "alignment validator is not configured"
                Exit Function
            End If
            Set allowedValues = discreteValidators(CStr(propertyId))
            If Not allowedValues.Exists(normalizedValue) Then
                outErrorText = "unsupported alignment value '" & propertyValue & "'"
                Exit Function
            End If

        Case Else
            outErrorText = "unsupported style property id '" & CStr(propertyId) & "'"
            Exit Function
    End Select

    mp_ValidatePropertyValue = True
End Function

Private Function mp_TryParseWidth(ByVal valueText As String, ByRef outWidth As Double) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    If Len(normalized) >= 2 Then
        If LCase$(Right$(normalized, 2)) = "px" Then
            normalized = Trim$(Left$(normalized, Len(normalized) - 2))
        End If
    End If

    If Not IsNumeric(normalized) Then Exit Function

    outWidth = CDbl(normalized)
    If outWidth <= 0 Then Exit Function

    mp_TryParseWidth = True
End Function

Private Sub mp_ApplyParsedStylesToColumn( _
    ByVal ws As Worksheet, _
    ByVal outCol As Long, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long, _
    ByVal parsedStyles As Object, _
    ByVal mapKey As String _
)
    Dim widthUnits As Double
    Dim minWidthUnits As Double
    Dim maxWidthUnits As Double
    Dim hasMinWidth As Boolean
    Dim hasMaxWidth As Boolean
    Dim currentWidth As Double
    Dim overflowValue As String
    Dim columnRange As Range
    Dim scopedRange As Range
    Dim autoHeightEnabled As Boolean
    Dim autoFitColumnsEnabled As Boolean
    Dim rowHeightValue As Double
    Dim boolValue As Boolean
    Dim fontSizeValue As Double
    Dim colorValue As Long
    Dim horizontalValue As String
    Dim verticalValue As String

    If ws Is Nothing Then Exit Sub
    If parsedStyles Is Nothing Then Exit Sub
    If outCol <= 0 Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    Set columnRange = ws.Columns(outCol)
    Set scopedRange = ws.Range(ws.Cells(rowStart, outCol), ws.Cells(rowEnd, outCol))

    If parsedStyles.Exists(STYLE_PROPERTY_WIDTH) Then
        If Not mp_TryParseWidth(CStr(parsedStyles(STYLE_PROPERTY_WIDTH)), widthUnits) Then
            Err.Raise vbObjectError + 1492, "ex_ConfigStylesParser", _
                "Invalid parsed width value for key '" & mapKey & "'."
        End If
        columnRange.ColumnWidth = widthUnits
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_MIN_WIDTH)) Then
        If Not mp_TryParseWidth(CStr(parsedStyles(LCase$(STYLE_PROPERTY_MIN_WIDTH))), minWidthUnits) Then
            Err.Raise vbObjectError + 1499, "ex_ConfigStylesParser", _
                "Invalid minWidth value for key '" & mapKey & "'."
        End If
        hasMinWidth = True
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_MAX_WIDTH)) Then
        If Not mp_TryParseWidth(CStr(parsedStyles(LCase$(STYLE_PROPERTY_MAX_WIDTH))), maxWidthUnits) Then
            Err.Raise vbObjectError + 1499, "ex_ConfigStylesParser", _
                "Invalid maxWidth value for key '" & mapKey & "'."
        End If
        hasMaxWidth = True
    End If

    If hasMinWidth Or hasMaxWidth Then
        currentWidth = columnRange.ColumnWidth
        If hasMinWidth Then
            If currentWidth < minWidthUnits Then
                columnRange.ColumnWidth = minWidthUnits
                currentWidth = minWidthUnits
            End If
        End If
        If hasMaxWidth Then
            If currentWidth > maxWidthUnits Then
                columnRange.ColumnWidth = maxWidthUnits
            End If
        End If
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_AUTO_FIT_COLUMNS)) Then
        If Not mp_TryParseBoolean(CStr(parsedStyles(LCase$(STYLE_PROPERTY_AUTO_FIT_COLUMNS))), autoFitColumnsEnabled) Then
            Err.Raise vbObjectError + 1499, "ex_ConfigStylesParser", _
                "Invalid autoFitColumns value for key '" & mapKey & "'."
        End If
        If autoFitColumnsEnabled Then columnRange.AutoFit
    End If

    If parsedStyles.Exists(STYLE_PROPERTY_OVERFLOW) Then
        overflowValue = LCase$(Trim$(CStr(parsedStyles(STYLE_PROPERTY_OVERFLOW))))
        Select Case overflowValue
            Case "wrap"
                scopedRange.WrapText = True
                scopedRange.ShrinkToFit = False
            Case "shrink"
                scopedRange.WrapText = False
                scopedRange.ShrinkToFit = True
            Case "clip"
                scopedRange.WrapText = False
                scopedRange.ShrinkToFit = False
            Case Else
                Err.Raise vbObjectError + 1493, "ex_ConfigStylesParser", _
                    "Unsupported overflow value for key '" & mapKey & "': " & overflowValue
        End Select
    End If

    If parsedStyles.Exists(STYLE_PROPERTY_AUTO_HEIGHT) Then
        If Not mp_TryParseBoolean(CStr(parsedStyles(STYLE_PROPERTY_AUTO_HEIGHT)), autoHeightEnabled) Then
            Err.Raise vbObjectError + 1494, "ex_ConfigStylesParser", _
                "Invalid autoHeight value for key '" & mapKey & "'."
        End If
        If autoHeightEnabled Then
            scopedRange.EntireRow.AutoFit
        End If
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_ROW_HEIGHT)) Then
        If Not mp_TryParsePositiveDouble(CStr(parsedStyles(LCase$(STYLE_PROPERTY_ROW_HEIGHT))), rowHeightValue) Then
            Err.Raise vbObjectError + 1499, "ex_ConfigStylesParser", _
                "Invalid rowHeight value for key '" & mapKey & "'."
        End If
        scopedRange.RowHeight = rowHeightValue
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_FONT_NAME)) Then
        scopedRange.Font.Name = CStr(parsedStyles(LCase$(STYLE_PROPERTY_FONT_NAME)))
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_FONT_SIZE)) Then
        If Not mp_TryParsePositiveDouble(CStr(parsedStyles(LCase$(STYLE_PROPERTY_FONT_SIZE))), fontSizeValue) Then
            Err.Raise vbObjectError + 1495, "ex_ConfigStylesParser", _
                "Invalid fontSize value for key '" & mapKey & "'."
        End If
        scopedRange.Font.Size = fontSizeValue
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_FONT_BOLD)) Then
        If Not mp_TryParseBoolean(CStr(parsedStyles(LCase$(STYLE_PROPERTY_FONT_BOLD))), boolValue) Then
            Err.Raise vbObjectError + 1496, "ex_ConfigStylesParser", _
                "Invalid fontBold value for key '" & mapKey & "'."
        End If
        scopedRange.Font.Bold = boolValue
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_BACK_COLOR)) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(parsedStyles(LCase$(STYLE_PROPERTY_BACK_COLOR))), colorValue) Then
            Err.Raise vbObjectError + 1497, "ex_ConfigStylesParser", _
                "Invalid backColor value for key '" & mapKey & "'."
        End If
        scopedRange.Interior.Color = colorValue
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_FONT_COLOR)) Then
        If Not ex_XmlCore.m_TryParseColor(CStr(parsedStyles(LCase$(STYLE_PROPERTY_FONT_COLOR))), colorValue) Then
            Err.Raise vbObjectError + 1498, "ex_ConfigStylesParser", _
                "Invalid fontColor value for key '" & mapKey & "'."
        End If
        scopedRange.Font.Color = colorValue
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_HORIZONTAL)) Then
        horizontalValue = LCase$(Trim$(CStr(parsedStyles(LCase$(STYLE_PROPERTY_HORIZONTAL)))))
        scopedRange.HorizontalAlignment = mp_ParseHorizontalAlignment(horizontalValue)
    End If

    If parsedStyles.Exists(LCase$(STYLE_PROPERTY_VERTICAL)) Then
        verticalValue = LCase$(Trim$(CStr(parsedStyles(LCase$(STYLE_PROPERTY_VERTICAL)))))
        scopedRange.VerticalAlignment = mp_ParseVerticalAlignment(verticalValue)
    End If
End Sub

Private Function mp_ParseHorizontalAlignment(ByVal valueText As String) As XlHAlign
    Select Case LCase$(Trim$(valueText))
        Case "left": mp_ParseHorizontalAlignment = xlLeft
        Case "center": mp_ParseHorizontalAlignment = xlCenter
        Case "right": mp_ParseHorizontalAlignment = xlRight
        Case "fill": mp_ParseHorizontalAlignment = xlFill
        Case "justify": mp_ParseHorizontalAlignment = xlJustify
        Case "distributed": mp_ParseHorizontalAlignment = xlDistributed
        Case Else: mp_ParseHorizontalAlignment = xlGeneral
    End Select
End Function

Private Function mp_ParseVerticalAlignment(ByVal valueText As String) As XlVAlign
    Select Case LCase$(Trim$(valueText))
        Case "top": mp_ParseVerticalAlignment = xlTop
        Case "center": mp_ParseVerticalAlignment = xlCenter
        Case "bottom": mp_ParseVerticalAlignment = xlBottom
        Case "justify": mp_ParseVerticalAlignment = xlJustify
        Case "distributed": mp_ParseVerticalAlignment = xlDistributed
        Case Else: mp_ParseVerticalAlignment = xlCenter
    End Select
End Function

Private Function mp_TryParseBoolean(ByVal valueText As String, ByRef outValue As Boolean) As Boolean
    Dim normalized As String

    normalized = LCase$(Trim$(valueText))
    If Len(normalized) = 0 Then Exit Function

    Select Case normalized
        Case "true", "1", "yes", "on"
            outValue = True
            mp_TryParseBoolean = True
        Case "false", "0", "no", "off"
            outValue = False
            mp_TryParseBoolean = True
    End Select
End Function

Private Function mp_TryParsePositiveDouble(ByVal valueText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseDouble(normalized, outValue) Then Exit Function
    If outValue <= 0 Then Exit Function
    mp_TryParsePositiveDouble = True
End Function

Private Function mp_UnquoteText(ByVal valueText As String) As String
    Dim normalized As String
    normalized = mp_NormalizeStyleToken(valueText)
    If Len(normalized) >= 2 Then
        If (Left$(normalized, 1) = "'" And Right$(normalized, 1) = "'") _
            Or (Left$(normalized, 1) = Chr$(34) And Right$(normalized, 1) = Chr$(34)) Then
            normalized = Mid$(normalized, 2, Len(normalized) - 2)
        End If
    End If
    mp_UnquoteText = mp_NormalizeStyleToken(normalized)
End Function

Private Function mp_NormalizeStyleToken(ByVal valueText As String) As String
    Dim normalized As String

    normalized = CStr(valueText)
    normalized = Replace(normalized, vbCr, " ")
    normalized = Replace(normalized, vbLf, " ")
    normalized = Replace(normalized, vbTab, " ")
    normalized = Replace(normalized, Chr$(160), " ")
    normalized = Replace(normalized, ChrW$(160), " ")

    mp_NormalizeStyleToken = Trim$(normalized)
End Function
