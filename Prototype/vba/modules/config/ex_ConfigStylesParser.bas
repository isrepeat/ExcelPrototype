Attribute VB_Name = "ex_ConfigStylesParser"
Option Explicit

Private Enum e_ConfigStylePropertyId
    cfgStylePropUnknown = 0
    cfgStylePropWidth = 1
    cfgStylePropOverflow = 2
    cfgStylePropAutoHeight = 3
End Enum

Private Const STYLE_PROPERTY_WIDTH As String = "width"
Private Const STYLE_PROPERTY_OVERFLOW As String = "overflow"
Private Const STYLE_PROPERTY_AUTO_HEIGHT As String = "autoHeight"

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

    sourceText = Trim$(noteText)
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

    blockText = Trim$(Mid$(sourceText, openPos + 1, closePos - openPos - 1))
    If Len(blockText) = 0 Then
        outErrorText = "style block is empty"
        Exit Function
    End If

    Set propertyIds = mp_BuildSupportedPropertyIds()
    Set discreteValidators = mp_BuildDiscreteValidators()

    pairs = Split(blockText, ";")
    For i = LBound(pairs) To UBound(pairs)
        pairText = Trim$(CStr(pairs(i)))
        If Len(pairText) = 0 Then GoTo ContinuePair

        colonPos = InStr(1, pairText, ":", vbBinaryCompare)
        If colonPos <= 1 Then
            outErrorText = "invalid style token '" & pairText & "'"
            Exit Function
        End If

        propertyName = LCase$(Trim$(Left$(pairText, colonPos - 1)))
        propertyValue = Trim$(Mid$(pairText, colonPos + 1))

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
    propertyIds(STYLE_PROPERTY_OVERFLOW) = cfgStylePropOverflow
    propertyIds(STYLE_PROPERTY_AUTO_HEIGHT) = cfgStylePropAutoHeight

    Set mp_BuildSupportedPropertyIds = propertyIds
End Function

Private Function mp_BuildDiscreteValidators() As Object
    Dim validators As Object
    Dim overflowAllowed As Object

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

    Select Case propertyId
        Case cfgStylePropWidth
            If Not mp_TryParseWidth(propertyValue, widthValue) Then
                outErrorText = "invalid width value '" & propertyValue & "' (expected positive number)"
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
    Dim overflowValue As String
    Dim columnRange As Range
    Dim scopedRange As Range
    Dim autoHeightEnabled As Boolean

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
End Sub

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
