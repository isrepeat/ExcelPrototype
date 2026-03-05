Attribute VB_Name = "ex_OutputPanel"
Option Explicit

Private Const PANEL_INPUT_NAME As String = "outPanelInputCell"
Private Const PANEL_INPUT_PREFIX As String = "outPanelInput_"
Private Const PANEL_BUTTON_PREFIX As String = "btnOutPanelSearch_"
Private Const PANEL_RANGE_NAME As String = "outPanelRange"
Private Const PANEL_ONCHANGE_BINDING_PREFIX As String = "chg::"

Private g_OnChangeBindings As Object
Public Sub m_RenderForSheet(ByVal ws As Worksheet, ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle)
    Dim topRow As Long
    Dim startCol As Long
    Dim rightCol As Long
    Dim bottomRow As Long
    Dim dataLastCol As Long
    Dim titleRange As Range
    Dim labelRange As Range
    Dim inputRange As Range
    Dim inputAnchor As Range
    Dim buttonShape As Shape
    Dim buttonName As String
    Dim buttonStartCol As Long
    Dim inputStartCol As Long
    Dim inputEndCol As Long
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    Dim buttonLeft As Double
    Dim buttonTop As Double
    Dim titleEndCol As Long
    Dim currentValue As String
    Dim panelWidth As Long
    Dim fieldsTopRow As Long
    Dim rowSpan As Long
    Dim fieldSpacing As Long
    Dim fieldIndex As Long
    Dim fieldTopRow As Long
    Dim fieldBottomRow As Long
    Dim buttonAnchorCol As Long
    Dim buttonWidthCols As Long
    Dim buttonCellRange As Range
    Dim panelAutoFitLastCol As Long
    Dim panelAutoFitCols As Long
    Dim panelRange As Range
    Dim panelRenderRightCol As Long
    Dim fieldError As String
    Dim buttonIndex As Long
    Dim buttonsCount As Long
    Dim buttonRowTop As Long
    Dim fieldRow As Long
    Dim buttonBackColor As Long
    Dim buttonTextColor As Long
    Dim buttonBorderColor As Long

    If ws Is Nothing Then Exit Sub

    mp_ClearPanelArtifacts ws
    If Not style.HasControlPanel Then Exit Sub

    topRow = style.PanelTopRow
    If topRow < 1 Then topRow = 1

    dataLastCol = mp_GetLastUsedColumn(ws)
    startCol = style.PanelStartColumn
    If startCol <= 0 Then
        startCol = dataLastCol + style.PanelOffsetColumns
        If startCol < style.PanelMinStartColumn Then startCol = style.PanelMinStartColumn
    End If
    If startCol < 1 Then startCol = 1

    If style.PanelFieldCount <= 0 Then Exit Sub

    rowSpan = style.PanelFieldRowSpan
    If rowSpan < 1 Then rowSpan = 2
    fieldSpacing = style.PanelFieldSpacingRows
    If fieldSpacing < 0 Then fieldSpacing = 0

    panelWidth = style.PanelLabelColumns + style.PanelValueColumns + 1
    If style.PanelWidthColumns > panelWidth Then panelWidth = style.PanelWidthColumns
    rightCol = startCol + panelWidth - 1

    fieldsTopRow = topRow + 1
    bottomRow = mp_GetControlPanelBottomRow(style, fieldsTopRow, rowSpan, fieldSpacing)

    inputStartCol = startCol + style.PanelLabelColumns
    inputEndCol = inputStartCol + style.PanelValueColumns - 1
    If inputEndCol < inputStartCol Then inputEndCol = inputStartCol
    titleEndCol = inputEndCol
    If titleEndCol < startCol Then titleEndCol = startCol
    buttonStartCol = inputEndCol + 1
    If buttonStartCol > rightCol Then buttonStartCol = rightCol

    buttonAnchorCol = style.PanelButtonAnchorColumn
    If buttonAnchorCol < 1 Then buttonAnchorCol = 4
    buttonWidthCols = style.PanelButtonWidthColumns
    If buttonWidthCols < 1 Then buttonWidthCols = 1

    panelRenderRightCol = inputEndCol
    If titleEndCol > panelRenderRightCol Then panelRenderRightCol = titleEndCol
    If (buttonAnchorCol + buttonWidthCols - 1) > panelRenderRightCol Then
        panelRenderRightCol = buttonAnchorCol + buttonWidthCols - 1
    End If
    If panelRenderRightCol < startCol Then panelRenderRightCol = startCol

    Set panelRange = ws.Range(ws.Cells(topRow, startCol), ws.Cells(bottomRow, panelRenderRightCol))
    panelRange.Interior.Pattern = xlSolid
    panelRange.Interior.Color = style.PanelBackColor

    ws.Columns(startCol).Resize(, style.PanelLabelColumns + style.PanelValueColumns).ColumnWidth = style.PanelColumnWidth

    Set titleRange = ws.Range(ws.Cells(topRow, startCol), ws.Cells(topRow, titleEndCol))
    titleRange.UnMerge
    titleRange.Merge
    titleRange.Value = style.PanelTitle
    titleRange.Font.Bold = True
    titleRange.Font.Color = style.PanelTitleColor
    titleRange.HorizontalAlignment = xlCenter
    titleRange.VerticalAlignment = xlCenter

    mp_DeletePanelButtons ws

    panelAutoFitLastCol = inputStartCol
    If (buttonAnchorCol + buttonWidthCols - 1) > panelAutoFitLastCol Then
        panelAutoFitLastCol = buttonAnchorCol + buttonWidthCols - 1
    End If
    If panelAutoFitLastCol >= startCol Then
        panelAutoFitCols = panelAutoFitLastCol - startCol + 1
        ws.Columns(startCol).Resize(, panelAutoFitCols).AutoFit
    End If

    For fieldIndex = 1 To style.PanelFieldCount
        fieldTopRow = mp_GetFieldTopRow(style, fieldsTopRow, rowSpan, fieldSpacing, fieldIndex)
        fieldBottomRow = fieldTopRow + mp_GetFieldEffectiveRowSpan(style, fieldIndex, rowSpan) - 1

        For fieldRow = fieldTopRow To fieldBottomRow
            ws.Rows(fieldRow).RowHeight = 32
        Next fieldRow

        Set labelRange = ws.Cells(fieldTopRow, startCol)
        labelRange.UnMerge
        labelRange.Value = style.PanelFields(fieldIndex).Label
        labelRange.Font.Bold = True
        labelRange.Font.Color = style.PanelLabelColor
        labelRange.HorizontalAlignment = xlCenter
        labelRange.VerticalAlignment = xlCenter

        Set inputRange = ws.Cells(fieldTopRow, inputStartCol)
        inputRange.UnMerge
        inputRange.Interior.Pattern = xlSolid
        inputRange.Interior.Color = style.PanelInputBackColor
        inputRange.Font.Color = style.PanelInputFontColor
        inputRange.HorizontalAlignment = xlCenter
        inputRange.VerticalAlignment = xlCenter
        inputRange.NumberFormat = "@"

        Set inputAnchor = inputRange.Cells(1, 1)
        mp_ApplyInputOverflowStyle inputAnchor, style.PanelFields(fieldIndex).InputOverflowStyle
        currentValue = Trim$(CStr(inputAnchor.Value))
        If Len(currentValue) = 0 Then
            inputAnchor.Value = ex_ConfigProvider.m_GetConfigValue(style.PanelFields(fieldIndex).InputConfigKey, vbNullString)
        End If
        If fieldIndex = 1 Then
            mp_SetPanelInputName ws, inputAnchor
        End If
        mp_SetPanelInputKeyName ws, inputAnchor, style.PanelFields(fieldIndex).InputName
        mp_RegisterOnChangeBinding ws, inputAnchor, style.PanelFields(fieldIndex).OnChangeMacroName

        fieldError = mp_GetConfigRefFieldError(style.PanelFields(fieldIndex), fieldIndex)
        If Len(fieldError) > 0 Then
            mp_RenderFieldInlineError inputAnchor, fieldError, style, style.PanelFields(fieldIndex).InputOverflowStyle
            GoTo ContinueField
        End If

        buttonsCount = style.PanelFields(fieldIndex).ButtonCount
        If buttonsCount < 1 Then buttonsCount = 1

        For buttonIndex = 1 To buttonsCount
            buttonRowTop = fieldTopRow + buttonIndex - 1
            Set buttonCellRange = ws.Range(ws.Cells(buttonRowTop, buttonAnchorCol), ws.Cells(buttonRowTop, buttonAnchorCol + buttonWidthCols - 1))
            buttonTop = buttonCellRange.Top
            buttonLeft = buttonCellRange.Left
            buttonWidth = buttonCellRange.Width
            buttonHeight = buttonCellRange.Height
            If buttonWidth < 8 Then buttonWidth = 8
            If buttonHeight < 8 Then buttonHeight = 8

            buttonName = mp_GetButtonName(ws, fieldIndex, buttonIndex)
            Set buttonShape = ws.Shapes.AddShape(msoShapeRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
            buttonShape.Name = buttonName
            buttonShape.TextFrame.Characters.Text = style.PanelFields(fieldIndex).Buttons(buttonIndex).Caption
            buttonBackColor = mp_GetButtonBackColor(style, fieldIndex, buttonIndex)
            buttonTextColor = mp_GetButtonTextColor(style, fieldIndex, buttonIndex)
            buttonBorderColor = mp_GetButtonBorderColor(style, fieldIndex, buttonIndex)
            buttonShape.Fill.Solid
            buttonShape.Fill.ForeColor.RGB = buttonBackColor
            buttonShape.Fill.Transparency = 0
            buttonShape.Line.ForeColor.RGB = buttonBorderColor
            buttonShape.Line.Weight = 1
            buttonShape.TextFrame.Characters.Font.Bold = True
            buttonShape.TextFrame.Characters.Font.Color = buttonTextColor
            buttonShape.TextFrame.Characters.Font.Name = style.FontName
            buttonShape.TextFrame.Characters.Font.Size = style.FontSize
            buttonShape.TextFrame.HorizontalAlignment = xlHAlignCenter
            buttonShape.TextFrame.VerticalAlignment = xlVAlignCenter
            buttonShape.Placement = xlMoveAndSize
            buttonShape.OnAction = "'" & ThisWorkbook.Name & "'!" & Trim$(style.PanelFields(fieldIndex).Buttons(buttonIndex).MacroName)
        Next buttonIndex

ContinueField:
    Next fieldIndex

    mp_SetPanelRangeName ws, ws.Range(ws.Cells(topRow, startCol), ws.Cells(bottomRow, panelRenderRightCol))

    mp_ApplyFixedControlPanelLayout ws, style, startCol, inputStartCol
End Sub

Public Sub m_HandleSheetInputChange(ByVal ws As Worksheet, ByVal target As Range)
    Dim macroName As String
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean

    If ws Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If target.CountLarge <> 1 Then Exit Sub

    macroName = mp_GetOnChangeMacroName(ws, target)
    If Len(macroName) = 0 Then Exit Sub

    On Error GoTo EH
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Run macroName
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Exit Sub

EH:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    On Error GoTo 0
    MsgBox "Failed to run onChange macro '" & macroName & "': " & Err.Description, vbExclamation
End Sub

Private Function mp_GetConfigRefFieldError( _
    ByRef fieldStyle As ex_SheetStylesXmlProvider.t_ControlPanelFieldStyle, _
    ByVal fieldIndex As Long _
) As String
    Dim buttonIndex As Long

    If Not fieldStyle.IsConfigRefField Then Exit Function

    If Len(Trim$(fieldStyle.Label)) = 0 Then
        mp_GetConfigRefFieldError = "Missing required attribute: inputConfigRefField@label (field " & CStr(fieldIndex) & ")."
        Exit Function
    End If
    If Len(Trim$(fieldStyle.InputConfigKey)) = 0 Then
        mp_GetConfigRefFieldError = "Missing required attribute: inputConfigRefField@inputConfigKey (field " & CStr(fieldIndex) & ")."
        Exit Function
    End If
    If Len(Trim$(fieldStyle.InputName)) = 0 Then
        mp_GetConfigRefFieldError = "Missing required attribute: inputConfigRefField@inputName (field " & CStr(fieldIndex) & ")."
        Exit Function
    End If

    If fieldStyle.ButtonCount <= 0 Then
        mp_GetConfigRefFieldError = "Missing required node: inputConfigRefField/button (field " & CStr(fieldIndex) & ")."
        Exit Function
    End If

    For buttonIndex = 1 To fieldStyle.ButtonCount
        If Len(Trim$(fieldStyle.Buttons(buttonIndex).Caption)) = 0 Then
            mp_GetConfigRefFieldError = "Missing required attribute: inputConfigRefField/button@caption (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ")."
            Exit Function
        End If
        If Len(Trim$(fieldStyle.Buttons(buttonIndex).MacroName)) = 0 Then
            mp_GetConfigRefFieldError = "Missing required attribute: inputConfigRefField/button@macro (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ")."
            Exit Function
        End If
    Next buttonIndex
End Function

Private Sub mp_RenderFieldInlineError( _
    ByVal inputAnchor As Range, _
    ByVal errorText As String, _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal overflowStyle As String _
)
    If inputAnchor Is Nothing Then Exit Sub

    inputAnchor.UnMerge
    inputAnchor.Value = errorText
    inputAnchor.Font.Bold = True
    inputAnchor.Font.Color = style.PanelErrorFontColor
    inputAnchor.Interior.Pattern = xlSolid
    inputAnchor.Interior.Color = style.PanelErrorBackColor
    inputAnchor.HorizontalAlignment = xlLeft
    inputAnchor.VerticalAlignment = xlCenter
    mp_ApplyInputOverflowStyle inputAnchor, overflowStyle
End Sub

Private Sub mp_ApplyInputOverflowStyle(ByVal targetCell As Range, ByVal overflowStyle As String)
    overflowStyle = LCase$(Trim$(overflowStyle))
    If Len(overflowStyle) = 0 Then overflowStyle = "overflow"

    Select Case overflowStyle
        Case "wrap"
            targetCell.WrapText = True
            targetCell.ShrinkToFit = False
        Case "shrink"
            targetCell.WrapText = False
            targetCell.ShrinkToFit = True
        Case Else
            targetCell.WrapText = False
            targetCell.ShrinkToFit = False
    End Select
End Sub

Public Function m_ReadSearchValue(ByVal ws As Worksheet) As String
    Dim inputCell As Range
    Set inputCell = mp_GetPanelInputCell(ws)
    If inputCell Is Nothing Then Exit Function
    m_ReadSearchValue = Trim$(CStr(inputCell.Value))
End Function

Public Function m_ReadFieldValue(ByVal ws As Worksheet, ByVal inputName As String) As String
    Dim inputCell As Range
    Set inputCell = mp_GetPanelInputCellByKey(ws, inputName)
    If inputCell Is Nothing Then Exit Function
    m_ReadFieldValue = Trim$(CStr(inputCell.Value))
End Function

Public Sub m_ApplyFixedWidthViewZoneLayer( _
    ByVal ws As Worksheet, _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal dataLastCol As Long _
)
    Dim panelStartCol As Long
    Dim keyCol As Long
    Dim valueCol As Long
    Dim buttonCol As Long
    Dim maxCol As Long
    Dim hasFixedColumn As Boolean
    Dim fixedFlags() As Boolean
    Dim c As Long
    Dim colRange As Range

    If ws Is Nothing Then Exit Sub
    If Not style.HasControlPanel Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub

    panelStartCol = style.PanelStartColumn
    If panelStartCol <= 0 Then
        panelStartCol = dataLastCol + style.PanelOffsetColumns
        If panelStartCol < style.PanelMinStartColumn Then panelStartCol = style.PanelMinStartColumn
    End If
    If panelStartCol < 1 Then panelStartCol = 1

    keyCol = panelStartCol
    valueCol = panelStartCol + style.PanelLabelColumns
    buttonCol = style.PanelButtonAnchorColumn
    If buttonCol < 1 Then buttonCol = 4

    maxCol = dataLastCol
    If style.PanelFixedWidthKey > 0 And keyCol > maxCol Then maxCol = keyCol
    If style.PanelFixedWidthValue > 0 And valueCol > maxCol Then maxCol = valueCol
    If style.PanelFixedWidthButton > 0 And buttonCol > maxCol Then maxCol = buttonCol
    If maxCol < 1 Then Exit Sub

    ReDim fixedFlags(1 To maxCol)
    If style.PanelFixedWidthKey > 0 And keyCol >= 1 And keyCol <= maxCol Then fixedFlags(keyCol) = True
    If style.PanelFixedWidthValue > 0 And valueCol >= 1 And valueCol <= maxCol Then fixedFlags(valueCol) = True
    If style.PanelFixedWidthButton > 0 And buttonCol >= 1 And buttonCol <= maxCol Then fixedFlags(buttonCol) = True

    For c = 1 To maxCol
        If fixedFlags(c) Then
            hasFixedColumn = True
            Set colRange = ws.Range(ws.Cells(viewStartRow, c), ws.Cells(viewEndRow, c))
            colRange.WrapText = True
        End If
    Next c

    If hasFixedColumn Then
        ws.Rows(CStr(viewStartRow) & ":" & CStr(viewEndRow)).AutoFit
    End If
End Sub

Public Function m_CreateRuntimeLayer( _
    ByVal ws As Worksheet, _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    Optional ByVal layerId As String = "runtime-control-panel", _
    Optional ByVal priority As Long = 800 _
) As obj_StyleLayer
    Dim panelRange As Range
    Dim runtimeLayer As obj_StyleLayer
    Dim declarations As Object
    Dim topRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim titleEndCol As Long
    Dim inputStartCol As Long
    Dim inputEndCol As Long
    Dim buttonCol As Long
    Dim buttonWidthCols As Long
    Dim autoFitLastCol As Long
    Dim rowSpan As Long
    Dim fieldSpacing As Long
    Dim fieldsTopRow As Long
    Dim fieldIndex As Long
    Dim fieldTopRow As Long
    Dim fieldBottomRow As Long
    Dim fixedFieldHeight As Double
    Dim ruleIndex As Long
    Dim fieldError As String
    Dim inputOverflow As String

    If ws Is Nothing Then Exit Function
    If Not style.HasControlPanel Then Exit Function
    If style.PanelFieldCount <= 0 Then Exit Function

    Set panelRange = mp_TryGetPanelRange(ws)
    If panelRange Is Nothing Then Exit Function

    topRow = panelRange.Row
    startCol = panelRange.Column
    endCol = panelRange.Column + panelRange.Columns.Count - 1

    inputStartCol = startCol + style.PanelLabelColumns
    inputEndCol = inputStartCol + style.PanelValueColumns - 1
    If inputEndCol < inputStartCol Then inputEndCol = inputStartCol
    titleEndCol = inputEndCol
    If titleEndCol < startCol Then titleEndCol = startCol
    If titleEndCol > endCol Then titleEndCol = endCol

    buttonCol = style.PanelButtonAnchorColumn
    If buttonCol < 1 Then buttonCol = 4
    buttonWidthCols = style.PanelButtonWidthColumns
    If buttonWidthCols < 1 Then buttonWidthCols = 1
    autoFitLastCol = inputStartCol
    If (buttonCol + buttonWidthCols - 1) > autoFitLastCol Then
        autoFitLastCol = buttonCol + buttonWidthCols - 1
    End If
    If autoFitLastCol < startCol Then autoFitLastCol = startCol

    rowSpan = style.PanelFieldRowSpan
    If rowSpan < 1 Then rowSpan = 1
    fieldSpacing = style.PanelFieldSpacingRows
    If fieldSpacing < 0 Then fieldSpacing = 0
    fieldsTopRow = topRow + 1

    fixedFieldHeight = style.PanelFixedFieldRowHeight
    If fixedFieldHeight <= 0 Then fixedFieldHeight = 32#

    Set runtimeLayer = New obj_StyleLayer
    runtimeLayer.Initialize layerId, priority, "runtime", True

    ruleIndex = ruleIndex + 1
    Set declarations = mp_CreateDeclarations()
    declarations("backColor") = mp_ColorToHex(style.PanelBackColor)
    declarations("borderColor") = mp_ColorToHex(style.PanelBorderColor)
    declarations("borderWeight") = "thin"
    mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "range", mp_BuildAddress(topRow, startCol, panelRange.Row + panelRange.Rows.Count - 1, endCol), declarations

    ruleIndex = ruleIndex + 1
    Set declarations = mp_CreateDeclarations()
    declarations("fontBold") = "true"
    declarations("fontColor") = mp_ColorToHex(style.PanelTitleColor)
    declarations("horizontal") = "center"
    declarations("vertical") = "center"
    mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "range", mp_BuildAddress(topRow, startCol, topRow, titleEndCol), declarations

    If style.PanelColumnWidth > 0 Then
        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("width") = mp_ToInvariantDoubleText(style.PanelColumnWidth)
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "column", mp_BuildColumnAddress(startCol, inputEndCol), declarations
    End If

    ruleIndex = ruleIndex + 1
    Set declarations = mp_CreateDeclarations()
    declarations("autoFitColumns") = "true"
    mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "column", mp_BuildColumnAddress(startCol, autoFitLastCol), declarations

    If style.PanelFixedWidthKey > 0 Then
        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("width") = mp_ToInvariantDoubleText(style.PanelFixedWidthKey)
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "column", mp_BuildColumnAddress(startCol, startCol), declarations
    End If
    If style.PanelFixedWidthValue > 0 Then
        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("width") = mp_ToInvariantDoubleText(style.PanelFixedWidthValue)
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "column", mp_BuildColumnAddress(inputStartCol, inputStartCol), declarations
    End If
    If style.PanelFixedWidthButton > 0 Then
        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("width") = mp_ToInvariantDoubleText(style.PanelFixedWidthButton)
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "column", mp_BuildColumnAddress(buttonCol, buttonCol), declarations
    End If

    For fieldIndex = 1 To style.PanelFieldCount
        fieldTopRow = mp_GetFieldTopRow(style, fieldsTopRow, rowSpan, fieldSpacing, fieldIndex)
        fieldBottomRow = fieldTopRow + mp_GetFieldEffectiveRowSpan(style, fieldIndex, rowSpan) - 1

        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("rowHeight") = mp_ToInvariantDoubleText(fixedFieldHeight)
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "row", "row=" & CStr(fieldTopRow) & ":" & CStr(fieldBottomRow) & ";col=" & CStr(startCol) & ":" & CStr(endCol), declarations

        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("fontBold") = "true"
        declarations("fontColor") = mp_ColorToHex(style.PanelLabelColor)
        declarations("horizontal") = "center"
        declarations("vertical") = "center"
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "cell", mp_BuildAddress(fieldTopRow, startCol, fieldTopRow, startCol), declarations

        fieldError = mp_GetConfigRefFieldError(style.PanelFields(fieldIndex), fieldIndex)
        inputOverflow = mp_NormalizeOverflowText(style.PanelFields(fieldIndex).InputOverflowStyle)
        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        If Len(fieldError) > 0 Then
            declarations("fontBold") = "true"
            declarations("backColor") = mp_ColorToHex(style.PanelErrorBackColor)
            declarations("fontColor") = mp_ColorToHex(style.PanelErrorFontColor)
            declarations("horizontal") = "left"
            declarations("vertical") = "center"
        Else
            declarations("backColor") = mp_ColorToHex(style.PanelInputBackColor)
            declarations("fontColor") = mp_ColorToHex(style.PanelInputFontColor)
            declarations("horizontal") = "center"
            declarations("vertical") = "center"
        End If
        declarations("overflow") = inputOverflow
        mp_AddRuntimeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), "cell", mp_BuildAddress(fieldTopRow, inputStartCol, fieldTopRow, inputStartCol), declarations
    Next fieldIndex

    Set m_CreateRuntimeLayer = runtimeLayer
End Function

Private Function mp_TryGetPanelRange(ByVal ws As Worksheet) As Range
    Dim panelName As Name

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set panelName = ws.Names(PANEL_RANGE_NAME)
    On Error GoTo 0
    If panelName Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_TryGetPanelRange = panelName.RefersToRange
    On Error GoTo 0
End Function

Private Sub mp_AddRuntimeRule( _
    ByVal layer As obj_StyleLayer, _
    ByVal ruleId As String, _
    ByVal targetName As String, _
    ByVal selectorAddress As String, _
    ByVal declarations As Object _
)
    Dim ruleObj As obj_StyleRule
    Dim selectorText As String

    If layer Is Nothing Then Exit Sub
    If declarations Is Nothing Then Exit Sub
    If Len(Trim$(ruleId)) = 0 Then Exit Sub
    If Len(Trim$(targetName)) = 0 Then Exit Sub
    If Len(Trim$(selectorAddress)) = 0 Then Exit Sub

    If StrComp(LCase$(Trim$(targetName)), "column", vbTextCompare) = 0 Or _
       StrComp(LCase$(Trim$(targetName)), "range", vbTextCompare) = 0 Or _
       StrComp(LCase$(Trim$(targetName)), "cell", vbTextCompare) = 0 Then
        selectorText = "address=" & selectorAddress
    Else
        selectorText = selectorAddress
    End If

    Set ruleObj = New obj_StyleRule
    ruleObj.Initialize ruleId, targetName, selectorText, declarations
    layer.AddRule ruleObj
End Sub

Private Function mp_CreateDeclarations() As Object
    Set mp_CreateDeclarations = CreateObject("Scripting.Dictionary")
    mp_CreateDeclarations.CompareMode = 1
End Function

Private Function mp_BuildAddress( _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
) As String
    If rowStart < 1 Then rowStart = 1
    If colStart < 1 Then colStart = 1
    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    mp_BuildAddress = mp_ToColumnLetter(colStart) & CStr(rowStart) & ":" & mp_ToColumnLetter(colEnd) & CStr(rowEnd)
End Function

Private Function mp_BuildColumnAddress(ByVal colStart As Long, ByVal colEnd As Long) As String
    If colStart < 1 Then colStart = 1
    If colEnd < colStart Then colEnd = colStart
    mp_BuildColumnAddress = mp_ToColumnLetter(colStart) & ":" & mp_ToColumnLetter(colEnd)
End Function

Private Function mp_ToColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long

    If columnIndex < 1 Then columnIndex = 1
    n = columnIndex
    Do While n > 0
        remainder = (n - 1) Mod 26
        mp_ToColumnLetter = Chr$(65 + remainder) & mp_ToColumnLetter
        n = (n - remainder - 1) \ 26
    Loop
End Function

Private Function mp_ColorToHex(ByVal colorValue As Long) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = colorValue Mod 256
    g = (colorValue \ 256) Mod 256
    b = (colorValue \ 65536) Mod 256
    mp_ColorToHex = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

Private Function mp_ToInvariantDoubleText(ByVal value As Double) As String
    mp_ToInvariantDoubleText = Replace$(Trim$(CStr(value)), ",", ".")
End Function

Private Function mp_NormalizeOverflowText(ByVal overflowStyle As String) As String
    overflowStyle = LCase$(Trim$(overflowStyle))
    Select Case overflowStyle
        Case "wrap", "shrink", "clip"
            mp_NormalizeOverflowText = overflowStyle
        Case Else
            mp_NormalizeOverflowText = "clip"
    End Select
End Function

Public Function m_TryGetClickedFieldIndex(ByVal ws As Worksheet, ByVal callerName As String, ByRef fieldIndex As Long) As Boolean
    Dim prefix As String
    Dim suffix As String
    Dim sepPos As Long
    Dim fieldToken As String

    If ws Is Nothing Then Exit Function
    callerName = Trim$(callerName)
    If Len(callerName) = 0 Then Exit Function

    prefix = PANEL_BUTTON_PREFIX & ws.CodeName & "_"
    If LCase$(Left$(callerName, Len(prefix))) <> LCase$(prefix) Then Exit Function

    suffix = Mid$(callerName, Len(prefix) + 1)
    sepPos = InStr(1, suffix, "_", vbTextCompare)
    If sepPos > 1 Then
        fieldToken = Left$(suffix, sepPos - 1)
    Else
        fieldToken = suffix
    End If

    If Not ex_XmlCore.m_TryParseLong(fieldToken, fieldIndex) Then Exit Function
    If fieldIndex < 1 Then Exit Function
    m_TryGetClickedFieldIndex = True
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then
        mp_GetLastUsedColumn = 1
        Exit Function
    End If

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        mp_GetLastUsedColumn = 1
    Else
        mp_GetLastUsedColumn = lastCell.Column
    End If
End Function

Private Sub mp_SetPanelInputName(ByVal ws As Worksheet, ByVal inputCell As Range)
    If ws Is Nothing Then Exit Sub
    If inputCell Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Names(PANEL_INPUT_NAME).Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Names.Add Name:=PANEL_INPUT_NAME, RefersTo:="=" & inputCell.Address(True, True, xlA1, True)
    On Error GoTo 0
End Sub

Private Sub mp_SetPanelInputKeyName(ByVal ws As Worksheet, ByVal inputCell As Range, ByVal inputName As String)
    Dim namedKey As String

    If ws Is Nothing Then Exit Sub
    If inputCell Is Nothing Then Exit Sub

    namedKey = mp_GetInputNameByKey(inputName)
    If Len(namedKey) = 0 Then Exit Sub

    On Error Resume Next
    ws.Names(namedKey).Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Names.Add Name:=namedKey, RefersTo:="=" & inputCell.Address(True, True, xlA1, True)
    On Error GoTo 0
End Sub

Private Function mp_GetPanelInputCell(ByVal ws As Worksheet) As Range
    Dim inputName As Name
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set inputName = ws.Names(PANEL_INPUT_NAME)
    On Error GoTo 0
    If inputName Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetPanelInputCell = inputName.RefersToRange
    On Error GoTo 0
End Function

Private Function mp_GetPanelInputCellByKey(ByVal ws As Worksheet, ByVal inputKey As String) As Range
    Dim inputName As Name
    Dim namedKey As String

    If ws Is Nothing Then Exit Function
    namedKey = mp_GetInputNameByKey(inputKey)
    If Len(namedKey) = 0 Then Exit Function

    On Error Resume Next
    Set inputName = ws.Names(namedKey)
    On Error GoTo 0
    If inputName Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetPanelInputCellByKey = inputName.RefersToRange
    On Error GoTo 0
End Function

Private Function mp_GetButtonName(ByVal ws As Worksheet, ByVal fieldIndex As Long, Optional ByVal buttonIndex As Long = 1) As String
    If ws Is Nothing Then
        mp_GetButtonName = PANEL_BUTTON_PREFIX
        Exit Function
    End If
    mp_GetButtonName = PANEL_BUTTON_PREFIX & ws.CodeName & "_" & CStr(fieldIndex) & "_" & CStr(buttonIndex)
End Function

Private Sub mp_ClearPanelArtifacts(ByVal ws As Worksheet)
    mp_ClearStoredPanelRange ws
    mp_DeletePanelButtons ws
    mp_DeletePanelInputNames ws
    mp_ClearOnChangeBindings ws
End Sub

Private Function mp_GetOnChangeMacroName(ByVal ws As Worksheet, ByVal target As Range) As String
    Dim registry As Object
    Dim key As String

    If ws Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function
    If g_OnChangeBindings Is Nothing Then Exit Function

    Set registry = mp_GetSheetOnChangeRegistry(ws, False)
    If registry Is Nothing Then Exit Function

    key = mp_GetOnChangeBindingKey(target)
    If Len(key) = 0 Then Exit Function
    If registry.Exists(key) Then
        mp_GetOnChangeMacroName = CStr(registry(key))
    End If
End Function

Private Sub mp_RegisterOnChangeBinding(ByVal ws As Worksheet, ByVal inputCell As Range, ByVal macroName As String)
    Dim registry As Object
    Dim key As String

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If inputCell Is Nothing Then Exit Sub

    Set registry = mp_GetSheetOnChangeRegistry(ws, True)
    If registry Is Nothing Then Exit Sub

    key = mp_GetOnChangeBindingKey(inputCell)
    If Len(key) = 0 Then Exit Sub
    registry(key) = macroName
End Sub

Private Sub mp_ClearOnChangeBindings(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If g_OnChangeBindings Is Nothing Then Exit Sub
    If g_OnChangeBindings.Exists(mp_GetSheetBindingKey(ws)) Then
        g_OnChangeBindings.Remove mp_GetSheetBindingKey(ws)
    End If
End Sub

Private Function mp_GetSheetOnChangeRegistry(ByVal ws As Worksheet, ByVal createIfMissing As Boolean) As Object
    Dim mapKey As String
    Dim registry As Object

    If ws Is Nothing Then Exit Function
    mapKey = mp_GetSheetBindingKey(ws)
    If Len(mapKey) = 0 Then Exit Function

    If g_OnChangeBindings Is Nothing Then
        If Not createIfMissing Then Exit Function
        Set g_OnChangeBindings = CreateObject("Scripting.Dictionary")
        g_OnChangeBindings.CompareMode = 1
    End If

    If g_OnChangeBindings.Exists(mapKey) Then
        Set mp_GetSheetOnChangeRegistry = g_OnChangeBindings(mapKey)
        Exit Function
    End If

    If Not createIfMissing Then Exit Function

    Set registry = CreateObject("Scripting.Dictionary")
    registry.CompareMode = 1
    g_OnChangeBindings.Add mapKey, registry
    Set mp_GetSheetOnChangeRegistry = registry
End Function

Private Function mp_GetSheetBindingKey(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Exit Function
    mp_GetSheetBindingKey = PANEL_ONCHANGE_BINDING_PREFIX & LCase$(Trim$(ws.CodeName))
End Function

Private Function mp_GetOnChangeBindingKey(ByVal target As Range) As String
    If target Is Nothing Then Exit Function
    mp_GetOnChangeBindingKey = LCase$(target.Address(False, False, xlA1))
End Function

Private Sub mp_SetPanelRangeName(ByVal ws As Worksheet, ByVal panelRange As Range)
    If ws Is Nothing Then Exit Sub
    If panelRange Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Names(PANEL_RANGE_NAME).Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Names.Add Name:=PANEL_RANGE_NAME, RefersTo:="=" & panelRange.Address(True, True, xlA1, True)
    On Error GoTo 0
End Sub

Private Sub mp_ClearStoredPanelRange(ByVal ws As Worksheet)
    Dim panelName As Name
    Dim panelRange As Range

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set panelName = ws.Names(PANEL_RANGE_NAME)
    On Error GoTo 0

    If panelName Is Nothing Then Exit Sub

    On Error Resume Next
    Set panelRange = panelName.RefersToRange
    On Error GoTo 0

    If Not panelRange Is Nothing Then
        On Error Resume Next
        panelRange.UnMerge
        panelRange.ClearContents
        On Error GoTo 0
    End If

    On Error Resume Next
    panelName.Delete
    On Error GoTo 0
End Sub

Private Sub mp_DeletePanelInputNames(ByVal ws As Worksheet)
    Dim i As Long
    Dim localName As String

    If ws Is Nothing Then Exit Sub

    For i = ws.Names.Count To 1 Step -1
        localName = mp_GetLocalNameToken(ws.Names(i).Name)
        If LCase$(localName) = LCase$(PANEL_INPUT_NAME) Or _
           LCase$(Left$(localName, Len(PANEL_INPUT_PREFIX))) = LCase$(PANEL_INPUT_PREFIX) Then
            On Error Resume Next
            ws.Names(i).Delete
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function mp_GetLocalNameToken(ByVal fullName As String) As String
    Dim bangPos As Long

    bangPos = InStrRev(fullName, "!")
    If bangPos > 0 Then
        mp_GetLocalNameToken = Mid$(fullName, bangPos + 1)
    Else
        mp_GetLocalNameToken = fullName
    End If
End Function

Private Sub mp_DeletePanelButtons(ByVal ws As Worksheet)
    Dim shapeIndex As Long
    Dim shp As Shape
    Dim prefix As String
    Dim commonPrefix As String
    Dim backPrefix As String

    If ws Is Nothing Then Exit Sub
    prefix = PANEL_BUTTON_PREFIX & ws.CodeName & "_"
    commonPrefix = PANEL_BUTTON_PREFIX
    backPrefix = "btnOutPanelBackToDev_" & ws.CodeName

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        If LCase$(Left$(shp.Name, Len(prefix))) = LCase$(prefix) _
           Or LCase$(Left$(shp.Name, Len(commonPrefix))) = LCase$(commonPrefix) _
           Or LCase$(Left$(shp.Name, Len(backPrefix))) = LCase$(backPrefix) Then
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
        End If
    Next shapeIndex
End Sub

Private Sub mp_ApplyFixedControlPanelLayout( _
    ByVal ws As Worksheet, _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal keyCol As Long, _
    ByVal valueCol As Long _
)
    Dim buttonCol As Long
    Dim fieldsTopRow As Long
    Dim rowSpan As Long
    Dim spacingRows As Long
    Dim fieldIndex As Long
    Dim fieldTopRow As Long
    Dim fieldBottomRow As Long
    Dim rowIndex As Long

    If ws Is Nothing Then Exit Sub

    buttonCol = style.PanelButtonAnchorColumn
    If buttonCol < 1 Then buttonCol = 4

    If style.PanelFixedWidthKey > 0 Then
        ws.Columns(keyCol).ColumnWidth = style.PanelFixedWidthKey
    End If
    If style.PanelFixedWidthValue > 0 Then
        ws.Columns(valueCol).ColumnWidth = style.PanelFixedWidthValue
    End If
    If style.PanelFixedWidthButton > 0 Then
        ws.Columns(buttonCol).ColumnWidth = style.PanelFixedWidthButton
    End If

    If style.PanelFixedFieldRowHeight <= 0 Then Exit Sub
    If style.PanelFieldCount <= 0 Then Exit Sub

    rowSpan = style.PanelFieldRowSpan
    If rowSpan < 1 Then rowSpan = 1
    spacingRows = style.PanelFieldSpacingRows
    If spacingRows < 0 Then spacingRows = 0
    fieldsTopRow = style.PanelTopRow + 1

    For fieldIndex = 1 To style.PanelFieldCount
        fieldTopRow = mp_GetFieldTopRow(style, fieldsTopRow, rowSpan, spacingRows, fieldIndex)
        fieldBottomRow = fieldTopRow + mp_GetFieldEffectiveRowSpan(style, fieldIndex, rowSpan) - 1
        For rowIndex = fieldTopRow To fieldBottomRow
            ws.Rows(rowIndex).RowHeight = style.PanelFixedFieldRowHeight
        Next rowIndex
    Next fieldIndex
End Sub

Private Function mp_GetControlPanelBottomRow( _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal fieldsTopRow As Long, _
    ByVal baseRowSpan As Long, _
    ByVal spacingRows As Long _
) As Long
    Dim totalRows As Long
    Dim fieldIndex As Long

    If style.PanelFieldCount <= 0 Then
        mp_GetControlPanelBottomRow = fieldsTopRow - 1
        Exit Function
    End If

    For fieldIndex = 1 To style.PanelFieldCount
        totalRows = totalRows + mp_GetFieldEffectiveRowSpan(style, fieldIndex, baseRowSpan)
        If fieldIndex < style.PanelFieldCount Then
            totalRows = totalRows + spacingRows
        End If
    Next fieldIndex

    mp_GetControlPanelBottomRow = fieldsTopRow + totalRows - 1
End Function

Private Function mp_GetFieldTopRow( _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal fieldsTopRow As Long, _
    ByVal baseRowSpan As Long, _
    ByVal spacingRows As Long, _
    ByVal fieldIndex As Long _
) As Long
    Dim i As Long
    Dim topRow As Long

    topRow = fieldsTopRow
    For i = 1 To fieldIndex - 1
        topRow = topRow + mp_GetFieldEffectiveRowSpan(style, i, baseRowSpan) + spacingRows
    Next i

    mp_GetFieldTopRow = topRow
End Function

Private Function mp_GetFieldEffectiveRowSpan( _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal fieldIndex As Long, _
    ByVal baseRowSpan As Long _
) As Long
    Dim effectiveSpan As Long

    effectiveSpan = baseRowSpan
    If effectiveSpan < 1 Then effectiveSpan = 1

    If fieldIndex >= 1 And fieldIndex <= style.PanelFieldCount Then
        If style.PanelFields(fieldIndex).ButtonCount > effectiveSpan Then
            effectiveSpan = style.PanelFields(fieldIndex).ButtonCount
        End If
    End If

    mp_GetFieldEffectiveRowSpan = effectiveSpan
End Function

Private Function mp_GetButtonBackColor( _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal fieldIndex As Long, _
    ByVal buttonIndex As Long _
) As Long
    mp_GetButtonBackColor = style.PanelButtonBackColor
    If fieldIndex < 1 Or fieldIndex > style.PanelFieldCount Then Exit Function
    If buttonIndex < 1 Or buttonIndex > style.PanelFields(fieldIndex).ButtonCount Then Exit Function
    If style.PanelFields(fieldIndex).Buttons(buttonIndex).HasBackColor Then
        mp_GetButtonBackColor = style.PanelFields(fieldIndex).Buttons(buttonIndex).BackColor
    End If
End Function

Private Function mp_GetButtonTextColor( _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal fieldIndex As Long, _
    ByVal buttonIndex As Long _
) As Long
    mp_GetButtonTextColor = style.PanelButtonTextColor
    If fieldIndex < 1 Or fieldIndex > style.PanelFieldCount Then Exit Function
    If buttonIndex < 1 Or buttonIndex > style.PanelFields(fieldIndex).ButtonCount Then Exit Function
    If style.PanelFields(fieldIndex).Buttons(buttonIndex).HasTextColor Then
        mp_GetButtonTextColor = style.PanelFields(fieldIndex).Buttons(buttonIndex).TextColor
    End If
End Function

Private Function mp_GetButtonBorderColor( _
    ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle, _
    ByVal fieldIndex As Long, _
    ByVal buttonIndex As Long _
) As Long
    mp_GetButtonBorderColor = style.PanelButtonBorderColor
    If fieldIndex < 1 Or fieldIndex > style.PanelFieldCount Then Exit Function
    If buttonIndex < 1 Or buttonIndex > style.PanelFields(fieldIndex).ButtonCount Then Exit Function
    If style.PanelFields(fieldIndex).Buttons(buttonIndex).HasBorderColor Then
        mp_GetButtonBorderColor = style.PanelFields(fieldIndex).Buttons(buttonIndex).BorderColor
    End If
End Function

Private Function mp_GetInputNameByKey(ByVal inputKey As String) As String
    Dim normalized As String
    normalized = mp_NormalizeNameToken(inputKey)
    If Len(normalized) = 0 Then Exit Function
    mp_GetInputNameByKey = PANEL_INPUT_PREFIX & normalized
End Function

Private Function mp_NormalizeNameToken(ByVal rawText As String) As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then Exit Function

    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            outText = outText & ch
        Else
            outText = outText & "_"
        End If
    Next i

    If Len(outText) = 0 Then Exit Function
    If Mid$(outText, 1, 1) >= "0" And Mid$(outText, 1, 1) <= "9" Then
        outText = "_" & outText
    End If
    mp_NormalizeNameToken = outText
End Function
