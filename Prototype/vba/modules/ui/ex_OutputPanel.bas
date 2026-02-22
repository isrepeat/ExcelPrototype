Attribute VB_Name = "ex_OutputPanel"
Option Explicit

Private Const PANEL_INPUT_NAME As String = "outPanelInputCell"
Private Const PANEL_INPUT_PREFIX As String = "outPanelInput_"
Private Const PANEL_BUTTON_PREFIX As String = "btnOutPanelSearch_"
Private Const PANEL_RANGE_NAME As String = "outPanelRange"
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
    Dim buttonWidthPoints As Double
    Dim anchorColumnWidthPoints As Double
    Dim buttonCellRange As Range
    Dim panelAutoFitLastCol As Long
    Dim panelAutoFitCols As Long
    Dim panelRange As Range
    Dim panelRenderRightCol As Long

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
    bottomRow = fieldsTopRow + (style.PanelFieldCount * rowSpan) + ((style.PanelFieldCount - 1) * fieldSpacing) - 1

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

    For fieldIndex = 1 To style.PanelFieldCount
        fieldTopRow = fieldsTopRow + ((fieldIndex - 1) * (rowSpan + fieldSpacing))
        fieldBottomRow = fieldTopRow + rowSpan - 1

        ws.Rows(fieldTopRow).RowHeight = 32

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
        currentValue = Trim$(CStr(inputAnchor.Value))
        If Len(currentValue) = 0 Then
            inputAnchor.Value = ex_ConfigProvider.m_GetConfigValue(style.PanelFields(fieldIndex).InputConfigKey, vbNullString)
        End If
        If fieldIndex = 1 Then
            mp_SetPanelInputName ws, inputAnchor
        End If
        mp_SetPanelInputKeyName ws, inputAnchor, style.PanelFields(fieldIndex).InputName

        panelAutoFitLastCol = inputStartCol
        If (buttonAnchorCol + buttonWidthCols - 1) > panelAutoFitLastCol Then
            panelAutoFitLastCol = buttonAnchorCol + buttonWidthCols - 1
        End If
        If panelAutoFitLastCol >= startCol Then
            panelAutoFitCols = panelAutoFitLastCol - startCol + 1
            ws.Columns(startCol).Resize(, panelAutoFitCols).AutoFit
        End If

        Set buttonCellRange = ws.Range(ws.Cells(fieldTopRow, buttonAnchorCol), ws.Cells(fieldBottomRow, buttonAnchorCol + buttonWidthCols - 1))
        anchorColumnWidthPoints = buttonCellRange.Width
        buttonWidthPoints = anchorColumnWidthPoints
        If buttonWidthPoints < 8 Then buttonWidthPoints = 8

        buttonTop = buttonCellRange.Top
        buttonLeft = buttonCellRange.Left
        buttonWidth = buttonWidthPoints
        buttonHeight = buttonCellRange.Height
        If buttonWidth < 8 Then buttonWidth = 8
        If buttonHeight < 8 Then buttonHeight = 8

        buttonName = mp_GetButtonName(ws, fieldIndex)
        Set buttonShape = ws.Shapes.AddShape(msoShapeRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
        buttonShape.Name = buttonName
        buttonShape.TextFrame.Characters.Text = style.PanelFields(fieldIndex).Button.Caption
        buttonShape.Fill.Solid
        buttonShape.Fill.ForeColor.RGB = style.PanelButtonBackColor
        buttonShape.Fill.Transparency = 0
        buttonShape.Line.ForeColor.RGB = style.PanelButtonBorderColor
        buttonShape.Line.Weight = 1
        buttonShape.TextFrame.Characters.Font.Bold = True
        buttonShape.TextFrame.Characters.Font.Color = style.PanelButtonTextColor
        buttonShape.TextFrame.Characters.Font.Name = style.FontName
        buttonShape.TextFrame.Characters.Font.Size = style.FontSize
        buttonShape.TextFrame.HorizontalAlignment = xlHAlignCenter
        buttonShape.TextFrame.VerticalAlignment = xlVAlignCenter
        buttonShape.Placement = xlMoveAndSize
        buttonShape.OnAction = "'" & ThisWorkbook.Name & "'!" & Trim$(style.PanelFields(fieldIndex).Button.MacroName)
    Next fieldIndex

    mp_SetPanelRangeName ws, ws.Range(ws.Cells(topRow, startCol), ws.Cells(bottomRow, panelRenderRightCol))

    mp_ApplyFixedControlPanelLayout ws, style, startCol, inputStartCol
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

Private Function mp_GetButtonName(ByVal ws As Worksheet, ByVal fieldIndex As Long) As String
    If ws Is Nothing Then
        mp_GetButtonName = PANEL_BUTTON_PREFIX
        Exit Function
    End If
    mp_GetButtonName = PANEL_BUTTON_PREFIX & ws.CodeName & "_" & CStr(fieldIndex)
End Function

Private Sub mp_ClearPanelArtifacts(ByVal ws As Worksheet)
    mp_ClearStoredPanelRange ws
    mp_DeletePanelButtons ws
    mp_DeletePanelInputNames ws
End Sub

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
    Dim shp As Shape
    Dim prefix As String
    Dim backPrefix As String

    If ws Is Nothing Then Exit Sub
    prefix = PANEL_BUTTON_PREFIX & ws.CodeName & "_"
    backPrefix = "btnOutPanelBackToDev_" & ws.CodeName

    For Each shp In ws.Shapes
        If LCase$(Left$(shp.Name, Len(prefix))) = LCase$(prefix) _
           Or LCase$(Left$(shp.Name, Len(backPrefix))) = LCase$(backPrefix) Then
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
        End If
    Next shp
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
        fieldTopRow = fieldsTopRow + ((fieldIndex - 1) * (rowSpan + spacingRows))
        ws.Rows(fieldTopRow).RowHeight = style.PanelFixedFieldRowHeight
    Next fieldIndex
End Sub

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
