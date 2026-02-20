Attribute VB_Name = "ex_OutputPanel"
Option Explicit

Private Const PANEL_INPUT_NAME As String = "outPanelInputCell"
Private Const PANEL_INPUT_PREFIX As String = "outPanelInput_"
Private Const PANEL_BUTTON_PREFIX As String = "btnOutPanelSearch_"
Private Const PANEL_BUTTON_GAP_PTS As Double = 6#

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
    Dim buttonAreaRight As Double
    Dim buttonAreaWidth As Double
    Dim titleEndCol As Long
    Dim currentValue As String
    Dim maxButtonCount As Long
    Dim panelWidth As Long
    Dim fieldsTopRow As Long
    Dim rowSpan As Long
    Dim fieldSpacing As Long
    Dim fieldIndex As Long
    Dim fieldTopRow As Long
    Dim fieldBottomRow As Long
    Dim buttonIndex As Long

    If ws Is Nothing Then Exit Sub
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
    maxButtonCount = mp_GetMaxButtonCount(style)
    If maxButtonCount < 1 Then maxButtonCount = 1

    rowSpan = style.PanelFieldRowSpan
    If rowSpan < 1 Then rowSpan = 2
    fieldSpacing = style.PanelFieldSpacingRows
    If fieldSpacing < 0 Then fieldSpacing = 0

    panelWidth = style.PanelLabelColumns + style.PanelValueColumns + maxButtonCount
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

        Set labelRange = ws.Range(ws.Cells(fieldTopRow, startCol), ws.Cells(fieldBottomRow, startCol + style.PanelLabelColumns - 1))
        labelRange.UnMerge
        labelRange.Merge
        labelRange.Value = style.PanelFields(fieldIndex).Label
        labelRange.Font.Bold = True
        labelRange.Font.Color = style.PanelLabelColor
        labelRange.HorizontalAlignment = xlCenter
        labelRange.VerticalAlignment = xlCenter

        Set inputRange = ws.Range(ws.Cells(fieldTopRow, inputStartCol), ws.Cells(fieldBottomRow, inputEndCol))
        inputRange.UnMerge
        inputRange.Merge
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

        buttonTop = ws.Cells(fieldTopRow, buttonStartCol).Top + 1
        buttonLeft = ws.Cells(fieldTopRow, buttonStartCol).Left + 1
        buttonAreaRight = ws.Cells(fieldTopRow, rightCol).Left + ws.Cells(fieldTopRow, rightCol).Width - 1
        buttonAreaWidth = buttonAreaRight - buttonLeft
        If maxButtonCount > 0 Then
            buttonWidth = (buttonAreaWidth - (PANEL_BUTTON_GAP_PTS * (maxButtonCount - 1))) / maxButtonCount
        Else
            buttonWidth = ws.Cells(fieldTopRow, buttonStartCol).Width - 2
        End If
        buttonHeight = ws.Range(ws.Cells(fieldTopRow, buttonStartCol), ws.Cells(fieldBottomRow, buttonStartCol)).Height - 2
        If buttonWidth < 8 Then buttonWidth = 8
        If buttonHeight < 8 Then buttonHeight = 8

        For buttonIndex = 1 To style.PanelFields(fieldIndex).ButtonCount
            If buttonLeft + buttonWidth > buttonAreaRight Then Exit For
            buttonName = mp_GetButtonName(ws, fieldIndex, buttonIndex)
            Set buttonShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, buttonLeft, buttonTop, buttonWidth, buttonHeight)
            buttonShape.Name = buttonName
            buttonShape.TextFrame.Characters.Text = style.PanelFields(fieldIndex).Buttons(buttonIndex).Caption
            buttonShape.Fill.ForeColor.RGB = style.PanelButtonBackColor
            buttonShape.Line.ForeColor.RGB = style.PanelButtonBorderColor
            buttonShape.Line.Weight = 1
            buttonShape.TextFrame.Characters.Font.Bold = True
            buttonShape.TextFrame.Characters.Font.Color = style.PanelButtonTextColor
            buttonShape.TextFrame.Characters.Font.Name = style.FontName
            buttonShape.TextFrame.Characters.Font.Size = style.FontSize
            buttonShape.Placement = xlMove
            buttonShape.OnAction = "'" & ThisWorkbook.Name & "'!" & Trim$(style.PanelFields(fieldIndex).Buttons(buttonIndex).MacroName)

            buttonLeft = buttonLeft + buttonWidth + PANEL_BUTTON_GAP_PTS
        Next buttonIndex
    Next fieldIndex
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

Public Function m_TryGetClickedFieldIndex(ByVal ws As Worksheet, ByVal callerName As String, ByRef fieldIndex As Long) As Boolean
    Dim prefix As String
    Dim suffix As String
    Dim sepPos As Long

    If ws Is Nothing Then Exit Function
    callerName = Trim$(callerName)
    If Len(callerName) = 0 Then Exit Function

    prefix = PANEL_BUTTON_PREFIX & ws.CodeName & "_"
    If LCase$(Left$(callerName, Len(prefix))) <> LCase$(prefix) Then Exit Function

    suffix = Mid$(callerName, Len(prefix) + 1)
    sepPos = InStr(1, suffix, "_", vbTextCompare)
    If sepPos <= 1 Then Exit Function

    If Not ex_XmlCore.m_TryParseLong(Left$(suffix, sepPos - 1), fieldIndex) Then Exit Function
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

Private Function mp_GetButtonName(ByVal ws As Worksheet, ByVal fieldIndex As Long, ByVal buttonIndex As Long) As String
    If ws Is Nothing Then
        mp_GetButtonName = PANEL_BUTTON_PREFIX
        Exit Function
    End If
    mp_GetButtonName = PANEL_BUTTON_PREFIX & ws.CodeName & "_" & CStr(fieldIndex) & "_" & CStr(buttonIndex)
End Function

Private Sub mp_DeletePanelButtons(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim prefix As String

    If ws Is Nothing Then Exit Sub
    prefix = PANEL_BUTTON_PREFIX & ws.CodeName & "_"

    For Each shp In ws.Shapes
        If LCase$(Left$(shp.Name, Len(prefix))) = LCase$(prefix) Then
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
        End If
    Next shp
End Sub

Private Function mp_GetMaxButtonCount(ByRef style As ex_SheetStylesXmlProvider.t_OutputSheetStyle) As Long
    Dim i As Long
    Dim maxCount As Long

    For i = 1 To style.PanelFieldCount
        If style.PanelFields(i).ButtonCount > maxCount Then
            maxCount = style.PanelFields(i).ButtonCount
        End If
    Next i

    mp_GetMaxButtonCount = maxCount
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
