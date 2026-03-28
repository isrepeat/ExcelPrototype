Attribute VB_Name = "ex_SheetStylesXmlProvider"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const OUTPUT_UI_FILE_SUFFIX As String = "UI.xml"
Private Const CONTROL_PANEL_BUTTON_TYPE_BUTTON As String = "button"
Private Const CONTROL_PANEL_BUTTON_TYPE_TOGGLE As String = "togglebutton"
Private Const DEFAULT_BANNER_COLUMNS As Long = 8
Private Const DEFAULT_ERROR_BANNER_ROWS As Long = 4
Private Const DEFAULT_WARNING_BANNER_ROWS As Long = 3

Public Const LAYER_BASE As String = "base"
Public Const LAYER_OUTPUT As String = "output"

Public Type t_ControlPanelToggleVariantStyle
    Value As String
    Caption As String
    HasBackColor As Boolean
    BackColor As Long
    HasTextColor As Boolean
    TextColor As Long
    HasBorderColor As Boolean
    BorderColor As Long
End Type

Public Type t_ControlPanelButtonStyle
    ButtonType As String
    Caption As String
    MacroName As String
    HasBackColor As Boolean
    BackColor As Long
    HasTextColor As Boolean
    TextColor As Long
    HasBorderColor As Boolean
    BorderColor As Long
    ToggleSource As String
    ToggleChangedMacroName As String
    ToggleVariantCount As Long
    ToggleVariants() As t_ControlPanelToggleVariantStyle
End Type

Public Type t_ControlPanelFieldStyle
    IsConfigRefField As Boolean
    ShowInput As Boolean
    OnChangeEnabled As Boolean
    Label As String
    InputConfigKey As String
    InputName As String
    OnChangeMacroName As String
    InputOverflowStyle As String
    ButtonCount As Long
    Buttons() As t_ControlPanelButtonStyle
End Type

Public Type t_BaseSheetStyle
    PriorityLayer As Long

    BaseBackColor As Long
    BaseFontColor As Long
    ShowGridlines As Boolean

    BackgroundExtraRows As Long
    BackgroundExtraCols As Long

    GridColor As Long
    GridWeight As Long
    GridExtraRows As Long
    GridExtraCols As Long
End Type

Public Type t_OutputSheetStyle
    PriorityLayer As Long
    OutputTopOffsetRows As Long
    HeaderRows As Long
    ViewStartRow As Long

    FontName As String
    FontSize As Double
    RowHeight As Double

    ContentColor As Long
    ContentBackColor As Long
    ContentWidth As Double
    ContentOverflow As String
    ContentAutoHeight As Boolean
    HeaderColor As Long
    HeaderBackColor As Long
    HeaderBold As Boolean
    HeaderWidth As Double
    HeaderOverflow As String
    HeaderAutoHeight As Boolean
    SectionColor As Long
    SectionBackColor As Long
    SectionBold As Boolean
    SectionMergeColumns As Long
    SectionWidth As Double
    SectionOverflow As String
    SectionAutoHeight As Boolean

    HorizontalAlignment As Long
    VerticalAlignment As Long

    HasStatusStyle As Boolean
    StatusColumnName As String
    StatusFontColor As Long
    StatusDefaultBackColor As Long
    StatusAddedBackColor As Long
    StatusChangedBackColor As Long
    StatusRemovedBackColor As Long

    HasControlPanel As Boolean
    PanelStartColumn As Long
    PanelMinStartColumn As Long
    PanelOffsetColumns As Long
    PanelWidthColumns As Long
    PanelHeightRows As Long
    PanelTopRow As Long
    PanelLabelColumns As Long
    PanelValueColumns As Long
    PanelFieldRowSpan As Long
    PanelFieldSpacingRows As Long
    PanelViewZoneGapRows As Long
    PanelColumnWidth As Double
    PanelTitle As String
    PanelBackColor As Long
    PanelBorderColor As Long
    PanelTitleColor As Long
    PanelLabelColor As Long
    PanelInputBackColor As Long
    PanelInputFontColor As Long
    PanelButtonBackColor As Long
    PanelButtonTextColor As Long
    PanelButtonBorderColor As Long
    PanelErrorBackColor As Long
    PanelErrorFontColor As Long
    PanelButtonAnchorColumn As Long
    PanelButtonWidthColumns As Long
    PanelFixedWidthKey As Double
    PanelFixedWidthValue As Double
    PanelFixedWidthButton As Double
    PanelFixedFieldRowHeight As Double
    PanelFieldCount As Long
    PanelFields() As t_ControlPanelFieldStyle
End Type

Public Type t_ErrorBannerStyle
    Columns As Long
    Rows As Long
    RowHeight As Double
    BackColor As Long
    FontColor As Long
    WrapText As Boolean
    TitleBold As Boolean
    ShowGrid As Boolean
    GridColor As Long
    HorizontalAlignment As Long
    VerticalAlignment As Long
End Type

Private g_IsInitialized As Boolean
Private g_BaseStyle As t_BaseSheetStyle
Private g_OutputStyle As t_OutputSheetStyle
Private g_ErrorBannerStyle As t_ErrorBannerStyle
Private g_WarningBannerStyle As t_ErrorBannerStyle
Private g_HasOutputStyle As Boolean
Private g_HasErrorBannerStyle As Boolean
Private g_HasWarningBannerStyle As Boolean
Private g_LastActiveModeKey As String

Public Function m_InitializeStyles(Optional ByVal wb As Workbook) As Boolean
    Dim modeUiDoc As Object
    Dim modeUiFilePath As String
    Dim activeModeKey As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Failed to initialize sheet styles: workbook is not specified.", vbExclamation
        Exit Function
    End If

    activeModeKey = mp_GetCurrentActiveModeKey()
    mp_ApplyDefaultBaseSheetStyle g_BaseStyle
    mp_ApplyDefaultOutputSheetStyle g_OutputStyle, activeModeKey
    mp_ApplyDefaultBannerStyles g_ErrorBannerStyle, g_WarningBannerStyle
    g_HasOutputStyle = True
    g_HasErrorBannerStyle = True
    g_HasWarningBannerStyle = True

    modeUiFilePath = mp_GetOutputUiFilePathByMode(wb)
    Set modeUiDoc = mp_LoadModeUiDomByFilePath(modeUiFilePath)
    If modeUiDoc Is Nothing Then Exit Function

    If Not mp_LoadControlPanelFromModeUi(modeUiDoc, g_OutputStyle) Then Exit Function

    g_LastActiveModeKey = mp_GetCurrentActiveModeKey()

    g_IsInitialized = True
    m_InitializeStyles = True
End Function

Private Sub mp_ApplyDefaultBaseSheetStyle(ByRef style As t_BaseSheetStyle)
    style.PriorityLayer = 100
    style.BaseBackColor = RGB(34, 34, 34)
    style.BaseFontColor = RGB(235, 235, 235)
    style.ShowGridlines = False
    style.BackgroundExtraRows = 200
    style.BackgroundExtraCols = 30
    style.GridColor = RGB(0, 0, 0)
    style.GridWeight = xlThin
    style.GridExtraRows = 200
    style.GridExtraCols = 50
End Sub

Private Sub mp_ApplyDefaultOutputSheetStyle(ByRef style As t_OutputSheetStyle, Optional ByVal activeModeKey As String = vbNullString)
    style.PriorityLayer = 200
    style.OutputTopOffsetRows = 4
    style.HeaderRows = 4
    style.ViewStartRow = 6

    style.FontName = "Times New Roman"
    style.FontSize = 12#
    style.RowHeight = 20#

    style.ContentColor = RGB(235, 235, 235)
    style.ContentBackColor = RGB(34, 34, 34)
    style.ContentWidth = 0#
    style.ContentOverflow = "wrap"
    style.ContentAutoHeight = False

    style.HeaderColor = RGB(104, 217, 71)
    style.HeaderBackColor = RGB(34, 34, 34)
    style.HeaderBold = True
    style.HeaderWidth = 0#
    style.HeaderOverflow = "wrap"
    style.HeaderAutoHeight = True

    style.SectionColor = RGB(215, 167, 99)
    style.SectionBackColor = RGB(22, 22, 22)
    style.SectionBold = True
    style.SectionMergeColumns = 2
    style.SectionWidth = 0#
    style.SectionOverflow = "clip"
    style.SectionAutoHeight = False

    style.HorizontalAlignment = xlCenter
    style.VerticalAlignment = xlCenter

    style.HasStatusStyle = False
    If StrComp(Trim$(activeModeKey), "TablesComparing", vbTextCompare) = 0 Then
        style.HasStatusStyle = True
        style.StatusColumnName = "Status"
        style.StatusFontColor = RGB(235, 235, 235)
        style.StatusDefaultBackColor = RGB(30, 30, 30)
        style.StatusAddedBackColor = RGB(46, 125, 50)
        style.StatusChangedBackColor = RGB(123, 31, 162)
        style.StatusRemovedBackColor = RGB(183, 28, 28)
    End If

    mp_ResetControlPanelStyle style
End Sub

Private Sub mp_ApplyDefaultBannerStyles(ByRef errorStyle As t_ErrorBannerStyle, ByRef warningStyle As t_ErrorBannerStyle)
    errorStyle.Columns = DEFAULT_BANNER_COLUMNS
    errorStyle.Rows = DEFAULT_ERROR_BANNER_ROWS
    errorStyle.RowHeight = 24#
    errorStyle.BackColor = RGB(90, 32, 32)
    errorStyle.FontColor = RGB(255, 225, 225)
    errorStyle.WrapText = True
    errorStyle.TitleBold = True
    errorStyle.ShowGrid = False
    errorStyle.GridColor = RGB(80, 80, 80)
    errorStyle.HorizontalAlignment = xlLeft
    errorStyle.VerticalAlignment = xlCenter

    warningStyle.Columns = DEFAULT_BANNER_COLUMNS
    warningStyle.Rows = DEFAULT_WARNING_BANNER_ROWS
    warningStyle.RowHeight = 24#
    warningStyle.BackColor = RGB(76, 63, 16)
    warningStyle.FontColor = RGB(255, 229, 153)
    warningStyle.WrapText = True
    warningStyle.TitleBold = True
    warningStyle.ShowGrid = False
    warningStyle.GridColor = RGB(80, 80, 80)
    warningStyle.HorizontalAlignment = xlLeft
    warningStyle.VerticalAlignment = xlCenter
End Sub

Public Function m_EnsureInitialized(Optional ByVal wb As Workbook) As Boolean
    If g_IsInitialized And StrComp(g_LastActiveModeKey, mp_GetCurrentActiveModeKey(), vbTextCompare) = 0 Then
        m_EnsureInitialized = True
        Exit Function
    End If

    m_EnsureInitialized = m_InitializeStyles(wb)
End Function

Public Function m_GetBaseSheetStyle(ByRef style As t_BaseSheetStyle, Optional ByVal wb As Workbook) As Boolean
    If Not m_EnsureInitialized(wb) Then Exit Function
    style = g_BaseStyle
    m_GetBaseSheetStyle = True
End Function

Public Function m_GetOutputSheetStyle(ByRef style As t_OutputSheetStyle, Optional ByVal wb As Workbook) As Boolean
    If Not m_EnsureInitialized(wb) Then Exit Function
    If Not g_HasOutputStyle Then Exit Function
    style = g_OutputStyle
    m_GetOutputSheetStyle = True
End Function

Public Function m_GetErrorBannerStyle(ByRef style As t_ErrorBannerStyle, Optional ByVal wb As Workbook) As Boolean
    If Not m_EnsureInitialized(wb) Then Exit Function
    If Not g_HasErrorBannerStyle Then Exit Function
    style = g_ErrorBannerStyle
    m_GetErrorBannerStyle = True
End Function

Public Function m_GetWarningBannerStyle(ByRef style As t_ErrorBannerStyle, Optional ByVal wb As Workbook) As Boolean
    If Not m_EnsureInitialized(wb) Then Exit Function
    If Not g_HasWarningBannerStyle Then Exit Function
    style = g_WarningBannerStyle
    m_GetWarningBannerStyle = True
End Function

Public Function m_GetOutputViewStartRow(Optional ByVal wb As Workbook) As Long
    Dim style As t_OutputSheetStyle
    Dim panelBottomRow As Long

    If Not m_GetOutputSheetStyle(style, wb) Then
        m_GetOutputViewStartRow = 1
        Exit Function
    End If

    panelBottomRow = mp_GetControlPanelBottomRow(style)
    If panelBottomRow > 0 Then
        ' Keep a configurable visual gap between control panel and data view.
        m_GetOutputViewStartRow = panelBottomRow + 1 + style.PanelViewZoneGapRows
    Else
        m_GetOutputViewStartRow = 1
    End If

    If m_GetOutputViewStartRow < 1 Then m_GetOutputViewStartRow = 1
End Function

Public Function m_GetOutputErrorBannerRangeAddress(Optional ByVal wb As Workbook) As String
    m_GetOutputErrorBannerRangeAddress = m_GetOutputBannerRangeAddress(wb, DEFAULT_ERROR_BANNER_ROWS)
End Function

Public Function m_GetOutputWarningBannerRangeAddress(Optional ByVal wb As Workbook) As String
    m_GetOutputWarningBannerRangeAddress = m_GetOutputBannerRangeAddress(wb, DEFAULT_WARNING_BANNER_ROWS)
End Function

Public Function m_GetOutputBannerRangeAddress( _
    Optional ByVal wb As Workbook, _
    Optional ByVal rowCount As Long = DEFAULT_WARNING_BANNER_ROWS, _
    Optional ByVal endCol As Long = DEFAULT_BANNER_COLUMNS _
) As String
    Dim startRow As Long

    startRow = m_GetOutputViewStartRow(wb)
    If startRow < 1 Then startRow = 1
    If endCol < 1 Then endCol = DEFAULT_BANNER_COLUMNS
    If rowCount < 1 Then rowCount = DEFAULT_WARNING_BANNER_ROWS

    m_GetOutputBannerRangeAddress = "A" & CStr(startRow) & ":" & mp_ColumnLetter(endCol) & CStr(startRow + rowCount - 1)
End Function

Public Function m_HasOutputSheetStyle(Optional ByVal wb As Workbook) As Boolean
    If Not m_EnsureInitialized(wb) Then Exit Function
    m_HasOutputSheetStyle = g_HasOutputStyle
End Function

Public Function m_GetLayerOrder( _
    ByVal includeOutputLayer As Boolean, _
    ByRef layerOrder As Variant, _
    Optional ByVal wb As Workbook _
) As Boolean
    Dim names(1 To 2) As String
    Dim priorities(1 To 2) As Long
    Dim itemCount As Long
    Dim i As Long
    Dim j As Long
    Dim tmpName As String
    Dim tmpPriority As Long
    Dim result() As String

    If Not m_EnsureInitialized(wb) Then Exit Function

    itemCount = 0
    mp_AddLayer names, priorities, itemCount, LAYER_BASE, g_BaseStyle.PriorityLayer

    If includeOutputLayer And g_HasOutputStyle Then
        mp_AddLayer names, priorities, itemCount, LAYER_OUTPUT, g_OutputStyle.PriorityLayer
    End If

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If priorities(j) < priorities(i) Then
                tmpPriority = priorities(i)
                priorities(i) = priorities(j)
                priorities(j) = tmpPriority

                tmpName = names(i)
                names(i) = names(j)
                names(j) = tmpName
            End If
        Next j
    Next i

    ReDim result(1 To itemCount)
    For i = 1 To itemCount
        result(i) = names(i)
    Next i

    layerOrder = result
    m_GetLayerOrder = True
End Function

Public Sub m_ApplyDarkThemeToSheet(ByVal ws As Worksheet)
    Dim baseStyle As t_BaseSheetStyle
    Dim rowCount As Long
    Dim colCount As Long

    If ws Is Nothing Then Exit Sub
    If Not m_GetBaseSheetStyle(baseStyle, ThisWorkbook) Then Exit Sub
    If Not m_GetUsedRangeSize(ws, rowCount, colCount) Then Exit Sub

    m_ApplyBaseLayer ws, rowCount, colCount, baseStyle
End Sub

Public Function m_GetUsedRangeSize(ByVal ws As Worksheet, ByRef rowCount As Long, ByRef colCount As Long) As Boolean
    Dim lastRowCell As Range
    Dim lastColCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastRowCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastRowCell Is Nothing Then Exit Function
    If lastColCell Is Nothing Then Exit Function

    rowCount = lastRowCell.Row
    colCount = lastColCell.Column
    m_GetUsedRangeSize = (rowCount > 0 And colCount > 0)
End Function

Public Sub m_ApplyBaseLayer(ByVal ws As Worksheet, ByVal rowCount As Long, ByVal colCount As Long, ByRef style As t_BaseSheetStyle)
    Dim visibleRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim bgRange As Range
    Dim gridRange As Range
    Dim gridRows As Long
    Dim gridCols As Long

    If ws Is Nothing Then Exit Sub
    If rowCount <= 0 Or colCount <= 0 Then Exit Sub

    ws.Activate
    Set visibleRange = ActiveWindow.VisibleRange

    lastRow = visibleRange.Row + visibleRange.Rows.Count - 1 + style.BackgroundExtraRows
    lastCol = visibleRange.Column + visibleRange.Columns.Count - 1 + style.BackgroundExtraCols

    If lastRow < rowCount + style.BackgroundExtraRows Then lastRow = rowCount + style.BackgroundExtraRows
    If lastCol < colCount + style.BackgroundExtraCols Then lastCol = colCount + style.BackgroundExtraCols

    Set bgRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    bgRange.Interior.Pattern = xlSolid
    bgRange.Interior.Color = style.BaseBackColor
    bgRange.Font.Color = style.BaseFontColor
    ActiveWindow.DisplayGridlines = style.ShowGridlines

    gridRows = rowCount + style.GridExtraRows
    gridCols = colCount + style.GridExtraCols
    If gridRows < 1 Then gridRows = 1
    If gridCols < 1 Then gridCols = 1

    Set gridRange = ws.Range(ws.Cells(1, 1), ws.Cells(gridRows, gridCols))
    With gridRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.Color = style.GridColor
        .Borders.Weight = style.GridWeight
    End With
End Sub

Public Sub m_ApplyStatusLayer(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal rowCount As Long, ByVal colCount As Long, ByRef style As t_OutputSheetStyle)
    Dim statusCol As Long
    Dim r As Long
    Dim statusValue As String
    Dim rowRange As Range
    Dim lastRow As Long

    If ws Is Nothing Then Exit Sub
    If Not style.HasStatusStyle Then Exit Sub
    If rowCount < 2 Or colCount < 1 Then Exit Sub
    If headerRow < 1 Then Exit Sub

    lastRow = headerRow + rowCount - 1

    statusCol = mp_FindColumnIndexInRow(ws, headerRow, colCount, style.StatusColumnName)
    If statusCol = 0 Then Exit Sub

    For r = headerRow + 1 To lastRow
        statusValue = Trim$(LCase$(CStr(ws.Cells(r, statusCol).Value)))
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, colCount))
        rowRange.Interior.Pattern = xlSolid

        Select Case statusValue
            Case "added"
                rowRange.Interior.Color = style.StatusAddedBackColor
            Case "changed"
                rowRange.Interior.Color = style.StatusChangedBackColor
            Case "removed"
                rowRange.Interior.Color = style.StatusRemovedBackColor
            Case Else
                rowRange.Interior.Color = style.StatusDefaultBackColor
        End Select

        rowRange.Font.Color = style.StatusFontColor
    Next r
End Sub

Public Sub m_DeleteResultSheets()
    Dim ws As Worksheet
    Dim i As Long

    On Error GoTo EH
    ex_ModePersonalCard.m_ResetResultPageSessionState
        ex_ModeHealthBenefits.m_ResetResultPageSessionState
    ex_SheetViewZoom.m_ResetZoomCache

    Application.DisplayAlerts = False

    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Left(ws.Name, 2) = "g_" Then ws.Delete
    Next i

    Application.DisplayAlerts = True
    Exit Sub

EH:
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "Clear failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
End Sub

Public Sub m_ApplyDefaultSheetView(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.Zoom = 115
End Sub

Private Function mp_LoadModeUiDomByFilePath(ByVal filePath As String) As Object
    If Len(Trim$(filePath)) = 0 Then Exit Function

    Set mp_LoadModeUiDomByFilePath = ex_XmlCore.m_LoadDomByFilePath( _
        filePath, _
        PROFILES_NS, _
        "Mode UI config file was not found: ", _
        "Failed to parse mode UI config file: ")
End Function

Private Function mp_GetCurrentActiveModeKey() As String
    Dim defaultModeKey As String

    On Error Resume Next
    mp_GetCurrentActiveModeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev))
    On Error GoTo 0

    If Len(mp_GetCurrentActiveModeKey) = 0 Then
        defaultModeKey = Trim$(ex_UiXmlProvider.m_GetDefaultModeKey(ThisWorkbook))
        If Len(defaultModeKey) > 0 Then
            mp_GetCurrentActiveModeKey = defaultModeKey
        Else
            mp_GetCurrentActiveModeKey = Trim$(ex_UiXmlProvider.m_GetModeKeyByIndex(1, ThisWorkbook))
        End If
    End If
End Function

Private Function mp_GetOutputUiFilePathByMode(ByVal wb As Workbook) As String
    Dim modeKey As String
    Dim profilesFilePath As String
    Dim slashPos As Long
    Dim modeDirPath As String
    Dim modeDirName As String

    modeKey = mp_GetCurrentActiveModeKey()
    If Len(modeKey) = 0 Then
        MsgBox "Active mode key is empty for mode UI mapping.", vbExclamation
        Exit Function
    End If

    profilesFilePath = Trim$(ex_UiXmlProvider.m_GetProfilesFilePathByMode(modeKey, wb, "profilesFileByMode"))
    If Len(profilesFilePath) = 0 Then
        MsgBox "Profiles file path is not resolved for active mode key '" & modeKey & "'.", vbExclamation
        Exit Function
    End If

    slashPos = InStrRev(profilesFilePath, "\")
    If slashPos <= 1 Then
        MsgBox "Invalid profiles file path for active mode key '" & modeKey & "': " & profilesFilePath, vbExclamation
        Exit Function
    End If
    modeDirPath = Left$(profilesFilePath, slashPos - 1)

    slashPos = InStrRev(modeDirPath, "\")
    If slashPos <= 0 Then
        modeDirName = modeDirPath
    Else
        modeDirName = Mid$(modeDirPath, slashPos + 1)
    End If
    modeDirName = Trim$(modeDirName)
    If Len(modeDirName) = 0 Then
        MsgBox "Invalid mode directory in profiles file path for active mode key '" & modeKey & "': " & profilesFilePath, vbExclamation
        Exit Function
    End If

    mp_GetOutputUiFilePathByMode = modeDirPath & "\" & modeDirName & OUTPUT_UI_FILE_SUFFIX
End Function

Private Function mp_LoadControlPanelFromModeUi(ByVal doc As Object, ByRef style As t_OutputSheetStyle) As Boolean
    Dim nodeControlPanel As Object
    Dim isPanelVisible As Boolean

    mp_ResetControlPanelStyle style

    Set nodeControlPanel = doc.selectSingleNode("/p:uiDefinition/p:controlPanel")
    If nodeControlPanel Is Nothing Then
        mp_LoadControlPanelFromModeUi = True
        Exit Function
    End If

    If Not mp_ReadOptionalAttrBoolean(nodeControlPanel, "visible", isPanelVisible, True, "controlPanel@visible") Then Exit Function
    If Not isPanelVisible Then
        mp_LoadControlPanelFromModeUi = True
        Exit Function
    End If

    style.HasControlPanel = True
    style.PanelTitle = mp_ReadOptionalAttrText(nodeControlPanel, "title", "Quick Search")

    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "startColumn", style.PanelStartColumn, 0, "controlPanel@startColumn") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "minStartColumn", style.PanelMinStartColumn, 8, "controlPanel@minStartColumn") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "offsetColumns", style.PanelOffsetColumns, 2, "controlPanel@offsetColumns") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "widthColumns", style.PanelWidthColumns, 6, "controlPanel@widthColumns") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "heightRows", style.PanelHeightRows, 3, "controlPanel@heightRows") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "topRow", style.PanelTopRow, 1, "controlPanel@topRow") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "labelColumns", style.PanelLabelColumns, 1, "controlPanel@labelColumns") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "valueColumns", style.PanelValueColumns, 2, "controlPanel@valueColumns") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "fieldRowSpan", style.PanelFieldRowSpan, 2, "controlPanel@fieldRowSpan") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "fieldSpacingRows", style.PanelFieldSpacingRows, 0, "controlPanel@fieldSpacingRows") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "viewZoneGapRows", style.PanelViewZoneGapRows, 2, "controlPanel@viewZoneGapRows") Then Exit Function
    If Not mp_ReadOptionalAttrDouble(nodeControlPanel, "panelColumnWidth", style.PanelColumnWidth, 12#, "controlPanel@panelColumnWidth") Then Exit Function

    If style.PanelStartColumn < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@startColumn': must be >= 0.", vbExclamation
        Exit Function
    End If
    If style.PanelMinStartColumn < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@minStartColumn': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelOffsetColumns < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@offsetColumns': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelWidthColumns < 4 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@widthColumns': must be >= 4.", vbExclamation
        Exit Function
    End If
    If style.PanelHeightRows < 3 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@heightRows': must be >= 3.", vbExclamation
        Exit Function
    End If
    If style.PanelTopRow < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@topRow': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelLabelColumns < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@labelColumns': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelValueColumns < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@valueColumns': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelFieldRowSpan < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@fieldRowSpan': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelFieldSpacingRows < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@fieldSpacingRows': must be >= 0.", vbExclamation
        Exit Function
    End If
    If style.PanelViewZoneGapRows < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@viewZoneGapRows': must be >= 0.", vbExclamation
        Exit Function
    End If
    If style.PanelColumnWidth <= 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@panelColumnWidth': must be > 0.", vbExclamation
        Exit Function
    End If

    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "panelBackColor", style.PanelBackColor, RGB(30, 30, 30), "controlPanel@panelBackColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "panelBorderColor", style.PanelBorderColor, RGB(80, 80, 80), "controlPanel@panelBorderColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "titleColor", style.PanelTitleColor, RGB(215, 167, 99), "controlPanel@titleColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "labelColor", style.PanelLabelColor, style.ContentColor, "controlPanel@labelColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "inputBackColor", style.PanelInputBackColor, RGB(38, 38, 38), "controlPanel@inputBackColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "inputFontColor", style.PanelInputFontColor, style.ContentColor, "controlPanel@inputFontColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "buttonBackColor", style.PanelButtonBackColor, RGB(31, 94, 156), "controlPanel@buttonBackColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "buttonTextColor", style.PanelButtonTextColor, RGB(255, 255, 255), "controlPanel@buttonTextColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "buttonBorderColor", style.PanelButtonBorderColor, RGB(22, 63, 105), "controlPanel@buttonBorderColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "errorBackColor", style.PanelErrorBackColor, style.PanelButtonBackColor, "controlPanel@errorBackColor") Then Exit Function
    If Not mp_ReadOptionalAttrHexColor(nodeControlPanel, "errorFontColor", style.PanelErrorFontColor, style.PanelButtonTextColor, "controlPanel@errorFontColor") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "buttonAnchorColumn", style.PanelButtonAnchorColumn, 4, "controlPanel@buttonAnchorColumn") Then Exit Function
    If Not mp_ReadOptionalAttrLong(nodeControlPanel, "buttonWidthColumns", style.PanelButtonWidthColumns, 1, "controlPanel@buttonWidthColumns") Then Exit Function
    If Not mp_ReadOptionalAttrDouble(nodeControlPanel, "fixedWidthKey", style.PanelFixedWidthKey, 0#, "controlPanel@fixedWidthKey") Then Exit Function
    If Not mp_ReadOptionalAttrDouble(nodeControlPanel, "fixedWidthValue", style.PanelFixedWidthValue, 0#, "controlPanel@fixedWidthValue") Then Exit Function
    If Not mp_ReadOptionalAttrDouble(nodeControlPanel, "fixedWidthButton", style.PanelFixedWidthButton, 0#, "controlPanel@fixedWidthButton") Then Exit Function
    If Not mp_ReadOptionalAttrDouble(nodeControlPanel, "fixedFieldRowHeight", style.PanelFixedFieldRowHeight, 0#, "controlPanel@fixedFieldRowHeight") Then Exit Function

    If style.PanelButtonAnchorColumn < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@buttonAnchorColumn': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelButtonWidthColumns < 1 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@buttonWidthColumns': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.PanelFixedWidthKey < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@fixedWidthKey': must be >= 0.", vbExclamation
        Exit Function
    End If
    If style.PanelFixedWidthValue < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@fixedWidthValue': must be >= 0.", vbExclamation
        Exit Function
    End If
    If style.PanelFixedWidthButton < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@fixedWidthButton': must be >= 0.", vbExclamation
        Exit Function
    End If
    If style.PanelFixedFieldRowHeight < 0 Then
        MsgBox "Invalid value for mode UI attribute 'controlPanel@fixedFieldRowHeight': must be >= 0.", vbExclamation
        Exit Function
    End If

    If Not mp_LoadControlPanelFields(nodeControlPanel, style) Then Exit Function

    mp_LoadControlPanelFromModeUi = True
End Function

Private Sub mp_ResetControlPanelStyle(ByRef style As t_OutputSheetStyle)
    style.HasControlPanel = False
    style.PanelFieldCount = 0
    Erase style.PanelFields
End Sub

Private Function mp_ReadOptionalAttrBoolean(ByVal node As Object, ByVal attrName As String, ByRef outValue As Boolean, ByVal defaultValue As Boolean, ByVal fieldName As String) As Boolean
    Dim textValue As String

    textValue = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(textValue) = 0 Then
        outValue = defaultValue
        mp_ReadOptionalAttrBoolean = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseBoolean(textValue, outValue) Then
        MsgBox "Invalid boolean output sheet style attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    mp_ReadOptionalAttrBoolean = True
End Function

Private Function mp_LoadControlPanelFields(ByVal nodeControlPanel As Object, ByRef style As t_OutputSheetStyle) As Boolean
    Dim fieldNodes As Object
    Dim fieldNode As Object
    Dim activeFieldNodes As Collection
    Dim i As Long
    Dim isNodeEnabled As Boolean

    Set fieldNodes = nodeControlPanel.selectNodes("p:fields/*[self::p:inputConfigRefField or self::p:field]")
    If fieldNodes Is Nothing Or fieldNodes.Length = 0 Then
        Set fieldNodes = nodeControlPanel.selectNodes("*[self::p:inputConfigRefField or self::p:field]")
    End If

    If fieldNodes Is Nothing Or fieldNodes.Length = 0 Then
        MsgBox "Invalid controlPanel layout: at least one field is required in 'controlPanel/fields/(inputConfigRefField|field)'.", vbExclamation
        Exit Function
    End If

    Set activeFieldNodes = New Collection
    For Each fieldNode In fieldNodes
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(fieldNode, isNodeEnabled, "condition", "controlPanel field node") Then Exit Function
        If isNodeEnabled Then activeFieldNodes.Add fieldNode
    Next fieldNode

    If activeFieldNodes.Count = 0 Then
        MsgBox "Invalid controlPanel layout: no active fields remain after applying conditions.", vbExclamation
        Exit Function
    End If

    style.PanelFieldCount = activeFieldNodes.Count
    ReDim style.PanelFields(1 To style.PanelFieldCount)

    For i = 1 To style.PanelFieldCount
        Set fieldNode = activeFieldNodes.Item(i)
        If Not mp_LoadControlPanelFieldNode(fieldNode, style.PanelFields(i), i) Then Exit Function
    Next i

    mp_LoadControlPanelFields = True
End Function

Private Function mp_LoadControlPanelFieldNode( _
    ByVal fieldNode As Object, _
    ByRef fieldStyle As t_ControlPanelFieldStyle, _
    ByVal fieldIndex As Long _
) As Boolean
    Dim buttonNodes As Object
    Dim buttonNode As Object
    Dim activeButtonNodes As Collection
    Dim fieldTagName As String
    Dim buttonIndex As Long
    Dim isNodeEnabled As Boolean
    Dim buttonType As String

    fieldTagName = LCase$(Trim$(CStr(fieldNode.baseName)))
    fieldStyle.IsConfigRefField = (fieldTagName = "inputconfigreffield")
    If Not mp_ReadOptionalAttrBoolean(fieldNode, "showInput", fieldStyle.ShowInput, True, "controlPanel field@showInput") Then Exit Function
    If Not mp_ReadOptionalAttrBoolean(fieldNode, "onChangeEnabled", fieldStyle.OnChangeEnabled, True, "controlPanel field@onChangeEnabled") Then Exit Function

    If fieldTagName <> "inputconfigreffield" And fieldTagName <> "field" Then
        MsgBox "Invalid control panel field tag: '" & fieldTagName & "'. Allowed tags: field, inputConfigRefField.", vbExclamation
        Exit Function
    End If

    fieldStyle.Label = mp_ReadOptionalAttrText(fieldNode, "label", vbNullString)
    fieldStyle.InputConfigKey = mp_ReadOptionalAttrText(fieldNode, "inputConfigKey", vbNullString)
    fieldStyle.InputName = mp_ReadOptionalAttrText(fieldNode, "inputName", vbNullString)
    fieldStyle.OnChangeMacroName = mp_ReadOptionalAttrText(fieldNode, "onChange", vbNullString)
    If Not fieldStyle.OnChangeEnabled Then fieldStyle.OnChangeMacroName = vbNullString
    fieldStyle.InputOverflowStyle = mp_ReadOptionalAttrText(fieldNode, "inputOverflowStyle", "wrap")

    If Not mp_NormalizeInputOverflowStyle(fieldStyle.InputOverflowStyle, fieldStyle.InputOverflowStyle, fieldTagName, fieldIndex) Then Exit Function

    If Not fieldStyle.IsConfigRefField Then
        If Len(Trim$(fieldStyle.Label)) = 0 Then fieldStyle.Label = "Key"
        If Len(Trim$(fieldStyle.InputConfigKey)) = 0 Then fieldStyle.InputConfigKey = "CommonKey"
        If Len(Trim$(fieldStyle.InputName)) = 0 Then
            fieldStyle.InputName = "Field" & CStr(fieldIndex)
        End If
    End If

    Set buttonNodes = fieldNode.selectNodes("p:button | p:toggleButton")
    If buttonNodes Is Nothing Or buttonNodes.Length = 0 Then
        MsgBox "Invalid control panel field: at least one 'button' or 'toggleButton' node is required.", vbExclamation
        Exit Function
    End If

    Set activeButtonNodes = New Collection
    For Each buttonNode In buttonNodes
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(buttonNode, isNodeEnabled, "condition", "controlPanel button node (field " & CStr(fieldIndex) & ")") Then Exit Function
        If isNodeEnabled Then activeButtonNodes.Add buttonNode
    Next buttonNode

    If activeButtonNodes.Count = 0 Then
        MsgBox "Invalid control panel field: no active 'button' nodes remain after applying conditions (field " & CStr(fieldIndex) & ").", vbExclamation
        Exit Function
    End If

    fieldStyle.ButtonCount = activeButtonNodes.Count
    ReDim fieldStyle.Buttons(1 To fieldStyle.ButtonCount)

    For buttonIndex = 1 To fieldStyle.ButtonCount
        Set buttonNode = activeButtonNodes.Item(buttonIndex)
        buttonType = mp_NormalizeControlPanelButtonType(CStr(buttonNode.baseName), mp_ReadOptionalAttrText(buttonNode, "type", vbNullString))
        If Len(buttonType) = 0 Then Exit Function

        fieldStyle.Buttons(buttonIndex).ButtonType = buttonType
        fieldStyle.Buttons(buttonIndex).Caption = mp_ReadOptionalAttrText(buttonNode, "caption", vbNullString)
        fieldStyle.Buttons(buttonIndex).MacroName = mp_ReadOptionalAttrText(buttonNode, "macro", vbNullString)
        fieldStyle.Buttons(buttonIndex).ToggleSource = mp_ReadOptionalAttrText(buttonNode, "source", vbNullString)
        fieldStyle.Buttons(buttonIndex).ToggleChangedMacroName = mp_ReadOptionalAttrText(buttonNode, "onToggle", vbNullString)
        fieldStyle.Buttons(buttonIndex).ToggleVariantCount = 0
        Erase fieldStyle.Buttons(buttonIndex).ToggleVariants
        If Not mp_ReadOptionalButtonHexColor(buttonNode, "buttonBackColor", fieldStyle.Buttons(buttonIndex).BackColor, fieldStyle.Buttons(buttonIndex).HasBackColor, "inputConfigRefField/button@buttonBackColor", fieldIndex, buttonIndex) Then Exit Function
        If Not mp_ReadOptionalButtonHexColor(buttonNode, "buttonTextColor", fieldStyle.Buttons(buttonIndex).TextColor, fieldStyle.Buttons(buttonIndex).HasTextColor, "inputConfigRefField/button@buttonTextColor", fieldIndex, buttonIndex) Then Exit Function
        If Not mp_ReadOptionalButtonHexColor(buttonNode, "buttonBorderColor", fieldStyle.Buttons(buttonIndex).BorderColor, fieldStyle.Buttons(buttonIndex).HasBorderColor, "inputConfigRefField/button@buttonBorderColor", fieldIndex, buttonIndex) Then Exit Function

        If StrComp(buttonType, CONTROL_PANEL_BUTTON_TYPE_TOGGLE, vbTextCompare) = 0 Then
            If Len(Trim$(fieldStyle.Buttons(buttonIndex).ToggleSource)) = 0 Then
                MsgBox "Invalid control panel toggleButton: missing required attribute 'source' (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ").", vbExclamation
                Exit Function
            End If

            If Not mp_LoadControlPanelToggleVariants(buttonNode, fieldStyle.Buttons(buttonIndex), fieldIndex, buttonIndex) Then Exit Function

            If Len(Trim$(fieldStyle.Buttons(buttonIndex).MacroName)) = 0 Then
                fieldStyle.Buttons(buttonIndex).MacroName = "ex_UIActions.m_OutputPanelToggleButton_OnClick"
            End If

            If Len(Trim$(fieldStyle.Buttons(buttonIndex).Caption)) = 0 Then
                fieldStyle.Buttons(buttonIndex).Caption = fieldStyle.Buttons(buttonIndex).ToggleVariants(1).Caption
                If Len(Trim$(fieldStyle.Buttons(buttonIndex).Caption)) = 0 Then
                    fieldStyle.Buttons(buttonIndex).Caption = fieldStyle.Buttons(buttonIndex).ToggleVariants(1).Value
                End If
            End If
        End If

        If Not fieldStyle.IsConfigRefField Then
            If Len(Trim$(fieldStyle.Buttons(buttonIndex).Caption)) = 0 Then
                fieldStyle.Buttons(buttonIndex).Caption = "Action " & CStr(buttonIndex)
            End If
            If Len(Trim$(fieldStyle.Buttons(buttonIndex).MacroName)) = 0 Then
                fieldStyle.Buttons(buttonIndex).MacroName = "ex_UIActions.m_OutputPanelStartSearch_OnClick"
            End If
        End If
    Next buttonIndex

    mp_LoadControlPanelFieldNode = True
End Function

Private Function mp_LoadControlPanelToggleVariants( _
    ByVal buttonNode As Object, _
    ByRef buttonStyle As t_ControlPanelButtonStyle, _
    ByVal fieldIndex As Long, _
    ByVal buttonIndex As Long _
) As Boolean
    Dim variantNodes As Object
    Dim variantNode As Object
    Dim activeVariantNodes As Collection
    Dim variantIdx As Long
    Dim isNodeEnabled As Boolean
    Dim valueText As String

    Set variantNodes = buttonNode.selectNodes("p:modeVariants/p:variant")
    If variantNodes Is Nothing Or variantNodes.Length = 0 Then
        MsgBox "Invalid control panel toggleButton: at least one 'modeVariants/variant' node is required (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ").", vbExclamation
        Exit Function
    End If

    Set activeVariantNodes = New Collection
    For Each variantNode In variantNodes
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(variantNode, isNodeEnabled, "condition", "controlPanel toggle variant node (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ")") Then Exit Function
        If isNodeEnabled Then activeVariantNodes.Add variantNode
    Next variantNode

    If activeVariantNodes.Count = 0 Then
        MsgBox "Invalid control panel toggleButton: no active 'variant' nodes remain after applying conditions (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ").", vbExclamation
        Exit Function
    End If

    buttonStyle.ToggleVariantCount = activeVariantNodes.Count
    ReDim buttonStyle.ToggleVariants(1 To buttonStyle.ToggleVariantCount)

    For variantIdx = 1 To buttonStyle.ToggleVariantCount
        Set variantNode = activeVariantNodes.Item(variantIdx)
        valueText = mp_ReadOptionalAttrText(variantNode, "value", vbNullString)
        If Len(valueText) = 0 Then
            MsgBox "Invalid control panel toggleButton variant: missing required attribute 'value' (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & ", variant " & CStr(variantIdx) & ").", vbExclamation
            Exit Function
        End If

        buttonStyle.ToggleVariants(variantIdx).Value = valueText
        buttonStyle.ToggleVariants(variantIdx).Caption = mp_ReadOptionalAttrText(variantNode, "caption", vbNullString)
        If Len(buttonStyle.ToggleVariants(variantIdx).Caption) = 0 Then
            buttonStyle.ToggleVariants(variantIdx).Caption = mp_ReadOptionalAttrText(variantNode, "display", vbNullString)
        End If

        If Not mp_ReadOptionalButtonHexColor(variantNode, "buttonBackColor", buttonStyle.ToggleVariants(variantIdx).BackColor, buttonStyle.ToggleVariants(variantIdx).HasBackColor, "toggleButton/modeVariants/variant@buttonBackColor", fieldIndex, buttonIndex) Then Exit Function
        If Not mp_ReadOptionalButtonHexColor(variantNode, "buttonTextColor", buttonStyle.ToggleVariants(variantIdx).TextColor, buttonStyle.ToggleVariants(variantIdx).HasTextColor, "toggleButton/modeVariants/variant@buttonTextColor", fieldIndex, buttonIndex) Then Exit Function
        If Not mp_ReadOptionalButtonHexColor(variantNode, "buttonBorderColor", buttonStyle.ToggleVariants(variantIdx).BorderColor, buttonStyle.ToggleVariants(variantIdx).HasBorderColor, "toggleButton/modeVariants/variant@buttonBorderColor", fieldIndex, buttonIndex) Then Exit Function
    Next variantIdx

    mp_LoadControlPanelToggleVariants = True
End Function

Private Function mp_NormalizeControlPanelButtonType(ByVal nodeTagName As String, ByVal rawType As String) As String
    nodeTagName = LCase$(Trim$(nodeTagName))
    rawType = LCase$(Trim$(rawType))

    If Len(rawType) = 0 Then
        If StrComp(nodeTagName, CONTROL_PANEL_BUTTON_TYPE_TOGGLE, vbTextCompare) = 0 Then
            rawType = CONTROL_PANEL_BUTTON_TYPE_TOGGLE
        Else
            rawType = CONTROL_PANEL_BUTTON_TYPE_BUTTON
        End If
    End If

    Select Case rawType
        Case CONTROL_PANEL_BUTTON_TYPE_BUTTON, CONTROL_PANEL_BUTTON_TYPE_TOGGLE
            mp_NormalizeControlPanelButtonType = rawType
        Case Else
            MsgBox "Unsupported control panel button type '" & rawType & "'. Allowed: button, toggleButton.", vbExclamation
    End Select
End Function

Private Function mp_NormalizeInputOverflowStyle( _
    ByVal rawValue As String, _
    ByRef normalized As String, _
    ByVal fieldTagName As String, _
    ByVal fieldIndex As Long _
) As Boolean
    rawValue = LCase$(Trim$(rawValue))
    If Len(rawValue) = 0 Then rawValue = "wrap"

    Select Case rawValue
        Case "wrap"
            normalized = "wrap"
        Case "shrink"
            normalized = "shrink"
        Case "overflow", "clip"
            normalized = "overflow"
        Case Else
            MsgBox "Invalid value for input overflow style: " & fieldTagName & "@inputOverflowStyle='" & rawValue & "' (field " & CStr(fieldIndex) & "). Allowed: wrap, shrink, overflow, clip.", vbExclamation
            Exit Function
    End Select

    mp_NormalizeInputOverflowStyle = True
End Function

Private Function mp_ReadOptionalAttrText(ByVal node As Object, ByVal attrName As String, ByVal defaultValue As String) As String
    mp_ReadOptionalAttrText = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(mp_ReadOptionalAttrText) = 0 Then
        mp_ReadOptionalAttrText = defaultValue
    End If
End Function

Private Function mp_ReadOptionalAttrLong(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long, ByVal defaultValue As Long, ByVal fieldName As String) As Boolean
    Dim textValue As String

    textValue = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(textValue) = 0 Then
        outValue = defaultValue
        mp_ReadOptionalAttrLong = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseLong(textValue, outValue) Then
        MsgBox "Invalid integer output sheet style attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    mp_ReadOptionalAttrLong = True
End Function

Private Function mp_ReadOptionalAttrDouble(ByVal node As Object, ByVal attrName As String, ByRef outValue As Double, ByVal defaultValue As Double, ByVal fieldName As String) As Boolean
    Dim textValue As String

    textValue = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(textValue) = 0 Then
        outValue = defaultValue
        mp_ReadOptionalAttrDouble = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseDouble(textValue, outValue, True) Then
        MsgBox "Invalid numeric output sheet style attribute '" & fieldName & "': " & textValue, vbExclamation
        Exit Function
    End If

    mp_ReadOptionalAttrDouble = True
End Function

Private Function mp_ReadOptionalAttrHexColor(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long, ByVal defaultValue As Long, ByVal fieldName As String) As Boolean
    Dim textValue As String

    textValue = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(textValue) = 0 Then
        outValue = defaultValue
        mp_ReadOptionalAttrHexColor = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseHexColor(textValue, outValue) Then
        MsgBox "Invalid color output sheet style attribute '" & fieldName & "': expected #RRGGBB, got " & textValue, vbExclamation
        Exit Function
    End If

    mp_ReadOptionalAttrHexColor = True
End Function

Private Function mp_ReadOptionalButtonHexColor( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByRef hasValue As Boolean, _
    ByVal fieldName As String, _
    ByVal fieldIndex As Long, _
    ByVal buttonIndex As Long _
) As Boolean
    Dim textValue As String

    textValue = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(textValue) = 0 Then
        hasValue = False
        mp_ReadOptionalButtonHexColor = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseHexColor(textValue, outValue) Then
        MsgBox "Invalid color output sheet style attribute '" & fieldName & "' (field " & CStr(fieldIndex) & ", button " & CStr(buttonIndex) & "): expected #RRGGBB, got " & textValue, vbExclamation
        Exit Function
    End If

    hasValue = True
    mp_ReadOptionalButtonHexColor = True
End Function

Private Function mp_ReadOptionalAttrHorizontalAlignment( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByVal defaultValue As Long, _
    ByVal fieldName As String _
) As Boolean
    Dim textValue As String

    textValue = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(node, attrName)))
    If Len(textValue) = 0 Then
        outValue = defaultValue
        mp_ReadOptionalAttrHorizontalAlignment = True
        Exit Function
    End If

    Select Case textValue
        Case "center": outValue = xlCenter
        Case "left": outValue = xlLeft
        Case "right": outValue = xlRight
        Case Else
            MsgBox "Invalid alignment value for '" & fieldName & "': " & textValue & ". Allowed: left, center, right.", vbExclamation
            Exit Function
    End Select

    mp_ReadOptionalAttrHorizontalAlignment = True
End Function

Private Function mp_ReadOptionalAttrVerticalAlignment( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByVal defaultValue As Long, _
    ByVal fieldName As String _
) As Boolean
    Dim textValue As String

    textValue = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(node, attrName)))
    If Len(textValue) = 0 Then
        outValue = defaultValue
        mp_ReadOptionalAttrVerticalAlignment = True
        Exit Function
    End If

    Select Case textValue
        Case "center": outValue = xlCenter
        Case "top": outValue = xlTop
        Case "bottom": outValue = xlBottom
        Case Else
            MsgBox "Invalid alignment value for '" & fieldName & "': " & textValue & ". Allowed: top, center, bottom.", vbExclamation
            Exit Function
    End Select

    mp_ReadOptionalAttrVerticalAlignment = True
End Function

Private Sub mp_AddLayer( _
    ByRef names() As String, _
    ByRef priorities() As Long, _
    ByRef itemCount As Long, _
    ByVal layerName As String, _
    ByVal layerPriority As Long _
)
    itemCount = itemCount + 1
    names(itemCount) = layerName
    priorities(itemCount) = layerPriority
End Sub

Private Function mp_FindColumnIndexInRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal colCount As Long, ByVal headerName As String) As Long
    Dim c As Long

    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Then Exit Function
    If colCount <= 0 Then Exit Function

    For c = 1 To colCount
        If StrComp(CStr(ws.Cells(headerRow, c).Value), headerName, vbTextCompare) = 0 Then
            mp_FindColumnIndexInRow = c
            Exit Function
        End If
    Next c
End Function

Private Function mp_ColumnLetter(ByVal colIndex As Long) As String
    Dim n As Long
    Dim part As Long

    n = colIndex
    If n < 1 Then n = 1

    Do While n > 0
        part = (n - 1) Mod 26
        mp_ColumnLetter = Chr$(65 + part) & mp_ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function mp_GetControlPanelBottomRow(ByRef style As t_OutputSheetStyle) As Long
    Dim fieldsTopRow As Long
    Dim baseRowSpan As Long
    Dim spacingRows As Long
    Dim totalRows As Long
    Dim fieldIndex As Long

    If Not style.HasControlPanel Then Exit Function
    If style.PanelFieldCount <= 0 Then Exit Function

    baseRowSpan = style.PanelFieldRowSpan
    If baseRowSpan < 1 Then baseRowSpan = 2

    spacingRows = style.PanelFieldSpacingRows
    If spacingRows < 0 Then spacingRows = 0

    fieldsTopRow = style.PanelTopRow + 1
    For fieldIndex = 1 To style.PanelFieldCount
        totalRows = totalRows + mp_GetPanelFieldEffectiveRowSpan(style, fieldIndex, baseRowSpan)
        If fieldIndex < style.PanelFieldCount Then
            totalRows = totalRows + spacingRows
        End If
    Next fieldIndex

    mp_GetControlPanelBottomRow = fieldsTopRow + totalRows - 1
End Function

Private Function mp_GetPanelFieldEffectiveRowSpan(ByRef style As t_OutputSheetStyle, ByVal fieldIndex As Long, ByVal baseRowSpan As Long) As Long
    Dim effectiveSpan As Long

    effectiveSpan = baseRowSpan
    If effectiveSpan < 1 Then effectiveSpan = 1

    If fieldIndex >= 1 And fieldIndex <= style.PanelFieldCount Then
        If style.PanelFields(fieldIndex).ButtonCount > effectiveSpan Then
            effectiveSpan = style.PanelFields(fieldIndex).ButtonCount
        End If
    End If

    mp_GetPanelFieldEffectiveRowSpan = effectiveSpan
End Function
