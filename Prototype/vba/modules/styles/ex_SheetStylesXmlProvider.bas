Attribute VB_Name = "ex_SheetStylesXmlProvider"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const SHEET_STYLES_REL_PATH As String = "config\SheetStyles.xml"
Private Const BASE_STYLE_LABEL As String = "base sheet style"
Private Const OUTPUT_STYLE_LABEL As String = "output sheet style"

Public Const LAYER_BASE As String = "base"
Public Const LAYER_OUTPUT As String = "output"

Public Type t_ControlPanelButtonStyle
    Caption As String
    MacroName As String
End Type

Public Type t_ControlPanelFieldStyle
    Label As String
    InputConfigKey As String
    InputName As String
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
    ErrorBannerColumns As Long

    FontName As String
    FontSize As Double
    RowHeight As Double

    ContentColor As Long
    ContentBackColor As Long
    HeaderColor As Long
    HeaderBackColor As Long
    HeaderBold As Boolean
    SectionColor As Long
    SectionBackColor As Long
    SectionBold As Boolean
    SectionMergeColumns As Long

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
    PanelFieldCount As Long
    PanelFields() As t_ControlPanelFieldStyle
End Type

Private g_IsInitialized As Boolean
Private g_BaseStyle As t_BaseSheetStyle
Private g_OutputStyle As t_OutputSheetStyle
Private g_HasOutputStyle As Boolean

Public Function m_InitializeStyles(Optional ByVal wb As Workbook) As Boolean
    Dim doc As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Failed to initialize sheet styles: workbook is not specified.", vbExclamation
        Exit Function
    End If

    Set doc = mp_LoadSheetStylesDom(wb)
    If doc Is Nothing Then Exit Function

    If Not mp_LoadBaseSheetStyleFromDom(doc, g_BaseStyle) Then Exit Function
    g_HasOutputStyle = mp_TryLoadOutputSheetStyleFromDom(doc, g_OutputStyle)

    g_IsInitialized = True
    m_InitializeStyles = True
End Function

Public Function m_EnsureInitialized(Optional ByVal wb As Workbook) As Boolean
    If g_IsInitialized Then
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
    ElseIf style.ViewStartRow > 0 Then
        m_GetOutputViewStartRow = style.ViewStartRow
    Else
        m_GetOutputViewStartRow = 1 + style.OutputTopOffsetRows
    End If

    If m_GetOutputViewStartRow < 1 Then m_GetOutputViewStartRow = 1
End Function

Public Function m_GetOutputErrorBannerRangeAddress(Optional ByVal wb As Workbook) As String
    Dim style As t_OutputSheetStyle
    Dim startRow As Long
    Dim endCol As Long

    If Not m_GetOutputSheetStyle(style, wb) Then
        m_GetOutputErrorBannerRangeAddress = "A1:H4"
        Exit Function
    End If

    startRow = m_GetOutputViewStartRow(wb)
    endCol = style.ErrorBannerColumns
    If endCol < 1 Then endCol = 8

    m_GetOutputErrorBannerRangeAddress = "A" & CStr(startRow) & ":" & mp_ColumnLetter(endCol) & CStr(startRow + 3)
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
    Dim usedRange As Range

    If ws Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Function
    If ws.UsedRange Is Nothing Then Exit Function

    Set usedRange = ws.UsedRange
    rowCount = usedRange.Rows.Count
    colCount = usedRange.Columns.Count
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

    Application.DisplayAlerts = False
    On Error Resume Next

    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Left(ws.Name, 2) = "g_" Then ws.Delete
    Next i

    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Public Sub m_ApplyDefaultSheetView(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.Zoom = 115
End Sub

Private Function mp_LoadSheetStylesDom(ByVal wb As Workbook) As Object
    Set mp_LoadSheetStylesDom = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        SHEET_STYLES_REL_PATH, _
        PRESETS_NS, _
        "SheetStyles config file was not found: ", _
        "Failed to parse SheetStyles config file: ")
End Function

Private Function mp_LoadBaseSheetStyleFromDom(ByVal doc As Object, ByRef style As t_BaseSheetStyle) As Boolean
    Dim rootNode As Object
    Dim nodeBase As Object
    Dim nodeBackground As Object
    Dim nodeGrid As Object
    Dim gridWeightText As String

    Set rootNode = doc.selectSingleNode("/p:SheetStyles/p:baseSheetStyle")
    If rootNode Is Nothing Then
        MsgBox "SheetStyles must contain '/SheetStyles/baseSheetStyle'.", vbExclamation
        Exit Function
    End If

    Set nodeBase = rootNode.selectSingleNode("p:base")
    Set nodeBackground = rootNode.selectSingleNode("p:background")
    Set nodeGrid = rootNode.selectSingleNode("p:grid")
    If nodeBase Is Nothing Or nodeBackground Is Nothing Or nodeGrid Is Nothing Then
        MsgBox "baseSheetStyle must contain nodes: base, background, grid.", vbExclamation
        Exit Function
    End If

    If Not ex_XmlCore.m_ReadRequiredAttrLong(rootNode, "priority", style.PriorityLayer, "baseSheetStyle@priority", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeBase, "backColor", style.BaseBackColor, "base@backColor", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeBase, "fontColor", style.BaseFontColor, "base@fontColor", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrBoolean(nodeBase, "showGridlines", style.ShowGridlines, "base@showGridlines", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrLong(nodeBackground, "extraRows", style.BackgroundExtraRows, "background@extraRows", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrLong(nodeBackground, "extraCols", style.BackgroundExtraCols, "background@extraCols", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeGrid, "color", style.GridColor, "grid@color", BASE_STYLE_LABEL) Then Exit Function

    gridWeightText = LCase$(ex_XmlCore.m_ReadRequiredAttrText(nodeGrid, "weight", "grid@weight", BASE_STYLE_LABEL))
    If Len(gridWeightText) = 0 Then Exit Function
    If Not mp_TryParseBorderWeight(gridWeightText, style.GridWeight) Then
        MsgBox "Invalid value for baseSheetStyle grid@weight: " & gridWeightText & ".", vbExclamation
        Exit Function
    End If

    If Not ex_XmlCore.m_ReadRequiredAttrLong(nodeGrid, "extraRows", style.GridExtraRows, "grid@extraRows", BASE_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrLong(nodeGrid, "extraCols", style.GridExtraCols, "grid@extraCols", BASE_STYLE_LABEL) Then Exit Function

    mp_LoadBaseSheetStyleFromDom = True
End Function

Private Function mp_TryLoadOutputSheetStyleFromDom(ByVal doc As Object, ByRef style As t_OutputSheetStyle) As Boolean
    Dim rootNode As Object
    Dim nodeFont As Object
    Dim nodeRows As Object
    Dim nodeContent As Object
    Dim nodeHeader As Object
    Dim nodeSection As Object
    Dim nodeAlignment As Object
    Dim nodeStatus As Object
    Dim nodeControlPanel As Object
    Dim sectionTitleColumnsText As String

    Set rootNode = doc.selectSingleNode("/p:SheetStyles/p:outputSheetStyle")
    If rootNode Is Nothing Then
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
        MsgBox "outputSheetStyle must contain nodes: font, rows, content, header, section, alignment.", vbExclamation
        Exit Function
    End If

    If Not ex_XmlCore.m_ReadRequiredAttrLong(rootNode, "priority", style.PriorityLayer, "outputSheetStyle@priority", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not mp_ReadOptionalAttrLong(rootNode, "outputTopOffsetRows", style.OutputTopOffsetRows, 4, "outputSheetStyle@outputTopOffsetRows") Then Exit Function
    If style.OutputTopOffsetRows < 0 Then
        MsgBox "Invalid value for output sheet style attribute 'outputSheetStyle@outputTopOffsetRows': must be >= 0.", vbExclamation
        Exit Function
    End If
    If Not mp_ReadOptionalAttrLong(rootNode, "headerRows", style.HeaderRows, 4, "outputSheetStyle@headerRows") Then Exit Function
    If Not mp_ReadOptionalAttrLong(rootNode, "viewStartRow", style.ViewStartRow, 6, "outputSheetStyle@viewStartRow") Then Exit Function
    If Not mp_ReadOptionalAttrLong(rootNode, "errorBannerColumns", style.ErrorBannerColumns, 8, "outputSheetStyle@errorBannerColumns") Then Exit Function
    If style.HeaderRows < 1 Then
        MsgBox "Invalid value for output sheet style attribute 'outputSheetStyle@headerRows': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.ViewStartRow < 1 Then
        MsgBox "Invalid value for output sheet style attribute 'outputSheetStyle@viewStartRow': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.ErrorBannerColumns < 1 Then
        MsgBox "Invalid value for output sheet style attribute 'outputSheetStyle@errorBannerColumns': must be >= 1.", vbExclamation
        Exit Function
    End If
    If style.ViewStartRow <= style.HeaderRows Then
        MsgBox "Invalid layout in output sheet style: viewStartRow must be greater than headerRows.", vbExclamation
        Exit Function
    End If

    style.FontName = ex_XmlCore.m_ReadRequiredAttrText(nodeFont, "name", "font@name", OUTPUT_STYLE_LABEL)
    If Len(style.FontName) = 0 Then Exit Function

    If Not ex_XmlCore.m_ReadRequiredAttrDouble(nodeFont, "size", style.FontSize, "font@size", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrDouble(nodeRows, "height", style.RowHeight, "rows@height", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeContent, "color", style.ContentColor, "content@color", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeContent, "backColor", style.ContentBackColor, "content@backColor", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeHeader, "color", style.HeaderColor, "header@color", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeHeader, "backColor", style.HeaderBackColor, "header@backColor", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrBoolean(nodeHeader, "bold", style.HeaderBold, "header@bold", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeSection, "color", style.SectionColor, "section@color", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeSection, "backColor", style.SectionBackColor, "section@backColor", OUTPUT_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrBoolean(nodeSection, "bold", style.SectionBold, "section@bold", OUTPUT_STYLE_LABEL) Then Exit Function
    sectionTitleColumnsText = Trim$(ex_XmlCore.m_NodeAttrText(nodeSection, "sectionTitleColumns"))
    If Len(sectionTitleColumnsText) = 0 Then
        sectionTitleColumnsText = Trim$(ex_XmlCore.m_NodeAttrText(nodeSection, "mergeColumns"))
    End If
    If Len(sectionTitleColumnsText) = 0 Then
        MsgBox "Missing required " & OUTPUT_STYLE_LABEL & " attribute: section@sectionTitleColumns", vbExclamation
        Exit Function
    End If
    If Not ex_XmlCore.m_TryParseLong(sectionTitleColumnsText, style.SectionMergeColumns) Then
        MsgBox "Invalid integer " & OUTPUT_STYLE_LABEL & " attribute 'section@sectionTitleColumns': " & sectionTitleColumnsText, vbExclamation
        Exit Function
    End If

    If style.SectionMergeColumns < 1 Then
        MsgBox "Invalid value for outputSheetStyle section@sectionTitleColumns: must be >= 1.", vbExclamation
        Exit Function
    End If

    If Not mp_ReadRequiredAttrHorizontalAlignment(nodeAlignment, "horizontal", style.HorizontalAlignment) Then Exit Function
    If Not mp_ReadRequiredAttrVerticalAlignment(nodeAlignment, "vertical", style.VerticalAlignment) Then Exit Function

    Set nodeStatus = rootNode.selectSingleNode("p:status")
    If Not nodeStatus Is Nothing Then
        style.HasStatusStyle = True
        style.StatusColumnName = ex_XmlCore.m_ReadRequiredAttrText(nodeStatus, "columnName", "status@columnName", OUTPUT_STYLE_LABEL)
        If Len(style.StatusColumnName) = 0 Then Exit Function
        If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeStatus, "fontColor", style.StatusFontColor, "status@fontColor", OUTPUT_STYLE_LABEL) Then Exit Function
        If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeStatus, "defaultBackColor", style.StatusDefaultBackColor, "status@defaultBackColor", OUTPUT_STYLE_LABEL) Then Exit Function
        If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeStatus, "addedBackColor", style.StatusAddedBackColor, "status@addedBackColor", OUTPUT_STYLE_LABEL) Then Exit Function
        If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeStatus, "changedBackColor", style.StatusChangedBackColor, "status@changedBackColor", OUTPUT_STYLE_LABEL) Then Exit Function
        If Not ex_XmlCore.m_ReadRequiredAttrHexColor(nodeStatus, "removedBackColor", style.StatusRemovedBackColor, "status@removedBackColor", OUTPUT_STYLE_LABEL) Then Exit Function
    End If

    style.HasControlPanel = False
    Set nodeControlPanel = rootNode.selectSingleNode("p:controlPanel")
    If Not nodeControlPanel Is Nothing Then
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
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@startColumn': must be >= 0.", vbExclamation
            Exit Function
        End If
        If style.PanelMinStartColumn < 1 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@minStartColumn': must be >= 1.", vbExclamation
            Exit Function
        End If
        If style.PanelOffsetColumns < 1 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@offsetColumns': must be >= 1.", vbExclamation
            Exit Function
        End If
        If style.PanelWidthColumns < 4 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@widthColumns': must be >= 4.", vbExclamation
            Exit Function
        End If
        If style.PanelHeightRows < 3 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@heightRows': must be >= 3.", vbExclamation
            Exit Function
        End If
        If style.PanelTopRow < 1 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@topRow': must be >= 1.", vbExclamation
            Exit Function
        End If
        If style.PanelLabelColumns < 1 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@labelColumns': must be >= 1.", vbExclamation
            Exit Function
        End If
        If style.PanelValueColumns < 1 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@valueColumns': must be >= 1.", vbExclamation
            Exit Function
        End If
        If style.PanelFieldRowSpan < 1 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@fieldRowSpan': must be >= 1.", vbExclamation
            Exit Function
        End If
        If style.PanelFieldSpacingRows < 0 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@fieldSpacingRows': must be >= 0.", vbExclamation
            Exit Function
        End If
        If style.PanelViewZoneGapRows < 0 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@viewZoneGapRows': must be >= 0.", vbExclamation
            Exit Function
        End If
        If style.PanelColumnWidth <= 0 Then
            MsgBox "Invalid value for output sheet style attribute 'controlPanel@panelColumnWidth': must be > 0.", vbExclamation
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

        If Not mp_LoadControlPanelFields(nodeControlPanel, style) Then Exit Function
    End If

    mp_TryLoadOutputSheetStyleFromDom = True
End Function

Private Function mp_LoadControlPanelFields(ByVal nodeControlPanel As Object, ByRef style As t_OutputSheetStyle) As Boolean
    Dim fieldNodes As Object
    Dim fieldNode As Object
    Dim i As Long

    Set fieldNodes = nodeControlPanel.selectNodes("p:fields/p:field")
    If fieldNodes Is Nothing Or fieldNodes.Length = 0 Then
        Set fieldNodes = nodeControlPanel.selectNodes("p:field")
    End If

    If fieldNodes Is Nothing Or fieldNodes.Length = 0 Then
        MsgBox "Invalid controlPanel layout: at least one field is required in 'controlPanel/fields/field'.", vbExclamation
        Exit Function
    End If

    style.PanelFieldCount = fieldNodes.Length
    ReDim style.PanelFields(1 To style.PanelFieldCount)

    For i = 1 To style.PanelFieldCount
        Set fieldNode = fieldNodes.Item(i - 1)
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
    Dim i As Long

    fieldStyle.Label = mp_ReadOptionalAttrText(fieldNode, "label", vbNullString)
    fieldStyle.InputConfigKey = mp_ReadOptionalAttrText(fieldNode, "inputConfigKey", vbNullString)
    fieldStyle.InputName = mp_ReadOptionalAttrText(fieldNode, "inputName", fieldStyle.InputConfigKey)
    If Len(Trim$(fieldStyle.Label)) = 0 Then fieldStyle.Label = "Key"
    If Len(Trim$(fieldStyle.InputConfigKey)) = 0 Then fieldStyle.InputConfigKey = "Context.PersonValue"
    If Len(Trim$(fieldStyle.InputName)) = 0 Then
        fieldStyle.InputName = "Field" & CStr(fieldIndex)
    End If

    Set buttonNodes = fieldNode.selectNodes("p:button")
    If buttonNodes Is Nothing Or buttonNodes.Length = 0 Then
        MsgBox "Invalid control panel field: at least one 'button' node is required.", vbExclamation
        Exit Function
    End If

    fieldStyle.ButtonCount = buttonNodes.Length
    If fieldStyle.ButtonCount <= 0 Then
        MsgBox "Invalid control panel field: at least one button is required.", vbExclamation
        Exit Function
    End If
    ReDim fieldStyle.Buttons(1 To fieldStyle.ButtonCount)

    For i = 1 To fieldStyle.ButtonCount
        Set buttonNode = buttonNodes.Item(i - 1)
        fieldStyle.Buttons(i).Caption = mp_ReadOptionalAttrText(buttonNode, "caption", "Action " & CStr(i))
        fieldStyle.Buttons(i).MacroName = mp_ReadOptionalAttrText(buttonNode, "macro", "ex_UIActions.m_OutputPanelStartSearch_OnClick")
        If Len(Trim$(fieldStyle.Buttons(i).MacroName)) = 0 Then
            fieldStyle.Buttons(i).MacroName = "ex_UIActions.m_OutputPanelStartSearch_OnClick"
        End If
    Next i

    mp_LoadControlPanelFieldNode = True
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

Private Function mp_ReadRequiredAttrHorizontalAlignment(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long) As Boolean
    Dim textValue As String

    textValue = LCase$(ex_XmlCore.m_ReadRequiredAttrText(node, attrName, "alignment@" & attrName, OUTPUT_STYLE_LABEL))
    If Len(textValue) = 0 Then Exit Function

    Select Case textValue
        Case "center": outValue = xlCenter
        Case "left": outValue = xlLeft
        Case "right": outValue = xlRight
        Case Else
            MsgBox "Invalid alignment value for '" & attrName & "': " & textValue & ". Allowed: left, center, right.", vbExclamation
            Exit Function
    End Select

    mp_ReadRequiredAttrHorizontalAlignment = True
End Function

Private Function mp_ReadRequiredAttrVerticalAlignment(ByVal node As Object, ByVal attrName As String, ByRef outValue As Long) As Boolean
    Dim textValue As String

    textValue = LCase$(ex_XmlCore.m_ReadRequiredAttrText(node, attrName, "alignment@" & attrName, OUTPUT_STYLE_LABEL))
    If Len(textValue) = 0 Then Exit Function

    Select Case textValue
        Case "center": outValue = xlCenter
        Case "top": outValue = xlTop
        Case "bottom": outValue = xlBottom
        Case Else
            MsgBox "Invalid alignment value for '" & attrName & "': " & textValue & ". Allowed: top, center, bottom.", vbExclamation
            Exit Function
    End Select

    mp_ReadRequiredAttrVerticalAlignment = True
End Function

Private Function mp_TryParseBorderWeight(ByVal value As String, ByRef outValue As Long) As Boolean
    Select Case LCase$(Trim$(value))
        Case "hairline": outValue = xlHairline
        Case "thin": outValue = xlThin
        Case "medium": outValue = xlMedium
        Case "thick": outValue = xlThick
        Case Else: Exit Function
    End Select

    mp_TryParseBorderWeight = True
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
    Dim rowSpan As Long
    Dim spacingRows As Long

    If Not style.HasControlPanel Then Exit Function
    If style.PanelFieldCount <= 0 Then Exit Function

    rowSpan = style.PanelFieldRowSpan
    If rowSpan < 1 Then rowSpan = 2

    spacingRows = style.PanelFieldSpacingRows
    If spacingRows < 0 Then spacingRows = 0

    fieldsTopRow = style.PanelTopRow + 1
    mp_GetControlPanelBottomRow = fieldsTopRow + (style.PanelFieldCount * rowSpan) + ((style.PanelFieldCount - 1) * spacingRows) - 1
End Function
