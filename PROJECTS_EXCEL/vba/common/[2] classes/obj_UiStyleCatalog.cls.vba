VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiStyleCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private Const COMMON_STYLE_CATALOG_FILE_NAME As String = "CommonControlStyles.xaml"
Private Const SHEET_SCOPE_MIN_COLUMN As Long = 40
Private Const SHEET_SCOPE_MIN_ROW As Long = 100
Private m_stylesByName As Object
Private m_pageDocument As Object
Private m_targetWorksheet As Worksheet
Private m_pipeline As obj_UiStylePipeline
Private m_overlays As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_stylesByName = Nothing
    Set m_pageDocument = Nothing
    Set m_targetWorksheet = Nothing
    If Not m_pipeline Is Nothing Then m_pipeline.Dispose
    Set m_pipeline = Nothing
    Set m_overlays = Nothing
End Sub

Public Sub BeginPage( _
    ByVal targetWorksheet As Worksheet, _
    ByVal pageDocument As Object, _
    ByVal uiFolderPath As String _
)
    Dim fileSystem As Object
    Dim commonStyleDocument As Object
    Dim commonStylePath As String
    Dim startedAt As Double

    startedAt = VBA.Timer
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
    Set m_pipeline = New obj_UiStylePipeline
    If Not m_pipeline.Initialize(pageDocument) Then ex_UiStyleDiagnostics.fn_Raise "Initialize"
    private_CompileRules
    ex_Core.fn_Diagnostic_WriteLog "STYLE_PAGE_READY | Sheet=" & targetWorksheet.Name & _
        " | StyleCount=" & VBA.CStr(m_stylesByName.Count)
    ex_Core.fn_Diagnostic_WritePerf "Style.BeginPage | Sheet=" & targetWorksheet.Name, startedAt
End Sub

Public Sub ApplyControlStyle( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal controlNode As Object, _
    ByVal uiBindingContext As obj_UiBindingContext _
)
    Dim styleName As String
    Dim styleProperties As Object
    Dim directProperties As Object
    Dim startedAt As Double

    startedAt = VBA.Timer
    If m_stylesByName Is Nothing Then Exit Sub
    If Not ex_UiBindingRuntime.fn_TryResolveText( _
            private_ReadAttribute(controlNode, "style"), uiBindingContext, styleName) Then Exit Sub
    If VBA.Len(styleName) > 0 Then
        If m_stylesByName.Exists(styleName) Then
            Set styleProperties = m_stylesByName(styleName)
            private_ApplyProperties targetRange, targetShape, styleProperties
        Else
            ex_Core.fn_Diagnostic_WriteLog "STYLE_NOT_FOUND | Sheet=" & _
                m_targetWorksheet.Name & " | Style=" & styleName
            ex_WindowsUi.fn_ShowMessage "Control style is not declared: " & styleName, _
                VBA.vbExclamation, "PersonalEventBuilder / Styles"
            VBA.Err.Raise 5, "obj_UiStyleCatalog.ApplyControlStyle", _
                "Control style is not declared: " & styleName
        End If
    End If

    Set directProperties = private_ReadVisualAttributes(controlNode)
    private_ApplyProperties targetRange, targetShape, directProperties
    m_pipeline.RegisterRegion controlNode, targetRange, targetShape, "control"
    If Not targetRange Is Nothing Then m_pipeline.RegisterRegion controlNode, targetRange, Nothing, "cell"
    If Not targetShape Is Nothing Then m_pipeline.RegisterRegion controlNode, Nothing, targetShape, "shape"
    If Not targetShape Is Nothing And VBA.LCase$(private_ReadAttribute(controlNode, "type")) = "select" Then
        m_pipeline.RegisterRegion controlNode, Nothing, targetShape, "header"
    End If
    ex_Core.fn_Diagnostic_WritePerf "Style.ApplyControl", startedAt
End Sub

Public Sub ApplyControlPartStyle( _
    ByVal targetShape As Object, _
    ByVal controlNode As Object, _
    ByVal uiBindingContext As obj_UiBindingContext, _
    ByVal styleAttributeName As String _
)
    Dim styleName As String
    Dim styleProperties As Object

    If targetShape Is Nothing Or m_stylesByName Is Nothing Then Exit Sub
    If Not m_pipeline Is Nothing Then
        m_pipeline.RegisterRegion controlNode, Nothing, targetShape, VBA.LCase$(VBA.Replace(styleAttributeName, "Style", ""))
    End If
    If Not ex_UiBindingRuntime.fn_TryResolveText( _
            private_ReadAttribute(controlNode, styleAttributeName), _
            uiBindingContext, styleName) Then Exit Sub
    If VBA.Len(styleName) = 0 Then Exit Sub
    If Not m_stylesByName.Exists(styleName) Then
        ex_Core.fn_Diagnostic_WriteLog "STYLE_NOT_FOUND | Sheet=" & _
            m_targetWorksheet.Name & " | Style=" & styleName
        ex_WindowsUi.fn_ShowMessage "Control style is not declared: " & styleName, _
            VBA.vbExclamation, "PersonalEventBuilder / Styles"
        Exit Sub
    End If
    Set styleProperties = m_stylesByName(styleName)
    private_ApplyProperties Nothing, targetShape, styleProperties
End Sub

Public Sub ApplyPagePipeline(ByVal targetWorksheet As Worksheet)
    Me.ApplyStage "default"
End Sub

Public Sub BeginRender()
    If m_pipeline Is Nothing Then Exit Sub
    m_pipeline.BeginRender
    Set m_overlays = VBA.CreateObject("Scripting.Dictionary")
    m_overlays.CompareMode = VBA.vbTextCompare
End Sub

Public Sub ApplyOverlay( _
    ByVal area As Range, _
    ByVal node As Object, _
    ByVal bindings As obj_UiBindingContext _
)
    Dim styleName As String
    Dim overlay As Object

    If Not ex_UiBindingRuntime.fn_TryResolveText(private_ReadAttribute(node, "style"), bindings, styleName) Then ex_UiStyleDiagnostics.fn_Raise "OverlayBinding"
    If Not m_stylesByName.Exists(styleName) Then ex_UiStyleDiagnostics.fn_Raise "OverlayMissing", styleName
    Set overlay = VBA.CreateObject("Scripting.Dictionary")
    overlay.Add "area", area
    overlay.Add "properties", m_stylesByName(styleName)
    If m_overlays Is Nothing Then
        Set m_overlays = VBA.CreateObject("Scripting.Dictionary")
        m_overlays.CompareMode = VBA.vbTextCompare
    End If
    Set m_overlays(private_ReadAttribute(node, "name")) = overlay
    private_ApplyProperties area, Nothing, m_stylesByName(styleName)
End Sub

Public Sub RegisterPart( _
    ByVal controlNode As Object, _
    ByVal area As Range, _
    ByVal part As String, _
    Optional ByVal columnAlias As String, _
    Optional ByVal sourceAlias As String, _
    Optional ByVal sourceAliasTemplate As String, _
    Optional ByVal tags As String _
)
    If m_pipeline Is Nothing Then Exit Sub
    m_pipeline.RegisterRegion controlNode, area, Nothing, part, columnAlias, sourceAlias, sourceAliasTemplate, tags
End Sub

Public Sub ApplyStage(ByVal stageName As String)
    Dim rule As Object
    Dim regions As Collection
    Dim region As Object
    Dim properties As Object
    Dim area As Range
    Dim pageScope As Range
    Dim shape As Object
    Dim selector As Object
    Dim startedAt As Double
    Dim ruleIndex As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim regionKey As Variant
    Dim operation As String

    On Error GoTo EH_STAGE
    startedAt = VBA.Timer
    If m_pipeline Is Nothing Then Exit Sub
    m_pipeline.ValidateStage stageName
    ex_Core.fn_Diagnostic_WriteLog "STYLE_STAGE_STARTED | Name=" & stageName
    Set pageScope = private_GetSheetScope(m_targetWorksheet)
    For Each rule In m_pipeline.Rules
        ruleIndex = ruleIndex + 1
        If rule("enabled") And VBA.StrComp(rule("stage"), stageName, VBA.vbTextCompare) = 0 Then
            ex_Core.fn_Diagnostic_WriteLog "STYLE_RULE_STARTED | Index=" & VBA.CStr(ruleIndex) & " | Target=" & rule("target")
            Set properties = rule("properties")
            Set selector = rule("selector")
            operation = "ResolveRegions"
            Set regions = m_pipeline.ResolveRegions(rule, m_targetWorksheet, pageScope)
            operation = "ApplyRegions"
            private_ApplyRegions regions, properties
            If rule("target") = "sheet" Then private_ApplyWorksheetProperties m_targetWorksheet, properties
        End If
    Next rule
    If Not m_overlays Is Nothing Then
        For Each regionKey In m_overlays.Keys
            Set region = m_overlays(regionKey)
            Set area = region("area")
            Set properties = region("properties")
            private_ApplyProperties area, Nothing, properties
        Next regionKey
    End If
    ex_Core.fn_Diagnostic_WritePerf "Style.ApplyStage | Name=" & stageName, startedAt
    Exit Sub
EH_STAGE:
    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "STYLE_STAGE_ERROR | Name=" & stageName & " | Rule=" & VBA.CStr(ruleIndex) & " | Operation=" & operation & " | Error=" & VBA.CStr(errorNumber) & " | Source=" & errorSource & " | Description=" & errorDescription
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub private_ApplyRegions( _
    ByVal regions As Collection, _
    ByVal properties As Object _
)
    Dim batch As Range
    Dim rowArea As Range
    Dim shape As Object
    Dim region As Object
    Dim count As Long

    For Each region In regions
        Set rowArea = region("area")
        Set shape = region("shape")
        If region("part") = "row" And Not rowArea Is Nothing And shape Is Nothing Then
            If batch Is Nothing Then
                Set batch = rowArea
            Else
                Set batch = Application.Union(batch, rowArea)
            End If
            count = count + 1
            If count = 64 Then
                private_ApplyProperties batch, Nothing, properties
                Set batch = Nothing
                count = 0
            End If
        Else
            If Not batch Is Nothing Then
                private_ApplyProperties batch, Nothing, properties
                Set batch = Nothing
                count = 0
            End If
            private_ApplyProperties rowArea, shape, properties
        End If
    Next region
    If Not batch Is Nothing Then private_ApplyProperties batch, Nothing, properties
End Sub

Public Function ResolveInlineStyle( _
    ByVal node As Object, _
    ByVal part As String, _
    ByVal stageName As String _
) As Object
    Dim result As Object
    Dim rule As Object
    Dim properties As Object
    Dim key As Variant
    Dim selector As Object

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    m_pipeline.ValidateStage stageName
    For Each rule In m_pipeline.Rules
        If rule("enabled") And rule("target") = "inlinepart" And VBA.StrComp(rule("stage"), stageName, VBA.vbTextCompare) = 0 Then
            If m_pipeline.MatchesInline(node, part, rule("selector")) Then
                Set selector = rule("selector")
                If selector.Exists("sheet") Then
                    If VBA.StrComp(selector("sheet"), m_targetWorksheet.Name, VBA.vbTextCompare) <> 0 Then GoTo NextInlineRule
                End If
                Set properties = rule("properties")
                For Each key In properties.Keys
                    result(key) = properties(key)
                Next key
            End If
        End If
NextInlineRule:
    Next rule
    Set ResolveInlineStyle = result
End Function

Private Sub private_CompileRules()
    Dim rule As Object
    Dim properties As Object
    Dim baseProperties As Object
    Dim key As Variant

    For Each rule In m_pipeline.Rules
        Set properties = private_ParseStyleDeclarations(rule("styles"))
        If VBA.Len(rule("style")) > 0 Then
            If Not m_stylesByName.Exists(rule("style")) Then ex_UiStyleDiagnostics.fn_Raise "RuleStyleMissing", rule("style")
            Set baseProperties = m_stylesByName(rule("style"))
            For Each key In baseProperties.Keys
                If Not properties.Exists(key) Then properties.Add key, baseProperties(key)
            Next key
        End If
        rule.Add "properties", properties
    Next rule
End Sub


' //
' // Private
' //
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
                 "overflow", "zoom", "gridlines", "minwidth", "maxwidth", "autofitcolumns", "celltype"
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
                If VBA.Len(propertyValue) = 0 Then ex_UiStyleDiagnostics.fn_Raise "PropertyEmpty", propertyName
                result(propertyName) = propertyValue
            Else
                ex_UiStyleDiagnostics.fn_Raise "Property", propertyName
            End If
        ElseIf VBA.Len(private_TrimXmlWhitespace(VBA.CStr(declarationPart))) > 0 Then
            ex_UiStyleDiagnostics.fn_Raise "Declaration", VBA.CStr(declarationPart)
        End If
    Next declarationPart
    Set private_ParseStyleDeclarations = result
End Function

Private Function private_IsVisualProperty(ByVal propertyName As String) As Boolean
    Select Case propertyName
            Case "backcolor", "textcolor", "fontcolor", "bordercolor", "borderweight", _
             "borderlinestyle", "fontbold", "fontitalic", "fontsize", _
                 "fontname", "horizontal", "vertical", "columnwidth", "width", "rowheight", _
                  "overflow", "zoom", "gridlines", "minwidth", "maxwidth", "autofitcolumns", "celltype"
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
        If targetShape Is Nothing Then
            targetRange.Font.Color = colorValue
        Else
            targetShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorValue
        End If
    End If
    If properties.Exists("fontbold") Then
        If targetShape Is Nothing Then
            targetRange.Font.Bold = private_ReadBoolean(properties("fontbold"))
        Else
            targetShape.TextFrame2.TextRange.Font.Bold = _
                IIf(private_ReadBoolean(properties("fontbold")), -1, 0)
        End If
    End If
    If properties.Exists("fontitalic") Then
        If targetShape Is Nothing Then
            targetRange.Font.Italic = private_ReadBoolean(properties("fontitalic"))
        Else
            targetShape.TextFrame2.TextRange.Font.Italic = _
                IIf(private_ReadBoolean(properties("fontitalic")), -1, 0)
        End If
    End If
    If properties.Exists("fontsize") Then
        If VBA.IsNumeric(properties("fontsize")) Then
            If targetShape Is Nothing Then
                targetRange.Font.Size = VBA.CDbl(properties("fontsize"))
            Else
                targetShape.TextFrame2.TextRange.Font.Size = VBA.CDbl(properties("fontsize"))
            End If
        End If
    End If
    If properties.Exists("fontname") Then
        If targetShape Is Nothing Then
            targetRange.Font.Name = properties("fontname")
        Else
            targetShape.TextFrame2.TextRange.Font.Name = properties("fontname")
        End If
    End If
    If properties.Exists("horizontal") Then private_ApplyHorizontalAlignment _
        targetRange, targetShape, properties("horizontal")
    If properties.Exists("vertical") Then private_ApplyVerticalAlignment _
        targetRange, targetShape, properties("vertical")
    If properties.Exists("columnwidth") Then
        If Not targetRange Is Nothing And VBA.IsNumeric(properties("columnwidth")) Then _
            targetRange.EntireColumn.ColumnWidth = VBA.CDbl(properties("columnwidth"))
    End If
    If properties.Exists("width") Then
        If Not targetRange Is Nothing And VBA.IsNumeric(properties("width")) Then _
            targetRange.EntireColumn.ColumnWidth = VBA.CDbl(properties("width"))
    End If
    If properties.Exists("rowheight") Then
        If Not targetRange Is Nothing Then
            If VBA.LCase$(properties("rowheight")) = "auto" Then
                targetRange.EntireRow.AutoFit
            ElseIf VBA.IsNumeric(properties("rowheight")) Then
                targetRange.EntireRow.RowHeight = VBA.CDbl(properties("rowheight"))
            Else
                ex_UiStyleDiagnostics.fn_Raise "RowHeight"
            End If
        End If
    End If
    If Not targetRange Is Nothing Then
        If properties.Exists("autofitcolumns") Then
            If private_ReadBoolean(properties("autofitcolumns")) Then targetRange.Columns.AutoFit
        End If
        If properties.Exists("minwidth") Then
            private_ClampWidth targetRange, properties("minwidth"), True
        End If
        If properties.Exists("maxwidth") Then
            private_ClampWidth targetRange, properties("maxwidth"), False
        End If
        If properties.Exists("celltype") Then
            Select Case VBA.LCase$(properties("celltype"))
                Case "text"
                    targetRange.NumberFormat = "@"
                Case "date"
                    targetRange.NumberFormat = "yyyy-mm-dd"
                Case "general"
                    targetRange.NumberFormat = "General"
                Case Else
                    ex_UiStyleDiagnostics.fn_Raise "CellType"
            End Select
        End If
    End If
    If properties.Exists("overflow") And Not targetRange Is Nothing Then _
        private_ApplyOverflow targetRange, properties("overflow")
    private_ApplyBorders targetRange, targetShape, properties
End Sub

Private Sub private_ClampWidth( _
    ByVal area As Range, _
    ByVal value As String, _
    ByVal minimum As Boolean _
)
    Dim column As Range
    Dim width As Double

    If Not VBA.IsNumeric(value) Then ex_UiStyleDiagnostics.fn_Raise "Width"
    width = VBA.CDbl(value)
    For Each column In area.Columns
        If (minimum And column.ColumnWidth < width) Or (Not minimum And column.ColumnWidth > width) Then column.ColumnWidth = width
    Next column
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


Private Sub private_ApplyOverflow(ByVal targetRange As Range, ByVal overflowText As String)
    Select Case VBA.LCase$(VBA.Trim$(overflowText))
        Case "wrap"
            targetRange.WrapText = True
        Case "clip"
            targetRange.WrapText = False
        Case Else
            ex_WindowsUi.fn_ShowMessage "Unsupported overflow value: " & overflowText, _
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
            If targetShape Is Nothing Then
                targetRange.HorizontalAlignment = xlLeft
            Else
                targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            End If
        Case "right"
            If targetShape Is Nothing Then
                targetRange.HorizontalAlignment = xlRight
            Else
                targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
            End If
        Case Else
            If targetShape Is Nothing Then
                targetRange.HorizontalAlignment = xlCenter
            Else
                targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End If
    End Select
End Sub

Private Sub private_ApplyVerticalAlignment( _
    ByVal targetRange As Range, _
    ByVal targetShape As Object, _
    ByVal alignmentText As String _
)
    Select Case VBA.LCase$(alignmentText)
        Case "top"
            If targetShape Is Nothing Then
                targetRange.VerticalAlignment = xlTop
            Else
                targetShape.TextFrame2.VerticalAnchor = msoAnchorTop
            End If
        Case "bottom"
            If targetShape Is Nothing Then
                targetRange.VerticalAlignment = xlBottom
            Else
                targetShape.TextFrame2.VerticalAnchor = msoAnchorBottom
            End If
        Case Else
            If targetShape Is Nothing Then
                targetRange.VerticalAlignment = xlCenter
            Else
                targetShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
            End If
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


Private Function private_ReadAttribute( _
    ByVal node As Object, _
    ByVal attributeName As String _
) As String
    Dim attributeValue As Variant

    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function