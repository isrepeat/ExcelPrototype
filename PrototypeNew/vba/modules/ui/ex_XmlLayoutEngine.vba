Attribute VB_Name = "ex_XmlLayoutEngine"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_XmlLayoutEngine.m_Module_Dispose"
#End If
End Sub

' Layout handlers and XML/binding utilities.
' This module also routes visual node rendering by node kind.

' //
' // API
' //
Public Function m_RenderNode( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal layoutNode As Object, _
    Optional ByVal rowStart As Long = 0, _
    Optional ByVal colStart As Long = 0, _
    Optional ByVal rowEnd As Long = 0, _
    Optional ByVal colEnd As Long = 0 _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nodeKind As String

    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function

    If layoutNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: layout node is not specified."
#End If
        Exit Function
    End If
    If layoutNode.NodeType <> 1 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: layout node must be an element."
#End If
        Exit Function
    End If

    nodeKind = VBA.LCase$(VBA.CStr(layoutNode.baseName))
    Select Case nodeKind
        Case "page"
            m_RenderNode = ex_LayoutPageRenderer.m_Render(renderCtx, layoutNode)

        Case "control", "stackpanel", "grid", "list", "itemcontrol"
            If rowStart <= 0 Or colStart <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: invalid layout node position."
#End If
                Exit Function
            End If
            If rowEnd < rowStart Or colEnd < colStart Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.m_Diagnostic_LogError "PrototypeNew: invalid layout node bounds."
#End If
                Exit Function
            End If

            Select Case nodeKind
                Case "control"
                    m_RenderNode = ex_LayoutControlRenderer.m_Render( _
                        renderCtx:=renderCtx, _
                        layoutNode:=layoutNode, _
                        rowStart:=rowStart, _
                        colStart:=colStart, _
                        rowEnd:=rowEnd, _
                        colEnd:=colEnd)

                Case "stackpanel"
                    m_RenderNode = ex_LayoutStackPanelRenderer.m_Render( _
                        renderCtx:=renderCtx, _
                        layoutNode:=layoutNode, _
                        rowStart:=rowStart, _
                        colStart:=colStart, _
                        rowEnd:=rowEnd, _
                        colEnd:=colEnd)

                Case "grid"
                    m_RenderNode = ex_LayoutGridRenderer.m_Render( _
                        renderCtx:=renderCtx, _
                        layoutNode:=layoutNode, _
                        rowStart:=rowStart, _
                        colStart:=colStart, _
                        rowEnd:=rowEnd, _
                        colEnd:=colEnd)

                Case "list"
                    m_RenderNode = ex_LayoutListRenderer.m_Render( _
                        renderCtx:=renderCtx, _
                        layoutNode:=layoutNode, _
                        rowStart:=rowStart, _
                        colStart:=colStart, _
                        rowEnd:=rowEnd, _
                        colEnd:=colEnd)

                Case "itemcontrol"
                    m_RenderNode = ex_LayoutItemControlRenderer.m_Render( _
                        renderCtx:=renderCtx, _
                        layoutNode:=layoutNode, _
                        rowStart:=rowStart, _
                        colStart:=colStart, _
                        rowEnd:=rowEnd, _
                        colEnd:=colEnd)
            End Select

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: unsupported layout node '" & VBA.CStr(layoutNode.baseName) & "'."
#End If
    End Select
End Function


Public Function m_RenderTemplateChildren( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal templateControlNode As Object, _
    Optional ByVal layoutRowStart As Long = 0, _
    Optional ByVal layoutColStart As Long = 0, _
    Optional ByVal layoutRowEnd As Long = 0, _
    Optional ByVal layoutColEnd As Long = 0 _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet

    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function
    If templateControlNode Is Nothing Then Exit Function

    m_RenderTemplateChildren = private_RenderContainerChildrenInBounds( _
        renderCtx, templateControlNode, _
        layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)
End Function


Public Function m_TryResolveNodeBoundsFromAnchor( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal node As Object, _
    ByVal anchorCellAddr As String, _
    ByRef outRow As Long, _
    ByRef outCol As Long, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim anchorCell As Range

    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function
    If node Is Nothing Then Exit Function

    anchorCellAddr = VBA.Trim$(anchorCellAddr)
    If VBA.Len(anchorCellAddr) = 0 Then anchorCellAddr = "A1"

    On Error GoTo EH_ANCHOR
    Set anchorCell = ws.Range(anchorCellAddr)
    On Error GoTo 0

    If Not private_TryResolveNodeCellPosition(node, anchorCell, outRow, outCol) Then Exit Function
    If Not private_TryGetEffectiveNodeSpan(renderCtx, node, outSpanRows, outSpanCols) Then Exit Function

    m_TryResolveNodeBoundsFromAnchor = True
    Exit Function

EH_ANCHOR:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "PrototypeNew: invalid anchorCell '" & anchorCellAddr & "'."
#End If
End Function


Public Function m_TryGetEffectiveNodeSpan( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal node As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    m_TryGetEffectiveNodeSpan = private_TryGetEffectiveNodeSpan(renderCtx, node, outSpanRows, outSpanCols, dataContext)
End Function


Public Function m_IsVisualLayoutNode(ByVal node As Object) As Boolean
    m_IsVisualLayoutNode = private_IsVisualLayoutNode(node)
End Function


Public Function m_RenderNodeBySpan( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal layoutNode As Object, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long, _
    ByVal spanRows As Long, _
    ByVal spanCols As Long _
) As Boolean
    If spanRows <= 0 Or spanCols <= 0 Then
        m_RenderNodeBySpan = True
        Exit Function
    End If

    m_RenderNodeBySpan = m_RenderNode( _
        renderCtx:=renderCtx, _
        layoutNode:=layoutNode, _
        rowStart:=rowIndex, _
        colStart:=colIndex, _
        rowEnd:=rowIndex + spanRows - 1, _
        colEnd:=colIndex + spanCols - 1)
End Function


Public Function m_RenderNodeInBounds( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal layoutNode As Object, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
) As Boolean
    m_RenderNodeInBounds = m_RenderNode( _
        renderCtx:=renderCtx, _
        layoutNode:=layoutNode, _
        rowStart:=rowStart, _
        colStart:=colStart, _
        rowEnd:=rowEnd, _
        colEnd:=colEnd)
End Function


Public Function m_RenderContainerNodeInBounds( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal containerNode As Object, _
    Optional ByVal layoutRowStart As Long = 0, _
    Optional ByVal layoutColStart As Long = 0, _
    Optional ByVal layoutRowEnd As Long = 0, _
    Optional ByVal layoutColEnd As Long = 0 _
) As Boolean
    m_RenderContainerNodeInBounds = private_RenderContainerChildrenInBounds( _
        renderCtx, containerNode, _
        layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)
End Function

' //
' // Internal
' //
Private Function private_TryIsNodeVisible( _
    ByVal node As Object, _
    ByVal dataContext As Object, _
    ByRef outVisible As Boolean _
) As Boolean
    Dim visibilityRaw As String

    If node Is Nothing Then
        outVisible = True
        private_TryIsNodeVisible = True
        Exit Function
    End If

    visibilityRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(node, "visibility")))
    If VBA.Len(visibilityRaw) = 0 Then
        outVisible = True
        private_TryIsNodeVisible = True
        Exit Function
    End If

    ' visibility всегда вычисляется относительно текущего dataContext узла.
    ' Для вложенных list/itemControl этот контекст приходит от родительского итема.
    If Not ex_BindingRuntime.m_TryResolveVisibilityBinding(visibilityRaw, dataContext, outVisible) Then Exit Function
    private_TryIsNodeVisible = True
End Function


Private Function private_RenderContainerChildrenInBounds( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal containerNode As Object, _
    Optional ByVal containerRowStart As Long = 0, _
    Optional ByVal containerColStart As Long = 0, _
    Optional ByVal containerRowEnd As Long = 0, _
    Optional ByVal containerColEnd As Long = 0 _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim visualCount As Long
    Dim maxRows As Long
    Dim maxCols As Long
    Dim childNode As Object
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim spanRows As Long
    Dim spanCols As Long
    Dim orientation As String
    Dim seqRow As Long
    Dim seqCol As Long
    Dim hasGridBounds As Boolean
    Dim childRowStart As Long
    Dim childColStart As Long
    Dim childRowEnd As Long
    Dim childColEnd As Long

    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function
    If containerNode Is Nothing Then
        private_RenderContainerChildrenInBounds = True
        Exit Function
    End If

    visualCount = private_CountVisualChildren(containerNode)
    If visualCount = 0 Then
        private_RenderContainerChildrenInBounds = True
        Exit Function
    End If

    orientation = private_GetContainerOrientation(containerNode)
    hasGridBounds = (containerRowStart > 0 And containerColStart > 0 And containerRowEnd >= containerRowStart And containerColEnd >= containerColStart)
    If Not hasGridBounds Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: container bounds are required for nested layout rendering."
#End If
        Exit Function
    End If

    seqRow = 1
    seqCol = 1

    For Each childNode In containerNode.ChildNodes
        If Not private_IsVisualLayoutNode(childNode) Then GoTo ContinueFirstPass

        If Not private_TryGetEffectiveNodeSpan(renderCtx, childNode, spanRows, spanCols) Then Exit Function

        If Not private_ResolveChildGridPosition(childNode, orientation, seqRow, seqCol, rowIdx, colIdx, spanRows, spanCols) Then Exit Function
        If spanRows <= 0 Or spanCols <= 0 Then GoTo ContinueFirstPass

        If rowIdx + spanRows - 1 > maxRows Then maxRows = rowIdx + spanRows - 1
        If colIdx + spanCols - 1 > maxCols Then maxCols = colIdx + spanCols - 1

ContinueFirstPass:
    Next childNode

    If maxRows <= 0 Then maxRows = 1
    If maxCols <= 0 Then maxCols = 1

    seqRow = 1
    seqCol = 1

    For Each childNode In containerNode.ChildNodes
        If Not private_IsVisualLayoutNode(childNode) Then GoTo ContinueSecondPass

        If Not private_TryGetEffectiveNodeSpan(renderCtx, childNode, spanRows, spanCols) Then Exit Function

        If Not private_ResolveChildGridPosition(childNode, orientation, seqRow, seqCol, rowIdx, colIdx, spanRows, spanCols) Then Exit Function
        If spanRows <= 0 Or spanCols <= 0 Then GoTo ContinueSecondPass

        childRowStart = containerRowStart + rowIdx - 1
        childColStart = containerColStart + colIdx - 1
        childRowEnd = childRowStart + spanRows - 1
        childColEnd = childColStart + spanCols - 1

        If Not m_RenderNodeInBounds( _
            renderCtx:=renderCtx, _
            layoutNode:=childNode, _
            rowStart:=childRowStart, _
            colStart:=childColStart, _
            rowEnd:=childRowEnd, _
            colEnd:=childColEnd) Then Exit Function

ContinueSecondPass:
    Next childNode

    private_RenderContainerChildrenInBounds = True
End Function


Private Function private_TryGetEffectiveNodeSpan( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal node As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    Dim nodeKind As String
    Dim isVisible As Boolean
    Dim explicitRows As Long
    Dim explicitCols As Long
    Dim measuredRows As Long
    Dim measuredCols As Long

    If node Is Nothing Then Exit Function
    If Not private_TryIsNodeVisible(node, dataContext, isVisible) Then Exit Function
    If Not isVisible Then
        outSpanRows = 0
        outSpanCols = 0
        private_TryGetEffectiveNodeSpan = True
        Exit Function
    End If

    explicitRows = private_ReadPositiveLongAttr(node, "spanRows", 0)
    explicitCols = private_ReadPositiveLongAttr(node, "spanCells", 0)
    If explicitRows > 0 And explicitCols > 0 Then
        outSpanRows = explicitRows
        outSpanCols = explicitCols
        private_TryGetEffectiveNodeSpan = True
        Exit Function
    End If

    nodeKind = VBA.LCase$(VBA.CStr(node.baseName))
    Select Case nodeKind
        Case "control"
            If explicitRows > 0 Then
                outSpanRows = explicitRows
            Else
                outSpanRows = 1
            End If

            If explicitCols > 0 Then
                outSpanCols = explicitCols
            Else
                outSpanCols = 1
            End If

        Case "stackpanel", "grid"
            ' Пробрасываем тот же dataContext в измерение контейнера:
            ' visibility/binding дочерних узлов должны оцениваться в том же контексте.
            If Not private_TryMeasureContainerContentSpan(renderCtx, node, measuredRows, measuredCols, dataContext) Then Exit Function

            If explicitRows > 0 Then
                outSpanRows = explicitRows
            Else
                outSpanRows = measuredRows
            End If

            If explicitCols > 0 Then
                outSpanCols = explicitCols
            Else
                outSpanCols = measuredCols
            End If

        Case "list"
            ' list измеряется на данных itemsSource; dataContext нужен для Binding-ветки.
            If Not ex_LayoutListRenderer.m_TryMeasureContentSpan(renderCtx, node, measuredRows, measuredCols, dataContext) Then Exit Function

            If explicitRows > 0 Then
                outSpanRows = explicitRows
            Else
                outSpanRows = measuredRows
            End If

            If explicitCols > 0 Then
                outSpanCols = explicitCols
            Else
                outSpanCols = measuredCols
            End If

        Case "itemcontrol"
            ' itemControl измеряется в контексте objectSource (или item dataContext).
            If Not ex_LayoutItemControlRenderer.m_TryMeasureContentSpan(renderCtx, node, measuredRows, measuredCols, dataContext) Then Exit Function

            If explicitRows > 0 Then
                outSpanRows = explicitRows
            Else
                outSpanRows = measuredRows
            End If

            If explicitCols > 0 Then
                outSpanCols = explicitCols
            Else
                outSpanCols = measuredCols
            End If

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: unsupported layout node '" & VBA.CStr(node.baseName) & "'."
#End If
            Exit Function
    End Select

    If VBA.StrComp(nodeKind, "itemcontrol", VBA.vbBinaryCompare) = 0 Then
        If outSpanRows < 0 Then outSpanRows = 0
        If outSpanCols < 0 Then outSpanCols = 0
    Else
        If outSpanRows <= 0 Then outSpanRows = 1
        If outSpanCols <= 0 Then outSpanCols = 1
    End If
    private_TryGetEffectiveNodeSpan = True
End Function


Private Function private_TryMeasureContainerContentSpan( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal containerNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    Dim orientation As String
    Dim seqRow As Long
    Dim seqCol As Long
    Dim childNode As Object
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim childRows As Long
    Dim childCols As Long
    Dim maxRows As Long
    Dim maxCols As Long

    If containerNode Is Nothing Then Exit Function

    orientation = private_GetContainerOrientation(containerNode)
    If VBA.StrComp(VBA.LCase$(VBA.CStr(containerNode.baseName)), "stackpanel", VBA.vbBinaryCompare) = 0 And VBA.Len(orientation) = 0 Then
        Exit Function
    End If

    seqRow = 1
    seqCol = 1

    For Each childNode In containerNode.ChildNodes
        If Not private_IsVisualLayoutNode(childNode) Then GoTo ContinueChild

        ' Передаем родительский dataContext вниз по дереву измерения.
        If Not private_TryGetEffectiveNodeSpan(renderCtx, childNode, childRows, childCols, dataContext) Then Exit Function
        If Not private_ResolveChildGridPosition(childNode, orientation, seqRow, seqCol, rowIdx, colIdx, childRows, childCols) Then Exit Function
        If childRows <= 0 Or childCols <= 0 Then GoTo ContinueChild

        If rowIdx + childRows - 1 > maxRows Then maxRows = rowIdx + childRows - 1
        If colIdx + childCols - 1 > maxCols Then maxCols = colIdx + childCols - 1

ContinueChild:
    Next childNode

    If maxRows <= 0 Then maxRows = 1
    If maxCols <= 0 Then maxCols = 1

    outSpanRows = maxRows
    outSpanCols = maxCols
    private_TryMeasureContainerContentSpan = True
End Function


Private Function private_ResolveChildGridPosition( _
    ByVal childNode As Object, _
    ByVal parentOrientation As String, _
    ByRef seqRow As Long, _
    ByRef seqCol As Long, _
    ByRef outRow As Long, _
    ByRef outCol As Long, _
    ByVal spanRows As Long, _
    ByVal spanCols As Long _
) As Boolean
    Dim atText As String

    atText = VBA.Trim$(ex_XmlCore.m_NodeAttrText(childNode, "at"))
    If VBA.Len(atText) > 0 Then
        If Not private_TryParseAtAddress(atText, outRow, outCol) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: invalid 'at' format '" & atText & "'. Expected format is rNcM."
#End If
            Exit Function
        End If
    Else
        Select Case parentOrientation
            Case "horizontal"
                outRow = 1
                outCol = seqCol
                seqCol = seqCol + spanCols
            Case "vertical"
                outRow = seqRow
                outCol = 1
                seqRow = seqRow + spanRows
            Case Else
                outRow = 1
                outCol = 1
        End Select
    End If

    private_ResolveChildGridPosition = True
End Function


Private Function private_GetContainerOrientation(ByVal node As Object) As String
    If node Is Nothing Then Exit Function

    If VBA.StrComp(VBA.LCase$(VBA.CStr(node.baseName)), "stackpanel", VBA.vbBinaryCompare) = 0 Then
        private_GetContainerOrientation = VBA.LCase$(VBA.Trim$(ex_XmlCore.m_NodeAttrText(node, "orientation")))
        If VBA.Len(private_GetContainerOrientation) = 0 Then
            private_GetContainerOrientation = "vertical"
        ElseIf VBA.StrComp(private_GetContainerOrientation, "vertical", VBA.vbBinaryCompare) <> 0 And _
               VBA.StrComp(private_GetContainerOrientation, "horizontal", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: stackPanel orientation must be 'vertical' or 'horizontal'."
#End If
            private_GetContainerOrientation = VBA.vbNullString
        End If
    End If
End Function


Private Function private_CountVisualChildren(ByVal node As Object) As Long
    Dim childNode As Object

    If node Is Nothing Then Exit Function

    For Each childNode In node.ChildNodes
        If private_IsVisualLayoutNode(childNode) Then
            private_CountVisualChildren = private_CountVisualChildren + 1
        End If
    Next childNode
End Function


Private Function private_IsVisualLayoutNode(ByVal node As Object) As Boolean
    If node Is Nothing Then Exit Function
    If node.NodeType <> 1 Then Exit Function

    Select Case VBA.LCase$(VBA.CStr(node.baseName))
        Case "control", "stackpanel", "grid", "list", "itemcontrol"
            private_IsVisualLayoutNode = True
    End Select
End Function


Private Function private_TryResolveNodeCellPosition( _
    ByVal node As Object, _
    ByVal anchorCell As Range, _
    ByRef outRow As Long, _
    ByRef outCol As Long _
) As Boolean
    Dim atText As String
    Dim relRow As Long
    Dim relCol As Long

    If node Is Nothing Then Exit Function
    If anchorCell Is Nothing Then Exit Function

    atText = VBA.Trim$(ex_XmlCore.m_NodeAttrText(node, "at"))
    If VBA.Len(atText) = 0 Then
        relRow = 1
        relCol = 1
    Else
        If Not private_TryParseAtAddress(atText, relRow, relCol) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: invalid 'at' format '" & atText & "'. Expected format is rNcM."
#End If
            Exit Function
        End If
    End If

    outRow = anchorCell.Row + relRow - 1
    outCol = anchorCell.Column + relCol - 1
    private_TryResolveNodeCellPosition = True
End Function


Private Function private_TryParseAtAddress( _
    ByVal atText As String, _
    ByRef outRelRow As Long, _
    ByRef outRelCol As Long _
) As Boolean
    Dim normalized As String
    Dim cPos As Long
    Dim rowText As String
    Dim colText As String

    normalized = VBA.LCase$(VBA.Trim$(atText))
    If VBA.Len(normalized) < 4 Then Exit Function
    If VBA.Left$(normalized, 1) <> "r" Then Exit Function

    cPos = VBA.InStr(2, normalized, "c", VBA.vbBinaryCompare)
    If cPos <= 2 Then Exit Function
    If cPos >= VBA.Len(normalized) Then Exit Function

    rowText = VBA.Mid$(normalized, 2, cPos - 2)
    colText = VBA.Mid$(normalized, cPos + 1)

    If Not VBA.IsNumeric(rowText) Then Exit Function
    If Not VBA.IsNumeric(colText) Then Exit Function

    outRelRow = VBA.CLng(rowText)
    outRelCol = VBA.CLng(colText)
    If outRelRow <= 0 Or outRelCol <= 0 Then Exit Function

    private_TryParseAtAddress = True
End Function


Private Function private_ReadPositiveLongAttr( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Long _
) As Long
    Dim rawText As String

    rawText = VBA.Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If VBA.Len(rawText) = 0 Then
        private_ReadPositiveLongAttr = defaultValue
        Exit Function
    End If

    If Not VBA.IsNumeric(rawText) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: attribute '" & attrName & "' must be numeric."
#End If
        Exit Function
    End If

    private_ReadPositiveLongAttr = VBA.CLng(rawText)
    If private_ReadPositiveLongAttr <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: attribute '" & attrName & "' must be greater than zero."
#End If
        private_ReadPositiveLongAttr = 0
    End If
End Function


Private Function private_TryGetPageRenderContext( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByRef outWb As Workbook, _
    ByRef outWs As Worksheet _
) As Boolean
    Set outWb = Nothing
    Set outWs = Nothing

    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    Set outWs = renderCtx.Worksheet
    If outWs Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: worksheet is not specified."
#End If
        Exit Function
    End If

    Set outWb = renderCtx.Workbook
    If outWb Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: workbook is not specified."
#End If
        Exit Function
    End If

    private_TryGetPageRenderContext = True
End Function
