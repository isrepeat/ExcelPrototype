Attribute VB_Name = "ex_XmlLayoutEngine"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private g_ListRuntimeSourceSeed As Long
Private g_ObjectRuntimeSourceSeed As Long

Public Function m_RenderPageLayout( _
    ByVal wb As Workbook, _
    ByVal defaultWs As Worksheet, _
    ByVal wsUiDoc As Object _
) As Boolean
    Dim gridNode As Object

    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Function
    End If
    If defaultWs Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified.", vbExclamation
        Exit Function
    End If
    If wsUiDoc Is Nothing Then
        MsgBox "PrototypeNew: page UI document is not specified.", vbExclamation
        Exit Function
    End If

    Set gridNode = wsUiDoc.selectSingleNode("/p:uiDefinition/p:layout/p:grid")
    If gridNode Is Nothing Then
        MsgBox "PrototypeNew: layout must contain <grid> node.", vbExclamation
        Exit Function
    End If

    ' Reuse runtime keys between renders to avoid unbounded key growth.
    g_ListRuntimeSourceSeed = 0
    g_ObjectRuntimeSourceSeed = 0

    m_RenderPageLayout = mp_RenderGridNodeOnWorksheet(wb, defaultWs, gridNode)
End Function

Public Function m_RenderTemplateChildren( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal templateControlNode As Object, _
    ByVal recursionDepth As Long, _
    Optional ByVal layoutRowStart As Long = 0, _
    Optional ByVal layoutColStart As Long = 0, _
    Optional ByVal layoutRowEnd As Long = 0, _
    Optional ByVal layoutColEnd As Long = 0 _
) As Boolean
    If wb Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If templateControlNode Is Nothing Then Exit Function

    m_RenderTemplateChildren = mp_RenderContainerChildrenInBounds( _
        wb, ws, templateControlNode, recursionDepth, _
        layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)
End Function

Private Function mp_RenderGridNodeOnWorksheet( _
    ByVal wb As Workbook, _
    ByVal defaultWs As Worksheet, _
    ByVal gridNode As Object _
) As Boolean
    Dim ws As Worksheet
    Dim anchorCell As Range
    Dim anchorCellAddr As String
    Dim childNode As Object

    If wb Is Nothing Then Exit Function
    If defaultWs Is Nothing Then Exit Function
    If gridNode Is Nothing Then Exit Function

    Set ws = defaultWs

    anchorCellAddr = Trim$(ex_XmlCore.m_NodeAttrText(gridNode, "anchorCell"))
    If Len(anchorCellAddr) = 0 Then anchorCellAddr = "A1"

    On Error GoTo EH_ANCHOR
    Set anchorCell = ws.Range(anchorCellAddr)
    On Error GoTo 0

    mp_ClearWorksheet ws

    For Each childNode In gridNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueNode

        Select Case LCase$(CStr(childNode.baseName))
            Case "stackpanel"
                If Not mp_RenderStackPanelOnWorksheet(wb, ws, anchorCell, childNode) Then Exit Function

            Case "control"
                If Not mp_RenderSingleControlOnWorksheet(wb, ws, anchorCell, childNode) Then Exit Function

            Case "list"
                If Not mp_RenderSingleListOnWorksheet(wb, ws, anchorCell, childNode) Then Exit Function

            Case "itemcontrol"
                If Not mp_RenderSingleItemControlOnWorksheet(wb, ws, anchorCell, childNode) Then Exit Function

            Case Else
                MsgBox "PrototypeNew: unsupported node '" & CStr(childNode.baseName) & "' inside <grid>.", vbExclamation
                Exit Function
        End Select

ContinueNode:
    Next childNode

    mp_RenderGridNodeOnWorksheet = True
    Exit Function

EH_ANCHOR:
    MsgBox "PrototypeNew: invalid grid anchorCell '" & anchorCellAddr & "'.", vbExclamation
End Function

Private Function mp_RenderStackPanelOnWorksheet( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal anchorCell As Range, _
    ByVal panelNode As Object _
) As Boolean
    Dim panelStartRow As Long
    Dim panelStartCol As Long
    Dim cursorRow As Long
    Dim cursorCol As Long
    Dim orientation As String
    Dim childNode As Object
    Dim spanRows As Long
    Dim spanCols As Long

    If wb Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If anchorCell Is Nothing Then Exit Function
    If panelNode Is Nothing Then Exit Function

    If Not mp_TryResolveNodeCellPosition(panelNode, anchorCell, panelStartRow, panelStartCol) Then Exit Function

    orientation = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(panelNode, "orientation")))
    If Len(orientation) = 0 Then orientation = "vertical"
    If StrComp(orientation, "vertical", vbBinaryCompare) <> 0 And _
       StrComp(orientation, "horizontal", vbBinaryCompare) <> 0 Then
        MsgBox "PrototypeNew: stackPanel orientation must be 'vertical' or 'horizontal'.", vbExclamation
        Exit Function
    End If

    cursorRow = panelStartRow
    cursorCol = panelStartCol

    For Each childNode In panelNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueNode

        If Not mp_IsVisualLayoutNode(childNode) Then
            MsgBox "PrototypeNew: stackPanel supports only <control>, <stackPanel>, <grid>, <list> or <itemControl> children.", vbExclamation
            Exit Function
        End If

        If Not mp_TryGetEffectiveNodeSpan(childNode, spanRows, spanCols) Then Exit Function

        If Not mp_RenderLayoutNodeByWorksheetSpan(wb, ws, childNode, cursorRow, cursorCol, spanRows, spanCols) Then Exit Function

        If StrComp(orientation, "horizontal", vbBinaryCompare) = 0 Then
            cursorCol = cursorCol + spanCols
        Else
            cursorRow = cursorRow + spanRows
        End If

ContinueNode:
    Next childNode

    mp_RenderStackPanelOnWorksheet = True
End Function

Private Function mp_RenderSingleListOnWorksheet( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal anchorCell As Range, _
    ByVal listNode As Object _
) As Boolean
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim spanRows As Long
    Dim spanCols As Long

    If Not mp_TryResolveNodeCellPosition(listNode, anchorCell, rowIndex, colIndex) Then Exit Function
    If Not mp_TryGetEffectiveNodeSpan(listNode, spanRows, spanCols) Then Exit Function

    mp_RenderSingleListOnWorksheet = mp_RenderLayoutNodeByWorksheetSpan( _
        wb, ws, listNode, rowIndex, colIndex, spanRows, spanCols)
End Function

Private Function mp_RenderSingleItemControlOnWorksheet( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal anchorCell As Range, _
    ByVal itemControlNode As Object _
) As Boolean
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim spanRows As Long
    Dim spanCols As Long

    If Not mp_TryResolveNodeCellPosition(itemControlNode, anchorCell, rowIndex, colIndex) Then Exit Function
    If Not mp_TryGetEffectiveNodeSpan(itemControlNode, spanRows, spanCols) Then Exit Function

    mp_RenderSingleItemControlOnWorksheet = mp_RenderLayoutNodeByWorksheetSpan( _
        wb, ws, itemControlNode, rowIndex, colIndex, spanRows, spanCols)
End Function

Private Function mp_RenderSingleControlOnWorksheet( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal anchorCell As Range, _
    ByVal controlNode As Object _
) As Boolean
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim spanRows As Long
    Dim spanCols As Long

    If Not mp_TryResolveNodeCellPosition(controlNode, anchorCell, rowIndex, colIndex) Then Exit Function

    spanRows = mp_ReadPositiveLongAttr(controlNode, "spanRows", 1)
    spanCols = mp_ReadPositiveLongAttr(controlNode, "spanCells", 1)
    If spanRows <= 0 Or spanCols <= 0 Then Exit Function

    mp_RenderSingleControlOnWorksheet = mp_RenderControlByWorksheetSpan(wb, ws, controlNode, rowIndex, colIndex, spanRows, spanCols)
End Function

Private Function mp_RenderControlByWorksheetSpan( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal controlNode As Object, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long, _
    ByVal spanRows As Long, _
    ByVal spanCols As Long _
) As Boolean
    Dim targetRange As Range

    If rowIndex <= 0 Or colIndex <= 0 Then
        MsgBox "PrototypeNew: invalid control position.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(rowIndex, colIndex), ws.Cells(rowIndex + spanRows - 1, colIndex + spanCols - 1))
    On Error GoTo 0

    mp_RenderControlByWorksheetSpan = ex_ControlRenderer.m_RenderControl( _
        wb:=wb, _
        ws:=ws, _
        layoutControlNode:=controlNode, _
        recursionDepth:=0, _
        layoutRowStart:=rowIndex, _
        layoutColStart:=colIndex, _
        layoutRowEnd:=rowIndex + spanRows - 1, _
        layoutColEnd:=colIndex + spanCols - 1)
    Exit Function

EH_RANGE:
    MsgBox "PrototypeNew: failed to resolve control range for row=" & CStr(rowIndex) & ", col=" & CStr(colIndex) & ".", vbExclamation
End Function

Private Function mp_RenderLayoutNodeByWorksheetSpan( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal layoutNode As Object, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long, _
    ByVal spanRows As Long, _
    ByVal spanCols As Long _
) As Boolean
    Dim targetRange As Range
    Dim nodeKind As String

    If rowIndex <= 0 Or colIndex <= 0 Then
        MsgBox "PrototypeNew: invalid layout node position.", vbExclamation
        Exit Function
    End If
    If spanRows <= 0 Or spanCols <= 0 Then
        mp_RenderLayoutNodeByWorksheetSpan = True
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(rowIndex, colIndex), ws.Cells(rowIndex + spanRows - 1, colIndex + spanCols - 1))
    On Error GoTo 0

    nodeKind = LCase$(CStr(layoutNode.baseName))
    Select Case nodeKind
        Case "control"
            mp_RenderLayoutNodeByWorksheetSpan = ex_ControlRenderer.m_RenderControl( _
                wb:=wb, _
                ws:=ws, _
                layoutControlNode:=layoutNode, _
                recursionDepth:=0, _
                layoutRowStart:=rowIndex, _
                layoutColStart:=colIndex, _
                layoutRowEnd:=rowIndex + spanRows - 1, _
                layoutColEnd:=colIndex + spanCols - 1)

        Case "stackpanel", "grid"
            mp_RenderLayoutNodeByWorksheetSpan = mp_RenderContainerChildrenInBounds( _
                wb:=wb, _
                ws:=ws, _
                containerNode:=layoutNode, _
                recursionDepth:=0, _
                containerRowStart:=rowIndex, _
                containerColStart:=colIndex, _
                containerRowEnd:=rowIndex + spanRows - 1, _
                containerColEnd:=colIndex + spanCols - 1)

        Case "list"
            mp_RenderLayoutNodeByWorksheetSpan = mp_RenderListInBounds( _
                wb:=wb, _
                ws:=ws, _
                listNode:=layoutNode, _
                recursionDepth:=0, _
                layoutRowStart:=rowIndex, _
                layoutColStart:=colIndex, _
                layoutRowEnd:=rowIndex + spanRows - 1, _
                layoutColEnd:=colIndex + spanCols - 1)

        Case "itemcontrol"
            mp_RenderLayoutNodeByWorksheetSpan = mp_RenderItemControlInBounds( _
                wb:=wb, _
                ws:=ws, _
                itemControlNode:=layoutNode, _
                recursionDepth:=0, _
                layoutRowStart:=rowIndex, _
                layoutColStart:=colIndex, _
                layoutRowEnd:=rowIndex + spanRows - 1, _
                layoutColEnd:=colIndex + spanCols - 1)

        Case Else
            MsgBox "PrototypeNew: unsupported layout node '" & CStr(layoutNode.baseName) & "'.", vbExclamation
    End Select
    Exit Function

EH_RANGE:
    MsgBox "PrototypeNew: failed to resolve layout range for row=" & CStr(rowIndex) & ", col=" & CStr(colIndex) & ".", vbExclamation
End Function

Private Function mp_RenderContainerChildrenInBounds( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal containerNode As Object, _
    ByVal recursionDepth As Long, _
    Optional ByVal containerRowStart As Long = 0, _
    Optional ByVal containerColStart As Long = 0, _
    Optional ByVal containerRowEnd As Long = 0, _
    Optional ByVal containerColEnd As Long = 0 _
) As Boolean
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

    If containerNode Is Nothing Then
        mp_RenderContainerChildrenInBounds = True
        Exit Function
    End If

    visualCount = mp_CountVisualChildren(containerNode)
    If visualCount = 0 Then
        mp_RenderContainerChildrenInBounds = True
        Exit Function
    End If

    orientation = mp_GetContainerOrientation(containerNode)
    hasGridBounds = (containerRowStart > 0 And containerColStart > 0 And containerRowEnd >= containerRowStart And containerColEnd >= containerColStart)
    If Not hasGridBounds Then
        MsgBox "PrototypeNew: container bounds are required for nested layout rendering.", vbExclamation
        Exit Function
    End If

    seqRow = 1
    seqCol = 1

    For Each childNode In containerNode.ChildNodes
        If Not mp_IsVisualLayoutNode(childNode) Then GoTo ContinueFirstPass

        If Not mp_TryGetEffectiveNodeSpan(childNode, spanRows, spanCols) Then Exit Function

        If Not mp_ResolveChildGridPosition(childNode, orientation, seqRow, seqCol, rowIdx, colIdx, spanRows, spanCols) Then Exit Function
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
        If Not mp_IsVisualLayoutNode(childNode) Then GoTo ContinueSecondPass

        If Not mp_TryGetEffectiveNodeSpan(childNode, spanRows, spanCols) Then Exit Function

        If Not mp_ResolveChildGridPosition(childNode, orientation, seqRow, seqCol, rowIdx, colIdx, spanRows, spanCols) Then Exit Function
        If spanRows <= 0 Or spanCols <= 0 Then GoTo ContinueSecondPass

        childRowStart = containerRowStart + rowIdx - 1
        childColStart = containerColStart + colIdx - 1
        childRowEnd = childRowStart + spanRows - 1
        childColEnd = childColStart + spanCols - 1

        Select Case LCase$(CStr(childNode.baseName))
            Case "control"
                If Not ex_ControlRenderer.m_RenderControl( _
                    wb, ws, childNode, recursionDepth, _
                    childRowStart, childColStart, childRowEnd, childColEnd) Then Exit Function

            Case "stackpanel", "grid"
                If Not mp_RenderContainerChildrenInBounds( _
                    wb, ws, childNode, recursionDepth, _
                    childRowStart, childColStart, childRowEnd, childColEnd) Then Exit Function

            Case "list"
                If Not mp_RenderListInBounds( _
                    wb, ws, childNode, recursionDepth, _
                    childRowStart, childColStart, childRowEnd, childColEnd) Then Exit Function

            Case "itemcontrol"
                If Not mp_RenderItemControlInBounds( _
                    wb, ws, childNode, recursionDepth, _
                    childRowStart, childColStart, childRowEnd, childColEnd) Then Exit Function
        End Select

ContinueSecondPass:
    Next childNode

    mp_RenderContainerChildrenInBounds = True
End Function

Private Function mp_TryGetEffectiveNodeSpan( _
    ByVal node As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal bindingSource As Object _
) As Boolean
    Dim nodeKind As String
    Dim explicitRows As Long
    Dim explicitCols As Long
    Dim measuredRows As Long
    Dim measuredCols As Long

    If node Is Nothing Then Exit Function

    explicitRows = mp_ReadPositiveLongAttr(node, "spanRows", 0)
    explicitCols = mp_ReadPositiveLongAttr(node, "spanCells", 0)
    If explicitRows > 0 And explicitCols > 0 Then
        outSpanRows = explicitRows
        outSpanCols = explicitCols
        mp_TryGetEffectiveNodeSpan = True
        Exit Function
    End If

    nodeKind = LCase$(CStr(node.baseName))
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
            If Not mp_TryMeasureContainerContentSpan(node, measuredRows, measuredCols, bindingSource) Then Exit Function

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
            If Not mp_TryMeasureListContentSpan(node, measuredRows, measuredCols, bindingSource) Then Exit Function

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
            If Not mp_TryMeasureItemControlContentSpan(node, measuredRows, measuredCols, bindingSource) Then Exit Function

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
            MsgBox "PrototypeNew: unsupported layout node '" & CStr(node.baseName) & "'.", vbExclamation
            Exit Function
    End Select

    If StrComp(nodeKind, "itemcontrol", vbBinaryCompare) = 0 Then
        If outSpanRows < 0 Then outSpanRows = 0
        If outSpanCols < 0 Then outSpanCols = 0
    Else
        If outSpanRows <= 0 Then outSpanRows = 1
        If outSpanCols <= 0 Then outSpanCols = 1
    End If
    mp_TryGetEffectiveNodeSpan = True
End Function

Private Function mp_TryMeasureContainerContentSpan( _
    ByVal containerNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal bindingSource As Object _
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

    orientation = mp_GetContainerOrientation(containerNode)
    If StrComp(LCase$(CStr(containerNode.baseName)), "stackpanel", vbBinaryCompare) = 0 And Len(orientation) = 0 Then
        Exit Function
    End If

    seqRow = 1
    seqCol = 1

    For Each childNode In containerNode.ChildNodes
        If Not mp_IsVisualLayoutNode(childNode) Then GoTo ContinueChild

        If Not mp_TryGetEffectiveNodeSpan(childNode, childRows, childCols, bindingSource) Then Exit Function
        If Not mp_ResolveChildGridPosition(childNode, orientation, seqRow, seqCol, rowIdx, colIdx, childRows, childCols) Then Exit Function
        If childRows <= 0 Or childCols <= 0 Then GoTo ContinueChild

        If rowIdx + childRows - 1 > maxRows Then maxRows = rowIdx + childRows - 1
        If colIdx + childCols - 1 > maxCols Then maxCols = colIdx + childCols - 1

ContinueChild:
    Next childNode

    If maxRows <= 0 Then maxRows = 1
    If maxCols <= 0 Then maxCols = 1

    outSpanRows = maxRows
    outSpanCols = maxCols
    mp_TryMeasureContainerContentSpan = True
End Function

Private Function mp_ResolveChildGridPosition( _
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

    atText = Trim$(ex_XmlCore.m_NodeAttrText(childNode, "at"))
    If Len(atText) > 0 Then
        If Not mp_TryParseAtAddress(atText, outRow, outCol) Then
            MsgBox "PrototypeNew: invalid 'at' format '" & atText & "'. Expected format is rNcM.", vbExclamation
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

    mp_ResolveChildGridPosition = True
End Function

Private Function mp_GetContainerOrientation(ByVal node As Object) As String
    If node Is Nothing Then Exit Function

    If StrComp(LCase$(CStr(node.baseName)), "stackpanel", vbBinaryCompare) = 0 Then
        mp_GetContainerOrientation = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(node, "orientation")))
        If Len(mp_GetContainerOrientation) = 0 Then
            mp_GetContainerOrientation = "vertical"
        ElseIf StrComp(mp_GetContainerOrientation, "vertical", vbBinaryCompare) <> 0 And _
               StrComp(mp_GetContainerOrientation, "horizontal", vbBinaryCompare) <> 0 Then
            MsgBox "PrototypeNew: stackPanel orientation must be 'vertical' or 'horizontal'.", vbExclamation
            mp_GetContainerOrientation = vbNullString
        End If
    End If
End Function

Private Function mp_CountVisualChildren(ByVal node As Object) As Long
    Dim childNode As Object

    If node Is Nothing Then Exit Function

    For Each childNode In node.ChildNodes
        If mp_IsVisualLayoutNode(childNode) Then
            mp_CountVisualChildren = mp_CountVisualChildren + 1
        End If
    Next childNode
End Function

Private Function mp_IsVisualLayoutNode(ByVal node As Object) As Boolean
    If node Is Nothing Then Exit Function
    If node.NodeType <> 1 Then Exit Function

    Select Case LCase$(CStr(node.baseName))
        Case "control", "stackpanel", "grid", "list", "itemcontrol"
            mp_IsVisualLayoutNode = True
    End Select
End Function

Private Function mp_RenderListInBounds( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal listNode As Object, _
    ByVal recursionDepth As Long, _
    Optional ByVal layoutRowStart As Long = 0, _
    Optional ByVal layoutColStart As Long = 0, _
    Optional ByVal layoutRowEnd As Long = 0, _
    Optional ByVal layoutColEnd As Long = 0 _
) As Boolean
    Dim items As Collection
    Dim templateRoot As Object
    Dim listOrientation As String
    Dim tempDoc As Object
    Dim syntheticRoot As Object
    Dim itemValue As Variant
    Dim clonedNode As Object
    Dim itemIndex As Long

    If wb Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If listNode Is Nothing Then Exit Function

    If Not ex_ListItemsSourceRuntime.m_TryResolveItemsSource( _
        ex_XmlCore.m_NodeAttrText(listNode, "itemsSource"), items) Then Exit Function

    If items Is Nothing Then
        MsgBox "PrototypeNew: list itemsSource resolved to Nothing.", vbExclamation
        Exit Function
    End If
    If items.Count = 0 Then
        mp_RenderListInBounds = True
        Exit Function
    End If

    If Not mp_TryResolveListTemplateRoot(listNode, templateRoot) Then Exit Function

    listOrientation = mp_GetListOrientation(listNode)
    If Len(listOrientation) = 0 Then Exit Function

    Set tempDoc = ex_XmlCore.m_CreateDom(UI_NS)
    mp_CopyTemplatesToTempListDoc tempDoc, listNode.OwnerDocument

    Set syntheticRoot = tempDoc.createNode(1, "stackPanel", UI_NS)
    syntheticRoot.setAttribute "orientation", listOrientation

    itemIndex = 0
    For Each itemValue In items
        itemIndex = itemIndex + 1

        Set clonedNode = tempDoc.importNode(templateRoot, True)
        If clonedNode Is Nothing Then
            MsgBox "PrototypeNew: failed to clone list template node.", vbExclamation
            Exit Function
        End If

        If Not mp_ApplyListItemBindings(clonedNode, itemValue) Then Exit Function
        mp_AppendSuffixToControlNames clonedNode, "_" & CStr(itemIndex)
        mp_ApplyListItemValueToTemplate clonedNode, itemValue
        syntheticRoot.appendChild clonedNode
    Next itemValue

    mp_RenderListInBounds = mp_RenderContainerChildrenInBounds( _
        wb, ws, syntheticRoot, recursionDepth, _
        layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)
End Function

Private Function mp_RenderItemControlInBounds( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal itemControlNode As Object, _
    ByVal recursionDepth As Long, _
    Optional ByVal layoutRowStart As Long = 0, _
    Optional ByVal layoutColStart As Long = 0, _
    Optional ByVal layoutRowEnd As Long = 0, _
    Optional ByVal layoutColEnd As Long = 0 _
) As Boolean
    Dim sourceObject As Object
    Dim templateRoot As Object
    Dim tempDoc As Object
    Dim syntheticRoot As Object
    Dim clonedNode As Object
    Dim suffixValue As Long

    If wb Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If itemControlNode Is Nothing Then Exit Function

    If Not ex_ObjectSourceRuntime.m_TryResolveObjectSource( _
        ex_XmlCore.m_NodeAttrText(itemControlNode, "objectSource"), _
        sourceObject, _
        True) Then Exit Function
    If sourceObject Is Nothing Then
        mp_RenderItemControlInBounds = True
        Exit Function
    End If

    If Not mp_TryResolveItemControlTemplateRoot(itemControlNode, templateRoot) Then Exit Function

    Set tempDoc = ex_XmlCore.m_CreateDom(UI_NS)
    mp_CopyTemplatesToTempListDoc tempDoc, itemControlNode.OwnerDocument

    Set syntheticRoot = tempDoc.createNode(1, "stackPanel", UI_NS)
    syntheticRoot.setAttribute "orientation", "vertical"

    Set clonedNode = tempDoc.importNode(templateRoot, True)
    If clonedNode Is Nothing Then
        MsgBox "PrototypeNew: failed to clone itemControl template node.", vbExclamation
        Exit Function
    End If

    If Not mp_ApplyNodeBindingsRecursive(clonedNode, sourceObject) Then Exit Function
    suffixValue = g_ObjectRuntimeSourceSeed + 1
    mp_AppendSuffixToControlNames clonedNode, "_obj" & CStr(suffixValue)
    syntheticRoot.appendChild clonedNode

    mp_RenderItemControlInBounds = mp_RenderContainerChildrenInBounds( _
        wb, ws, syntheticRoot, recursionDepth, _
        layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)
End Function

Private Sub mp_CopyTemplatesToTempListDoc(ByVal targetDoc As Object, ByVal sourceDoc As Object)
    Dim targetRoot As Object
    Dim targetTemplatesNode As Object
    Dim srcTemplateNodes As Object
    Dim srcTemplateNode As Object
    Dim templateName As String

    If targetDoc Is Nothing Then Exit Sub
    If sourceDoc Is Nothing Then Exit Sub

    Set targetRoot = targetDoc.selectSingleNode("/p:uiDefinition")
    If targetRoot Is Nothing Then
        Set targetRoot = targetDoc.createNode(1, "uiDefinition", UI_NS)
        targetRoot.setAttribute "version", "1"
        targetDoc.appendChild targetRoot
    End If

    Set targetTemplatesNode = targetRoot.selectSingleNode("p:templates")
    If targetTemplatesNode Is Nothing Then
        Set targetTemplatesNode = targetDoc.createNode(1, "templates", UI_NS)
        targetRoot.appendChild targetTemplatesNode
    End If

    Set srcTemplateNodes = sourceDoc.selectNodes("/p:uiDefinition/p:templates/p:template")
    If srcTemplateNodes Is Nothing Then Exit Sub

    For Each srcTemplateNode In srcTemplateNodes
        templateName = Trim$(ex_XmlCore.m_NodeAttrText(srcTemplateNode, "name"))
        If Len(templateName) = 0 Then GoTo ContinueTemplate

        If Not targetTemplatesNode.selectSingleNode("p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]") Is Nothing Then
            GoTo ContinueTemplate
        End If

        targetTemplatesNode.appendChild targetDoc.importNode(srcTemplateNode, True)

ContinueTemplate:
    Next srcTemplateNode
End Sub

Private Function mp_TryResolveListTemplateRoot(ByVal listNode As Object, ByRef outTemplateRoot As Object) As Boolean
    Dim templateName As String
    Dim ownerDoc As Object
    Dim templateNode As Object
    Dim rootNodes As Object

    templateName = Trim$(ex_XmlCore.m_NodeAttrText(listNode, "itemsSourceTemplate"))
    If Len(templateName) = 0 Then
        MsgBox "PrototypeNew: list requires non-empty attribute 'itemsSourceTemplate'.", vbExclamation
        Exit Function
    End If

    Set ownerDoc = listNode.OwnerDocument
    If ownerDoc Is Nothing Then
        MsgBox "PrototypeNew: failed to resolve owner document for list template.", vbExclamation
        Exit Function
    End If

    Set templateNode = ownerDoc.selectSingleNode( _
        "/p:uiDefinition/p:templates/p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]")
    If templateNode Is Nothing Then
        MsgBox "PrototypeNew: list references missing template '" & templateName & "'.", vbExclamation
        Exit Function
    End If

    Set rootNodes = templateNode.selectNodes("p:control | p:stackPanel | p:grid | p:list | p:itemControl")
    If rootNodes Is Nothing Or rootNodes.Length = 0 Then
        MsgBox "PrototypeNew: template '" & templateName & "' has no visual root node.", vbExclamation
        Exit Function
    End If
    If rootNodes.Length <> 1 Then
        MsgBox "PrototypeNew: template '" & templateName & "' must contain exactly one visual root node.", vbExclamation
        Exit Function
    End If

    Set outTemplateRoot = rootNodes.Item(0)
    mp_TryResolveListTemplateRoot = True
End Function

Private Function mp_TryResolveItemControlTemplateRoot(ByVal itemControlNode As Object, ByRef outTemplateRoot As Object) As Boolean
    Dim templateName As String
    Dim ownerDoc As Object
    Dim templateNode As Object
    Dim rootNodes As Object

    templateName = Trim$(ex_XmlCore.m_NodeAttrText(itemControlNode, "objectSourceTemplate"))
    If Len(templateName) = 0 Then
        MsgBox "PrototypeNew: itemControl requires non-empty attribute 'objectSourceTemplate'.", vbExclamation
        Exit Function
    End If

    Set ownerDoc = itemControlNode.OwnerDocument
    If ownerDoc Is Nothing Then
        MsgBox "PrototypeNew: failed to resolve owner document for itemControl template.", vbExclamation
        Exit Function
    End If

    Set templateNode = ownerDoc.selectSingleNode( _
        "/p:uiDefinition/p:templates/p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]")
    If templateNode Is Nothing Then
        MsgBox "PrototypeNew: itemControl references missing template '" & templateName & "'.", vbExclamation
        Exit Function
    End If

    Set rootNodes = templateNode.selectNodes("p:control | p:stackPanel | p:grid | p:list | p:itemControl")
    If rootNodes Is Nothing Or rootNodes.Length = 0 Then
        MsgBox "PrototypeNew: template '" & templateName & "' has no visual root node.", vbExclamation
        Exit Function
    End If
    If rootNodes.Length <> 1 Then
        MsgBox "PrototypeNew: template '" & templateName & "' must contain exactly one visual root node.", vbExclamation
        Exit Function
    End If

    Set outTemplateRoot = rootNodes.Item(0)
    mp_TryResolveItemControlTemplateRoot = True
End Function

Private Function mp_GetListOrientation(ByVal listNode As Object) As String
    mp_GetListOrientation = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(listNode, "orientation")))
    If Len(mp_GetListOrientation) = 0 Then mp_GetListOrientation = "vertical"

    If StrComp(mp_GetListOrientation, "vertical", vbBinaryCompare) <> 0 And _
       StrComp(mp_GetListOrientation, "horizontal", vbBinaryCompare) <> 0 Then
        MsgBox "PrototypeNew: list orientation must be 'vertical' or 'horizontal'.", vbExclamation
        mp_GetListOrientation = vbNullString
    End If
End Function

Private Function mp_TryMeasureListContentSpan( _
    ByVal listNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal bindingSource As Object _
) As Boolean
    Dim items As Collection
    Dim templateRoot As Object
    Dim itemRows As Long
    Dim itemCols As Long
    Dim orientation As String
    Dim itemValue As Variant
    Dim itemBindingSource As Object

    If listNode Is Nothing Then Exit Function

    If Not mp_TryResolveItemsSourceForMeasure( _
        ex_XmlCore.m_NodeAttrText(listNode, "itemsSource"), bindingSource, items) Then Exit Function

    If items Is Nothing Or items.Count = 0 Then
        outSpanRows = 1
        outSpanCols = 1
        mp_TryMeasureListContentSpan = True
        Exit Function
    End If

    If Not mp_TryResolveListTemplateRoot(listNode, templateRoot) Then Exit Function

    orientation = mp_GetListOrientation(listNode)
    If Len(orientation) = 0 Then Exit Function

    outSpanRows = 0
    outSpanCols = 0

    For Each itemValue In items
        Set itemBindingSource = Nothing
        If Not mp_TryCreateListItemBindingSource(itemValue, itemBindingSource) Then Exit Function
        If Not mp_TryGetEffectiveNodeSpan(templateRoot, itemRows, itemCols, itemBindingSource) Then Exit Function

        If StrComp(orientation, "horizontal", vbBinaryCompare) = 0 Then
            If itemRows > outSpanRows Then outSpanRows = itemRows
            outSpanCols = outSpanCols + itemCols
        Else
            outSpanRows = outSpanRows + itemRows
            If itemCols > outSpanCols Then outSpanCols = itemCols
        End If
    Next itemValue

    If outSpanRows <= 0 Then outSpanRows = 1
    If outSpanCols <= 0 Then outSpanCols = 1
    mp_TryMeasureListContentSpan = True
End Function

Private Function mp_TryMeasureItemControlContentSpan( _
    ByVal itemControlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal bindingSource As Object _
) As Boolean
    Dim sourceObject As Object
    Dim templateRoot As Object

    If itemControlNode Is Nothing Then Exit Function

    If Not mp_TryResolveObjectSourceForMeasure( _
        ex_XmlCore.m_NodeAttrText(itemControlNode, "objectSource"), _
        bindingSource, _
        sourceObject) Then Exit Function

    If sourceObject Is Nothing Then
        outSpanRows = 0
        outSpanCols = 0
        mp_TryMeasureItemControlContentSpan = True
        Exit Function
    End If

    If Not mp_TryResolveItemControlTemplateRoot(itemControlNode, templateRoot) Then Exit Function
    If Not mp_TryGetEffectiveNodeSpan(templateRoot, outSpanRows, outSpanCols, sourceObject) Then Exit Function
    mp_TryMeasureItemControlContentSpan = True
End Function

Private Function mp_TryResolveItemsSourceForMeasure( _
    ByVal rawItemsSource As String, _
    ByVal bindingSource As Object, _
    ByRef outItems As Collection _
) As Boolean
    Dim sourceText As String
    Dim resolvedValue As Variant

    sourceText = Trim$(rawItemsSource)
    If Len(sourceText) = 0 Then
        MsgBox "PrototypeNew: list itemsSource is required.", vbExclamation
        Exit Function
    End If

    If mp_IsBindingExpression(sourceText) Then
        If bindingSource Is Nothing Then
            MsgBox "PrototypeNew: list itemsSource binding requires item context during layout measurement.", vbExclamation
            Exit Function
        End If

        If Not ex_BindingRuntime.m_TryResolveValueBinding(sourceText, bindingSource, resolvedValue) Then Exit Function

        If IsObject(resolvedValue) Then
            If TypeName(resolvedValue) <> "Collection" Then
                MsgBox "PrototypeNew: list itemsSource binding must resolve to Collection.", vbExclamation
                Exit Function
            End If

            Set outItems = resolvedValue
            mp_TryResolveItemsSourceForMeasure = True
            Exit Function
        End If

        sourceText = Trim$(CStr(resolvedValue))
        If Len(sourceText) = 0 Then
            MsgBox "PrototypeNew: list itemsSource binding resolved to empty value.", vbExclamation
            Exit Function
        End If
    End If

    If Not ex_ListItemsSourceRuntime.m_TryResolveItemsSource(sourceText, outItems) Then Exit Function
    mp_TryResolveItemsSourceForMeasure = True
End Function

Private Function mp_TryResolveObjectSourceForMeasure( _
    ByVal rawObjectSource As String, _
    ByVal bindingSource As Object, _
    ByRef outObject As Object _
) As Boolean
    Dim sourceText As String
    Dim resolvedValue As Variant

    sourceText = Trim$(rawObjectSource)
    If Len(sourceText) = 0 Then
        mp_TryResolveObjectSourceForMeasure = True
        Exit Function
    End If

    If mp_IsBindingExpression(sourceText) Then
        If bindingSource Is Nothing Then
            MsgBox "PrototypeNew: itemControl objectSource binding requires item context during layout measurement.", vbExclamation
            Exit Function
        End If

        If Not ex_BindingRuntime.m_TryResolveValueBinding(sourceText, bindingSource, resolvedValue) Then Exit Function

        If IsObject(resolvedValue) Then
            Set outObject = resolvedValue
            mp_TryResolveObjectSourceForMeasure = True
            Exit Function
        End If

        sourceText = Trim$(CStr(resolvedValue))
    End If

    If Len(sourceText) = 0 Then
        mp_TryResolveObjectSourceForMeasure = True
        Exit Function
    End If

    If Not ex_ObjectSourceRuntime.m_TryResolveObjectSource(sourceText, outObject, True) Then Exit Function
    mp_TryResolveObjectSourceForMeasure = True
End Function

Private Function mp_IsBindingExpression(ByVal rawText As String) As Boolean
    Dim normalized As String

    normalized = Trim$(rawText)
    If Len(normalized) < 10 Then Exit Function
    If StrComp(Left$(normalized, 9), "{Binding ", vbTextCompare) <> 0 Then Exit Function
    If Right$(normalized, 1) <> "}" Then Exit Function

    mp_IsBindingExpression = True
End Function

Private Sub mp_AppendSuffixToControlNames(ByVal rootNode As Object, ByVal suffix As String)
    Dim childNode As Object
    Dim baseName As String
    Dim nodeName As String

    If rootNode Is Nothing Then Exit Sub
    If rootNode.NodeType <> 1 Then Exit Sub

    baseName = LCase$(CStr(rootNode.baseName))
    If StrComp(baseName, "control", vbBinaryCompare) = 0 Then
        nodeName = Trim$(ex_XmlCore.m_NodeAttrText(rootNode, "name"))
        If Len(nodeName) = 0 Then nodeName = "item"
        rootNode.setAttribute "name", nodeName & suffix
    End If

    For Each childNode In rootNode.ChildNodes
        mp_AppendSuffixToControlNames childNode, suffix
    Next childNode
End Sub

Private Sub mp_ApplyListItemValueToTemplate(ByVal templateRoot As Object, ByVal itemValue As Variant)
    Dim captionText As String
    Dim targetControl As Object
    Dim existingCaption As String

    If IsObject(itemValue) Then Exit Sub

    captionText = CStr(itemValue)
    Set targetControl = mp_FindFirstControlNode(templateRoot)
    If targetControl Is Nothing Then Exit Sub

    existingCaption = Trim$(ex_XmlCore.m_NodeAttrText(targetControl, "caption"))
    If Len(existingCaption) > 0 Then Exit Sub

    targetControl.setAttribute "caption", captionText
End Sub

Private Function mp_ApplyListItemBindings(ByVal templateRoot As Object, ByVal itemValue As Variant) As Boolean
    Dim bindingSource As Object

    If Not mp_TryCreateListItemBindingSource(itemValue, bindingSource) Then Exit Function
    If Not mp_ApplyNodeBindingsRecursive(templateRoot, bindingSource) Then Exit Function

    mp_ApplyListItemBindings = True
End Function

Private Function mp_TryCreateListItemBindingSource(ByVal itemValue As Variant, ByRef outSource As Object) As Boolean
    Dim scalarSource As Object
    Dim scalarText As String

    If IsObject(itemValue) Then
        If itemValue Is Nothing Then
            MsgBox "PrototypeNew: list item object is Nothing.", vbExclamation
            Exit Function
        End If

        Set outSource = itemValue
        mp_TryCreateListItemBindingSource = True
        Exit Function
    End If

    scalarText = CStr(itemValue)
    Set scalarSource = CreateObject("Scripting.Dictionary")
    scalarSource.CompareMode = 1
    scalarSource("Value") = scalarText
    scalarSource("Text") = scalarText

    Set outSource = scalarSource
    mp_TryCreateListItemBindingSource = True
End Function

Private Function mp_ApplyNodeBindingsRecursive(ByVal rootNode As Object, ByVal bindingSource As Object) As Boolean
    Dim attrs As Object
    Dim attrNode As Object
    Dim attrName As String
    Dim rawText As String
    Dim resolvedValue As Variant
    Dim childNode As Object
    Dim runtimeListSourceKey As String
    Dim runtimeObjectSourceKey As String
    Dim runtimeItems As Collection
    Dim resolvedObject As Object
    Dim rootNodeName As String

    If rootNode Is Nothing Then
        mp_ApplyNodeBindingsRecursive = True
        Exit Function
    End If
    If rootNode.NodeType <> 1 Then
        mp_ApplyNodeBindingsRecursive = True
        Exit Function
    End If

    Set attrs = rootNode.selectNodes("@*")
    If Not attrs Is Nothing Then
        For Each attrNode In attrs
            attrName = CStr(attrNode.nodeName)
            If LCase$(Left$(attrName, 5)) = "xmlns" Then GoTo ContinueAttr

            rawText = CStr(attrNode.Text)
            If InStr(1, rawText, "{Binding ", vbTextCompare) = 0 Then GoTo ContinueAttr

            If Not ex_BindingRuntime.m_TryResolveValueBinding(rawText, bindingSource, resolvedValue) Then Exit Function

            If IsObject(resolvedValue) Then
                rootNodeName = LCase$(CStr(rootNode.baseName))

                If StrComp(LCase$(attrName), "itemssource", vbBinaryCompare) = 0 And _
                   (StrComp(rootNodeName, "list", vbBinaryCompare) = 0 Or _
                    StrComp(rootNodeName, "control", vbBinaryCompare) = 0) Then

                    If TypeName(resolvedValue) = "Collection" Then
                        Set runtimeItems = resolvedValue
                    ElseIf StrComp(rootNodeName, "control", vbBinaryCompare) = 0 Then
                        Set runtimeItems = New Collection
                        Set resolvedObject = resolvedValue
                        If resolvedObject Is Nothing Then
                            MsgBox "PrototypeNew: itemsSource binding resolved to Nothing object.", vbExclamation
                            Exit Function
                        End If
                        runtimeItems.Add resolvedObject
                    Else
                        MsgBox "PrototypeNew: list itemsSource binding must resolve to Collection.", vbExclamation
                        Exit Function
                    End If

                    runtimeListSourceKey = mp_RegisterRuntimeListItemsSourceKey(runtimeItems)
                    If Len(runtimeListSourceKey) = 0 Then Exit Function
                    rootNode.setAttribute attrName, runtimeListSourceKey
                ElseIf StrComp(LCase$(attrName), "objectsource", vbBinaryCompare) = 0 And _
                       StrComp(rootNodeName, "itemcontrol", vbBinaryCompare) = 0 Then

                    Set resolvedObject = resolvedValue
                    If resolvedObject Is Nothing Then
                        rootNode.setAttribute attrName, vbNullString
                    Else
                        runtimeObjectSourceKey = mp_RegisterRuntimeObjectSourceKey(resolvedObject)
                        If Len(runtimeObjectSourceKey) = 0 Then Exit Function
                        rootNode.setAttribute attrName, runtimeObjectSourceKey
                    End If
                Else
                    MsgBox "PrototypeNew: template binding for attribute '" & attrName & "' must resolve to scalar value.", vbExclamation
                    Exit Function
                End If
            Else
                rootNode.setAttribute attrName, CStr(resolvedValue)
            End If

ContinueAttr:
        Next attrNode
    End If

    For Each childNode In rootNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueChild
        If Not mp_ApplyNodeBindingsRecursive(childNode, bindingSource) Then Exit Function
ContinueChild:
    Next childNode

    mp_ApplyNodeBindingsRecursive = True
End Function

Private Function mp_RegisterRuntimeListItemsSourceKey(ByVal items As Collection) As String
    Dim sourceKey As String

    If items Is Nothing Then
        MsgBox "PrototypeNew: runtime list items source is Nothing.", vbExclamation
        Exit Function
    End If

    g_ListRuntimeSourceSeed = g_ListRuntimeSourceSeed + 1
    sourceKey = "__list_runtime_" & CStr(g_ListRuntimeSourceSeed)

    If Not ex_ListItemsSourceRuntime.m_SetItemsSource(sourceKey, items) Then Exit Function
    mp_RegisterRuntimeListItemsSourceKey = sourceKey
End Function

Private Function mp_RegisterRuntimeObjectSourceKey(ByVal sourceObject As Object) As String
    Dim sourceKey As String

    If sourceObject Is Nothing Then Exit Function

    g_ObjectRuntimeSourceSeed = g_ObjectRuntimeSourceSeed + 1
    sourceKey = "__object_runtime_" & CStr(g_ObjectRuntimeSourceSeed)

    If Not ex_ObjectSourceRuntime.m_SetObjectSource(sourceKey, sourceObject) Then Exit Function
    mp_RegisterRuntimeObjectSourceKey = sourceKey
End Function

Private Function mp_FindFirstControlNode(ByVal rootNode As Object) As Object
    Dim childNode As Object

    If rootNode Is Nothing Then Exit Function
    If rootNode.NodeType <> 1 Then Exit Function

    If StrComp(LCase$(CStr(rootNode.baseName)), "control", vbBinaryCompare) = 0 Then
        Set mp_FindFirstControlNode = rootNode
        Exit Function
    End If

    For Each childNode In rootNode.ChildNodes
        Set mp_FindFirstControlNode = mp_FindFirstControlNode(childNode)
        If Not mp_FindFirstControlNode Is Nothing Then Exit Function
    Next childNode
End Function

Private Function mp_TryResolveNodeCellPosition( _
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

    atText = Trim$(ex_XmlCore.m_NodeAttrText(node, "at"))
    If Len(atText) = 0 Then
        relRow = 1
        relCol = 1
    Else
        If Not mp_TryParseAtAddress(atText, relRow, relCol) Then
            MsgBox "PrototypeNew: invalid 'at' format '" & atText & "'. Expected format is rNcM.", vbExclamation
            Exit Function
        End If
    End If

    outRow = anchorCell.Row + relRow - 1
    outCol = anchorCell.Column + relCol - 1
    mp_TryResolveNodeCellPosition = True
End Function

Private Function mp_TryParseAtAddress( _
    ByVal atText As String, _
    ByRef outRelRow As Long, _
    ByRef outRelCol As Long _
) As Boolean
    Dim normalized As String
    Dim cPos As Long
    Dim rowText As String
    Dim colText As String

    normalized = LCase$(Trim$(atText))
    If Len(normalized) < 4 Then Exit Function
    If Left$(normalized, 1) <> "r" Then Exit Function

    cPos = InStr(2, normalized, "c", vbBinaryCompare)
    If cPos <= 2 Then Exit Function
    If cPos >= Len(normalized) Then Exit Function

    rowText = Mid$(normalized, 2, cPos - 2)
    colText = Mid$(normalized, cPos + 1)

    If Not IsNumeric(rowText) Then Exit Function
    If Not IsNumeric(colText) Then Exit Function

    outRelRow = CLng(rowText)
    outRelCol = CLng(colText)
    If outRelRow <= 0 Or outRelCol <= 0 Then Exit Function

    mp_TryParseAtAddress = True
End Function

Private Function mp_ReadPositiveLongAttr( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Long _
) As Long
    Dim rawText As String

    rawText = Trim$(ex_XmlCore.m_NodeAttrText(node, attrName))
    If Len(rawText) = 0 Then
        mp_ReadPositiveLongAttr = defaultValue
        Exit Function
    End If

    If Not IsNumeric(rawText) Then
        MsgBox "PrototypeNew: attribute '" & attrName & "' must be numeric.", vbExclamation
        Exit Function
    End If

    mp_ReadPositiveLongAttr = CLng(rawText)
    If mp_ReadPositiveLongAttr <= 0 Then
        MsgBox "PrototypeNew: attribute '" & attrName & "' must be greater than zero.", vbExclamation
        mp_ReadPositiveLongAttr = 0
    End If
End Function

Private Sub mp_ClearWorksheet(ByVal ws As Worksheet)
    Dim i As Long
    Dim clearRange As Range

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set clearRange = ws.UsedRange
    If Not clearRange Is Nothing Then clearRange.Clear
    On Error GoTo 0

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If mp_IsGeneratedRuntimeShape(ws.Shapes(i)) Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function mp_IsGeneratedRuntimeShape(ByVal shp As Shape) As Boolean
    Dim tagValue As String

    If shp Is Nothing Then Exit Function

    On Error Resume Next
    tagValue = CStr(shp.Tags("pn.control"))
    If Err.Number <> 0 Then
        Err.Clear
        tagValue = vbNullString
    End If
    On Error GoTo 0

    If Len(Trim$(tagValue)) > 0 Then
        mp_IsGeneratedRuntimeShape = True
        Exit Function
    End If

    If LCase$(Left$(shp.Name, 4)) = "btn_" Then
        mp_IsGeneratedRuntimeShape = True
    End If
End Function
