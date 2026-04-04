Attribute VB_Name = "ex_ResultLayoutXmlEngine"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const INPUT_KEY_LAYOUT_ROWKINDS As String = "__ResultLayoutRowKinds"
Private Const INPUT_KEY_LAYOUT_FIELDRANGES As String = "__ResultLayoutFieldRanges"
Private Const INPUT_KEY_LAYOUT_ITEMSOURCES As String = "__ResultLayoutItemsSources"
Private Const META_KEY_VIRTUAL_FIELD_ALIASES As String = "VirtualFieldAliases"
Private Const META_KEY_VIRTUAL_FIELD_KIND As String = "VirtualFieldKind"
Private Const TABLE_ATTR_ROW_KIND_HEADER As String = "rowKindHeader"
Private Const TABLE_ATTR_ROW_KIND_CONTENT As String = "rowKindContent"
Private Const TABLE_ATTR_FIELD_KIND_HEADER As String = "fieldKindHeader"
Private Const TABLE_ATTR_FIELD_KIND_CONTENT As String = "fieldKindContent"
Private Const TABLE_ATTR_FIELD_KIND_VIRTUAL As String = "fieldKindVirtual"
Private Const DEFAULT_TABLE_ROW_KIND_HEADER As String = "header"
Private Const DEFAULT_TABLE_ROW_KIND_CONTENT As String = "content"
Private Const DEBUG_LOG_PATH As String = "Logs\layout_engine.log"
Private Const DEBUG_LOG_ENABLED As Boolean = False
Private Const ROOT_BOUNDS_FIELD_ROW As String = "row"
Private Const ROOT_BOUNDS_FIELD_COL As String = "col"
Private Const ROOT_BOUNDS_FIELD_WIDTH As String = "width"
Private Const ROOT_BOUNDS_FIELD_HEIGHT As String = "height"
Private g_ButtonStylesMap As Object
Private g_RootBoundsBySheet As Object
Private g_VirtualAliasesLookupByText As Object

Public Function m_ApplyResultLayoutFromDom( _
    ByVal doc As Object, _
    ByVal ws As Worksheet, _
    ByVal resultTables As Collection, _
    ByVal inputObject As Object, _
    ByRef outErrorText As String _
) As Boolean

    Dim gridNodes As Object
    Dim gridNode As Object
    Dim rootNodes As Object
    Dim rootNode As Object
    Dim anchorCell As Range
    Dim rootWidth As Long
    Dim rootHeight As Long
    Dim rootRow As Long
    Dim rootCol As Long
    Dim rootContext As Object
    Dim rowKinds As Object
    Dim resultFieldRanges As Collection
    Dim styleLoadError As String
    Dim stageName As String
    Dim gridOrdinal As Long
    Dim rootOrdinal As Long
    Dim rootRenderId As String

    On Error GoTo EH
    stageName = "validate-input"
    mp_DebugLog "m_ApplyResultLayoutFromDom: start."

    If doc Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If resultTables Is Nothing Then
        outErrorText = "ResultTables are required for XML layout rendering."
        Exit Function
    End If

    stageName = "load-styles"
    Set g_ButtonStylesMap = ex_UiXmlProvider.m_ReadButtonStylesFromDoc(doc, styleLoadError)
    If g_ButtonStylesMap Is Nothing Then
        If Len(styleLoadError) = 0 Then styleLoadError = "Failed to read styles map from result layout XML."
        outErrorText = styleLoadError
        mp_DebugLog "m_ApplyResultLayoutFromDom: style load failed. " & outErrorText
        Exit Function
    End If
    mp_DebugLog "m_ApplyResultLayoutFromDom: styles count=" & mp_TryGetDictionaryCountText(g_ButtonStylesMap)

    stageName = "load-grid"
    Set gridNodes = doc.selectNodes("/p:uiDefinition/p:layout/p:grid")
    If gridNodes Is Nothing Then Exit Function
    If gridNodes.Length = 0 Then Exit Function
    mp_DebugLog "m_ApplyResultLayoutFromDom: grid count=" & CStr(gridNodes.Length)

    stageName = "init-rowKinds"
    Set g_VirtualAliasesLookupByText = Nothing
    Set rowKinds = CreateObject("Scripting.Dictionary")
    rowKinds.CompareMode = 1

    Set resultFieldRanges = New Collection

    stageName = "init-root-context"
    Set rootContext = CreateObject("Scripting.Dictionary")
    rootContext.CompareMode = 1
    rootContext.Add "ResultTables", mp_BuildResultTableItems(resultTables)
    If Not mp_MergeInputItemsSources(rootContext, inputObject, outErrorText) Then Exit Function
    If Not mp_EnsureRootItemsPanelSources(doc, rootContext, outErrorText) Then Exit Function

    stageName = "prepare-sheet"
    mp_PrepareResultSheet ws
    mp_ClearRootBoundsForSheet ws

    stageName = "apply-layout"
    gridOrdinal = 0
    For Each gridNode In gridNodes
        gridOrdinal = gridOrdinal + 1
        If Not mp_TryResolveGridAnchorCell(ws, gridNode, anchorCell, outErrorText) Then Exit Function
        If Not mp_ApplyGridColumns(ws, gridNode, anchorCell, outErrorText) Then Exit Function

        Set rootNodes = gridNode.selectNodes("p:stackPanel | p:border | p:control")
        If rootNodes Is Nothing Then GoTo ContinueGrid

        rootOrdinal = 0
        For Each rootNode In rootNodes
            rootOrdinal = rootOrdinal + 1
            rootRenderId = mp_BuildRootRenderId(gridOrdinal, rootOrdinal)
            mp_DebugLog "apply-layout: root begin tag='" & mp_NodeTag(rootNode) & "' at='" & Trim$(mp_NodeAttrText(rootNode, "at")) & "'."
            If Not mp_MeasureNode(rootNode, ws, doc, rootContext, anchorCell.Row, anchorCell.Column, False, 0, False, 0, rootWidth, rootHeight, outErrorText) Then Exit Function
            If rootWidth < 0 Or rootHeight < 0 Then
                outErrorText = "Layout node '" & mp_NodeTag(rootNode) & "' produced negative size."
                Exit Function
            End If
            mp_DebugLog "apply-layout: root measured tag='" & mp_NodeTag(rootNode) & "' w=" & CStr(rootWidth) & " h=" & CStr(rootHeight) & "."
            If Not mp_TryResolveNodeStart(rootNode, anchorCell.Row, anchorCell.Column, anchorCell.Row, anchorCell.Column, True, rootRow, rootCol, outErrorText) Then Exit Function
            If Not mp_ApplyNode(rootNode, ws, doc, rootContext, anchorCell.Row, anchorCell.Column, rootRow, rootCol, rootWidth, rootHeight, False, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
            mp_RecordRootBounds ws, rootRenderId, rootRow, rootCol, rootWidth, rootHeight
            mp_DebugLog "apply-layout: root done tag='" & mp_NodeTag(rootNode) & "'."
        Next rootNode
ContinueGrid:
    Next gridNode

    If Not inputObject Is Nothing Then
        ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_ROWKINDS, rowKinds
        ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_FIELDRANGES, resultFieldRanges
    End If

    m_ApplyResultLayoutFromDom = True
    mp_DebugLog "m_ApplyResultLayoutFromDom: completed."
    Exit Function
EH:
    outErrorText = "[" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    mp_DebugLog "m_ApplyResultLayoutFromDom: fail stage='" & stageName & "' error='" & outErrorText & "'."
End Function

Public Function m_ApplyResultLayoutPartialFromDom( _
    ByVal doc As Object, _
    ByVal ws As Worksheet, _
    ByVal resultTables As Collection, _
    ByVal inputObject As Object, _
    ByVal changedItemsSourceKey As String, _
    ByRef outErrorText As String _
) As Boolean

    Dim gridNodes As Object
    Dim gridNode As Object
    Dim rootNodes As Object
    Dim rootNode As Object
    Dim anchorCell As Range
    Dim rootWidth As Long
    Dim rootHeight As Long
    Dim rootRow As Long
    Dim rootCol As Long
    Dim rootContext As Object
    Dim rowKinds As Object
    Dim resultFieldRanges As Collection
    Dim styleLoadError As String
    Dim stageName As String
    Dim changedKey As String
    Dim gridOrdinal As Long
    Dim rootOrdinal As Long
    Dim rootRenderId As String
    Dim oldRow As Long
    Dim oldCol As Long
    Dim oldWidth As Long
    Dim oldHeight As Long
    Dim clearRowStart As Long
    Dim clearColStart As Long
    Dim clearRowEnd As Long
    Dim clearColEnd As Long

    On Error GoTo EH
    stageName = "validate-input"

    changedKey = Trim$(changedItemsSourceKey)
    If Len(changedKey) = 0 Then
        m_ApplyResultLayoutPartialFromDom = m_ApplyResultLayoutFromDom(doc, ws, resultTables, inputObject, outErrorText)
        Exit Function
    End If

    If doc Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If resultTables Is Nothing Then
        outErrorText = "ResultTables are required for XML layout partial refresh."
        Exit Function
    End If

    stageName = "load-styles"
    Set g_ButtonStylesMap = ex_UiXmlProvider.m_ReadButtonStylesFromDoc(doc, styleLoadError)
    If g_ButtonStylesMap Is Nothing Then
        If Len(styleLoadError) = 0 Then styleLoadError = "Failed to read styles map from result layout XML."
        outErrorText = styleLoadError
        Exit Function
    End If

    stageName = "load-grid"
    Set gridNodes = doc.selectNodes("/p:uiDefinition/p:layout/p:grid")
    If gridNodes Is Nothing Then Exit Function
    If gridNodes.Length = 0 Then Exit Function

    stageName = "init-root-context"
    Set rootContext = CreateObject("Scripting.Dictionary")
    rootContext.CompareMode = 1
    rootContext.Add "ResultTables", mp_BuildResultTableItems(resultTables)
    If Not mp_MergeInputItemsSources(rootContext, inputObject, outErrorText) Then Exit Function
    If Not mp_EnsureRootItemsPanelSources(doc, rootContext, outErrorText) Then Exit Function

    stageName = "init-stub-context"
    Set g_VirtualAliasesLookupByText = Nothing
    Set rowKinds = CreateObject("Scripting.Dictionary")
    rowKinds.CompareMode = 1
    Set resultFieldRanges = New Collection

    stageName = "apply-partial-layout"
    gridOrdinal = 0
    For Each gridNode In gridNodes
        gridOrdinal = gridOrdinal + 1
        If Not mp_TryResolveGridAnchorCell(ws, gridNode, anchorCell, outErrorText) Then Exit Function
        If Not mp_ApplyGridColumns(ws, gridNode, anchorCell, outErrorText) Then Exit Function

        Set rootNodes = gridNode.selectNodes("p:stackPanel | p:border | p:control")
        If rootNodes Is Nothing Then GoTo ContinueGrid

        rootOrdinal = 0
        For Each rootNode In rootNodes
            rootOrdinal = rootOrdinal + 1
            If Not mp_NodeDependsOnItemsSourceKey(rootNode, changedKey) Then GoTo ContinueRoot

            rootRenderId = mp_BuildRootRenderId(gridOrdinal, rootOrdinal)
            If Not mp_MeasureNode(rootNode, ws, doc, rootContext, anchorCell.Row, anchorCell.Column, False, 0, False, 0, rootWidth, rootHeight, outErrorText) Then Exit Function
            If rootWidth < 0 Or rootHeight < 0 Then
                outErrorText = "Layout node '" & mp_NodeTag(rootNode) & "' produced negative size."
                Exit Function
            End If
            If Not mp_TryResolveNodeStart(rootNode, anchorCell.Row, anchorCell.Column, anchorCell.Row, anchorCell.Column, True, rootRow, rootCol, outErrorText) Then Exit Function

            clearRowStart = rootRow
            clearColStart = rootCol
            clearRowEnd = rootRow + rootHeight - 1
            clearColEnd = rootCol + rootWidth - 1
            If mp_TryGetRootBounds(ws, rootRenderId, oldRow, oldCol, oldWidth, oldHeight) Then
                If oldRow < clearRowStart Then clearRowStart = oldRow
                If oldCol < clearColStart Then clearColStart = oldCol
                If oldRow + oldHeight - 1 > clearRowEnd Then clearRowEnd = oldRow + oldHeight - 1
                If oldCol + oldWidth - 1 > clearColEnd Then clearColEnd = oldCol + oldWidth - 1
            End If
            mp_ClearCellsRegion ws, clearRowStart, clearColStart, clearRowEnd, clearColEnd

            If Not mp_ApplyNode(rootNode, ws, doc, rootContext, anchorCell.Row, anchorCell.Column, rootRow, rootCol, rootWidth, rootHeight, False, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
            mp_RecordRootBounds ws, rootRenderId, rootRow, rootCol, rootWidth, rootHeight
ContinueRoot:
        Next rootNode
ContinueGrid:
    Next gridNode

    m_ApplyResultLayoutPartialFromDom = True
    Exit Function
EH:
    outErrorText = "[" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Public Function m_PartialRefreshMayAffectTable( _
    ByVal doc As Object, _
    ByVal changedItemsSourceKey As String _
) As Boolean
    Dim gridNodes As Object
    Dim gridNode As Object
    Dim rootNodes As Object
    Dim rootNode As Object
    Dim changedKey As String

    changedKey = Trim$(changedItemsSourceKey)
    If Len(changedKey) = 0 Then
        m_PartialRefreshMayAffectTable = True
        Exit Function
    End If
    If doc Is Nothing Then Exit Function

    On Error GoTo EH
    Set gridNodes = doc.selectNodes("/p:uiDefinition/p:layout/p:grid")
    If gridNodes Is Nothing Then Exit Function

    For Each gridNode In gridNodes
        Set rootNodes = gridNode.selectNodes("p:stackPanel | p:border | p:control")
        If rootNodes Is Nothing Then GoTo ContinueGrid
        For Each rootNode In rootNodes
            If mp_NodeDependsOnItemsSourceKey(rootNode, changedKey) Then
                If mp_NodeContainsTableControl(rootNode, doc) Then
                    m_PartialRefreshMayAffectTable = True
                    Exit Function
                End If
            End If
        Next rootNode
ContinueGrid:
    Next gridNode
    Exit Function
EH:
    m_PartialRefreshMayAffectTable = True
End Function

Private Sub mp_PrepareResultSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Cells.Clear
    ex_Messaging.m_ClearBannerAnchors ws
    ex_Messaging.m_ClearResultTableAnchors ws
    ex_Messaging.m_ClearResultRowAnchors ws
    ex_LayoutBindingsRuntime.m_ClearSheetBindings ws
    On Error GoTo 0
End Sub

Private Function mp_BuildResultTableItems(ByVal resultTables As Collection) As Collection
    Dim items As Collection
    Dim i As Long
    Dim tableObj As obj_ResultTable
    Dim item As Object
    Dim metaInfo As Object
    Dim metaName As String

    Set items = New Collection
    If resultTables Is Nothing Then
        Set mp_BuildResultTableItems = items
        Exit Function
    End If

    For i = 1 To resultTables.Count
        Set tableObj = resultTables(i)
        If tableObj Is Nothing Then GoTo ContinueTable

        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = 1
        item.Add "__raw", tableObj
        item("TableRef") = CStr(tableObj.TableRef)
        item.Add "Rows", tableObj.Rows
        item("Count") = CLng(tableObj.Count)

        Set metaInfo = mp_GetOrCreateResultTableMetaInfo(tableObj)
        metaName = mp_GetMetaInfoName(metaInfo, CStr(tableObj.TableRef))
        metaInfo("Name") = metaName
        item.Add "MetaInfo", metaInfo

        items.Add item
ContinueTable:
    Next i

    Set mp_BuildResultTableItems = items
End Function

Private Function mp_MergeInputItemsSources( _
    ByVal rootContext As Object, _
    ByVal inputObject As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim itemsMap As Object
    Dim sourceKey As Variant
    Dim sourceKeyText As String
    Dim sourceObj As Object
    Dim sourceValue As Variant

    If rootContext Is Nothing Then
        outErrorText = "Root context is not initialized for layout itemsSource merge."
        Exit Function
    End If

    If inputObject Is Nothing Then
        mp_MergeInputItemsSources = True
        Exit Function
    End If

    If Not ex_ScriptIO.m_TryGetObject(inputObject, INPUT_KEY_LAYOUT_ITEMSOURCES, itemsMap) Then
        mp_MergeInputItemsSources = True
        Exit Function
    End If
    If itemsMap Is Nothing Then
        mp_MergeInputItemsSources = True
        Exit Function
    End If
    If Not (TypeName(itemsMap) = "Dictionary" Or TypeName(itemsMap) = "Scripting.Dictionary") Then
        outErrorText = "Input key '" & INPUT_KEY_LAYOUT_ITEMSOURCES & "' must be Dictionary."
        Exit Function
    End If

    On Error GoTo EH
    For Each sourceKey In itemsMap.Keys
        sourceKeyText = Trim$(CStr(sourceKey))
        If Len(sourceKeyText) = 0 Then GoTo ContinueKey

        If rootContext.Exists(sourceKeyText) Then
            rootContext.Remove sourceKeyText
        End If

        Set sourceObj = Nothing
        On Error Resume Next
        Set sourceObj = itemsMap(sourceKey)
        On Error GoTo EH
        If Not sourceObj Is Nothing Then
            ex_ScriptIO.m_SetObject rootContext, sourceKeyText, sourceObj
        Else
            sourceValue = itemsMap(sourceKey)
            rootContext(sourceKeyText) = sourceValue
        End If
ContinueKey:
    Next sourceKey

    mp_MergeInputItemsSources = True
    Exit Function
EH:
    outErrorText = "Failed to merge runtime itemsSource map: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_EnsureRootItemsPanelSources( _
    ByVal doc As Object, _
    ByVal rootContext As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim panelNodes As Object
    Dim panelNode As Object
    Dim sourceText As String
    Dim emptyItems As Collection

    If doc Is Nothing Then
        outErrorText = "Layout DOM is not available for itemsSource pre-initialization."
        Exit Function
    End If
    If rootContext Is Nothing Then
        outErrorText = "Root context is not available for itemsSource pre-initialization."
        Exit Function
    End If

    On Error GoTo EH
    Set panelNodes = doc.selectNodes("//*[local-name()='control'][translate(@type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='itemspanel'][@itemsSource]")
    If panelNodes Is Nothing Then
        mp_EnsureRootItemsPanelSources = True
        Exit Function
    End If

    For Each panelNode In panelNodes
        sourceText = Trim$(mp_NodeAttrText(panelNode, "itemsSource"))
        If Len(sourceText) = 0 Then GoTo ContinueNode
        If mp_TryExtractBindingPath(sourceText, sourceText) Then GoTo ContinueNode
        If rootContext.Exists(sourceText) Then GoTo ContinueNode

        Set emptyItems = New Collection
        ex_ScriptIO.m_SetObject rootContext, sourceText, emptyItems
ContinueNode:
    Next panelNode

    mp_EnsureRootItemsPanelSources = True
    Exit Function
EH:
    outErrorText = "Failed to pre-initialize itemsPanel sources: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_GetOrCreateResultTableMetaInfo(ByVal tableObj As obj_ResultTable) As Object
    Dim metaInfo As Object

    If tableObj Is Nothing Then GoTo CreateDefault

    On Error Resume Next
    Set metaInfo = tableObj.MetaInfo
    On Error GoTo 0
    If Not metaInfo Is Nothing Then
        Set mp_GetOrCreateResultTableMetaInfo = metaInfo
        Exit Function
    End If

CreateDefault:
    Set metaInfo = CreateObject("Scripting.Dictionary")
    metaInfo.CompareMode = 1
    Set mp_GetOrCreateResultTableMetaInfo = metaInfo
End Function

Private Function mp_GetMetaInfoName(ByVal metaInfo As Object, ByVal defaultName As String) As String
    If Not metaInfo Is Nothing Then
        On Error Resume Next
        mp_GetMetaInfoName = Trim$(CStr(metaInfo("Name")))
        If Err.Number <> 0 Then
            Err.Clear
            mp_GetMetaInfoName = vbNullString
        End If
        On Error GoTo 0
    End If

    If Len(mp_GetMetaInfoName) = 0 Then mp_GetMetaInfoName = Trim$(defaultName)
End Function

Private Function mp_TryResolveGridAnchorCell(ByVal ws As Worksheet, ByVal gridNode As Object, ByRef anchorCell As Range, ByRef outErrorText As String) As Boolean
    Dim anchorCellText As String

    anchorCellText = Trim$(mp_NodeAttrText(gridNode, "anchorCell"))
    If Len(anchorCellText) = 0 Then anchorCellText = "A1"

    On Error GoTo EH
    Set anchorCell = ws.Range(anchorCellText)
    On Error GoTo 0

    mp_TryResolveGridAnchorCell = True
    Exit Function
EH:
    outErrorText = "Invalid grid@anchorCell value '" & anchorCellText & "'."
End Function

Private Function mp_ApplyGridColumns(ByVal ws As Worksheet, ByVal gridNode As Object, ByVal anchorCell As Range, ByRef outErrorText As String) As Boolean
    Dim colNodes As Object
    Dim colNode As Object
    Dim iText As String
    Dim widthText As String
    Dim colIndexRel As Long
    Dim widthUnits As Double
    Dim targetColAbs As Long

    Set colNodes = gridNode.selectNodes("p:columns/p:col")
    If colNodes Is Nothing Then
        mp_ApplyGridColumns = True
        Exit Function
    End If

    For Each colNode In colNodes
        iText = Trim$(mp_NodeAttrText(colNode, "i"))
        If Len(iText) = 0 Then
            outErrorText = "Grid column entry is missing required attribute 'i'."
            Exit Function
        End If
        If Not ex_XmlCore.m_TryParseLong(iText, colIndexRel) Then
            outErrorText = "Grid column has non-numeric i='" & iText & "'."
            Exit Function
        End If
        If colIndexRel < 1 Then
            outErrorText = "Grid column index must be >= 1, got: " & CStr(colIndexRel) & "."
            Exit Function
        End If

        widthText = Trim$(mp_NodeAttrText(colNode, "width"))
        If Len(widthText) = 0 Then GoTo ContinueCol
        If Not ex_XmlCore.m_TryParseDouble(widthText, widthUnits) Then
            outErrorText = "Grid column has invalid width='" & widthText & "' for i=" & CStr(colIndexRel) & "."
            Exit Function
        End If
        If widthUnits <= 0 Then
            outErrorText = "Grid column width must be > 0 for i=" & CStr(colIndexRel) & "."
            Exit Function
        End If

        targetColAbs = anchorCell.Column + colIndexRel - 1
        ws.Columns(targetColAbs).ColumnWidth = widthUnits
ContinueCol:
    Next colNode

    mp_ApplyGridColumns = True
End Function

Private Function mp_ApplyNode( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal isFlowChild As Boolean, _
    ByVal rowKinds As Object, _
    ByVal resultFieldRanges As Collection, _
    ByRef outErrorText As String _
) As Boolean

    Dim tagName As String
    Dim atText As String

    tagName = mp_NodeTag(node)
    atText = Trim$(mp_NodeAttrText(node, "at"))
    If isFlowChild And Len(atText) > 0 Then
        outErrorText = "Layout node '" & tagName & "' cannot define 'at' inside stack/border flow."
        Exit Function
    End If

    Select Case tagName
        Case "stackpanel"
            mp_ApplyNode = mp_ApplyStackPanel(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, nodeWidth, nodeHeight, rowKinds, resultFieldRanges, outErrorText)
        Case "border"
            mp_ApplyNode = mp_ApplyBorder(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, nodeWidth, nodeHeight, rowKinds, resultFieldRanges, outErrorText)
        Case "control"
            mp_ApplyNode = mp_ApplyControl(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, nodeWidth, nodeHeight, isFlowChild, rowKinds, resultFieldRanges, outErrorText)
        Case Else
            outErrorText = "Unsupported layout node '" & tagName & "'. Allowed: stackPanel, border, control."
    End Select
End Function

Private Function mp_MeasureNode( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long, _
    ByRef outErrorText As String _
) As Boolean

    Dim tagName As String

    tagName = mp_NodeTag(node)
    Select Case tagName
        Case "stackpanel"
            mp_MeasureNode = mp_MeasureStackPanel(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight, outErrorText)
        Case "border"
            mp_MeasureNode = mp_MeasureBorder(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight, outErrorText)
        Case "control"
            mp_MeasureNode = mp_MeasureControl(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight, outErrorText)
        Case Else
            outErrorText = "Unsupported layout node '" & tagName & "' during measurement."
    End Select
End Function

Private Function mp_MeasureStackPanel( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long, _
    ByRef outErrorText As String _
) As Boolean

    Dim orientation As String
    Dim hasWidthSpec As Boolean
    Dim hasHeightSpec As Boolean
    Dim widthAuto As Boolean
    Dim heightAuto As Boolean
    Dim widthValue As Long
    Dim heightValue As Long
    Dim knownWidth As Long
    Dim knownHeight As Long
    Dim hasKnownWidth As Boolean
    Dim hasKnownHeight As Boolean
    Dim children As Collection
    Dim childNode As Object
    Dim childW As Long
    Dim childH As Long
    Dim contentW As Long
    Dim contentH As Long

    orientation = LCase$(Trim$(mp_NodeAttrText(node, "orientation")))
    If orientation <> "vertical" And orientation <> "horizontal" Then
        outErrorText = "stackPanel must define orientation='vertical' or 'horizontal'."
        Exit Function
    End If

    If Not mp_TryReadLayoutSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "stackPanel@spanCells", outErrorText) Then Exit Function
    If Not mp_TryReadLayoutSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "stackPanel@spanRows", outErrorText) Then Exit Function

    If hasWidthSpec And Not widthAuto Then
        knownWidth = widthValue
        hasKnownWidth = True
    ElseIf Not hasWidthSpec And hasParentWidth Then
        knownWidth = parentWidth
        hasKnownWidth = True
    End If

    If hasHeightSpec And Not heightAuto Then
        knownHeight = heightValue
        hasKnownHeight = True
    ElseIf Not hasHeightSpec And hasParentHeight Then
        knownHeight = parentHeight
        hasKnownHeight = True
    End If

    If Not mp_CollectActiveChildren(node, children, "stackPanel children", outErrorText) Then Exit Function
    If Not children Is Nothing Then
        For Each childNode In children
            If orientation = "vertical" Then
                If Not mp_MeasureNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, hasKnownWidth, knownWidth, False, 0, childW, childH, outErrorText) Then Exit Function
                contentH = contentH + childH
                If childW > contentW Then contentW = childW
            Else
                If Not mp_MeasureNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, False, 0, hasKnownHeight, knownHeight, childW, childH, outErrorText) Then Exit Function
                contentW = contentW + childW
                If childH > contentH Then contentH = childH
            End If
        Next childNode
    End If

    If hasWidthSpec Then
        If widthAuto Then
            outWidth = contentW
        Else
            outWidth = widthValue
        End If
    ElseIf hasParentWidth Then
        outWidth = parentWidth
    Else
        outWidth = contentW
    End If

    If hasHeightSpec Then
        If heightAuto Then
            outHeight = contentH
        Else
            outHeight = heightValue
        End If
    ElseIf hasParentHeight Then
        outHeight = parentHeight
    Else
        outHeight = contentH
    End If

    mp_MeasureStackPanel = True
End Function

Private Function mp_MeasureBorder( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long, _
    ByRef outErrorText As String _
) As Boolean

    Dim hasWidthSpec As Boolean
    Dim hasHeightSpec As Boolean
    Dim widthAuto As Boolean
    Dim heightAuto As Boolean
    Dim widthValue As Long
    Dim heightValue As Long
    Dim knownWidth As Long
    Dim knownHeight As Long
    Dim hasKnownWidth As Boolean
    Dim hasKnownHeight As Boolean
    Dim children As Collection
    Dim childNode As Object
    Dim childW As Long
    Dim childH As Long
    Dim contentW As Long
    Dim contentH As Long

    If Not mp_TryReadLayoutSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "border@spanCells", outErrorText) Then Exit Function
    If Not mp_TryReadLayoutSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "border@spanRows", outErrorText) Then Exit Function

    If hasWidthSpec And Not widthAuto Then
        knownWidth = widthValue
        hasKnownWidth = True
    ElseIf Not hasWidthSpec And hasParentWidth Then
        knownWidth = parentWidth
        hasKnownWidth = True
    End If

    If hasHeightSpec And Not heightAuto Then
        knownHeight = heightValue
        hasKnownHeight = True
    ElseIf Not hasHeightSpec And hasParentHeight Then
        knownHeight = parentHeight
        hasKnownHeight = True
    End If

    If Not mp_CollectActiveChildren(node, children, "border children", outErrorText) Then Exit Function
    If Not children Is Nothing Then
        For Each childNode In children
            If Not mp_MeasureNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, hasKnownWidth, knownWidth, hasKnownHeight, knownHeight, childW, childH, outErrorText) Then Exit Function
            If childW > contentW Then contentW = childW
            If childH > contentH Then contentH = childH
        Next childNode
    End If

    If hasWidthSpec Then
        If widthAuto Then
            outWidth = contentW
        Else
            outWidth = widthValue
        End If
    ElseIf hasParentWidth Then
        outWidth = parentWidth
    Else
        outWidth = contentW
    End If

    If hasHeightSpec Then
        If heightAuto Then
            outHeight = contentH
        Else
            outHeight = heightValue
        End If
    ElseIf hasParentHeight Then
        outHeight = parentHeight
    Else
        outHeight = contentH
    End If

    mp_MeasureBorder = True
End Function

Private Function mp_MeasureControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long, _
    ByRef outErrorText As String _
) As Boolean

    Dim controlType As String
    Dim hasWidthSpec As Boolean
    Dim hasHeightSpec As Boolean
    Dim widthAuto As Boolean
    Dim heightAuto As Boolean
    Dim widthValue As Long
    Dim heightValue As Long
    Dim measureH As Long

    controlType = LCase$(Trim$(mp_NodeAttrText(node, "type")))
    If Len(controlType) = 0 Then
        outErrorText = "Layout control node is missing required attribute 'type'."
        Exit Function
    End If

    If Not mp_TryReadLayoutSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "control@spanCells", outErrorText) Then Exit Function
    If Not mp_TryReadLayoutSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "control@spanRows", outErrorText) Then Exit Function

    If widthAuto Then
        outErrorText = "Control type='" & controlType & "' does not support spanCells='auto'."
        Exit Function
    End If

    If hasWidthSpec Then
        outWidth = widthValue
    ElseIf hasParentWidth Then
        outWidth = parentWidth
    Else
        outErrorText = "Control type='" & controlType & "' must define spanCells or be hosted by parent with known width."
        Exit Function
    End If

    Select Case controlType
        Case "label"
            If hasHeightSpec Then
                If heightAuto Then
                    outHeight = 1
                Else
                    outHeight = heightValue
                End If
            ElseIf hasParentHeight Then
                outHeight = parentHeight
            Else
                outHeight = 1
            End If
        Case "table"
            If hasHeightSpec Then
                If heightAuto Then
                    If Not mp_MeasureTableHeight(node, dataItem, measureH, outErrorText) Then Exit Function
                    outHeight = measureH
                Else
                    outHeight = heightValue
                End If
            ElseIf hasParentHeight Then
                outHeight = parentHeight
            Else
                If Not mp_MeasureTableHeight(node, dataItem, measureH, outErrorText) Then Exit Function
                outHeight = measureH
            End If
        Case "itemspanel"
            mp_MeasureControl = mp_MeasureItemsPanelControl(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight, outErrorText)
            Exit Function
        Case "button", "dropdownbutton"
            If hasHeightSpec Then
                If heightAuto Then
                    outHeight = 1
                Else
                    outHeight = heightValue
                End If
            ElseIf hasParentHeight Then
                outHeight = parentHeight
            Else
                outHeight = 1
            End If
        Case "input"
            If hasHeightSpec Then
                If heightAuto Then
                    outHeight = 1
                Else
                    outHeight = heightValue
                End If
            ElseIf hasParentHeight Then
                outHeight = parentHeight
            Else
                outHeight = 1
            End If
        Case Else
            outErrorText = "Unsupported result-layout control type='" & controlType & "'."
            Exit Function
    End Select

    If outWidth < 0 Or outHeight < 0 Then
        outErrorText = "Control type='" & controlType & "' produced negative size."
        Exit Function
    End If

    mp_MeasureControl = True
End Function

Private Function mp_MeasureItemsPanelControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long, _
    ByRef outErrorText As String _
) As Boolean

    Dim orientation As String
    Dim templateNode As Object
    Dim items As Collection
    Dim item As Variant
    Dim itemData As Object
    Dim itemW As Long
    Dim itemH As Long
    Dim contentW As Long
    Dim contentH As Long
    Dim hasWidthSpec As Boolean
    Dim hasHeightSpec As Boolean
    Dim widthAuto As Boolean
    Dim heightAuto As Boolean
    Dim widthValue As Long
    Dim heightValue As Long

    On Error GoTo EH_MEASURE_ITEMS

    orientation = LCase$(Trim$(mp_NodeAttrText(node, "orientation")))
    If orientation <> "vertical" And orientation <> "horizontal" Then
        outErrorText = "itemsPanel must define orientation='vertical' or 'horizontal'."
        Exit Function
    End If

    If Not mp_TryReadLayoutSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "itemsPanel@spanCells", outErrorText) Then Exit Function
    If Not mp_TryReadLayoutSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "itemsPanel@spanRows", outErrorText) Then Exit Function
    If widthAuto Then
        outErrorText = "itemsPanel does not support spanCells='auto'."
        Exit Function
    End If

    If hasWidthSpec Then
        outWidth = widthValue
    ElseIf hasParentWidth Then
        outWidth = parentWidth
    Else
        outErrorText = "itemsPanel must define spanCells or be hosted by parent with known width."
        Exit Function
    End If

    If Not mp_GetItemsPanelSourceItems(node, dataItem, items, outErrorText) Then Exit Function
    If Not mp_GetItemsPanelTemplateNode(doc, node, templateNode, outErrorText) Then Exit Function

    If Not items Is Nothing Then
        For Each item In items
            If Not IsObject(item) Then
                outErrorText = "itemsPanel itemsSource must contain object items compatible with bindings."
                Exit Function
            End If
            Set itemData = item
            If Not mp_MeasureNode(templateNode, ws, doc, itemData, gridAnchorRow, gridAnchorCol, True, outWidth, False, 0, itemW, itemH, outErrorText) Then Exit Function
            If orientation = "vertical" Then
                contentH = contentH + itemH
                If itemW > contentW Then contentW = itemW
            Else
                contentW = contentW + itemW
                If itemH > contentH Then contentH = itemH
            End If
        Next item
    End If

    If hasHeightSpec Then
        If heightAuto Then
            outHeight = contentH
        Else
            outHeight = heightValue
        End If
    ElseIf hasParentHeight Then
        outHeight = parentHeight
    Else
        outHeight = contentH
    End If

    If contentW > outWidth Then
        outErrorText = "itemsPanel template content exceeds allocated width."
        Exit Function
    End If

    mp_MeasureItemsPanelControl = True
    Exit Function

EH_MEASURE_ITEMS:
    outErrorText = "itemsPanel measure failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_GetItemsPanelTemplateNode(ByVal doc As Object, ByVal itemsPanelNode As Object, ByRef outTemplateRoot As Object, ByRef outErrorText As String) As Boolean
    Dim templateName As String
    Dim templateNode As Object
    Dim rootNodes As Object

    templateName = Trim$(mp_NodeAttrText(itemsPanelNode, "itemTemplate"))
    If Len(templateName) = 0 Then
        outErrorText = "itemsPanel requires non-empty attribute 'itemTemplate'."
        Exit Function
    End If

    Set templateNode = doc.selectSingleNode("/p:uiDefinition/p:templates/p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]")
    If templateNode Is Nothing Then
        outErrorText = "itemsPanel references missing template '" & templateName & "'."
        Exit Function
    End If

    Set rootNodes = templateNode.selectNodes("p:stackPanel | p:border | p:control")
    If rootNodes Is Nothing Then
        outErrorText = "Template '" & templateName & "' does not define root layout node."
        Exit Function
    End If
    If rootNodes.Length <> 1 Then
        outErrorText = "Template '" & templateName & "' must contain exactly one root node (stackPanel|border|control)."
        Exit Function
    End If

    Set outTemplateRoot = rootNodes.Item(0)
    mp_GetItemsPanelTemplateNode = True
End Function

Private Function mp_GetItemsPanelSourceItems(ByVal itemsPanelNode As Object, ByVal dataItem As Object, ByRef outItems As Collection, ByRef outErrorText As String) As Boolean
    Dim sourceValue As Variant
    Dim sourceText As String
    Dim sourceItems As Collection
    Dim finalItems As Collection
    Dim filterPattern As String
    Dim filterPath As String

    sourceText = Trim$(mp_NodeAttrText(itemsPanelNode, "itemsSource"))
    If Len(sourceText) = 0 Then
        outErrorText = "itemsPanel requires non-empty attribute 'itemsSource'."
        Exit Function
    End If

    sourceValue = mp_ResolveBindingValue(sourceText, dataItem)
    If IsObject(sourceValue) Then
        If TypeName(sourceValue) = "Collection" Then
            Set sourceItems = sourceValue
        End If
    ElseIf Not dataItem Is Nothing Then
        ' Non-binding shorthand: itemsSource="SomeKey" -> context("SomeKey")
        If mp_TryGetDictionaryObject(dataItem, sourceText, sourceValue) Then
            If IsObject(sourceValue) Then
                If TypeName(sourceValue) = "Collection" Then
                    Set sourceItems = sourceValue
                End If
            End If
        End If
    End If

    If sourceItems Is Nothing Then
        outErrorText = "itemsPanel itemsSource='" & sourceText & "' must resolve to Collection."
        If IsObject(sourceValue) Then
            outErrorText = outErrorText & " Resolved type: " & TypeName(sourceValue) & "."
        End If
        Exit Function
    End If

    filterPattern = Trim$(mp_NodeAttrText(itemsPanelNode, "itemsSourceFilter"))
    If Len(filterPattern) = 0 Then
        Set finalItems = sourceItems
    Else
        filterPath = Trim$(mp_NodeAttrText(itemsPanelNode, "itemsSourceFilterBind"))
        If Len(filterPath) = 0 Then
            ' Backward-compatible alias.
            filterPath = Trim$(mp_NodeAttrText(itemsPanelNode, "itemsSourceFilterPath"))
        End If

        If Not mp_FilterItemsByRegex(sourceItems, filterPattern, filterPath, finalItems, outErrorText) Then Exit Function
    End If

    Set outItems = finalItems
    mp_GetItemsPanelSourceItems = True
End Function

Private Function mp_FilterItemsByRegex( _
    ByVal sourceItems As Collection, _
    ByVal patternText As String, _
    ByVal filterPath As String, _
    ByRef outItems As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim rx As Object
    Dim item As Variant
    Dim itemObj As Object
    Dim targetText As String

    On Error GoTo EH_REGEX
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = CStr(patternText)
    On Error GoTo 0

    Set outItems = New Collection
    If sourceItems Is Nothing Then
        mp_FilterItemsByRegex = True
        Exit Function
    End If

    For Each item In sourceItems
        If Not IsObject(item) Then GoTo ContinueItem
        Set itemObj = item
        targetText = mp_ResolveItemsFilterTargetText(itemObj, filterPath)
        If Len(targetText) = 0 Then GoTo ContinueItem
        If rx.Test(targetText) Then outItems.Add itemObj
ContinueItem:
    Next item

    mp_FilterItemsByRegex = True
    Exit Function
EH_REGEX:
    outErrorText = "Invalid itemsSourceFilter regex: '" & patternText & "'."
End Function

Private Function mp_ResolveItemsFilterTargetText(ByVal itemObj As Object, ByVal filterPath As String) As String
    Dim normalizedPath As String
    Dim bindingPath As String
    Dim resolvedValue As Variant

    If itemObj Is Nothing Then Exit Function

    normalizedPath = Trim$(filterPath)
    If Len(normalizedPath) = 0 Then
        mp_ResolveItemsFilterTargetText = mp_TryGetItemMetaName(itemObj)
        Exit Function
    End If

    If mp_TryExtractBindingPath(normalizedPath, bindingPath) Then
        normalizedPath = bindingPath
    End If

    resolvedValue = mp_ReadBindingPathValue(itemObj, normalizedPath)
    If IsObject(resolvedValue) Then
        mp_ResolveItemsFilterTargetText = vbNullString
    Else
        mp_ResolveItemsFilterTargetText = Trim$(CStr(resolvedValue))
    End If
End Function

Private Function mp_TryGetItemMetaName(ByVal itemObj As Object) As String
    Dim metaInfo As Object

    If itemObj Is Nothing Then Exit Function

    On Error Resume Next
    If TypeName(itemObj) = "Dictionary" Or TypeName(itemObj) = "Scripting.Dictionary" Then
        Call mp_TryGetDictionaryObjectRef(itemObj, "MetaInfo", metaInfo)
    End If
    On Error GoTo 0

    If metaInfo Is Nothing Then
        On Error Resume Next
        Set metaInfo = itemObj.MetaInfo
        On Error GoTo 0
    End If

    If Not metaInfo Is Nothing Then
        On Error Resume Next
        mp_TryGetItemMetaName = Trim$(CStr(metaInfo("Name")))
        If Err.Number <> 0 Then
            Err.Clear
            mp_TryGetItemMetaName = vbNullString
        End If
        On Error GoTo 0
    End If
End Function

Private Function mp_MeasureTableHeight(ByVal tableNode As Object, ByVal dataItem As Object, ByRef outHeight As Long, ByRef outErrorText As String) As Boolean
    Dim tableObj As obj_ResultTable
    Dim rowsObj As Collection
    Dim rowsCount As Long
    Dim hasHeader As Boolean
    Dim stageName As String

    On Error GoTo EH_MEASURE_TABLE

    stageName = "read-showHeader"
    hasHeader = True
    If Len(Trim$(mp_NodeAttrText(tableNode, "showHeader"))) > 0 Then
        If Not ex_XmlCore.m_TryParseBoolean(Trim$(mp_NodeAttrText(tableNode, "showHeader")), hasHeader) Then
            outErrorText = "Invalid boolean value for table@showHeader."
            Exit Function
        End If
    End If

    stageName = "resolve-rows-object"
    Set rowsObj = mp_ResolveRowsObjectForTableNode(tableNode, dataItem)
    If Not rowsObj Is Nothing Then
        rowsCount = rowsObj.Count
    Else
        stageName = "resolve-table-object"
        Set tableObj = mp_ResolveTableObjectForNode(tableNode, dataItem)
        If Not tableObj Is Nothing Then
            stageName = "read-table-rows-count"
            Set rowsObj = tableObj.Rows
            If Not rowsObj Is Nothing Then
                rowsCount = rowsObj.Count
            Else
                rowsCount = 0
            End If
        End If
    End If

    stageName = "compute-height"
    outHeight = rowsCount
    If hasHeader Then outHeight = outHeight + 1
    If outHeight <= 0 Then outHeight = 1

    mp_MeasureTableHeight = True
    Exit Function

EH_MEASURE_TABLE:
    outErrorText = "table measure failed at stage '" & stageName & "': [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_ApplyStackPanel( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal rowKinds As Object, _
    ByVal resultFieldRanges As Collection, _
    ByRef outErrorText As String _
) As Boolean

    Dim orientation As String
    Dim children As Collection
    Dim childNode As Object
    Dim childW As Long
    Dim childH As Long
    Dim cursorRow As Long
    Dim cursorCol As Long
    Dim usedMain As Long

    orientation = LCase$(Trim$(mp_NodeAttrText(node, "orientation")))
    If orientation <> "vertical" And orientation <> "horizontal" Then
        outErrorText = "stackPanel must define orientation='vertical' or 'horizontal'."
        Exit Function
    End If

    If Not mp_CollectActiveChildren(node, children, "stackPanel children", outErrorText) Then Exit Function

    cursorRow = nodeRow
    cursorCol = nodeCol
    If Not children Is Nothing Then
        For Each childNode In children
            If orientation = "vertical" Then
                If Not mp_MeasureNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, True, nodeWidth, False, 0, childW, childH, outErrorText) Then Exit Function
                If childW > nodeWidth Then
                    outErrorText = "Child node '" & mp_NodeTag(childNode) & "' exceeds stackPanel width."
                    Exit Function
                End If
                If Not mp_ApplyNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, cursorRow, nodeCol, childW, childH, True, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
                cursorRow = cursorRow + childH
                usedMain = usedMain + childH
            Else
                If Not mp_MeasureNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, False, 0, True, nodeHeight, childW, childH, outErrorText) Then Exit Function
                If childH > nodeHeight Then
                    outErrorText = "Child node '" & mp_NodeTag(childNode) & "' exceeds stackPanel height."
                    Exit Function
                End If
                If Not mp_ApplyNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, nodeRow, cursorCol, childW, childH, True, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
                cursorCol = cursorCol + childW
                usedMain = usedMain + childW
            End If
        Next childNode
    End If

    If orientation = "vertical" Then
        If usedMain > nodeHeight Then
            outErrorText = "stackPanel content height exceeds allocated height."
            Exit Function
        End If
    Else
        If usedMain > nodeWidth Then
            outErrorText = "stackPanel content width exceeds allocated width."
            Exit Function
        End If
    End If

    mp_ApplyStackPanel = True
End Function

Private Function mp_ApplyBorder( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal rowKinds As Object, _
    ByVal resultFieldRanges As Collection, _
    ByRef outErrorText As String _
) As Boolean

    Dim children As Collection
    Dim childNode As Object
    Dim childW As Long
    Dim childH As Long

    If Not mp_CollectActiveChildren(node, children, "border children", outErrorText) Then Exit Function
    If children Is Nothing Then
        mp_ApplyBorder = True
        Exit Function
    End If

    For Each childNode In children
        If Not mp_MeasureNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, True, nodeWidth, True, nodeHeight, childW, childH, outErrorText) Then Exit Function
        If childW > nodeWidth Or childH > nodeHeight Then
            outErrorText = "Child node '" & mp_NodeTag(childNode) & "' exceeds border bounds."
            Exit Function
        End If
        If Not mp_ApplyNode(childNode, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, childW, childH, True, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
    Next childNode

    mp_ApplyBorder = True
End Function

Private Function mp_ApplyControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal isFlowChild As Boolean, _
    ByVal rowKinds As Object, _
    ByVal resultFieldRanges As Collection, _
    ByRef outErrorText As String _
) As Boolean

    Dim controlType As String
    Dim atText As String
    Dim controlName As String

    controlType = LCase$(Trim$(mp_NodeAttrText(node, "type")))
    If Len(controlType) = 0 Then
        outErrorText = "Layout control node is missing required attribute 'type'."
        Exit Function
    End If

    atText = Trim$(mp_NodeAttrText(node, "at"))
    If isFlowChild And Len(atText) > 0 Then
        outErrorText = "Control type='" & controlType & "' cannot define 'at' when hosted in stack/border flow."
        Exit Function
    End If

    If nodeWidth < 0 Or nodeHeight < 0 Then
        outErrorText = "Control type='" & controlType & "' has negative size after layout resolution."
        Exit Function
    End If

    controlName = Trim$(mp_NodeAttrText(node, "name"))
    If Len(controlName) = 0 Then controlName = "<anon>"
    mp_DebugLog "mp_ApplyControl: begin type='" & controlType & "' name='" & controlName & "' row=" & CStr(nodeRow) & " col=" & CStr(nodeCol) & " w=" & CStr(nodeWidth) & " h=" & CStr(nodeHeight) & "."

    On Error GoTo EH_CONTROL

    Select Case controlType
        Case "label"
            mp_ApplyControl = mp_ApplyLabelControl(node, ws, dataItem, nodeRow, nodeCol, nodeWidth, nodeHeight, rowKinds, outErrorText)
        Case "table"
            mp_ApplyControl = mp_ApplyTableControl(node, ws, dataItem, nodeRow, nodeCol, nodeWidth, nodeHeight, rowKinds, resultFieldRanges, outErrorText)
        Case "itemspanel"
            mp_ApplyControl = mp_ApplyItemsPanelControl(node, ws, doc, dataItem, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, nodeWidth, nodeHeight, rowKinds, resultFieldRanges, outErrorText)
        Case "button", "dropdownbutton"
            mp_ApplyControl = mp_ApplyShapeControl(node, ws, dataItem, nodeRow, nodeCol, nodeWidth, nodeHeight, controlType, outErrorText)
        Case "input"
            mp_ApplyControl = mp_ApplyInputControl(node, ws, dataItem, nodeRow, nodeCol, nodeWidth, nodeHeight, rowKinds, outErrorText)
        Case Else
            outErrorText = "Unsupported result-layout control type='" & controlType & "'."
    End Select

    If mp_ApplyControl Then
        mp_DebugLog "mp_ApplyControl: done type='" & controlType & "' name='" & controlName & "'."
    ElseIf Len(outErrorText) > 0 Then
        mp_DebugLog "mp_ApplyControl: failed type='" & controlType & "' name='" & controlName & "' reason='" & outErrorText & "'."
    End If
    Exit Function

EH_CONTROL:
    outErrorText = "Control type='" & controlType & "' name='" & controlName & "' failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    mp_DebugLog "mp_ApplyControl: exception type='" & controlType & "' name='" & controlName & "' error='" & outErrorText & "'."
End Function

Private Function mp_ApplyLabelControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal dataItem As Object, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal rowKinds As Object, _
    ByRef outErrorText As String _
) As Boolean

    Dim targetRange As Range
    Dim textValue As String
    Dim rowKindText As String

    If nodeWidth <= 0 Then
        outErrorText = "Label control requires width > 0."
        Exit Function
    End If
    If nodeHeight <= 0 Then
        nodeHeight = 1
    End If

    If Not mp_TryBuildRangeByTracks(ws, nodeRow, nodeCol, nodeHeight, nodeWidth, targetRange) Then
        outErrorText = "Label control resolved to invalid grid bounds."
        Exit Function
    End If

    textValue = mp_ResolveTextValue(mp_NodeAttrText(node, "text"), dataItem)

    targetRange.UnMerge
    If targetRange.Cells.CountLarge > 1 Then targetRange.Merge
    targetRange.Value = textValue
    targetRange.WrapText = True
    targetRange.HorizontalAlignment = xlLeft
    targetRange.VerticalAlignment = xlTop

    rowKindText = Trim$(mp_ResolveTemplateText(mp_NodeAttrText(node, "rowKind"), dataItem))
    If Len(rowKindText) > 0 Then
        mp_RegisterRowKinds rowKindText, targetRange.Row, rowKinds
    End If

    mp_ApplyLabelControl = True
End Function

Private Function mp_ApplyInputControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal dataItem As Object, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal rowKinds As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim targetRange As Range
    Dim inputCell As Range
    Dim valueText As String
    Dim currentValue As String
    Dim inputConfigKey As String
    Dim inputBind As String
    Dim resolvedConfigKey As String
    Dim inputName As String
    Dim onChangeMacro As String
    Dim rowKindText As String
    Dim isPrimaryInput As Boolean
    Dim styleName As String
    Dim controlName As String

    If nodeWidth <= 0 Then
        outErrorText = "Input control requires width > 0."
        Exit Function
    End If
    If nodeHeight <= 0 Then
        nodeHeight = 1
    End If

    If Not mp_TryBuildRangeByTracks(ws, nodeRow, nodeCol, nodeHeight, nodeWidth, targetRange) Then
        outErrorText = "Input control resolved to invalid grid bounds."
        Exit Function
    End If

    targetRange.UnMerge
    If targetRange.Cells.CountLarge > 1 Then targetRange.Merge
    Set inputCell = targetRange.Cells(1, 1)

    inputCell.NumberFormat = "@"
    inputCell.HorizontalAlignment = xlLeft
    inputCell.VerticalAlignment = xlCenter

    styleName = Trim$(mp_NodeAttrText(node, "style"))
    If Len(styleName) > 0 Then
        controlName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(controlName) = 0 Then controlName = "input"
        If Not mp_ApplyCellStyleByName(targetRange, styleName, mp_GetButtonStylesMap(), controlName, outErrorText) Then Exit Function
    End If

    valueText = mp_ResolveTextValue(mp_NodeAttrText(node, "value"), dataItem)
    If Len(valueText) = 0 Then valueText = mp_ResolveTextValue(mp_NodeAttrText(node, "text"), dataItem)
    inputConfigKey = Trim$(mp_NodeAttrText(node, "inputConfigKey"))
    inputBind = Trim$(mp_NodeAttrText(node, "bind"))
    If Len(inputBind) = 0 Then inputBind = inputConfigKey

    currentValue = Trim$(CStr(inputCell.Value))
    If Len(currentValue) = 0 Then
        If Len(valueText) > 0 Then
            inputCell.Value = valueText
        ElseIf ex_LayoutBindingsRuntime.m_TryResolveConfigKeyFromBindSpec(inputBind, resolvedConfigKey) Then
            inputCell.Value = ex_ConfigProvider.m_GetConfigValue(resolvedConfigKey, vbNullString)
        End If
    End If

    inputName = Trim$(mp_NodeAttrText(node, "inputName"))
    If Len(inputName) = 0 Then inputName = Trim$(mp_NodeAttrText(node, "name"))
    If Len(inputName) = 0 Then inputName = inputConfigKey

    onChangeMacro = Trim$(mp_NodeAttrText(node, "onChange"))
    If Len(onChangeMacro) = 0 Then onChangeMacro = Trim$(mp_NodeAttrText(node, "onChangeMacro"))
    If Not mp_TryReadOptionalBoolean(node, "inputPrimary", False, isPrimaryInput, "input control", outErrorText) Then Exit Function

    ex_LayoutBindingsRuntime.m_RegisterInputBinding ws, inputCell, inputName, inputBind, onChangeMacro, isPrimaryInput

    rowKindText = Trim$(mp_ResolveTemplateText(mp_NodeAttrText(node, "rowKind"), dataItem))
    If Len(rowKindText) > 0 Then
        mp_RegisterRowKinds rowKindText, targetRange.Row, rowKinds
    End If

    mp_ApplyInputControl = True
End Function

Private Function mp_ApplyTableControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal dataItem As Object, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal rowKinds As Object, _
    ByVal resultFieldRanges As Collection, _
    ByRef outErrorText As String _
) As Boolean

    Dim tableObj As obj_ResultTable
    Dim rowsObj As Collection
    Dim fieldAliases As Collection
    Dim headerRow As Long
    Dim rowIndex As Long
    Dim rowObj As obj_ResultRow
    Dim i As Long
    Dim colIndex As Long
    Dim fieldAlias As String
    Dim valueText As String
    Dim showHeader As Boolean
    Dim tableRef As String
    Dim mapKey As String
    Dim rowAnchorName As String
    Dim rowKindText As String
    Dim rowObjectKind As String
    Dim headerRowKind As String
    Dim contentRowKind As String
    Dim headerFieldKind As String
    Dim contentFieldKind As String
    Dim virtualFieldKind As String
    Dim virtualAliasLookup As Object
    Dim fieldCount As Long
    Dim rowCount As Long
    Dim dataStartRow As Long
    Dim headerValues() As Variant
    Dim bodyValues() As Variant
    Dim writeRange As Range
    Dim contentRowKindCache As Object
    Dim rowKindCacheKey As String

    showHeader = True
    If Len(Trim$(mp_NodeAttrText(node, "showHeader"))) > 0 Then
        If Not ex_XmlCore.m_TryParseBoolean(Trim$(mp_NodeAttrText(node, "showHeader")), showHeader) Then
            outErrorText = "Invalid boolean value for table@showHeader."
            Exit Function
        End If
    End If

    Set rowsObj = mp_ResolveRowsObjectForTableNode(node, dataItem)
    Set tableObj = mp_ResolveTableObjectForNode(node, dataItem)
    mp_ResolveTableKindConfig node, dataItem, tableObj, headerRowKind, contentRowKind, headerFieldKind, contentFieldKind, virtualFieldKind
    Set virtualAliasLookup = mp_GetResultTableVirtualAliasLookup(tableObj)
    Set contentRowKindCache = CreateObject("Scripting.Dictionary")
    contentRowKindCache.CompareMode = 1
    contentRowKindCache(vbNullString) = contentRowKind

    If rowsObj Is Nothing Then
        outErrorText = "Table control must resolve to rows collection via itemsSource or bound table object."
        Exit Function
    End If

    Set fieldAliases = mp_GetFieldAliasesForTable(tableObj, rowsObj)
    fieldCount = fieldAliases.Count
    rowCount = rowsObj.Count

    tableRef = mp_ResolveTableRef(tableObj, dataItem)
    If Len(tableRef) = 0 Then tableRef = "Table"

    rowIndex = nodeRow
    If showHeader Then
        headerRow = rowIndex
        If fieldCount > 0 Then
            ReDim headerValues(1 To 1, 1 To fieldCount)
            For colIndex = 1 To fieldCount
                headerValues(1, colIndex) = CStr(fieldAliases(colIndex))
            Next colIndex
            Set writeRange = ws.Range(ws.Cells(headerRow, nodeCol), ws.Cells(headerRow, nodeCol + fieldCount - 1))
            writeRange.Value = headerValues
        End If
        If Len(headerRowKind) > 0 Then
            mp_RegisterRowKinds headerRowKind, headerRow, rowKinds
        End If
        rowIndex = rowIndex + 1
    End If

    dataStartRow = rowIndex
    If rowCount > 0 And fieldCount > 0 Then ReDim bodyValues(1 To rowCount, 1 To fieldCount)

    For i = 1 To rowCount
        Set rowObj = rowsObj(i)
        For colIndex = 1 To fieldCount
            fieldAlias = CStr(fieldAliases(colIndex))
            valueText = vbNullString
            If Not rowObj Is Nothing Then
                On Error Resume Next
                If rowObj.HasAlias(fieldAlias) Then valueText = CStr(rowObj.Column(fieldAlias))
                On Error GoTo 0
            End If
            If fieldCount > 0 Then bodyValues(i, colIndex) = valueText
        Next colIndex

        rowAnchorName = vbNullString
        If Not rowObj Is Nothing Then
            On Error Resume Next
            rowAnchorName = Trim$(CStr(rowObj.RowAnchorName))
            On Error GoTo 0
        End If
        If Len(rowAnchorName) = 0 Then
            rowAnchorName = ex_Messaging.m_BuildResultRowAnchorName(tableRef, i)
            If Len(rowAnchorName) > 0 Then
                If Not rowObj Is Nothing Then
                    On Error Resume Next
                    rowObj.RowAnchorName = rowAnchorName
                    On Error GoTo 0
                End If
            End If
        End If
        If Len(rowAnchorName) > 0 Then
            ex_Messaging.m_RegisterResultRowAnchor ws, rowAnchorName, rowIndex
        End If

        rowObjectKind = vbNullString
        If Not rowObj Is Nothing Then
            On Error Resume Next
            rowObjectKind = CStr(rowObj.Kind)
            On Error GoTo 0
        End If
        rowKindCacheKey = LCase$(Trim$(rowObjectKind))
        If contentRowKindCache.Exists(rowKindCacheKey) Then
            rowKindText = CStr(contentRowKindCache(rowKindCacheKey))
        Else
            rowKindText = mp_CombineKindTags(contentRowKind, rowObjectKind)
            contentRowKindCache(rowKindCacheKey) = rowKindText
        End If
        If Len(rowKindText) > 0 Then
            mp_RegisterRowKinds rowKindText, rowIndex, rowKinds
        End If
        rowIndex = rowIndex + 1
    Next i

    If rowCount > 0 And fieldCount > 0 Then
        Set writeRange = ws.Range(ws.Cells(dataStartRow, nodeCol), ws.Cells(dataStartRow + rowCount - 1, nodeCol + fieldCount - 1))
        writeRange.Value = bodyValues
    End If

    If showHeader Then
        mp_AddFieldRangesFromAliases resultFieldRanges, tableObj, fieldAliases, nodeCol, headerRow, headerRow, headerFieldKind, virtualFieldKind, virtualAliasLookup
        If rowIndex - 1 >= headerRow + 1 Then
            mp_AddFieldRangesFromAliases resultFieldRanges, tableObj, fieldAliases, nodeCol, headerRow + 1, rowIndex - 1, contentFieldKind, virtualFieldKind, virtualAliasLookup
        End If
    Else
        mp_AddFieldRangesFromAliases resultFieldRanges, tableObj, fieldAliases, nodeCol, nodeRow, rowIndex - 1, contentFieldKind, virtualFieldKind, virtualAliasLookup
    End If

    ex_Messaging.m_RegisterResultTableAnchor ws, tableRef, nodeRow, rowIndex - 1

    mp_ApplyTableControl = True
End Function

Private Function mp_ApplyItemsPanelControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal doc As Object, _
    ByVal dataItem As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal rowKinds As Object, _
    ByVal resultFieldRanges As Collection, _
    ByRef outErrorText As String _
) As Boolean

    Dim orientation As String
    Dim items As Collection
    Dim templateNode As Object
    Dim cursorRow As Long
    Dim cursorCol As Long
    Dim item As Variant
    Dim itemData As Object
    Dim itemW As Long
    Dim itemH As Long
    Dim usedMain As Long

    orientation = LCase$(Trim$(mp_NodeAttrText(node, "orientation")))
    If orientation <> "vertical" And orientation <> "horizontal" Then
        outErrorText = "itemsPanel must define orientation='vertical' or 'horizontal'."
        Exit Function
    End If

    If Not mp_GetItemsPanelSourceItems(node, dataItem, items, outErrorText) Then Exit Function
    If Not mp_GetItemsPanelTemplateNode(doc, node, templateNode, outErrorText) Then Exit Function

    cursorRow = nodeRow
    cursorCol = nodeCol

    If Not items Is Nothing Then
        For Each item In items
            If Not IsObject(item) Then
                outErrorText = "itemsPanel itemsSource must contain object items compatible with bindings."
                Exit Function
            End If
            Set itemData = item
            If orientation = "vertical" Then
                If Not mp_MeasureNode(templateNode, ws, doc, itemData, gridAnchorRow, gridAnchorCol, True, nodeWidth, False, 0, itemW, itemH, outErrorText) Then Exit Function
                If itemW > nodeWidth Then
                    outErrorText = "itemsPanel item template width exceeds itemsPanel width."
                    Exit Function
                End If
                If Not mp_ApplyNode(templateNode, ws, doc, itemData, gridAnchorRow, gridAnchorCol, cursorRow, nodeCol, itemW, itemH, True, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
                cursorRow = cursorRow + itemH
                usedMain = usedMain + itemH
            Else
                If Not mp_MeasureNode(templateNode, ws, doc, itemData, gridAnchorRow, gridAnchorCol, False, 0, True, nodeHeight, itemW, itemH, outErrorText) Then Exit Function
                If itemH > nodeHeight Then
                    outErrorText = "itemsPanel item template height exceeds itemsPanel height."
                    Exit Function
                End If
                If Not mp_ApplyNode(templateNode, ws, doc, itemData, gridAnchorRow, gridAnchorCol, nodeRow, cursorCol, itemW, itemH, True, rowKinds, resultFieldRanges, outErrorText) Then Exit Function
                cursorCol = cursorCol + itemW
                usedMain = usedMain + itemW
            End If
        Next item
    End If

    If orientation = "vertical" Then
        If usedMain > nodeHeight Then
            outErrorText = "itemsPanel content height exceeds allocated height."
            Exit Function
        End If
    Else
        If usedMain > nodeWidth Then
            outErrorText = "itemsPanel content width exceeds allocated width."
            Exit Function
        End If
    End If

    mp_ApplyItemsPanelControl = True
End Function

Private Function mp_ApplyShapeControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal dataItem As Object, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal controlType As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim controlName As String
    Dim targetRange As Range
    Dim shp As Shape
    Dim normalizedType As String

    If ws Is Nothing Then
        outErrorText = "Shape control requires worksheet context."
        Exit Function
    End If

    controlName = Trim$(mp_NodeAttrText(node, "name"))
    If Len(controlName) = 0 Then
        outErrorText = "Control type='" & controlType & "' requires non-empty attribute 'name'."
        Exit Function
    End If

    If Not mp_TryBuildRangeByTracks(ws, nodeRow, nodeCol, nodeHeight, nodeWidth, targetRange) Then
        outErrorText = "Control '" & controlName & "' resolved to invalid grid bounds."
        Exit Function
    End If

    normalizedType = LCase$(Trim$(controlType))
    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        If Not mp_TryCreateShapeControl(ws, normalizedType, controlName, targetRange, shp, outErrorText) Then Exit Function
    End If
    If shp Is Nothing Then
        outErrorText = "Control '" & controlName & "' was not created/resolved."
        Exit Function
    End If

    shp.Left = targetRange.Left
    shp.Top = targetRange.Top
    shp.Width = targetRange.Width
    shp.Height = targetRange.Height

    If Not mp_ApplyShapeVisibleAttr(node, shp, outErrorText) Then Exit Function
    If Not mp_ApplyShapePlacementAttr(node, shp, outErrorText) Then Exit Function
    If Not mp_ApplyShapeMacroAttr(node, shp, outErrorText) Then Exit Function

    Select Case normalizedType
        Case "button"
            If Not mp_ApplyButtonShapeVisuals(node, shp, dataItem, controlName, outErrorText) Then Exit Function
        Case "dropdownbutton"
            If Not mp_ApplyButtonShapeVisuals(node, shp, dataItem, controlName, outErrorText) Then Exit Function
            If Not mp_ApplyManagedDropdownButton(node, ws, shp, dataItem, controlName, outErrorText) Then Exit Function
        Case Else
            outErrorText = "Unsupported shape control type='" & controlType & "'."
            Exit Function
    End Select

    mp_ApplyShapeControl = True
End Function

Private Function mp_TryCreateShapeControl( _
    ByVal ws As Worksheet, _
    ByVal normalizedType As String, _
    ByVal controlName As String, _
    ByVal targetRange As Range, _
    ByRef outShape As Shape, _
    ByRef outErrorText As String _
) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function

    On Error GoTo EH_CREATE
    Select Case normalizedType
        Case "button", "dropdownbutton"
            Set outShape = ws.Shapes.AddShape(msoShapeRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
            outShape.Name = controlName
        Case Else
            outErrorText = "Unsupported shape control type='" & normalizedType & "' for control '" & controlName & "'."
            Exit Function
    End Select

    mp_TryCreateShapeControl = True
    Exit Function
EH_CREATE:
    outErrorText = "Failed to create control '" & controlName & "': " & Err.Description
End Function

Private Function mp_ApplyShapeVisibleAttr(ByVal node As Object, ByVal shp As Shape, ByRef outErrorText As String) As Boolean
    Dim valueText As String
    Dim valueBool As Boolean

    valueText = Trim$(mp_NodeAttrText(node, "visible"))
    If Len(valueText) = 0 Then
        mp_ApplyShapeVisibleAttr = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseBoolean(valueText, valueBool) Then
        outErrorText = "Invalid boolean value for control '" & shp.Name & "' attribute visible='" & valueText & "'."
        Exit Function
    End If

    shp.Visible = IIf(valueBool, msoTrue, msoFalse)
    mp_ApplyShapeVisibleAttr = True
End Function

Private Function mp_ApplyShapePlacementAttr(ByVal node As Object, ByVal shp As Shape, ByRef outErrorText As String) As Boolean
    Dim placementText As String
    Dim placementValue As XlPlacement

    placementText = Trim$(mp_NodeAttrText(node, "placement"))
    If Len(placementText) = 0 Then
        mp_ApplyShapePlacementAttr = True
        Exit Function
    End If

    If Not mp_TryParsePlacement(placementText, placementValue) Then
        outErrorText = "Invalid control placement value '" & placementText & "' for '" & shp.Name & "'."
        Exit Function
    End If

    shp.Placement = placementValue
    mp_ApplyShapePlacementAttr = True
End Function

Private Function mp_ApplyShapeMacroAttr(ByVal node As Object, ByVal shp As Shape, ByRef outErrorText As String) As Boolean
    Dim macroName As String
    Dim onActionText As String

    macroName = Trim$(mp_NodeAttrText(node, "macro"))
    If Len(macroName) = 0 Then
        mp_ApplyShapeMacroAttr = True
        Exit Function
    End If

    onActionText = "'" & ThisWorkbook.Name & "'!" & macroName
    On Error GoTo EH_ASSIGN
    shp.OnAction = onActionText
    mp_ApplyShapeMacroAttr = True
    Exit Function
EH_ASSIGN:
    outErrorText = "Failed to assign macro '" & macroName & "' to control '" & shp.Name & "': " & Err.Description
End Function

Private Function mp_ApplyButtonShapeVisuals( _
    ByVal node As Object, _
    ByVal shp As Shape, _
    ByVal dataItem As Object, _
    ByVal controlName As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim captionText As String
    Dim styleName As String
    Dim stylesMap As Object

    captionText = mp_ResolveControlCaptionText(node, dataItem, controlName)

    On Error Resume Next
    shp.AutoShapeType = msoShapeRectangle
    On Error GoTo EH_APPLY

    shp.TextFrame.Characters.Text = captionText

    styleName = Trim$(mp_NodeAttrText(node, "style"))
    If Len(styleName) > 0 Then
        Set stylesMap = mp_GetButtonStylesMap()
        If stylesMap Is Nothing Then
            outErrorText = "Button style '" & styleName & "' requested for control '" & controlName & "', but styles map is unavailable."
            Exit Function
        End If
        If Not ex_UiXmlProvider.m_ApplyButtonStyleByName(shp, styleName, stylesMap) Then
            outErrorText = "Failed to apply button style '" & styleName & "' to control '" & controlName & "'."
            Exit Function
        End If
    End If

    mp_ApplyButtonShapeVisuals = True
    Exit Function
EH_APPLY:
    outErrorText = "Failed to apply button visuals for control '" & controlName & "': " & Err.Description
End Function

Private Function mp_ResolveControlCaptionText(ByVal node As Object, ByVal dataItem As Object, ByVal defaultText As String) As String
    Dim rawText As String
    Dim resolvedText As String

    rawText = Trim$(mp_NodeAttrText(node, "caption"))
    If Len(rawText) = 0 Then rawText = Trim$(mp_NodeAttrText(node, "text"))
    If Len(rawText) = 0 Then
        mp_ResolveControlCaptionText = defaultText
        Exit Function
    End If

    resolvedText = mp_ResolveTextValue(rawText, dataItem)
    If Len(resolvedText) = 0 Then
        mp_ResolveControlCaptionText = defaultText
    Else
        mp_ResolveControlCaptionText = resolvedText
    End If
End Function

Private Function mp_ApplyDropdownShapeItems( _
    ByVal node As Object, _
    ByVal shp As Shape, _
    ByVal dataItem As Object, _
    ByVal controlName As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim cf As Object
    Dim itemRecords As Variant
    Dim selectedRaw As String
    Dim selectedText As String
    Dim selectedIndex As Long
    Dim lowerRow As Long
    Dim upperRow As Long
    Dim i As Long
    Dim itemText As String

    On Error Resume Next
    Set cf = shp.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then
        outErrorText = "Control '" & controlName & "' is not a dropdown/combo Form Control."
        Exit Function
    End If

    On Error GoTo EH_CLEAR
    cf.RemoveAllItems
    On Error GoTo 0

    If Not mp_BuildDropdownItemRecords(node, dataItem, controlName, itemRecords, outErrorText) Then Exit Function
    If ex_UiXmlProvider.m_HasDropdownItemRecords(itemRecords) Then
        lowerRow = LBound(itemRecords, 1)
        upperRow = UBound(itemRecords, 1)
        For i = lowerRow To upperRow
            itemText = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
            If Len(itemText) = 0 Then itemText = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
            If Len(itemText) = 0 Then GoTo ContinueAdd
            On Error GoTo EH_ADD
            cf.AddItem itemText
            On Error GoTo 0
ContinueAdd:
        Next i
    End If

    selectedRaw = Trim$(mp_NodeAttrText(node, "selectedItem"))
    If Len(selectedRaw) > 0 Then
        selectedText = mp_ResolveTextValue(selectedRaw, dataItem)
        selectedIndex = mp_FindDropdownRecordIndex(itemRecords, selectedText)
        If selectedIndex = 0 Then
            outErrorText = "selectedItem '" & selectedText & "' was not found for dropdown control '" & controlName & "'."
            Exit Function
        End If
        On Error GoTo EH_SELECT
        cf.Value = selectedIndex - LBound(itemRecords, 1) + 1
        On Error GoTo 0
    End If

    mp_ApplyDropdownShapeItems = True
    Exit Function
EH_CLEAR:
    outErrorText = "Failed to clear dropdown items for control '" & controlName & "': " & Err.Description
    Exit Function
EH_ADD:
    outErrorText = "Failed to add dropdown item for control '" & controlName & "': " & Err.Description
    Exit Function
EH_SELECT:
    outErrorText = "Failed to select dropdown item '" & selectedText & "' for control '" & controlName & "': " & Err.Description
End Function

Private Function mp_ApplyManagedDropdownButton( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal headerShape As Shape, _
    ByVal dataItem As Object, _
    ByVal controlName As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim itemRecords As Variant
    Dim itemStyleName As String
    Dim selectionChangedMacro As String
    Dim selectedItem As String
    Dim itemMarginLeft As Double
    Dim itemFirstGap As Double
    Dim itemGap As Double
    Dim itemHeight As Double
    Dim itemMatchWidth As Boolean
    Dim headerShowsSelection As Boolean
    Dim stylesMap As Object

    If Not mp_BuildDropdownItemRecords(node, dataItem, controlName, itemRecords, outErrorText) Then Exit Function
    If Not ex_UiXmlProvider.m_HasDropdownItemRecords(itemRecords) Then
        outErrorText = "DropdownButton control '" & controlName & "' did not resolve any items."
        Exit Function
    End If

    itemStyleName = Trim$(mp_NodeAttrText(node, "itemStyle"))
    selectionChangedMacro = Trim$(mp_NodeAttrText(node, "selectionChangedMacro"))
    selectedItem = mp_ResolveTextValue(mp_NodeAttrText(node, "selectedItem"), dataItem)

    If Not mp_TryReadOptionalDouble(node, "itemMarginLeft", 0, itemMarginLeft, "control '" & controlName & "'", outErrorText) Then Exit Function
    If Not mp_TryReadOptionalDouble(node, "itemFirstGap", 2, itemFirstGap, "control '" & controlName & "'", outErrorText) Then Exit Function
    If Not mp_TryReadOptionalDouble(node, "itemGap", 2, itemGap, "control '" & controlName & "'", outErrorText) Then Exit Function
    If Not mp_TryReadOptionalDouble(node, "itemHeight", 16, itemHeight, "control '" & controlName & "'", outErrorText) Then Exit Function
    If itemHeight <= 0 Then
        outErrorText = "Control '" & controlName & "' has invalid itemHeight <= 0."
        Exit Function
    End If
    If Not mp_TryReadOptionalBoolean(node, "itemMatchWidth", True, itemMatchWidth, "control '" & controlName & "'", outErrorText) Then Exit Function
    If Not mp_TryReadOptionalBoolean(node, "headerShowsSelection", True, headerShowsSelection, "control '" & controlName & "'", outErrorText) Then Exit Function

    Set stylesMap = Nothing
    If Len(itemStyleName) > 0 Then
        Set stylesMap = mp_GetButtonStylesMap()
        If stylesMap Is Nothing Then
            outErrorText = "DropdownButton '" & controlName & "' requested itemStyle='" & itemStyleName & "', but styles map is unavailable."
            Exit Function
        End If
    End If

    If Not ex_ManagedDropdownRuntime.m_RebuildDropdownButton( _
        ws, _
        headerShape, _
        controlName, _
        itemRecords, _
        itemStyleName, _
        itemMarginLeft, _
        itemFirstGap, _
        itemGap, _
        itemHeight, _
        itemMatchWidth, _
        selectionChangedMacro, _
        "ex_UIActions.m_SelectDropdownOption_OnClick", _
        stylesMap, _
        selectedItem, _
        headerShowsSelection, _
        outErrorText) Then
        Exit Function
    End If

    mp_ApplyManagedDropdownButton = True
End Function

Private Function mp_BuildDropdownItemRecords( _
    ByVal node As Object, _
    ByVal dataItem As Object, _
    ByVal controlName As String, _
    ByRef outRecords As Variant, _
    ByRef outErrorText As String _
) As Boolean
    Dim sourceText As String
    Dim sourceValue As Variant
    Dim bindingPath As String
    Dim ownerDoc As Object
    Dim modeKey As String

    sourceText = Trim$(mp_NodeAttrText(node, "itemsSource"))

    If Len(sourceText) > 0 Then
        If mp_TryExtractBindingPath(sourceText, bindingPath) Then
            sourceValue = mp_ResolveBindingValue(sourceText, dataItem)
            If Not mp_CollectDropdownRecordsFromValue(sourceValue, outRecords, outErrorText) Then Exit Function
            mp_BuildDropdownItemRecords = True
            Exit Function
        End If

        If Not dataItem Is Nothing Then
            If mp_TryGetDictionaryObject(dataItem, sourceText, sourceValue) Then
                If mp_CollectDropdownRecordsFromValue(sourceValue, outRecords, outErrorText) Then
                    mp_BuildDropdownItemRecords = True
                    Exit Function
                End If
            End If
        End If
    End If

    On Error Resume Next
    Set ownerDoc = node.OwnerDocument
    On Error GoTo 0

    modeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey())
    outRecords = ex_UiXmlProvider.m_GetDropdownItemRecordsFromControlNode(node, ThisWorkbook, modeKey, ownerDoc)
    If ex_UiXmlProvider.m_HasDropdownItemRecords(outRecords) Then
        mp_BuildDropdownItemRecords = True
        Exit Function
    End If

    If Len(sourceText) > 0 Then
        outErrorText = "Dropdown control '" & controlName & "' did not resolve itemsSource '" & sourceText & "'."
        Exit Function
    End If

    mp_BuildDropdownItemRecords = True
End Function

Private Function mp_CollectDropdownRecordsFromValue( _
    ByVal sourceValue As Variant, _
    ByRef outRecords As Variant, _
    ByRef outErrorText As String _
) As Boolean
    Dim item As Variant
    Dim rows As Collection
    Dim i As Long

    Set rows = New Collection

    If IsObject(sourceValue) Then
        If TypeName(sourceValue) = "Collection" Then
            For Each item In sourceValue
                mp_AddDropdownRecordRow rows, item
            Next item
            GoTo BuildRecords
        End If
    End If

    If IsArray(sourceValue) Then
        For i = LBound(sourceValue) To UBound(sourceValue)
            mp_AddDropdownRecordRow rows, sourceValue(i)
        Next i
        GoTo BuildRecords
    End If

    mp_AddDropdownRecordRow rows, sourceValue

BuildRecords:
    If Not mp_BuildDropdownRecordsArray(rows, outRecords, outErrorText) Then Exit Function
    mp_CollectDropdownRecordsFromValue = True
End Function

Private Sub mp_AddDropdownRecordRow(ByVal rows As Collection, ByVal valueItem As Variant)
    Dim rowObj As Object
    Dim keyText As String
    Dim captionText As String
    Dim targetText As String
    Dim setContextText As String
    Dim actionKeyText As String
    Dim macroText As String
    Dim candidateValue As Variant
    Dim candidate As Variant

    If rows Is Nothing Then Exit Sub

    If IsObject(valueItem) Then
        If valueItem Is Nothing Then Exit Sub

        For Each candidate In Array("Key", "Value", "Name", "TableRef", "Caption", "Text")
            candidateValue = mp_ReadObjectMember(valueItem, CStr(candidate))
            If Not IsObject(candidateValue) Then
                keyText = Trim$(CStr(candidateValue))
                If Len(keyText) > 0 Then Exit For
            End If
        Next candidate
        For Each candidate In Array("Caption", "Text", "Value", "Name", "Key", "TableRef")
            candidateValue = mp_ReadObjectMember(valueItem, CStr(candidate))
            If Not IsObject(candidateValue) Then
                captionText = Trim$(CStr(candidateValue))
                If Len(captionText) > 0 Then Exit For
            End If
        Next candidate

        candidateValue = mp_ReadObjectMember(valueItem, "Target")
        If Not IsObject(candidateValue) Then targetText = Trim$(CStr(candidateValue))
        candidateValue = mp_ReadObjectMember(valueItem, "SetContext")
        If Not IsObject(candidateValue) Then setContextText = Trim$(CStr(candidateValue))
        candidateValue = mp_ReadObjectMember(valueItem, "ActionKey")
        If Not IsObject(candidateValue) Then actionKeyText = Trim$(CStr(candidateValue))
        candidateValue = mp_ReadObjectMember(valueItem, "Macro")
        If Not IsObject(candidateValue) Then macroText = Trim$(CStr(candidateValue))
    Else
        If IsNull(valueItem) Then Exit Sub
        If IsEmpty(valueItem) Then Exit Sub
        keyText = Trim$(CStr(valueItem))
        captionText = keyText
    End If

    If Len(keyText) = 0 And Len(captionText) = 0 Then Exit Sub
    If Len(captionText) = 0 Then captionText = keyText
    If Len(keyText) = 0 Then keyText = captionText

    Set rowObj = CreateObject("Scripting.Dictionary")
    rowObj.CompareMode = 1
    rowObj("key") = keyText
    rowObj("caption") = captionText
    rowObj("target") = targetText
    rowObj("setContext") = setContextText
    rowObj("actionKey") = actionKeyText
    rowObj("macro") = macroText
    rows.Add rowObj
End Sub

Private Function mp_BuildDropdownRecordsArray( _
    ByVal rows As Collection, _
    ByRef outRecords As Variant, _
    ByRef outErrorText As String _
) As Boolean
    Dim result() As Variant
    Dim i As Long
    Dim rowObj As Object

    If rows Is Nothing Then
        outRecords = Array()
        mp_BuildDropdownRecordsArray = True
        Exit Function
    End If
    If rows.Count = 0 Then
        outRecords = Array()
        mp_BuildDropdownRecordsArray = True
        Exit Function
    End If

    ReDim result(1 To rows.Count, 1 To ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO)
    For i = 1 To rows.Count
        Set rowObj = rows(i)
        result(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY) = CStr(rowObj("key"))
        result(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION) = CStr(rowObj("caption"))
        result(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_TARGET) = CStr(rowObj("target"))
        result(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_SET_CONTEXT) = CStr(rowObj("setContext"))
        result(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_ACTION_KEY) = CStr(rowObj("actionKey"))
        result(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_MACRO) = CStr(rowObj("macro"))
    Next i

    outRecords = result
    mp_BuildDropdownRecordsArray = True
End Function

Private Function mp_FindDropdownRecordIndex(ByVal itemRecords As Variant, ByVal selectedText As String) As Long
    Dim i As Long
    Dim lowerRow As Long
    Dim upperRow As Long
    Dim keyText As String
    Dim captionText As String
    Dim targetText As String

    selectedText = Trim$(selectedText)
    If Len(selectedText) = 0 Then Exit Function
    If Not ex_UiXmlProvider.m_HasDropdownItemRecords(itemRecords) Then Exit Function

    lowerRow = LBound(itemRecords, 1)
    upperRow = UBound(itemRecords, 1)
    For i = lowerRow To upperRow
        keyText = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
        captionText = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_CAPTION)))
        targetText = Trim$(CStr(itemRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_TARGET)))

        If StrComp(keyText, selectedText, vbTextCompare) = 0 Then
            mp_FindDropdownRecordIndex = i
            Exit Function
        End If
        If StrComp(captionText, selectedText, vbTextCompare) = 0 Then
            mp_FindDropdownRecordIndex = i
            Exit Function
        End If
        If StrComp(targetText, selectedText, vbTextCompare) = 0 Then
            mp_FindDropdownRecordIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function mp_FindDropdownControlItemIndex(ByVal cf As Object, ByVal itemText As String) As Long
    Dim itemCount As Long
    Dim i As Long

    itemText = Trim$(itemText)
    If Len(itemText) = 0 Then Exit Function

    On Error Resume Next
    itemCount = CLng(cf.ListCount)
    On Error GoTo 0
    If itemCount <= 0 Then Exit Function

    For i = 1 To itemCount
        On Error Resume Next
        If StrComp(CStr(cf.List(i)), itemText, vbTextCompare) = 0 Then
            On Error GoTo 0
            mp_FindDropdownControlItemIndex = i
            Exit Function
        End If
        On Error GoTo 0
    Next i
End Function

Private Function mp_TryReadOptionalDouble( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Double, _
    ByRef outValue As Double, _
    ByVal contextText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim valueText As String
    Dim parsedValue As Double

    outValue = defaultValue
    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        mp_TryReadOptionalDouble = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseDouble(valueText, parsedValue, True) Then
        outErrorText = "Invalid numeric value '" & valueText & "' for attribute '" & attrName & "' in " & contextText & "."
        Exit Function
    End If

    outValue = parsedValue
    mp_TryReadOptionalDouble = True
End Function

Private Function mp_TryReadOptionalBoolean( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Boolean, _
    ByRef outValue As Boolean, _
    ByVal contextText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim valueText As String
    Dim parsedValue As Boolean

    outValue = defaultValue
    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        mp_TryReadOptionalBoolean = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        outErrorText = "Invalid boolean value '" & valueText & "' for attribute '" & attrName & "' in " & contextText & "."
        Exit Function
    End If

    outValue = parsedValue
    mp_TryReadOptionalBoolean = True
End Function

Private Function mp_TryParsePlacement(ByVal valueText As String, ByRef result As XlPlacement) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "absolute", "free", "freefloating"
            result = xlFreeFloating
            mp_TryParsePlacement = True
        Case "move", "movewithcells"
            result = xlMove
            mp_TryParsePlacement = True
        Case "moveandsize", "move_and_size", "move-size", "moveandresize"
            result = xlMoveAndSize
            mp_TryParsePlacement = True
    End Select
End Function

Private Function mp_ArrayHasItems(ByVal values As Variant) As Boolean
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayHasItems = (UBound(values) >= LBound(values))
    End If
    Exit Function
EH:
    mp_ArrayHasItems = False
End Function

Private Function mp_GetButtonStylesMap() As Object
    If g_ButtonStylesMap Is Nothing Then
        Set g_ButtonStylesMap = CreateObject("Scripting.Dictionary")
    End If
    Set mp_GetButtonStylesMap = g_ButtonStylesMap
End Function

Private Function mp_ApplyCellStyleByName( _
    ByVal targetRange As Range, _
    ByVal styleName As String, _
    ByVal stylesMap As Object, _
    ByVal controlName As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim styleData As Object
    Dim colorValue As Long
    Dim numberValue As Double
    Dim boolValue As Boolean

    If targetRange Is Nothing Then Exit Function

    styleName = Trim$(styleName)
    If Len(styleName) = 0 Then
        mp_ApplyCellStyleByName = True
        Exit Function
    End If

    If stylesMap Is Nothing Then
        outErrorText = "Control '" & controlName & "' references style '" & styleName & "', but styles map is unavailable."
        mp_DebugLog "mp_ApplyCellStyleByName: styles map unavailable for control='" & controlName & "', style='" & styleName & "'."
        Exit Function
    End If
    If Not stylesMap.Exists(styleName) Then
        outErrorText = "Control '" & controlName & "' references missing style '" & styleName & "'."
        mp_DebugLog "mp_ApplyCellStyleByName: missing style='" & styleName & "' for control='" & controlName & "'."
        Exit Function
    End If

    Set styleData = stylesMap(styleName)
    On Error GoTo EH

    If styleData.Exists("backColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("backColor")), colorValue) Then
            outErrorText = "Invalid style backColor for style '" & styleName & "'."
            Exit Function
        End If
        targetRange.Interior.Pattern = xlSolid
        targetRange.Interior.Color = colorValue
    End If

    If styleData.Exists("textColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("textColor")), colorValue) Then
            outErrorText = "Invalid style textColor for style '" & styleName & "'."
            Exit Function
        End If
        targetRange.Font.Color = colorValue
    End If

    If styleData.Exists("fontName") Then
        targetRange.Font.Name = CStr(styleData("fontName"))
    End If

    If styleData.Exists("fontSize") Then
        If Not ex_XmlCore.m_TryParseDouble(CStr(styleData("fontSize")), numberValue, True) Then
            outErrorText = "Invalid style fontSize for style '" & styleName & "'."
            Exit Function
        End If
        targetRange.Font.Size = numberValue
    End If

    If styleData.Exists("fontBold") Then
        If Not ex_XmlCore.m_TryParseBoolean(CStr(styleData("fontBold")), boolValue) Then
            outErrorText = "Invalid style fontBold for style '" & styleName & "'."
            Exit Function
        End If
        targetRange.Font.Bold = boolValue
    End If

    If styleData.Exists("borderColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("borderColor")), colorValue) Then
            outErrorText = "Invalid style borderColor for style '" & styleName & "'."
            Exit Function
        End If
        targetRange.Borders.LineStyle = xlContinuous
        targetRange.Borders.Color = colorValue
    End If

    If styleData.Exists("borderWeight") Then
        If Not ex_XmlCore.m_TryParseDouble(CStr(styleData("borderWeight")), numberValue, True) Then
            outErrorText = "Invalid style borderWeight for style '" & styleName & "'."
            Exit Function
        End If
        targetRange.Borders.Weight = numberValue
    End If

    mp_ApplyCellStyleByName = True
    mp_DebugLog "mp_ApplyCellStyleByName: applied style='" & styleName & "' to control='" & controlName & "'."
    Exit Function
EH:
    outErrorText = "Failed to apply style '" & styleName & "' to control '" & controlName & "': " & Err.Description
    mp_DebugLog "mp_ApplyCellStyleByName: failed style='" & styleName & "' control='" & controlName & "' error='" & Err.Description & "'."
End Function

Private Function mp_ResolveTableObjectForNode(ByVal tableNode As Object, ByVal dataItem As Object) As obj_ResultTable
    Dim itemsSourceText As String
    Dim bindingPath As String
    Dim sourceValue As Variant
    Dim rawObj As Object
    Dim sourceTypeName As String

    On Error GoTo EH_RESOLVE_TABLE_OBJ

    If dataItem Is Nothing Then Exit Function

    itemsSourceText = Trim$(mp_NodeAttrText(tableNode, "itemsSource"))
    mp_DebugLog "mp_ResolveTableObjectForNode: itemsSource='" & itemsSourceText & "'."
    If mp_TryExtractBindingPath(itemsSourceText, bindingPath) Then
        If StrComp(Trim$(bindingPath), "Rows", vbTextCompare) = 0 Then
            If (TypeName(dataItem) = "Dictionary" Or TypeName(dataItem) = "Scripting.Dictionary") Then
                If mp_TryGetDictionaryObjectRef(dataItem, "__raw", rawObj) Then
                    If TypeName(rawObj) = "obj_ResultTable" Then
                        Set mp_ResolveTableObjectForNode = rawObj
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    If Len(itemsSourceText) > 0 Then
        sourceValue = mp_ResolveBindingValue(itemsSourceText, dataItem)
        If IsObject(sourceValue) Then
            sourceTypeName = TypeName(sourceValue)
            mp_DebugLog "mp_ResolveTableObjectForNode: resolved binding object type='" & sourceTypeName & "'."
            If TypeName(sourceValue) = "obj_ResultTable" Then
                Set mp_ResolveTableObjectForNode = sourceValue
                Exit Function
            End If
        Else
            mp_DebugLog "mp_ResolveTableObjectForNode: resolved binding scalar."
        End If
    End If

    On Error Resume Next
    If (TypeName(dataItem) = "Dictionary" Or TypeName(dataItem) = "Scripting.Dictionary") Then
        If mp_TryGetDictionaryObjectRef(dataItem, "__raw", rawObj) Then
            If TypeName(rawObj) = "obj_ResultTable" Then
                Set mp_ResolveTableObjectForNode = rawObj
            End If
        End If
    End If
    On Error GoTo 0
    Exit Function

EH_RESOLVE_TABLE_OBJ:
    mp_DebugLog "mp_ResolveTableObjectForNode: error [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_ResolveRowsObjectForTableNode(ByVal tableNode As Object, ByVal dataItem As Object) As Collection
    Dim tableObj As obj_ResultTable
    Dim itemsSourceText As String
    Dim rowsObjRef As Object
    Dim sourceValue As Variant

    If Not dataItem Is Nothing Then
        If (TypeName(dataItem) = "Dictionary" Or TypeName(dataItem) = "Scripting.Dictionary") Then
            If mp_TryGetDictionaryObjectRef(dataItem, "Rows", rowsObjRef) Then
                If TypeName(rowsObjRef) = "Collection" Then
                    Set mp_ResolveRowsObjectForTableNode = rowsObjRef
                    Exit Function
                End If
            End If
        End If
    End If

    Set tableObj = mp_ResolveTableObjectForNode(tableNode, dataItem)
    If Not tableObj Is Nothing Then
        Set mp_ResolveRowsObjectForTableNode = tableObj.Rows
        Exit Function
    End If

    itemsSourceText = Trim$(mp_NodeAttrText(tableNode, "itemsSource"))
    If Len(itemsSourceText) = 0 Then Exit Function

    sourceValue = mp_ResolveBindingValue(itemsSourceText, dataItem)
    If IsObject(sourceValue) Then
        If TypeName(sourceValue) = "Collection" Then
            Set mp_ResolveRowsObjectForTableNode = sourceValue
        End If
    End If
End Function

Private Function mp_ResolveTableRef(ByVal tableObj As obj_ResultTable, ByVal dataItem As Object) As String
    Dim tableRefValue As Variant

    If Not tableObj Is Nothing Then
        mp_ResolveTableRef = CStr(tableObj.TableRef)
        Exit Function
    End If

    If dataItem Is Nothing Then Exit Function
    On Error Resume Next
    If (TypeName(dataItem) = "Dictionary" Or TypeName(dataItem) = "Scripting.Dictionary") Then
        If mp_TryGetDictionaryObject(dataItem, "TableRef", tableRefValue) Then
            If Not IsObject(tableRefValue) Then
                mp_ResolveTableRef = Trim$(CStr(tableRefValue))
            End If
        End If
    End If
    On Error GoTo 0
End Function

Private Function mp_GetFieldAliasesForTable(ByVal tableObj As obj_ResultTable, ByVal rowsObj As Collection) As Collection
    Dim aliases As Collection
    Dim key As Variant
    Dim firstRow As obj_ResultRow
    Dim colObj As obj_ResultColumn

    Set aliases = New Collection

    If Not tableObj Is Nothing Then
        If Not tableObj.FieldMapByAlias Is Nothing Then
            For Each key In tableObj.FieldMapByAlias.Keys
                aliases.Add CStr(key)
            Next key
        End If
    End If

    If aliases.Count = 0 Then
        If Not rowsObj Is Nothing Then
            If rowsObj.Count > 0 Then
                Set firstRow = rowsObj(1)
                If Not firstRow Is Nothing Then
                    For Each colObj In firstRow.Columns
                        aliases.Add CStr(colObj.Alias)
                    Next colObj
                End If
            End If
        End If
    End If

    Set mp_GetFieldAliasesForTable = aliases
End Function

Private Sub mp_ResolveTableKindConfig( _
    ByVal tableNode As Object, _
    ByVal dataItem As Object, _
    ByVal tableObj As obj_ResultTable, _
    ByRef outHeaderRowKind As String, _
    ByRef outContentRowKind As String, _
    ByRef outHeaderFieldKind As String, _
    ByRef outContentFieldKind As String, _
    ByRef outVirtualFieldKind As String _
)
    outHeaderRowKind = mp_ResolveTableKindAttr(tableNode, dataItem, TABLE_ATTR_ROW_KIND_HEADER, DEFAULT_TABLE_ROW_KIND_HEADER)
    outContentRowKind = mp_ResolveTableKindAttr(tableNode, dataItem, TABLE_ATTR_ROW_KIND_CONTENT, DEFAULT_TABLE_ROW_KIND_CONTENT)
    outHeaderFieldKind = mp_ResolveTableKindAttr(tableNode, dataItem, TABLE_ATTR_FIELD_KIND_HEADER, outHeaderRowKind)
    outContentFieldKind = mp_ResolveTableKindAttr(tableNode, dataItem, TABLE_ATTR_FIELD_KIND_CONTENT, outContentRowKind)
    outVirtualFieldKind = mp_ResolveTableVirtualFieldKind(tableNode, dataItem, tableObj)
End Sub

Private Function mp_ResolveTableKindAttr( _
    ByVal tableNode As Object, _
    ByVal dataItem As Object, _
    ByVal attrName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim rawText As String

    rawText = Trim$(mp_NodeAttrText(tableNode, attrName))
    If Len(rawText) > 0 Then
        mp_ResolveTableKindAttr = mp_NormalizeKindTags(mp_ResolveTemplateText(rawText, dataItem))
    Else
        mp_ResolveTableKindAttr = mp_NormalizeKindTags(defaultValue)
    End If
End Function

Private Function mp_ResolveTableVirtualFieldKind( _
    ByVal tableNode As Object, _
    ByVal dataItem As Object, _
    ByVal tableObj As obj_ResultTable _
) As String
    Dim rawText As String

    rawText = Trim$(mp_NodeAttrText(tableNode, TABLE_ATTR_FIELD_KIND_VIRTUAL))
    If Len(rawText) > 0 Then
        mp_ResolveTableVirtualFieldKind = mp_NormalizeKindTags(mp_ResolveTemplateText(rawText, dataItem))
        Exit Function
    End If

    If Not tableObj Is Nothing Then
        mp_ResolveTableVirtualFieldKind = mp_NormalizeKindTags(tableObj.GetMetaInfoValue(META_KEY_VIRTUAL_FIELD_KIND, vbNullString))
    End If
End Function

Private Function mp_GetResultTableVirtualAliasLookup(ByVal tableObj As obj_ResultTable) As Object
    Dim aliasesText As String

    If tableObj Is Nothing Then Exit Function

    aliasesText = tableObj.GetMetaInfoValue(META_KEY_VIRTUAL_FIELD_ALIASES, vbNullString)
    If Len(Trim$(aliasesText)) = 0 Then Exit Function

    Set mp_GetResultTableVirtualAliasLookup = mp_GetVirtualAliasesLookupByText(aliasesText)
End Function

Private Sub mp_AddFieldRangesFromAliases( _
    ByVal resultFieldRanges As Collection, _
    ByVal tableObj As obj_ResultTable, _
    ByVal fieldAliases As Collection, _
    ByVal startCol As Long, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long, _
    Optional ByVal baseKind As String = vbNullString, _
    Optional ByVal virtualKind As String = vbNullString, _
    Optional ByVal virtualAliasLookup As Object = Nothing _
)
    Dim i As Long
    Dim fieldAlias As String
    Dim normalizedAlias As String
    Dim mapKey As String
    Dim targetKind As String
    Dim normalizedBaseKind As String
    Dim normalizedVirtualKind As String
    Dim normalizedVirtualTargetKind As String
    Dim fieldMapByAlias As Object
    Dim hasVirtualKindOverlay As Boolean

    If resultFieldRanges Is Nothing Then Exit Sub
    If fieldAliases Is Nothing Then Exit Sub
    If fieldAliases.Count = 0 Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then rowEnd = rowStart

    normalizedBaseKind = mp_NormalizeKindTags(baseKind)
    normalizedVirtualKind = mp_NormalizeKindTags(virtualKind)
    If Len(normalizedVirtualKind) > 0 Then
        normalizedVirtualTargetKind = mp_CombineKindTags(normalizedBaseKind, normalizedVirtualKind)
        hasVirtualKindOverlay = Not virtualAliasLookup Is Nothing
    Else
        normalizedVirtualTargetKind = normalizedBaseKind
    End If

    If Not tableObj Is Nothing Then
        On Error Resume Next
        Set fieldMapByAlias = tableObj.FieldMapByAlias
        On Error GoTo 0
    End If

    For i = 1 To fieldAliases.Count
        fieldAlias = CStr(fieldAliases(i))
        mapKey = fieldAlias
        If Not fieldMapByAlias Is Nothing Then
            If fieldMapByAlias.Exists(fieldAlias) Then
                mapKey = CStr(fieldMapByAlias(fieldAlias))
            End If
        End If

        targetKind = normalizedBaseKind
        If hasVirtualKindOverlay Then
            normalizedAlias = LCase$(Trim$(fieldAlias))
            If Len(normalizedAlias) > 0 Then
                If virtualAliasLookup.Exists(normalizedAlias) Then
                    targetKind = normalizedVirtualTargetKind
                End If
            End If
        End If

        mp_AddRenderedFieldTarget resultFieldRanges, mapKey, startCol + i - 1, rowStart, rowEnd, targetKind
    Next i
End Sub

Private Function mp_GetVirtualAliasesLookupByText(ByVal aliasesText As String) As Object
    Dim lookupKey As String
    Dim aliasLookup As Object
    Dim aliasTokens As Variant
    Dim token As Variant
    Dim tokenText As String

    lookupKey = LCase$(Trim$(aliasesText))
    If Len(lookupKey) = 0 Then Exit Function

    If g_VirtualAliasesLookupByText Is Nothing Then
        Set g_VirtualAliasesLookupByText = CreateObject("Scripting.Dictionary")
        g_VirtualAliasesLookupByText.CompareMode = 1
    End If

    If g_VirtualAliasesLookupByText.Exists(lookupKey) Then
        Set mp_GetVirtualAliasesLookupByText = g_VirtualAliasesLookupByText(lookupKey)
        Exit Function
    End If

    Set aliasLookup = CreateObject("Scripting.Dictionary")
    aliasLookup.CompareMode = 1

    aliasTokens = Split(lookupKey, "|")
    For Each token In aliasTokens
        tokenText = LCase$(Trim$(CStr(token)))
        If Len(tokenText) > 0 Then aliasLookup(tokenText) = True
    Next token

    Set g_VirtualAliasesLookupByText(lookupKey) = aliasLookup
    Set mp_GetVirtualAliasesLookupByText = aliasLookup
End Function

Private Sub mp_AddRenderedFieldTarget( _
    ByVal resultFieldRanges As Collection, _
    ByVal mapKey As String, _
    ByVal columnIndex As Long, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long, _
    Optional ByVal targetKind As String = vbNullString _
)
    Dim target As Object

    If resultFieldRanges Is Nothing Then Exit Sub
    If columnIndex <= 0 Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then rowEnd = rowStart

    Set target = CreateObject("Scripting.Dictionary")
    target.CompareMode = 1
    target("MapKey") = CStr(mapKey)
    target("ColumnIndex") = CLng(columnIndex)
    target("RowStart") = CLng(rowStart)
    target("RowEnd") = CLng(rowEnd)
    If Len(targetKind) > 0 Then target("Kind") = targetKind
    resultFieldRanges.Add target
End Sub

Private Function mp_CombineKindTags( _
    ByVal primaryTag As String, _
    Optional ByVal secondaryTag As String = vbNullString _
) As String
    Dim merged As String

    merged = mp_AppendKindTokens(vbNullString, primaryTag)
    merged = mp_AppendKindTokens(merged, secondaryTag)
    mp_CombineKindTags = merged
End Function

Private Function mp_NormalizeKindTags(ByVal rawKinds As String) As String
    mp_NormalizeKindTags = mp_AppendKindTokens(vbNullString, rawKinds)
End Function

Private Function mp_AppendKindTokens(ByVal existingKinds As String, ByVal rawKinds As String) As String
    Dim parts As Variant
    Dim i As Long
    Dim tokenText As String

    rawKinds = Trim$(rawKinds)
    If Len(rawKinds) = 0 Then
        mp_AppendKindTokens = existingKinds
        Exit Function
    End If

    parts = Split(rawKinds, "|")
    For i = LBound(parts) To UBound(parts)
        tokenText = LCase$(Trim$(CStr(parts(i))))
        If Len(tokenText) = 0 Then GoTo ContinueToken

        If Len(existingKinds) = 0 Then
            existingKinds = tokenText
        ElseIf InStr(1, "|" & existingKinds & "|", "|" & tokenText & "|", vbBinaryCompare) = 0 Then
            existingKinds = existingKinds & "|" & tokenText
        End If
ContinueToken:
    Next i

    mp_AppendKindTokens = existingKinds
End Function

Private Sub mp_RegisterRowKinds(ByVal rowKindText As String, ByVal rowIndex As Long, ByVal rowKinds As Object)
    Dim tokens As Variant
    Dim token As Variant
    Dim tokenText As String
    Dim rowsCollection As Collection
    Dim pipePos As Long

    If rowKinds Is Nothing Then Exit Sub
    If rowIndex <= 0 Then Exit Sub

    rowKindText = LCase$(Trim$(rowKindText))
    If Len(rowKindText) = 0 Then Exit Sub

    pipePos = InStr(1, rowKindText, "|", vbBinaryCompare)
    If pipePos <= 0 Then
        If rowKinds.Exists(rowKindText) Then
            Set rowsCollection = rowKinds(rowKindText)
        Else
            Set rowsCollection = New Collection
            Set rowKinds(rowKindText) = rowsCollection
        End If
        rowsCollection.Add CLng(rowIndex)
        Exit Sub
    End If

    tokens = Split(rowKindText, "|")
    For Each token In tokens
        tokenText = Trim$(CStr(token))
        If Len(tokenText) = 0 Then GoTo ContinueToken

        If rowKinds.Exists(tokenText) Then
            Set rowsCollection = rowKinds(tokenText)
        Else
            Set rowsCollection = New Collection
            Set rowKinds(tokenText) = rowsCollection
        End If

        rowsCollection.Add CLng(rowIndex)
ContinueToken:
    Next token
End Sub

Private Function mp_TryResolveNodeStart( _
    ByVal node As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal defaultRow As Long, _
    ByVal defaultCol As Long, _
    ByVal allowAt As Boolean, _
    ByRef outRow As Long, _
    ByRef outCol As Long, _
    ByRef outErrorText As String _
) As Boolean

    Dim atText As String
    Dim relRow As Long
    Dim relCol As Long

    outRow = defaultRow
    outCol = defaultCol

    atText = Trim$(mp_NodeAttrText(node, "at"))
    If Len(atText) = 0 Then
        mp_TryResolveNodeStart = True
        Exit Function
    End If

    If Not allowAt Then
        outErrorText = "Layout node '" & mp_NodeTag(node) & "' cannot define 'at' inside flow layout."
        Exit Function
    End If

    If Not mp_TryParseAtText(atText, relRow, relCol) Then
        outErrorText = "Invalid layout coordinate '" & atText & "'. Expected format like r2c17."
        Exit Function
    End If

    outRow = gridAnchorRow + relRow - 1
    outCol = gridAnchorCol + relCol - 1
    mp_TryResolveNodeStart = True
End Function

Private Function mp_TryParseAtText(ByVal atText As String, ByRef outRow As Long, ByRef outCol As Long) As Boolean
    Dim normalized As String
    Dim cPos As Long
    Dim rowText As String
    Dim colText As String

    normalized = LCase$(Trim$(atText))
    normalized = Replace(normalized, " ", "")
    If Len(normalized) < 4 Then Exit Function
    If Left$(normalized, 1) <> "r" Then Exit Function
    cPos = InStr(1, normalized, "c", vbBinaryCompare)
    If cPos <= 2 Then Exit Function
    If cPos >= Len(normalized) Then Exit Function

    rowText = Mid$(normalized, 2, cPos - 2)
    colText = Mid$(normalized, cPos + 1)
    If Not ex_XmlCore.m_TryParseLong(rowText, outRow) Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(colText, outCol) Then Exit Function
    If outRow < 1 Or outCol < 1 Then Exit Function

    mp_TryParseAtText = True
End Function

Private Function mp_TryReadTrackSize( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByRef hasValue As Boolean, _
    ByRef isAuto As Boolean, _
    ByRef trackValue As Long, _
    ByVal contextName As String, _
    ByRef outErrorText As String _
) As Boolean

    Dim textValue As String
    Dim parsed As Long

    textValue = LCase$(Trim$(mp_NodeAttrText(node, attrName)))
    If Len(textValue) = 0 Then
        hasValue = False
        isAuto = False
        trackValue = 0
        mp_TryReadTrackSize = True
        Exit Function
    End If

    hasValue = True
    If StrComp(textValue, "auto", vbTextCompare) = 0 Then
        isAuto = True
        trackValue = 0
        mp_TryReadTrackSize = True
        Exit Function
    End If

    If StrComp(textValue, "*", vbTextCompare) = 0 Then
        outErrorText = "Unsupported value '*' for " & contextName & ". Use numeric tracks or 'auto'."
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseLong(textValue, parsed) Then
        outErrorText = "Invalid numeric value '" & textValue & "' for " & contextName & "."
        Exit Function
    End If
    If parsed < 0 Then
        outErrorText = "Value for " & contextName & " must be >= 0."
        Exit Function
    End If

    isAuto = False
    trackValue = parsed
    mp_TryReadTrackSize = True
End Function

Private Function mp_TryReadLayoutSpanSize( _
    ByVal node As Object, _
    ByVal primaryAttrName As String, _
    ByVal legacyAttrName As String, _
    ByRef hasValue As Boolean, _
    ByRef isAuto As Boolean, _
    ByRef trackValue As Long, _
    ByVal contextName As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim primaryText As String
    Dim legacyText As String

    primaryText = Trim$(mp_NodeAttrText(node, primaryAttrName))
    legacyText = Trim$(mp_NodeAttrText(node, legacyAttrName))

    If Len(legacyText) > 0 Then
        outErrorText = "Attribute '" & legacyAttrName & "' is no longer supported for " & contextName & ". Use '" & primaryAttrName & "'."
        Exit Function
    End If

    mp_TryReadLayoutSpanSize = mp_TryReadTrackSize(node, primaryAttrName, hasValue, isAuto, trackValue, contextName, outErrorText)
End Function

Private Function mp_TryBuildRangeByTracks( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long, _
    ByVal heightTracks As Long, _
    ByVal widthTracks As Long, _
    ByRef outRange As Range) As Boolean

    If ws Is Nothing Then Exit Function
    If startRow < 1 Or startCol < 1 Then Exit Function
    If heightTracks <= 0 Or widthTracks <= 0 Then Exit Function
    If startRow + heightTracks - 1 > ws.Rows.Count Then Exit Function
    If startCol + widthTracks - 1 > ws.Columns.Count Then Exit Function

    Set outRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + heightTracks - 1, startCol + widthTracks - 1))
    mp_TryBuildRangeByTracks = True
End Function

Private Function mp_NodeTag(ByVal node As Object) As String
    On Error Resume Next
    mp_NodeTag = LCase$(Trim$(CStr(node.baseName)))
    On Error GoTo 0
End Function

Private Function mp_CollectActiveChildren(ByVal parentNode As Object, ByRef outChildren As Collection, ByVal contextPrefix As String, ByRef outErrorText As String) As Boolean
    Dim childNodes As Object
    Dim childNode As Object
    Dim isNodeEnabled As Boolean
    Dim contextText As String

    Set childNodes = parentNode.selectNodes("p:stackPanel | p:border | p:control")
    If childNodes Is Nothing Then
        mp_CollectActiveChildren = True
        Exit Function
    End If
    If childNodes.Length = 0 Then
        mp_CollectActiveChildren = True
        Exit Function
    End If

    Set outChildren = New Collection
    For Each childNode In childNodes
        contextText = contextPrefix & " '" & mp_NodeTag(childNode) & "'"
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(childNode, isNodeEnabled, "condition", contextText) Then
            outErrorText = "Invalid condition in " & contextText & "."
            Exit Function
        End If
        If isNodeEnabled Then outChildren.Add childNode
    Next childNode

    mp_CollectActiveChildren = True
End Function

Private Function mp_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    mp_NodeAttrText = CStr(node.Attributes.getNamedItem(attrName).Text)
    If Err.Number <> 0 Then
        Err.Clear
        mp_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

Private Function mp_ResolveTextValue(ByVal rawText As String, ByVal dataItem As Object) As String
    Dim valueObj As Variant

    valueObj = mp_ResolveBindingValue(rawText, dataItem)
    If IsObject(valueObj) Then
        mp_ResolveTextValue = vbNullString
    Else
        mp_ResolveTextValue = CStr(valueObj)
    End If
End Function

Private Function mp_ResolveTemplateText(ByVal rawText As String, ByVal dataItem As Object) As String
    Dim resolved As String
    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim tokenText As String
    Dim replacement As String
    Dim i As Long

    resolved = Trim$(rawText)
    If Len(resolved) = 0 Then Exit Function

    For i = 1 To 20
        tokenStart = InStr(1, resolved, "{Binding ", vbBinaryCompare)
        If tokenStart <= 0 Then Exit For

        tokenEnd = InStr(tokenStart, resolved, "}", vbBinaryCompare)
        If tokenEnd <= tokenStart Then Exit For

        tokenText = Mid$(resolved, tokenStart, tokenEnd - tokenStart + 1)
        replacement = mp_ResolveTextValue(tokenText, dataItem)
        resolved = Left$(resolved, tokenStart - 1) & replacement & Mid$(resolved, tokenEnd + 1)
    Next i

    mp_ResolveTemplateText = Trim$(resolved)
End Function

Private Function mp_ResolveBindingValue(ByVal rawText As String, ByVal dataItem As Object) As Variant
    Dim bindingPath As String
    Dim resolvedValue As Variant

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then
        mp_ResolveBindingValue = vbNullString
        Exit Function
    End If

    If Not mp_TryExtractBindingPath(rawText, bindingPath) Then
        mp_ResolveBindingValue = rawText
        Exit Function
    End If

    resolvedValue = mp_ReadBindingPathValue(dataItem, bindingPath)
    If IsObject(resolvedValue) Then
        Set mp_ResolveBindingValue = resolvedValue
    Else
        mp_ResolveBindingValue = resolvedValue
    End If
End Function

Private Function mp_TryExtractBindingPath(ByVal rawText As String, ByRef outPath As String) As Boolean
    Dim normalized As String

    normalized = Trim$(rawText)
    If Len(normalized) < 10 Then Exit Function
    If Left$(normalized, 9) <> "{Binding " Then Exit Function
    If Right$(normalized, 1) <> "}" Then Exit Function

    outPath = Trim$(Mid$(normalized, 10, Len(normalized) - 10))
    If Len(outPath) = 0 Then outPath = "."
    mp_TryExtractBindingPath = True
End Function

Private Function mp_ReadBindingPathValue(ByVal baseObject As Object, ByVal bindingPath As String) As Variant
    Dim segments As Variant
    Dim segment As Variant
    Dim currentValue As Variant
    Dim currentObject As Object
    Dim configKey As String

    bindingPath = Trim$(bindingPath)
    If Len(bindingPath) = 0 Or bindingPath = "." Then
        Set mp_ReadBindingPathValue = baseObject
        Exit Function
    End If

    If StrComp(LCase$(Left$(bindingPath, 7)), "config.", vbBinaryCompare) = 0 Then
        configKey = Trim$(Mid$(bindingPath, 8))
        If Len(configKey) = 0 Then
            mp_ReadBindingPathValue = vbNullString
        Else
            mp_ReadBindingPathValue = ex_ConfigProvider.m_GetConfigValue(configKey, vbNullString)
        End If
        Exit Function
    End If

    Set currentObject = baseObject
    segments = Split(bindingPath, ".")

    For Each segment In segments
        If Len(Trim$(CStr(segment))) = 0 Then GoTo ContinueSegment

        If currentObject Is Nothing Then
            mp_ReadBindingPathValue = vbNullString
            Exit Function
        End If

        currentValue = mp_ReadObjectMember(currentObject, CStr(segment))
        If IsObject(currentValue) Then
            Set currentObject = currentValue
        Else
            If StrComp(CStr(segment), CStr(segments(UBound(segments))), vbBinaryCompare) = 0 Then
                mp_ReadBindingPathValue = currentValue
                Exit Function
            Else
                mp_ReadBindingPathValue = vbNullString
                Exit Function
            End If
        End If
ContinueSegment:
    Next segment

    If Not currentObject Is Nothing Then
        Set mp_ReadBindingPathValue = currentObject
    Else
        mp_ReadBindingPathValue = vbNullString
    End If
End Function

Private Function mp_ReadObjectMember(ByVal sourceObject As Object, ByVal memberName As String) As Variant
    Dim dictObj As Object
    Dim valueObj As Variant

    memberName = Trim$(memberName)
    If Len(memberName) = 0 Then
        mp_ReadObjectMember = vbNullString
        Exit Function
    End If

    On Error Resume Next
    If TypeName(sourceObject) = "Dictionary" Or TypeName(sourceObject) = "Scripting.Dictionary" Then
        Set dictObj = sourceObject
        If dictObj.Exists(memberName) Then
            Set mp_ReadObjectMember = dictObj.Item(memberName)
            If Err.Number = 0 Then
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear

            valueObj = dictObj.Item(memberName)
            If Err.Number <> 0 Then
                Err.Clear
                mp_ReadObjectMember = vbNullString
            Else
                mp_ReadObjectMember = valueObj
            End If
            On Error GoTo 0
            Exit Function
        End If
    End If
    On Error GoTo 0

    On Error Resume Next
    valueObj = CallByName(sourceObject, memberName, VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        mp_ReadObjectMember = vbNullString
    ElseIf IsObject(valueObj) Then
        Set mp_ReadObjectMember = valueObj
    Else
        mp_ReadObjectMember = valueObj
    End If
    On Error GoTo 0
End Function

Private Function mp_TryGetDictionaryObject(ByVal sourceObject As Object, ByVal keyName As String, ByRef outValue As Variant) As Boolean
    If sourceObject Is Nothing Then Exit Function

    On Error Resume Next
    If TypeName(sourceObject) = "Dictionary" Or TypeName(sourceObject) = "Scripting.Dictionary" Then
        If sourceObject.Exists(keyName) Then
            Set outValue = sourceObject.Item(keyName)
            If Err.Number = 0 Then
                mp_TryGetDictionaryObject = True
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear

            outValue = sourceObject.Item(keyName)
            If Err.Number = 0 Then
                mp_TryGetDictionaryObject = True
            Else
                Err.Clear
            End If
        End If
    End If
    On Error GoTo 0
End Function

Private Function mp_TryGetDictionaryObjectRef(ByVal sourceObject As Object, ByVal keyName As String, ByRef outObject As Object) As Boolean
    Dim dictObj As Object

    If sourceObject Is Nothing Then Exit Function

    On Error Resume Next
    If TypeName(sourceObject) = "Dictionary" Or TypeName(sourceObject) = "Scripting.Dictionary" Then
        Set dictObj = sourceObject
        If Not dictObj Is Nothing Then
            If dictObj.Exists(keyName) Then
                Set outObject = dictObj.Item(keyName)
                If Err.Number = 0 Then
                    If Not outObject Is Nothing Then
                        mp_TryGetDictionaryObjectRef = True
                    End If
                Else
                    Err.Clear
                End If
            End If
        End If
    End If
    On Error GoTo 0
End Function

Private Function mp_TryGetDictionaryCountText(ByVal dictObj As Object) As String
    On Error Resume Next
    If dictObj Is Nothing Then
        mp_TryGetDictionaryCountText = "0"
    Else
        mp_TryGetDictionaryCountText = CStr(dictObj.Count)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        mp_TryGetDictionaryCountText = "n/a"
    End If
    On Error GoTo 0
End Function

Private Function mp_NodeDependsOnItemsSourceKey(ByVal node As Object, ByVal itemsSourceKey As String) As Boolean
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim sourceText As String

    itemsSourceKey = Trim$(itemsSourceKey)
    If node Is Nothing Then Exit Function
    If Len(itemsSourceKey) = 0 Then Exit Function

    If StrComp(mp_NodeTag(node), "control", vbTextCompare) = 0 Then
        sourceText = Trim$(mp_NodeAttrText(node, "itemsSource"))
        If mp_SourceDependsOnItemsSourceKey(sourceText, itemsSourceKey) Then
            mp_NodeDependsOnItemsSourceKey = True
            Exit Function
        End If
    End If

    Set controlNodes = node.selectNodes(".//*[local-name()='control'][@itemsSource]")
    If controlNodes Is Nothing Then Exit Function

    For Each controlNode In controlNodes
        sourceText = Trim$(mp_NodeAttrText(controlNode, "itemsSource"))
        If mp_SourceDependsOnItemsSourceKey(sourceText, itemsSourceKey) Then
            mp_NodeDependsOnItemsSourceKey = True
            Exit Function
        End If
    Next controlNode
End Function

Private Function mp_SourceDependsOnItemsSourceKey(ByVal sourceText As String, ByVal itemsSourceKey As String) As Boolean
    Dim bindingPath As String

    sourceText = Trim$(sourceText)
    itemsSourceKey = Trim$(itemsSourceKey)

    If Len(sourceText) = 0 Then Exit Function
    If Len(itemsSourceKey) = 0 Then Exit Function

    If StrComp(sourceText, itemsSourceKey, vbTextCompare) = 0 Then
        mp_SourceDependsOnItemsSourceKey = True
        Exit Function
    End If

    If mp_TryExtractBindingPath(sourceText, bindingPath) Then
        If StrComp(bindingPath, itemsSourceKey, vbTextCompare) = 0 Then
            mp_SourceDependsOnItemsSourceKey = True
        End If
    End If
End Function

Private Function mp_NodeContainsTableControl(ByVal node As Object, Optional ByVal doc As Object = Nothing) As Boolean
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim controlType As String

    If node Is Nothing Then Exit Function

    If StrComp(mp_NodeTag(node), "control", vbTextCompare) = 0 Then
        controlType = LCase$(Trim$(mp_NodeAttrText(node, "type")))
        If StrComp(controlType, "table", vbTextCompare) = 0 Then
            mp_NodeContainsTableControl = True
            Exit Function
        End If
        If StrComp(controlType, "itemspanel", vbTextCompare) = 0 Then
            If mp_ItemsPanelTemplateContainsTable(doc, node) Then
                mp_NodeContainsTableControl = True
                Exit Function
            End If
        End If
    End If

    Set controlNodes = node.selectNodes(".//*[local-name()='control']")
    If controlNodes Is Nothing Then Exit Function

    For Each controlNode In controlNodes
        controlType = LCase$(Trim$(mp_NodeAttrText(controlNode, "type")))
        If StrComp(controlType, "table", vbTextCompare) = 0 Then
            mp_NodeContainsTableControl = True
            Exit Function
        End If
        If StrComp(controlType, "itemspanel", vbTextCompare) = 0 Then
            If mp_ItemsPanelTemplateContainsTable(doc, controlNode) Then
                mp_NodeContainsTableControl = True
                Exit Function
            End If
        End If
    Next controlNode
End Function

Private Function mp_ItemsPanelTemplateContainsTable(ByVal doc As Object, ByVal itemsPanelNode As Object) As Boolean
    Dim templateName As String
    Dim templateNode As Object
    Dim templateRoots As Object
    Dim templateRoot As Object
    Dim templateControlNodes As Object
    Dim templateControlNode As Object
    Dim controlType As String

    If itemsPanelNode Is Nothing Then Exit Function
    If doc Is Nothing Then
        mp_ItemsPanelTemplateContainsTable = True
        Exit Function
    End If

    templateName = Trim$(mp_NodeAttrText(itemsPanelNode, "itemTemplate"))
    If Len(templateName) = 0 Then
        mp_ItemsPanelTemplateContainsTable = True
        Exit Function
    End If

    On Error GoTo EH
    Set templateNode = doc.selectSingleNode("/p:uiDefinition/p:templates/p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]")
    If templateNode Is Nothing Then
        mp_ItemsPanelTemplateContainsTable = True
        Exit Function
    End If

    Set templateRoots = templateNode.selectNodes("p:stackPanel | p:border | p:control")
    If templateRoots Is Nothing Then Exit Function

    For Each templateRoot In templateRoots
        If StrComp(mp_NodeTag(templateRoot), "control", vbTextCompare) = 0 Then
            controlType = LCase$(Trim$(mp_NodeAttrText(templateRoot, "type")))
            If StrComp(controlType, "table", vbTextCompare) = 0 Then
                mp_ItemsPanelTemplateContainsTable = True
                Exit Function
            End If
        End If

        Set templateControlNodes = templateRoot.selectNodes(".//*[local-name()='control']")
        If templateControlNodes Is Nothing Then GoTo ContinueTemplateRoot
        For Each templateControlNode In templateControlNodes
            controlType = LCase$(Trim$(mp_NodeAttrText(templateControlNode, "type")))
            If StrComp(controlType, "table", vbTextCompare) = 0 Then
                mp_ItemsPanelTemplateContainsTable = True
                Exit Function
            End If
        Next templateControlNode
ContinueTemplateRoot:
    Next templateRoot

    Exit Function
EH:
    mp_ItemsPanelTemplateContainsTable = True
End Function

Private Sub mp_ClearCellsRegion( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
)
    If ws Is Nothing Then Exit Sub
    If rowStart < 1 Then rowStart = 1
    If colStart < 1 Then colStart = 1
    If rowEnd < rowStart Then Exit Sub
    If colEnd < colStart Then Exit Sub
    If rowEnd > ws.Rows.Count Then rowEnd = ws.Rows.Count
    If colEnd > ws.Columns.Count Then colEnd = ws.Columns.Count

    With ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
        .UnMerge
        .ClearContents
    End With
End Sub

Private Function mp_BuildRootRenderId(ByVal gridOrdinal As Long, ByVal rootOrdinal As Long) As String
    mp_BuildRootRenderId = "g" & CStr(gridOrdinal) & ":r" & CStr(rootOrdinal)
End Function

Private Sub mp_EnsureRootBoundsStore()
    If Not g_RootBoundsBySheet Is Nothing Then Exit Sub
    Set g_RootBoundsBySheet = CreateObject("Scripting.Dictionary")
    g_RootBoundsBySheet.CompareMode = 1
End Sub

Private Function mp_GetSheetBoundsKey(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Exit Function
    mp_GetSheetBoundsKey = LCase$(Trim$(ws.CodeName))
    If Len(mp_GetSheetBoundsKey) = 0 Then mp_GetSheetBoundsKey = LCase$(Trim$(ws.Name))
End Function

Private Function mp_GetOrCreateSheetRootBounds(ByVal ws As Worksheet) As Object
    Dim sheetKey As String
    Dim sheetBounds As Object

    sheetKey = mp_GetSheetBoundsKey(ws)
    If Len(sheetKey) = 0 Then Exit Function

    mp_EnsureRootBoundsStore
    If g_RootBoundsBySheet.Exists(sheetKey) Then
        Set mp_GetOrCreateSheetRootBounds = g_RootBoundsBySheet(sheetKey)
        Exit Function
    End If

    Set sheetBounds = CreateObject("Scripting.Dictionary")
    sheetBounds.CompareMode = 1
    Set g_RootBoundsBySheet(sheetKey) = sheetBounds
    Set mp_GetOrCreateSheetRootBounds = sheetBounds
End Function

Private Sub mp_ClearRootBoundsForSheet(ByVal ws As Worksheet)
    Dim sheetKey As String

    sheetKey = mp_GetSheetBoundsKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    mp_EnsureRootBoundsStore
    If g_RootBoundsBySheet.Exists(sheetKey) Then g_RootBoundsBySheet.Remove sheetKey
End Sub

Private Sub mp_RecordRootBounds( _
    ByVal ws As Worksheet, _
    ByVal rootRenderId As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal widthCells As Long, _
    ByVal heightRows As Long _
)
    Dim sheetBounds As Object
    Dim bounds As Object

    rootRenderId = Trim$(rootRenderId)
    If Len(rootRenderId) = 0 Then Exit Sub
    If widthCells <= 0 Or heightRows <= 0 Then Exit Sub

    Set sheetBounds = mp_GetOrCreateSheetRootBounds(ws)
    If sheetBounds Is Nothing Then Exit Sub

    Set bounds = CreateObject("Scripting.Dictionary")
    bounds.CompareMode = 1
    bounds(ROOT_BOUNDS_FIELD_ROW) = CLng(rowStart)
    bounds(ROOT_BOUNDS_FIELD_COL) = CLng(colStart)
    bounds(ROOT_BOUNDS_FIELD_WIDTH) = CLng(widthCells)
    bounds(ROOT_BOUNDS_FIELD_HEIGHT) = CLng(heightRows)
    Set sheetBounds(rootRenderId) = bounds
End Sub

Private Function mp_TryGetRootBounds( _
    ByVal ws As Worksheet, _
    ByVal rootRenderId As String, _
    ByRef outRow As Long, _
    ByRef outCol As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long _
) As Boolean
    Dim sheetBounds As Object
    Dim bounds As Object

    rootRenderId = Trim$(rootRenderId)
    If Len(rootRenderId) = 0 Then Exit Function

    Set sheetBounds = mp_GetOrCreateSheetRootBounds(ws)
    If sheetBounds Is Nothing Then Exit Function
    If Not sheetBounds.Exists(rootRenderId) Then Exit Function

    Set bounds = sheetBounds(rootRenderId)
    If bounds Is Nothing Then Exit Function

    outRow = CLng(bounds(ROOT_BOUNDS_FIELD_ROW))
    outCol = CLng(bounds(ROOT_BOUNDS_FIELD_COL))
    outWidth = CLng(bounds(ROOT_BOUNDS_FIELD_WIDTH))
    outHeight = CLng(bounds(ROOT_BOUNDS_FIELD_HEIGHT))
    mp_TryGetRootBounds = True
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ResultLayoutXmlEngine] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
