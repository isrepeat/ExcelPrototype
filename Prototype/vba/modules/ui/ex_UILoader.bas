Attribute VB_Name = "ex_UILoader"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const UI_CONFIG_REL_PATH As String = "config\DevUI.xml"
Private Const DEFAULT_SHEET_NAME As String = "Dev"
Private Const UI_BLOCK_GROUP_NAME As String = "grpUiBlock"
Private Const DEBUG_LOG_PATH As String = "Logs\layout_engine.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Public Sub m_LoadUiFromConfig(Optional ByVal wb As Workbook)
    Dim doc As Object
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim stylesMap As Object
    Dim ws As Worksheet
    Dim controlName As String
    Dim controlType As String
    Dim sheetName As String
    Dim isRequired As Boolean
    Dim createIfMissing As Boolean
    Dim didUngroupUiBlock As Boolean
    Dim regroupSheets As Object
    Dim regroupSheetName As Variant
    Dim clearedInputSheets As Object
    Dim isNodeEnabled As Boolean

    On Error GoTo EH_LOAD

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    If wb Is Nothing Then
        MsgBox "Failed to load UI from config: workbook is not specified.", vbExclamation
        Exit Sub
    End If
    mp_DebugLog "m_LoadUiFromConfig: start workbook='" & wb.Name & "'."

    Set doc = mp_LoadUiDom(wb)
    If doc Is Nothing Then Exit Sub

    If Not mp_ApplyGridLayouts(doc, wb) Then Exit Sub

    Set controlNodes = mp_SelectUiControlNodes(doc)
    If controlNodes Is Nothing Then
        MsgBox "Invalid UI config format. Expected controls in '/uiDefinition/layout//control'.", vbExclamation
        Exit Sub
    End If
    If controlNodes.Length = 0 Then
        MsgBox "UI config is empty: no control nodes found in " & ex_XmlCore.m_CombineBasePath(wb, UI_CONFIG_REL_PATH), vbExclamation
        Exit Sub
    End If
    mp_DebugLog "m_LoadUiFromConfig: control count=" & CStr(controlNodes.Length)

    Set stylesMap = ex_UiXmlProvider.m_ReadButtonStyles(wb)
    mp_DebugLog "m_LoadUiFromConfig: styles count=" & mp_TryGetDictionaryCountText(stylesMap)

    If Not mp_RemoveButtonsMissingInConfig(wb, controlNodes) Then Exit Sub

    Set regroupSheets = CreateObject("Scripting.Dictionary")
    Set clearedInputSheets = CreateObject("Scripting.Dictionary")

    For Each controlNode In controlNodes
        controlName = Trim$(mp_NodeAttrText(controlNode, "name"))
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(controlNode, isNodeEnabled, "condition", "control '" & controlName & "'") Then Exit Sub
        If Not isNodeEnabled Then GoTo NextControl

        If Len(controlName) = 0 Then
            MsgBox "UI config contains <control> without 'name' attribute.", vbExclamation
            Exit Sub
        End If

        controlType = mp_NormalizeControlType(mp_NodeAttrText(controlNode, "type"))
        If Len(controlType) = 0 Then
            controlType = "button"
        End If
        If Not mp_IsSupportedControlType(controlType) Then
            MsgBox "Unsupported UI control type '" & controlType & "' for control '" & controlName & "'.", vbExclamation
            Exit Sub
        End If

        sheetName = Trim$(mp_NodeAttrText(controlNode, "sheet"))
        If Len(sheetName) = 0 Then
            sheetName = DEFAULT_SHEET_NAME
        End If
        Set ws = mp_GetWorksheetByName(wb, sheetName)
        If ws Is Nothing Then
            MsgBox "Sheet '" & sheetName & "' for control '" & controlName & "' was not found in workbook '" & wb.Name & "'.", vbExclamation
            Exit Sub
        End If
        mp_DebugLog "m_LoadUiFromConfig: apply control='" & controlName & "' type='" & controlType & "' sheet='" & ws.Name & "'."

        isRequired = mp_NodeAttrBool(controlNode, "required", True)
        createIfMissing = mp_NodeAttrBool(controlNode, "createIfMissing", False)

        If StrComp(controlType, "input", vbTextCompare) = 0 Then
            If Not clearedInputSheets.Exists(ws.Name) Then
                ex_LayoutBindingsRuntime.m_ClearSheetBindings ws
                clearedInputSheets.Add ws.Name, True
            End If
            If Not mp_ApplyInputCellControl(ws, controlNode, controlName, stylesMap) Then Exit Sub
            GoTo NextControl
        End If

        If Not mp_EnsureControlShape(ws, controlNode, controlName, controlType, createIfMissing, isRequired) Then
            Exit Sub
        End If
        If ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName) Is Nothing Then
            If isRequired Then
                MsgBox "Control '" & controlName & "' is required but was not resolved on sheet '" & ws.Name & "'.", vbExclamation
                Exit Sub
            End If
            GoTo NextControl
        End If

        If Not mp_ApplyControlAttributes(ws, controlNode, controlName, controlType, stylesMap) Then
            Exit Sub
        End If

        If Not mp_AssignShapeMacro(ws, controlNode, controlName, wb, isRequired, didUngroupUiBlock) Then
            Exit Sub
        End If

        If StrComp(controlType, "dropdownbutton", vbTextCompare) = 0 Then
            If Not mp_RebuildManagedDropdownButton(ws, controlNode, controlName, stylesMap) Then Exit Sub
        End If

        If didUngroupUiBlock Then
            If Not regroupSheets.Exists(ws.Name) Then
                regroupSheets.Add ws.Name, True
            End If
        End If
NextControl:
    Next controlNode

    For Each regroupSheetName In regroupSheets.Keys
        Set ws = mp_GetWorksheetByName(wb, CStr(regroupSheetName))
        If Not ws Is Nothing Then
            mp_TryRegroupUiBlock ws
        End If
    Next regroupSheetName
    mp_DebugLog "m_LoadUiFromConfig: completed."
    Exit Sub

EH_LOAD:
    mp_DebugLog "m_LoadUiFromConfig: fail source='" & Err.Source & "' code=" & CStr(Err.Number) & " description='" & Err.Description & "'."
    MsgBox "Update UI failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
End Sub

Private Function mp_SelectUiControlNodes(ByVal doc As Object) As Object
    Set mp_SelectUiControlNodes = doc.selectNodes("/p:uiDefinition/p:layout//p:control")
End Function

Private Function mp_LoadUiDom(ByVal wb As Workbook) As Object
    Set mp_LoadUiDom = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        UI_CONFIG_REL_PATH, _
        PROFILES_NS, _
        "UI config file was not found: ", _
        "Failed to parse UI config file: ")
End Function

Private Function mp_ApplyGridLayouts(ByVal doc As Object, ByVal wb As Workbook) As Boolean
    Dim gridNodes As Object
    Dim gridNode As Object
    Dim rootNodes As Object
    Dim rootNode As Object
    Dim ws As Worksheet
    Dim sheetName As String
    Dim anchorCell As Range
    Dim isNodeEnabled As Boolean
    Dim contextText As String

    Set gridNodes = doc.selectNodes("/p:uiDefinition/p:layout/p:grid")
    If gridNodes Is Nothing Then
        mp_ApplyGridLayouts = True
        Exit Function
    End If
    If gridNodes.Length = 0 Then
        mp_ApplyGridLayouts = True
        Exit Function
    End If

    For Each gridNode In gridNodes
        sheetName = Trim$(mp_NodeAttrText(gridNode, "sheet"))
        If Len(sheetName) = 0 Then sheetName = DEFAULT_SHEET_NAME
        Set ws = mp_GetWorksheetByName(wb, sheetName)
        If ws Is Nothing Then
            MsgBox "Grid layout references missing sheet '" & sheetName & "' in DevUI.xml.", vbExclamation
            Exit Function
        End If

        If Not mp_LayoutTryResolveGridAnchorCell(ws, gridNode, anchorCell) Then Exit Function
        If Not mp_LayoutApplyGridColumns(ws, gridNode, anchorCell) Then Exit Function

        Set rootNodes = gridNode.selectNodes("p:stackPanel | p:border | p:control")
        If rootNodes Is Nothing Then GoTo ContinueGrid

        For Each rootNode In rootNodes
            contextText = "layout root node '" & mp_LayoutNodeTag(rootNode) & "'"
            If Not ex_XmlCore.m_TryEvaluateNodeCondition(rootNode, isNodeEnabled, "condition", contextText) Then Exit Function
            If Not isNodeEnabled Then GoTo ContinueRootNode
            If Not mp_LayoutApplyRootNode(rootNode, ws, anchorCell.Row, anchorCell.Column) Then Exit Function
ContinueRootNode:
        Next rootNode
ContinueGrid:
    Next gridNode

    mp_ApplyGridLayouts = True
End Function

Private Function mp_LayoutTryResolveGridAnchorCell(ByVal ws As Worksheet, ByVal gridNode As Object, ByRef anchorCell As Range) As Boolean
    Dim anchorCellText As String

    anchorCellText = Trim$(mp_NodeAttrText(gridNode, "anchorCell"))
    If Len(anchorCellText) = 0 Then anchorCellText = "A1"

    On Error GoTo EH_ANCHOR
    Set anchorCell = ws.Range(anchorCellText)
    On Error GoTo 0

    mp_LayoutTryResolveGridAnchorCell = True
    Exit Function
EH_ANCHOR:
    MsgBox "Invalid grid@anchorCell value '" & anchorCellText & "' in DevUI.xml.", vbExclamation
End Function

Private Function mp_LayoutApplyGridColumns(ByVal ws As Worksheet, ByVal gridNode As Object, ByVal anchorCell As Range) As Boolean
    Dim colNodes As Object
    Dim colNode As Object
    Dim iText As String
    Dim widthText As String
    Dim colIndexRel As Long
    Dim widthUnits As Double
    Dim targetColAbs As Long

    Set colNodes = gridNode.selectNodes("p:columns/p:col")
    If colNodes Is Nothing Then
        mp_LayoutApplyGridColumns = True
        Exit Function
    End If

    For Each colNode In colNodes
        iText = Trim$(mp_NodeAttrText(colNode, "i"))
        If Len(iText) = 0 Then
            MsgBox "Grid column entry is missing required attribute 'i'.", vbExclamation
            Exit Function
        End If
        If Not ex_XmlCore.m_TryParseLong(iText, colIndexRel) Then
            MsgBox "Grid column has non-numeric i='" & iText & "'.", vbExclamation
            Exit Function
        End If
        If colIndexRel < 1 Then
            MsgBox "Grid column index must be >= 1, got: " & CStr(colIndexRel) & ".", vbExclamation
            Exit Function
        End If

        widthText = Trim$(mp_NodeAttrText(colNode, "width"))
        If Len(widthText) = 0 Then GoTo ContinueCol
        If Not mp_TryParseDouble(widthText, widthUnits) Then
            MsgBox "Grid column has invalid width='" & widthText & "' for i=" & CStr(colIndexRel) & ".", vbExclamation
            Exit Function
        End If
        If widthUnits <= 0 Then
            MsgBox "Grid column width must be > 0 for i=" & CStr(colIndexRel) & ".", vbExclamation
            Exit Function
        End If

        targetColAbs = anchorCell.Column + colIndexRel - 1
        ws.Columns(targetColAbs).ColumnWidth = widthUnits
ContinueCol:
    Next colNode

    mp_LayoutApplyGridColumns = True
End Function

Private Function mp_LayoutApplyRootNode( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long) As Boolean

    Dim rootRow As Long
    Dim rootCol As Long
    Dim rootWidth As Long
    Dim rootHeight As Long

    If Not mp_LayoutMeasureNode(node, ws, gridAnchorRow, gridAnchorCol, False, 0, False, 0, rootWidth, rootHeight) Then Exit Function
    If rootWidth < 0 Or rootHeight < 0 Then
        MsgBox "Layout node '" & mp_LayoutNodeTag(node) & "' produced negative size.", vbExclamation
        Exit Function
    End If
    If Not mp_LayoutTryResolveNodeStart(node, gridAnchorRow, gridAnchorCol, gridAnchorRow, gridAnchorCol, True, rootRow, rootCol) Then Exit Function
    If Not mp_LayoutApplyNode(node, ws, gridAnchorRow, gridAnchorCol, rootRow, rootCol, rootWidth, rootHeight, False) Then Exit Function

    mp_LayoutApplyRootNode = True
End Function

Private Function mp_LayoutApplyNode( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal isFlowChild As Boolean) As Boolean

    Dim tagName As String
    Dim atText As String

    tagName = mp_LayoutNodeTag(node)
    atText = Trim$(mp_NodeAttrText(node, "at"))
    If isFlowChild And Len(atText) > 0 Then
        MsgBox "Layout node '" & tagName & "' cannot define 'at' inside stack/border flow.", vbExclamation
        Exit Function
    End If

    Select Case tagName
        Case "stackpanel"
            mp_LayoutApplyNode = mp_LayoutApplyStackPanel(node, ws, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, nodeWidth, nodeHeight)
        Case "border"
            mp_LayoutApplyNode = mp_LayoutApplyBorder(node, ws, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, nodeWidth, nodeHeight)
        Case "control"
            mp_LayoutApplyNode = mp_LayoutApplyControl(node, ws, nodeRow, nodeCol, nodeWidth, nodeHeight, isFlowChild)
        Case Else
            MsgBox "Unsupported layout node '" & tagName & "'. Allowed: stackPanel, border, control.", vbExclamation
    End Select
End Function

Private Function mp_LayoutMeasureNode( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long) As Boolean

    Dim tagName As String

    tagName = mp_LayoutNodeTag(node)
    Select Case tagName
        Case "stackpanel"
            mp_LayoutMeasureNode = mp_LayoutMeasureStackPanel(node, ws, gridAnchorRow, gridAnchorCol, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight)
        Case "border"
            mp_LayoutMeasureNode = mp_LayoutMeasureBorder(node, ws, gridAnchorRow, gridAnchorCol, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight)
        Case "control"
            mp_LayoutMeasureNode = mp_LayoutMeasureControl(node, hasParentWidth, parentWidth, hasParentHeight, parentHeight, outWidth, outHeight)
        Case Else
            MsgBox "Unsupported layout node '" & tagName & "' during measurement.", vbExclamation
    End Select
End Function

Private Function mp_LayoutMeasureStackPanel( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long) As Boolean

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
        MsgBox "stackPanel must define orientation='vertical' or 'horizontal'.", vbExclamation
        Exit Function
    End If

    If Not mp_LayoutTryReadSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "stackPanel@spanCells") Then Exit Function
    If Not mp_LayoutTryReadSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "stackPanel@spanRows") Then Exit Function

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

    If Not mp_LayoutCollectActiveChildren(node, children, "stackPanel children") Then Exit Function
    If Not children Is Nothing Then
        For Each childNode In children
            If orientation = "vertical" Then
                If Not mp_LayoutMeasureNode(childNode, ws, gridAnchorRow, gridAnchorCol, hasKnownWidth, knownWidth, False, 0, childW, childH) Then Exit Function
                contentH = contentH + childH
                If childW > contentW Then contentW = childW
            Else
                If Not mp_LayoutMeasureNode(childNode, ws, gridAnchorRow, gridAnchorCol, False, 0, hasKnownHeight, knownHeight, childW, childH) Then Exit Function
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

    mp_LayoutMeasureStackPanel = True
End Function

Private Function mp_LayoutMeasureBorder( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long) As Boolean

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

    If Not mp_LayoutTryReadSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "border@spanCells") Then Exit Function
    If Not mp_LayoutTryReadSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "border@spanRows") Then Exit Function

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

    If Not mp_LayoutCollectActiveChildren(node, children, "border children") Then Exit Function
    If Not children Is Nothing Then
        For Each childNode In children
            If Not mp_LayoutMeasureNode(childNode, ws, gridAnchorRow, gridAnchorCol, hasKnownWidth, knownWidth, hasKnownHeight, knownHeight, childW, childH) Then Exit Function
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

    mp_LayoutMeasureBorder = True
End Function

Private Function mp_LayoutMeasureControl( _
    ByVal node As Object, _
    ByVal hasParentWidth As Boolean, _
    ByVal parentWidth As Long, _
    ByVal hasParentHeight As Boolean, _
    ByVal parentHeight As Long, _
    ByRef outWidth As Long, _
    ByRef outHeight As Long) As Boolean

    Dim hasWidthSpec As Boolean
    Dim hasHeightSpec As Boolean
    Dim widthAuto As Boolean
    Dim heightAuto As Boolean
    Dim widthValue As Long
    Dim heightValue As Long
    Dim controlName As String

    controlName = Trim$(mp_NodeAttrText(node, "name"))
    If Len(controlName) = 0 Then
        MsgBox "Layout control node is missing required attribute 'name'.", vbExclamation
        Exit Function
    End If

    If Not mp_LayoutTryReadSpanSize(node, "spanCells", "width", hasWidthSpec, widthAuto, widthValue, "control@spanCells (" & controlName & ")") Then Exit Function
    If Not mp_LayoutTryReadSpanSize(node, "spanRows", "height", hasHeightSpec, heightAuto, heightValue, "control@spanRows (" & controlName & ")") Then Exit Function

    If widthAuto Then
        MsgBox "Control '" & controlName & "' does not support spanCells='auto'. Use numeric spanCells in grid tracks.", vbExclamation
        Exit Function
    End If
    If heightAuto Then
        MsgBox "Control '" & controlName & "' does not support spanRows='auto'. Use numeric spanRows in grid tracks.", vbExclamation
        Exit Function
    End If

    If hasWidthSpec Then
        outWidth = widthValue
    ElseIf hasParentWidth Then
        outWidth = parentWidth
    Else
        MsgBox "Control '" & controlName & "' must define spanCells or be hosted by a parent with known width.", vbExclamation
        Exit Function
    End If

    If hasHeightSpec Then
        outHeight = heightValue
    ElseIf hasParentHeight Then
        outHeight = parentHeight
    Else
        MsgBox "Control '" & controlName & "' must define spanRows or be hosted by a parent with known height.", vbExclamation
        Exit Function
    End If

    mp_LayoutMeasureControl = True
End Function

Private Function mp_LayoutApplyStackPanel( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long) As Boolean

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
        MsgBox "stackPanel must define orientation='vertical' or 'horizontal'.", vbExclamation
        Exit Function
    End If

    If Not mp_LayoutCollectActiveChildren(node, children, "stackPanel children") Then Exit Function

    cursorRow = nodeRow
    cursorCol = nodeCol
    If Not children Is Nothing Then
        For Each childNode In children
            If orientation = "vertical" Then
                If Not mp_LayoutMeasureNode(childNode, ws, gridAnchorRow, gridAnchorCol, True, nodeWidth, False, 0, childW, childH) Then Exit Function
                If childW > nodeWidth Then
                    MsgBox "Child node '" & mp_LayoutNodeTag(childNode) & "' exceeds stackPanel width.", vbExclamation
                    Exit Function
                End If
                If Not mp_LayoutApplyNode(childNode, ws, gridAnchorRow, gridAnchorCol, cursorRow, nodeCol, childW, childH, True) Then Exit Function
                cursorRow = cursorRow + childH
                usedMain = usedMain + childH
            Else
                If Not mp_LayoutMeasureNode(childNode, ws, gridAnchorRow, gridAnchorCol, False, 0, True, nodeHeight, childW, childH) Then Exit Function
                If childH > nodeHeight Then
                    MsgBox "Child node '" & mp_LayoutNodeTag(childNode) & "' exceeds stackPanel height.", vbExclamation
                    Exit Function
                End If
                If Not mp_LayoutApplyNode(childNode, ws, gridAnchorRow, gridAnchorCol, nodeRow, cursorCol, childW, childH, True) Then Exit Function
                cursorCol = cursorCol + childW
                usedMain = usedMain + childW
            End If
        Next childNode
    End If

    If orientation = "vertical" Then
        If usedMain > nodeHeight Then
            MsgBox "stackPanel content height exceeds allocated height.", vbExclamation
            Exit Function
        End If
    Else
        If usedMain > nodeWidth Then
            MsgBox "stackPanel content width exceeds allocated width.", vbExclamation
            Exit Function
        End If
    End If

    mp_LayoutApplyStackPanel = True
End Function

Private Function mp_LayoutApplyBorder( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long) As Boolean

    Dim children As Collection
    Dim childNode As Object
    Dim childW As Long
    Dim childH As Long

    If Not mp_LayoutCollectActiveChildren(node, children, "border children") Then Exit Function
    If children Is Nothing Then
        mp_LayoutApplyBorder = True
        Exit Function
    End If

    For Each childNode In children
        If Not mp_LayoutMeasureNode(childNode, ws, gridAnchorRow, gridAnchorCol, True, nodeWidth, True, nodeHeight, childW, childH) Then Exit Function
        If childW > nodeWidth Or childH > nodeHeight Then
            MsgBox "Child node '" & mp_LayoutNodeTag(childNode) & "' exceeds border bounds.", vbExclamation
            Exit Function
        End If
        If Not mp_LayoutApplyNode(childNode, ws, gridAnchorRow, gridAnchorCol, nodeRow, nodeCol, childW, childH, True) Then Exit Function
    Next childNode

    mp_LayoutApplyBorder = True
End Function

Private Function mp_LayoutApplyControl( _
    ByVal node As Object, _
    ByVal ws As Worksheet, _
    ByVal nodeRow As Long, _
    ByVal nodeCol As Long, _
    ByVal nodeWidth As Long, _
    ByVal nodeHeight As Long, _
    ByVal isFlowChild As Boolean) As Boolean

    Dim controlName As String
    Dim controlRange As Range
    Dim atText As String
    Dim sheetName As String

    controlName = Trim$(mp_NodeAttrText(node, "name"))
    If Len(controlName) = 0 Then
        MsgBox "Layout control node is missing required attribute 'name'.", vbExclamation
        Exit Function
    End If

    atText = Trim$(mp_NodeAttrText(node, "at"))
    If isFlowChild And Len(atText) > 0 Then
        MsgBox "Control '" & controlName & "' cannot define 'at' when hosted in stack/border flow.", vbExclamation
        Exit Function
    End If

    If nodeWidth <= 0 Then
        MsgBox "Control '" & controlName & "' has non-positive width after layout resolution.", vbExclamation
        Exit Function
    End If
    If nodeHeight <= 0 Then
        MsgBox "Control '" & controlName & "' has non-positive height after layout resolution.", vbExclamation
        Exit Function
    End If

    If Not mp_LayoutTryBuildRangeByTracks(ws, nodeRow, nodeCol, nodeHeight, nodeWidth, controlRange) Then
        MsgBox "Control '" & controlName & "' resolved to invalid grid bounds.", vbExclamation
        Exit Function
    End If

    mp_SetNodeAttrDouble node, "left", controlRange.Left
    mp_SetNodeAttrDouble node, "top", controlRange.Top
    mp_SetNodeAttrDouble node, "width", controlRange.Width
    mp_SetNodeAttrDouble node, "height", controlRange.Height
    mp_SetNodeAttrLong node, "gridRow", nodeRow
    mp_SetNodeAttrLong node, "gridCol", nodeCol
    mp_SetNodeAttrLong node, "gridWidth", nodeWidth
    mp_SetNodeAttrLong node, "gridHeight", nodeHeight

    sheetName = Trim$(mp_NodeAttrText(node, "sheet"))
    If Len(sheetName) = 0 Then node.setAttribute "sheet", ws.Name

    mp_LayoutApplyControl = True
End Function

Private Function mp_LayoutCollectActiveChildren(ByVal parentNode As Object, ByRef outChildren As Collection, ByVal contextPrefix As String) As Boolean
    Dim childNodes As Object
    Dim childNode As Object
    Dim isNodeEnabled As Boolean
    Dim contextText As String

    Set childNodes = parentNode.selectNodes("p:stackPanel | p:border | p:control")
    If childNodes Is Nothing Then
        mp_LayoutCollectActiveChildren = True
        Exit Function
    End If
    If childNodes.Length = 0 Then
        mp_LayoutCollectActiveChildren = True
        Exit Function
    End If

    Set outChildren = New Collection
    For Each childNode In childNodes
        contextText = contextPrefix & " '" & mp_LayoutNodeTag(childNode) & "'"
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(childNode, isNodeEnabled, "condition", contextText) Then Exit Function
        If isNodeEnabled Then outChildren.Add childNode
    Next childNode

    mp_LayoutCollectActiveChildren = True
End Function

Private Function mp_LayoutTryResolveNodeStart( _
    ByVal node As Object, _
    ByVal gridAnchorRow As Long, _
    ByVal gridAnchorCol As Long, _
    ByVal defaultRow As Long, _
    ByVal defaultCol As Long, _
    ByVal allowAt As Boolean, _
    ByRef outRow As Long, _
    ByRef outCol As Long) As Boolean

    Dim atText As String
    Dim relRow As Long
    Dim relCol As Long

    outRow = defaultRow
    outCol = defaultCol

    atText = Trim$(mp_NodeAttrText(node, "at"))
    If Len(atText) = 0 Then
        mp_LayoutTryResolveNodeStart = True
        Exit Function
    End If

    If Not allowAt Then
        MsgBox "Layout node '" & mp_LayoutNodeTag(node) & "' cannot define 'at' inside flow layout.", vbExclamation
        Exit Function
    End If

    If Not mp_LayoutTryParseAtText(atText, relRow, relCol) Then
        MsgBox "Invalid layout coordinate '" & atText & "'. Expected format like r2c17.", vbExclamation
        Exit Function
    End If

    outRow = gridAnchorRow + relRow - 1
    outCol = gridAnchorCol + relCol - 1
    mp_LayoutTryResolveNodeStart = True
End Function

Private Function mp_LayoutTryParseAtText(ByVal atText As String, ByRef outRow As Long, ByRef outCol As Long) As Boolean
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

    mp_LayoutTryParseAtText = True
End Function

Private Function mp_LayoutTryReadTrackSize( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByRef hasValue As Boolean, _
    ByRef isAuto As Boolean, _
    ByRef trackValue As Long, _
    ByVal contextName As String) As Boolean

    Dim textValue As String
    Dim parsed As Long

    textValue = LCase$(Trim$(mp_NodeAttrText(node, attrName)))
    If Len(textValue) = 0 Then
        hasValue = False
        isAuto = False
        trackValue = 0
        mp_LayoutTryReadTrackSize = True
        Exit Function
    End If

    hasValue = True
    If StrComp(textValue, "auto", vbTextCompare) = 0 Then
        isAuto = True
        trackValue = 0
        mp_LayoutTryReadTrackSize = True
        Exit Function
    End If

    If StrComp(textValue, "*", vbTextCompare) = 0 Then
        MsgBox "Unsupported value '*' for " & contextName & ". Use numeric tracks or 'auto'.", vbExclamation
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseLong(textValue, parsed) Then
        MsgBox "Invalid numeric value '" & textValue & "' for " & contextName & ".", vbExclamation
        Exit Function
    End If
    If parsed < 0 Then
        MsgBox "Value for " & contextName & " must be >= 0.", vbExclamation
        Exit Function
    End If

    isAuto = False
    trackValue = parsed
    mp_LayoutTryReadTrackSize = True
End Function

Private Function mp_LayoutTryReadSpanSize( _
    ByVal node As Object, _
    ByVal primaryAttrName As String, _
    ByVal legacyAttrName As String, _
    ByRef hasValue As Boolean, _
    ByRef isAuto As Boolean, _
    ByRef trackValue As Long, _
    ByVal contextName As String) As Boolean
    Dim legacyText As String

    legacyText = Trim$(mp_NodeAttrText(node, legacyAttrName))
    If Len(legacyText) > 0 Then
        MsgBox "Attribute '" & legacyAttrName & "' is no longer supported for " & contextName & ". Use '" & primaryAttrName & "'.", vbExclamation
        Exit Function
    End If

    mp_LayoutTryReadSpanSize = mp_LayoutTryReadTrackSize(node, primaryAttrName, hasValue, isAuto, trackValue, contextName)
End Function

Private Function mp_LayoutTryBuildRangeByTracks( _
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
    mp_LayoutTryBuildRangeByTracks = True
End Function

Private Function mp_LayoutNodeTag(ByVal node As Object) As String
    On Error Resume Next
    mp_LayoutNodeTag = LCase$(Trim$(CStr(node.baseName)))
    On Error GoTo 0
End Function

Private Sub mp_SetNodeAttrDouble(ByVal node As Object, ByVal attrName As String, ByVal valueNumber As Double)
    On Error Resume Next
    node.setAttribute attrName, mp_ToInvariantDoubleText(valueNumber)
    On Error GoTo 0
End Sub

Private Sub mp_SetNodeAttrLong(ByVal node As Object, ByVal attrName As String, ByVal valueNumber As Long)
    On Error Resume Next
    node.setAttribute attrName, CStr(valueNumber)
    On Error GoTo 0
End Sub

Private Function mp_ToInvariantDoubleText(ByVal valueNumber As Double) As String
    mp_ToInvariantDoubleText = Replace$(Trim$(CStr(valueNumber)), ",", ".")
End Function

Private Function mp_IsSupportedControlType(ByVal controlType As String) As Boolean
    controlType = mp_NormalizeControlType(controlType)
    Select Case controlType
        Case "button", "dropdownbutton", "input"
            mp_IsSupportedControlType = True
    End Select
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set mp_GetWorksheetByName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function mp_EnsureControlShape(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByVal controlType As String, ByVal createIfMissing As Boolean, ByVal isRequired As Boolean) As Boolean
    Dim shp As Shape

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
    If Not shp Is Nothing Then
        mp_EnsureControlShape = True
        Exit Function
    End If

    If Not createIfMissing Then
        If isRequired Then
            MsgBox "Control '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Else
            mp_EnsureControlShape = True
        End If
        Exit Function
    End If

    controlType = mp_NormalizeControlType(controlType)
    Select Case controlType
        Case "button", "dropdownbutton"
            If Not mp_TryCreateMissingButton(ws, controlNode, controlName, shp) Then Exit Function
        Case Else
            If isRequired Then
                MsgBox "Auto-create is not supported for type='" & controlType & "'. Control '" & controlName & "' cannot be created.", vbExclamation
            Else
                mp_EnsureControlShape = True
            End If
            Exit Function
    End Select

    mp_EnsureControlShape = True
End Function

Private Function mp_TryCreateMissingButton(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByRef createdShape As Shape) As Boolean
    Dim captionText As String
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim templateName As String
    Dim templateShape As Shape

    captionText = Trim$(mp_NodeAttrText(controlNode, "caption"))
    If Len(captionText) = 0 Then
        captionText = controlName
    End If

    If Not mp_ReadRequiredControlRect(ws, controlNode, controlName, leftPos, topPos, widthVal, heightVal) Then Exit Function

    On Error GoTo EH_CREATE
    Set createdShape = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, widthVal, heightVal)

    templateName = Trim$(mp_NodeAttrText(controlNode, "template"))
    If Len(templateName) > 0 Then
        Set templateShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, templateName)
    End If

    If Not templateShape Is Nothing Then
        On Error Resume Next
        templateShape.PickUp
        createdShape.Apply
        On Error GoTo 0
    End If

    createdShape.Left = leftPos
    createdShape.Top = topPos
    createdShape.Width = widthVal
    createdShape.Height = heightVal
    createdShape.Name = controlName
    createdShape.TextFrame.Characters.Text = captionText
    createdShape.Placement = xlFreeFloating
    mp_TryCreateMissingButton = True
    Exit Function

EH_CREATE:
    MsgBox "Failed to auto-create button control '" & controlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_ReadRequiredControlRect(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByRef leftPos As Double, ByRef topPos As Double, ByRef widthVal As Double, ByRef heightVal As Double) As Boolean
    Dim hasLeft As Boolean
    Dim hasTop As Boolean
    Dim hasWidth As Boolean
    Dim hasHeight As Boolean
    Dim relativeToName As String
    Dim baseShape As Shape
    Dim marginLeft As Double
    Dim marginTop As Double

    hasLeft = mp_ReadRequiredNumber(controlNode, "left", leftPos)
    hasTop = mp_ReadRequiredNumber(controlNode, "top", topPos)
    hasWidth = mp_ReadRequiredNumber(controlNode, "width", widthVal)
    hasHeight = mp_ReadRequiredNumber(controlNode, "height", heightVal)

    If hasLeft And hasTop And hasWidth And hasHeight Then
        mp_ReadRequiredControlRect = True
        Exit Function
    End If

    If mp_TryReadGridResolvedRect(ws, controlNode, controlName, leftPos, topPos, widthVal, heightVal) Then
        mp_ReadRequiredControlRect = True
        Exit Function
    End If

    If Not hasLeft Or Not hasTop Then
        relativeToName = Trim$(mp_NodeAttrText(controlNode, "relativeTo"))
        If Len(relativeToName) = 0 Then
            MsgBox "Control '" & controlName & "' with createIfMissing='true' must define either numeric left/top/width/height, or relative layout ('relativeTo' + 'marginLeft' + 'marginTop' + 'height' + optional 'width'/'matchWidthToRelative') in DevUI.xml.", vbExclamation
            Exit Function
        End If

        Set baseShape = ex_ConfigProfilesManager.m_GetShapeByName(ws, relativeToName)
        If baseShape Is Nothing Then
            MsgBox "Control '" & controlName & "' references relativeTo='" & relativeToName & "' but anchor shape was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Function
        End If

        If Not mp_ReadRequiredNumber(controlNode, "marginLeft", marginLeft) Then
            MsgBox "Control '" & controlName & "' with relative layout must define numeric 'marginLeft' in DevUI.xml.", vbExclamation
            Exit Function
        End If
        If Not mp_ReadRequiredNumber(controlNode, "marginTop", marginTop) Then
            MsgBox "Control '" & controlName & "' with relative layout must define numeric 'marginTop' in DevUI.xml.", vbExclamation
            Exit Function
        End If

        leftPos = baseShape.Left + marginLeft
        topPos = baseShape.Top + baseShape.Height + marginTop
    End If

    If Not hasWidth Then
        If Not baseShape Is Nothing Then
            If mp_NodeAttrBool(controlNode, "matchWidthToRelative", False) Then
                widthVal = baseShape.Width
                hasWidth = True
            End If
        End If
    End If

    If Not hasWidth Then
        MsgBox "Control '" & controlName & "' with createIfMissing='true' must define numeric 'width' or set matchWidthToRelative='true' with valid relativeTo in DevUI.xml.", vbExclamation
        Exit Function
    End If
    If Not hasHeight Then
        MsgBox "Control '" & controlName & "' with createIfMissing='true' must define numeric 'height' in DevUI.xml.", vbExclamation
        Exit Function
    End If

    mp_ReadRequiredControlRect = True
End Function

Private Function mp_TryReadGridResolvedRect( _
    ByVal ws As Worksheet, _
    ByVal controlNode As Object, _
    ByVal controlName As String, _
    ByRef leftPos As Double, _
    ByRef topPos As Double, _
    ByRef widthVal As Double, _
    ByRef heightVal As Double _
) As Boolean
    Dim rowText As String
    Dim colText As String
    Dim widthText As String
    Dim heightText As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim widthTracks As Long
    Dim heightTracks As Long
    Dim targetRange As Range

    If ws Is Nothing Then Exit Function
    If controlNode Is Nothing Then Exit Function

    rowText = Trim$(mp_NodeAttrText(controlNode, "gridRow"))
    colText = Trim$(mp_NodeAttrText(controlNode, "gridCol"))
    widthText = Trim$(mp_NodeAttrText(controlNode, "gridWidth"))
    heightText = Trim$(mp_NodeAttrText(controlNode, "gridHeight"))
    If Len(rowText) = 0 Or Len(colText) = 0 Or Len(widthText) = 0 Or Len(heightText) = 0 Then Exit Function

    If Not ex_XmlCore.m_TryParseLong(rowText, rowIndex) Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(colText, colIndex) Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(widthText, widthTracks) Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(heightText, heightTracks) Then Exit Function
    If rowIndex < 1 Or colIndex < 1 Or widthTracks <= 0 Or heightTracks <= 0 Then Exit Function

    If Not mp_LayoutTryBuildRangeByTracks(ws, rowIndex, colIndex, heightTracks, widthTracks, targetRange) Then
        MsgBox "Control '" & controlName & "' has invalid resolved grid bounds (gridRow/gridCol/gridWidth/gridHeight).", vbExclamation
        Exit Function
    End If

    leftPos = targetRange.Left
    topPos = targetRange.Top
    widthVal = targetRange.Width
    heightVal = targetRange.Height
    mp_TryReadGridResolvedRect = True
End Function

Private Function mp_ReadRequiredNumber(ByVal node As Object, ByVal attrName As String, ByRef result As Double) As Boolean
    Dim valueText As String

    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then Exit Function
    mp_ReadRequiredNumber = mp_TryParseDouble(valueText, result)
End Function

Private Function mp_ApplyControlAttributes(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByVal controlType As String, ByVal stylesMap As Object) As Boolean
    Dim shp As Shape
    Dim styleName As String

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        MsgBox "Control '" & controlName & "' was not found while applying attributes on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    If Not mp_ApplyShapeVisible(controlNode, shp) Then Exit Function
    If Not mp_ApplyShapePlacement(controlNode, shp, ws) Then Exit Function
    If Not mp_ApplyShapeGeometry(controlNode, shp, ws) Then Exit Function

    controlType = mp_NormalizeControlType(controlType)
    If StrComp(controlType, "button", vbTextCompare) = 0 Or StrComp(controlType, "dropdownbutton", vbTextCompare) = 0 Then
        On Error Resume Next
        shp.AutoShapeType = msoShapeRectangle
        On Error GoTo 0

        If Not mp_ApplyButtonCaption(controlNode, shp) Then Exit Function

        styleName = Trim$(mp_NodeAttrText(controlNode, "style"))
        If Len(styleName) > 0 Then
            If Not ex_UiXmlProvider.m_ApplyButtonStyleByName(shp, styleName, stylesMap) Then Exit Function
        End If
    End If

    mp_ApplyControlAttributes = True
End Function

Private Function mp_ApplyInputCellControl( _
    ByVal ws As Worksheet, _
    ByVal controlNode As Object, _
    ByVal controlName As String, _
    ByVal stylesMap As Object) As Boolean

    Dim rowText As String
    Dim colText As String
    Dim widthText As String
    Dim heightText As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim widthTracks As Long
    Dim heightTracks As Long
    Dim targetRange As Range
    Dim inputCell As Range
    Dim inputValue As String
    Dim inputName As String
    Dim inputConfigKey As String
    Dim inputBind As String
    Dim resolvedConfigKey As String
    Dim onChangeMacro As String
    Dim isPrimaryInput As Boolean
    Dim currentValue As String
    Dim styleName As String

    rowText = Trim$(mp_NodeAttrText(controlNode, "gridRow"))
    colText = Trim$(mp_NodeAttrText(controlNode, "gridCol"))
    widthText = Trim$(mp_NodeAttrText(controlNode, "gridWidth"))
    heightText = Trim$(mp_NodeAttrText(controlNode, "gridHeight"))

    If Len(rowText) = 0 Or Len(colText) = 0 Or Len(widthText) = 0 Or Len(heightText) = 0 Then
        MsgBox "Input control '" & controlName & "' must be defined inside grid layout (missing gridRow/gridCol/gridWidth/gridHeight).", vbExclamation
        Exit Function
    End If
    If Not ex_XmlCore.m_TryParseLong(rowText, rowIndex) Then
        MsgBox "Input control '" & controlName & "' has invalid gridRow='" & rowText & "'.", vbExclamation
        Exit Function
    End If
    If Not ex_XmlCore.m_TryParseLong(colText, colIndex) Then
        MsgBox "Input control '" & controlName & "' has invalid gridCol='" & colText & "'.", vbExclamation
        Exit Function
    End If
    If Not ex_XmlCore.m_TryParseLong(widthText, widthTracks) Then
        MsgBox "Input control '" & controlName & "' has invalid gridWidth='" & widthText & "'.", vbExclamation
        Exit Function
    End If
    If Not ex_XmlCore.m_TryParseLong(heightText, heightTracks) Then
        MsgBox "Input control '" & controlName & "' has invalid gridHeight='" & heightText & "'.", vbExclamation
        Exit Function
    End If

    If Not mp_LayoutTryBuildRangeByTracks(ws, rowIndex, colIndex, heightTracks, widthTracks, targetRange) Then
        MsgBox "Input control '" & controlName & "' resolved to invalid grid bounds.", vbExclamation
        Exit Function
    End If

    targetRange.UnMerge
    If targetRange.Cells.CountLarge > 1 Then targetRange.Merge
    Set inputCell = targetRange.Cells(1, 1)

    inputCell.NumberFormat = "@"
    inputCell.HorizontalAlignment = xlLeft
    inputCell.VerticalAlignment = xlCenter

    styleName = Trim$(mp_NodeAttrText(controlNode, "style"))
    If Len(styleName) > 0 Then
        If Not mp_ApplyCellStyleByName(targetRange, styleName, stylesMap, controlName) Then Exit Function
    End If

    inputValue = Trim$(mp_NodeAttrText(controlNode, "value"))
    If Len(inputValue) = 0 Then inputValue = Trim$(mp_NodeAttrText(controlNode, "text"))
    inputConfigKey = Trim$(mp_NodeAttrText(controlNode, "inputConfigKey"))
    inputBind = Trim$(mp_NodeAttrText(controlNode, "bind"))
    If Len(inputBind) = 0 Then inputBind = inputConfigKey

    currentValue = Trim$(CStr(inputCell.Value))
    If Len(currentValue) = 0 Then
        If Len(inputValue) > 0 Then
            inputCell.Value = inputValue
        ElseIf ex_LayoutBindingsRuntime.m_TryResolveConfigKeyFromBindSpec(inputBind, resolvedConfigKey) Then
            inputCell.Value = ex_ConfigProvider.m_GetConfigValue(resolvedConfigKey, vbNullString)
        End If
    End If

    inputName = Trim$(mp_NodeAttrText(controlNode, "inputName"))
    If Len(inputName) = 0 Then inputName = controlName

    onChangeMacro = Trim$(mp_NodeAttrText(controlNode, "onChange"))
    If Len(onChangeMacro) = 0 Then onChangeMacro = Trim$(mp_NodeAttrText(controlNode, "onChangeMacro"))
    isPrimaryInput = mp_NodeAttrBool(controlNode, "inputPrimary", False)

    ex_LayoutBindingsRuntime.m_RegisterInputBinding ws, inputCell, inputName, inputBind, onChangeMacro, isPrimaryInput

    mp_ApplyInputCellControl = True
End Function

Private Function mp_ApplyCellStyleByName( _
    ByVal targetRange As Range, _
    ByVal styleName As String, _
    ByVal stylesMap As Object, _
    ByVal controlName As String) As Boolean

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
        mp_DebugLog "mp_ApplyCellStyleByName: styles map is unavailable for control='" & controlName & "', style='" & styleName & "'."
        MsgBox "Input control '" & controlName & "' references style '" & styleName & "', but styles map is unavailable.", vbExclamation
        Exit Function
    End If
    If Not stylesMap.Exists(styleName) Then
        mp_DebugLog "mp_ApplyCellStyleByName: style '" & styleName & "' is missing for control='" & controlName & "'."
        MsgBox "Input control '" & controlName & "' references missing style '" & styleName & "'.", vbExclamation
        Exit Function
    End If

    Set styleData = stylesMap(styleName)
    On Error GoTo EH

    If styleData.Exists("backColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("backColor")), colorValue) Then
            MsgBox "Invalid style backColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        targetRange.Interior.Pattern = xlSolid
        targetRange.Interior.Color = colorValue
    End If

    If styleData.Exists("textColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("textColor")), colorValue) Then
            MsgBox "Invalid style textColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        targetRange.Font.Color = colorValue
    End If

    If styleData.Exists("fontName") Then
        targetRange.Font.Name = CStr(styleData("fontName"))
    End If

    If styleData.Exists("fontSize") Then
        If Not ex_XmlCore.m_TryParseDouble(CStr(styleData("fontSize")), numberValue) Then
            MsgBox "Invalid style fontSize for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        targetRange.Font.Size = numberValue
    End If

    If styleData.Exists("fontBold") Then
        If Not ex_XmlCore.m_TryParseBoolean(CStr(styleData("fontBold")), boolValue) Then
            MsgBox "Invalid style fontBold for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        targetRange.Font.Bold = boolValue
    End If

    If styleData.Exists("borderColor") Then
        If Not ex_XmlCore.m_TryParseColor(CStr(styleData("borderColor")), colorValue) Then
            MsgBox "Invalid style borderColor for style '" & styleName & "'.", vbExclamation
            Exit Function
        End If
        targetRange.Borders.LineStyle = xlContinuous
        targetRange.Borders.Color = colorValue
    End If

    mp_ApplyCellStyleByName = True
    mp_DebugLog "mp_ApplyCellStyleByName: applied style='" & styleName & "' to control='" & controlName & "'."
    Exit Function
EH:
    mp_DebugLog "mp_ApplyCellStyleByName: failed style='" & styleName & "' control='" & controlName & "' error='" & Err.Description & "'."
    MsgBox "Failed to apply style '" & styleName & "' to input control '" & controlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_RebuildManagedDropdownButton( _
    ByVal ws As Worksheet, _
    ByVal controlNode As Object, _
    ByVal controlName As String, _
    ByVal stylesMap As Object) As Boolean

    Dim shp As Shape
    Dim itemRecords As Variant
    Dim resolvedModeKey As String
    Dim itemStyleName As String
    Dim selectionChangedMacro As String
    Dim selectedItem As String
    Dim itemMarginLeft As Double
    Dim itemFirstGap As Double
    Dim itemGap As Double
    Dim itemHeight As Double
    Dim itemMatchWidth As Boolean
    Dim headerShowsSelection As Boolean
    Dim errText As String

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        MsgBox "DropdownButton '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    resolvedModeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws))
    itemRecords = ex_UiXmlProvider.m_GetDropdownItemRecordsFromControlNode(controlNode, ThisWorkbook, resolvedModeKey)
    If Not ex_UiXmlProvider.m_HasDropdownItemRecords(itemRecords) Then
        MsgBox "DropdownButton '" & controlName & "' did not resolve any items (itemsSource or inline <items>).", vbExclamation
        Exit Function
    End If

    itemStyleName = Trim$(mp_NodeAttrText(controlNode, "itemStyle"))
    selectionChangedMacro = Trim$(mp_NodeAttrText(controlNode, "selectionChangedMacro"))
    selectedItem = Trim$(mp_NodeAttrText(controlNode, "selectedItem"))

    If Not mp_TryReadOptionalNodeDouble(controlNode, "itemMarginLeft", 0, itemMarginLeft, controlName) Then Exit Function
    If Not mp_TryReadOptionalNodeDouble(controlNode, "itemFirstGap", 2, itemFirstGap, controlName) Then Exit Function
    If Not mp_TryReadOptionalNodeDouble(controlNode, "itemGap", 2, itemGap, controlName) Then Exit Function
    If Not mp_TryReadOptionalNodeDouble(controlNode, "itemHeight", 16, itemHeight, controlName) Then Exit Function
    If itemHeight <= 0 Then
        MsgBox "Control '" & controlName & "' has invalid attribute 'itemHeight'. Value must be > 0.", vbExclamation
        Exit Function
    End If
    If Not mp_TryReadOptionalNodeBoolean(controlNode, "itemMatchWidth", True, itemMatchWidth, controlName) Then Exit Function
    If Not mp_TryReadOptionalNodeBoolean(controlNode, "headerShowsSelection", True, headerShowsSelection, controlName) Then Exit Function

    If Not ex_ManagedDropdownRuntime.m_RebuildDropdownButton( _
        ws, _
        shp, _
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
        errText) Then
        If Len(errText) = 0 Then errText = "Unknown dropdown rebuild error."
        MsgBox "Failed to rebuild dropdown options for control '" & controlName & "': " & errText, vbExclamation
        Exit Function
    End If

    mp_RebuildManagedDropdownButton = True
End Function

Private Function mp_TryReadOptionalNodeDouble( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Double, _
    ByRef outValue As Double, _
    ByVal controlName As String) As Boolean

    Dim valueText As String
    Dim parsedValue As Double

    outValue = defaultValue
    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        mp_TryReadOptionalNodeDouble = True
        Exit Function
    End If

    If Not mp_TryParseDouble(valueText, parsedValue) Then
        MsgBox "Invalid numeric value '" & valueText & "' for attribute '" & attrName & "' on control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = parsedValue
    mp_TryReadOptionalNodeDouble = True
End Function

Private Function mp_TryReadOptionalNodeBoolean( _
    ByVal node As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Boolean, _
    ByRef outValue As Boolean, _
    ByVal controlName As String) As Boolean

    Dim valueText As String
    Dim parsedValue As Boolean

    outValue = defaultValue
    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        mp_TryReadOptionalNodeBoolean = True
        Exit Function
    End If

    If Not ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        MsgBox "Invalid boolean value '" & valueText & "' for attribute '" & attrName & "' on control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = parsedValue
    mp_TryReadOptionalNodeBoolean = True
End Function

Private Function mp_ApplyButtonCaption(ByVal node As Object, ByVal shp As Shape) As Boolean
    Dim captionText As String

    captionText = mp_NodeAttrText(node, "caption")
    If Len(captionText) = 0 Then
        mp_ApplyButtonCaption = True
        Exit Function
    End If

    On Error GoTo EH
    shp.TextFrame.Characters.Text = captionText
    mp_ApplyButtonCaption = True
    Exit Function
EH:
    MsgBox "Failed to apply caption for control '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Public Function m_GetDropdownItemsByName(ByVal controlName As String, Optional ByVal wb As Workbook, Optional ByVal modeKey As String = vbNullString) As Variant
    m_GetDropdownItemsByName = ex_UiXmlProvider.m_GetDropdownItemsByName(controlName, wb, modeKey)
End Function

Public Function m_GetProfilesFilePathByMode(Optional ByVal modeKey As String = vbNullString, Optional ByVal wb As Workbook, Optional ByVal sourceName As String = "profilesFileByMode") As String
    m_GetProfilesFilePathByMode = ex_UiXmlProvider.m_GetProfilesFilePathByMode(modeKey, wb, sourceName)
End Function

Public Function m_GetModeVariantsByControl(ByVal controlName As String, Optional ByVal wb As Workbook) As Variant
    m_GetModeVariantsByControl = ex_UiXmlProvider.m_GetModeVariantsByControl(controlName, wb)
End Function

Public Function m_GetControlAttribute(ByVal controlName As String, ByVal attrName As String, Optional ByVal wb As Workbook) As String
    m_GetControlAttribute = ex_UiXmlProvider.m_GetControlAttribute(controlName, attrName, wb)
End Function

Public Function m_ApplyControlStyleByName(ByVal ws As Worksheet, ByVal controlName As String, ByVal styleName As String, Optional ByVal wb As Workbook) As Boolean
    m_ApplyControlStyleByName = ex_UiXmlProvider.m_ApplyControlStyleByName(ws, controlName, styleName, wb)
End Function

Private Function mp_AssignShapeMacro(ByVal ws As Worksheet, ByVal controlNode As Object, ByVal controlName As String, ByVal wb As Workbook, ByVal isRequired As Boolean, ByRef didUngroupUiBlock As Boolean) As Boolean
    Dim shp As Shape
    Dim macroName As String
    Dim onActionText As String
    Dim assignmentErr As String
    Dim parentGroupName As String

    macroName = Trim$(mp_NodeAttrText(controlNode, "macro"))
    didUngroupUiBlock = False

    If Len(macroName) = 0 Then
        If isRequired Then
            MsgBox "UI config control '" & controlName & "' does not define required attribute 'macro'.", vbExclamation
        Else
            mp_AssignShapeMacro = True
        End If
        Exit Function
    End If

    onActionText = "'" & wb.Name & "'!" & macroName
    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
    If shp Is Nothing Then
        If isRequired Then
            MsgBox "Control '" & controlName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Else
            mp_AssignShapeMacro = True
        End If
        Exit Function
    End If

    On Error GoTo EH_ASSIGN_FIRST
    shp.OnAction = onActionText
    mp_AssignShapeMacro = True
    Exit Function

EH_ASSIGN_FIRST:
    assignmentErr = Err.Description
    Err.Clear

    If mp_TryUngroupParentShape(shp, parentGroupName) Then
        If StrComp(parentGroupName, UI_BLOCK_GROUP_NAME, vbTextCompare) = 0 Then
            didUngroupUiBlock = True
        End If

        Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, controlName)
        If shp Is Nothing Then
            MsgBox "Control '" & controlName & "' is unavailable after ungroup on sheet '" & ws.Name & "'.", vbExclamation
            Exit Function
        End If

        On Error GoTo EH_ASSIGN_SECOND
        shp.OnAction = onActionText
        mp_AssignShapeMacro = True
        Exit Function
    End If

    If isRequired Then
        MsgBox "Failed to assign macro '" & macroName & "' to control '" & controlName & "': " & assignmentErr, vbExclamation
    Else
        mp_AssignShapeMacro = True
    End If
    Exit Function

EH_ASSIGN_SECOND:
    If isRequired Then
        MsgBox "Failed to assign macro '" & macroName & "' to control '" & controlName & "' after ungroup retry: " & Err.Description, vbExclamation
    Else
        mp_AssignShapeMacro = True
    End If
End Function

Private Function mp_TryUngroupParentShape(ByVal shp As Shape, ByRef parentGroupName As String) As Boolean
    Dim parentGroup As Shape

    parentGroupName = vbNullString
    On Error Resume Next
    Set parentGroup = shp.ParentGroup
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If parentGroup Is Nothing Then Exit Function

    parentGroupName = parentGroup.Name
    On Error GoTo EH_UNGROUP
    parentGroup.Ungroup
    mp_TryUngroupParentShape = True
    Exit Function

EH_UNGROUP:
    MsgBox "Failed to ungroup parent group for control '" & shp.Name & "': " & Err.Description, vbExclamation
End Function

Private Sub mp_TryRegroupUiBlock(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error GoTo EH_REGROUP
    ex_ConfigProfilesManager.m_InitUiBlockLayoutAndGroup ws
    Exit Sub
EH_REGROUP:
    MsgBox "Failed to regroup UI block '" & UI_BLOCK_GROUP_NAME & "' on sheet '" & ws.Name & "': " & Err.Description, vbExclamation
End Sub


Private Function mp_ApplyShapeVisible(ByVal node As Object, ByVal shp As Shape) As Boolean
    Dim valueText As String
    Dim valueBool As Boolean

    valueText = Trim$(mp_NodeAttrText(node, "visible"))
    If Len(valueText) = 0 Then
        mp_ApplyShapeVisible = True
        Exit Function
    End If

    If Not mp_TryParseBoolean(valueText, valueBool) Then
        MsgBox "Invalid boolean value for UI attribute 'visible' on shape '" & shp.Name & "': " & valueText, vbExclamation
        Exit Function
    End If

    shp.Visible = IIf(valueBool, msoTrue, msoFalse)
    mp_ApplyShapeVisible = True
End Function

Private Function mp_ApplyShapePlacement(ByVal node As Object, ByVal shp As Shape, ByVal ws As Worksheet) As Boolean
    Dim placementText As String
    Dim placementValue As XlPlacement

    placementText = Trim$(mp_NodeAttrText(node, "placement"))
    If Len(placementText) > 0 Then
        If Not mp_TryParsePlacement(placementText, placementValue) Then
            MsgBox "Invalid UI placement value on shape '" & shp.Name & "': " & placementText, vbExclamation
            Exit Function
        End If
        shp.Placement = placementValue
    End If

    mp_ApplyShapePlacement = True
End Function

Private Function mp_ApplyShapeGeometry(ByVal node As Object, ByVal shp As Shape, ByVal ws As Worksheet) As Boolean
    If Not mp_ApplySingleGeometryAttribute(node, shp, "left") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "top") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "width") Then Exit Function
    If Not mp_ApplySingleGeometryAttribute(node, shp, "height") Then Exit Function

    mp_ApplyShapeGeometry = True
End Function

Private Function mp_ApplySingleGeometryAttribute(ByVal node As Object, ByVal shp As Shape, ByVal attrName As String) As Boolean
    Dim valueText As String
    Dim valueNumber As Double

    valueText = Trim$(mp_NodeAttrText(node, attrName))
    If Len(valueText) = 0 Then
        mp_ApplySingleGeometryAttribute = True
        Exit Function
    End If

    If Not mp_TryParseDouble(valueText, valueNumber) Then
        MsgBox "Invalid numeric value for UI attribute '" & attrName & "' on shape '" & shp.Name & "': " & valueText, vbExclamation
        Exit Function
    End If

    Select Case LCase$(attrName)
        Case "left": shp.Left = valueNumber
        Case "top": shp.Top = valueNumber
        Case "width": shp.Width = valueNumber
        Case "height": shp.Height = valueNumber
    End Select

    mp_ApplySingleGeometryAttribute = True
End Function

Private Function mp_RemoveButtonsMissingInConfig(ByVal wb As Workbook, ByVal controlNodes As Object) As Boolean
    Dim allowedButtons As Object
    Dim ws As Worksheet
    Dim existingButtons As Object
    Dim controlNode As Object
    Dim controlName As String
    Dim controlType As String
    Dim sheetName As String
    Dim buttonName As Variant
    Dim isNodeEnabled As Boolean

    Set allowedButtons = CreateObject("Scripting.Dictionary")

    For Each controlNode In controlNodes
        controlName = Trim$(mp_NodeAttrText(controlNode, "name"))
        If Not ex_XmlCore.m_TryEvaluateNodeCondition(controlNode, isNodeEnabled, "condition", "control '" & controlName & "'") Then Exit Function
        If Not isNodeEnabled Then GoTo NextNode

        controlType = mp_NormalizeControlType(mp_NodeAttrText(controlNode, "type"))
        If Len(controlType) = 0 Then
            controlType = "button"
        End If
        If StrComp(controlType, "button", vbTextCompare) <> 0 And StrComp(controlType, "dropdownbutton", vbTextCompare) <> 0 Then GoTo NextNode
        If Not mp_IsButtonShapeName(controlName) Then GoTo NextNode

        sheetName = Trim$(mp_NodeAttrText(controlNode, "sheet"))
        If Len(sheetName) = 0 Then
            sheetName = DEFAULT_SHEET_NAME
        End If
        allowedButtons(mp_SheetButtonKey(sheetName, controlName)) = True
NextNode:
    Next controlNode

    For Each ws In wb.Worksheets
        Set existingButtons = CreateObject("Scripting.Dictionary")
        mp_CollectButtonShapeNamesInContainer ws.Shapes, existingButtons

        For Each buttonName In existingButtons.Keys
            If Not allowedButtons.Exists(mp_SheetButtonKey(ws.Name, CStr(buttonName))) Then
                If Not mp_DeleteShapeByName(ws, CStr(buttonName)) Then Exit Function
            End If
NextButtonName:
        Next buttonName
    Next ws

    mp_RemoveButtonsMissingInConfig = True
End Function

Private Function mp_NormalizeControlType(ByVal rawType As String) As String
    rawType = LCase$(Trim$(rawType))
    mp_NormalizeControlType = rawType
End Function

Private Sub mp_CollectButtonShapeNamesInContainer(ByVal shapeContainer As Object, ByVal names As Object)
    Dim shp As Shape
    Dim groupItem As Shape

    For Each shp In shapeContainer
        If mp_IsButtonShapeName(shp.Name) Then
            names(shp.Name) = True
        End If

        If shp.Type = msoGroup Then
            For Each groupItem In shp.GroupItems
                If mp_IsButtonShapeName(groupItem.Name) Then
                    names(groupItem.Name) = True
                End If
                If groupItem.Type = msoGroup Then
                    mp_CollectButtonShapeNamesInContainer groupItem.GroupItems, names
                End If
            Next groupItem
        End If
    Next shp
End Sub

Private Function mp_DeleteShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shp As Shape
    Dim parentGroupName As String

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, shapeName)
    If shp Is Nothing Then
        mp_DeleteShapeByName = True
        Exit Function
    End If

    On Error GoTo EH_DELETE
    shp.Delete
    mp_DeleteShapeByName = True
    Exit Function
EH_DELETE:
    Err.Clear

    If mp_TryUngroupParentShape(shp, parentGroupName) Then
        Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then
            mp_DeleteShapeByName = True
            Exit Function
        End If

        On Error GoTo EH_DELETE_RETRY
        shp.Delete
        mp_DeleteShapeByName = True
        Exit Function
    End If

    MsgBox "Failed to delete UI shape '" & shapeName & "' from sheet '" & ws.Name & "'.", vbExclamation
    Exit Function
EH_DELETE_RETRY:
    MsgBox "Failed to delete UI shape '" & shapeName & "' from sheet '" & ws.Name & "' after ungroup retry: " & Err.Description, vbExclamation
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

Private Function mp_NodeAttrBool(ByVal node As Object, ByVal attrName As String, ByVal defaultValue As Boolean) As Boolean
    Dim valueText As String

    valueText = LCase$(Trim$(mp_NodeAttrText(node, attrName)))
    If Len(valueText) = 0 Then
        mp_NodeAttrBool = defaultValue
        Exit Function
    End If

    Select Case valueText
        Case "true", "1", "yes"
            mp_NodeAttrBool = True
        Case "false", "0", "no"
            mp_NodeAttrBool = False
        Case Else
            mp_NodeAttrBool = defaultValue
    End Select
End Function

Private Function mp_IsButtonShapeName(ByVal shapeName As String) As Boolean
    mp_IsButtonShapeName = (LCase$(Left$(Trim$(shapeName), 3)) = "btn")
End Function

Private Function mp_SheetButtonKey(ByVal sheetName As String, ByVal controlName As String) As String
    mp_SheetButtonKey = LCase$(Trim$(sheetName)) & "|" & LCase$(Trim$(controlName))
End Function

Private Function mp_TryParseBoolean(ByVal valueText As String, ByRef result As Boolean) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "1", "true", "yes"
            result = True
            mp_TryParseBoolean = True
        Case "0", "false", "no"
            result = False
            mp_TryParseBoolean = True
    End Select
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

Private Function mp_TryParseDouble(ByVal valueText As String, ByRef result As Double) As Boolean
    Dim normalized As String
    Dim decSep As String
    Dim altSep As String

    On Error GoTo EH

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    decSep = CStr(Application.International(xlDecimalSeparator))
    If decSep = "." Then
        altSep = ","
    Else
        altSep = "."
    End If

    normalized = Replace(normalized, altSep, decSep)
    If Not IsNumeric(normalized) Then Exit Function

    result = CDbl(normalized)
    mp_TryParseDouble = True
    Exit Function
EH:
    mp_TryParseDouble = False
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

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_UILoader] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
