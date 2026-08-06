Attribute VB_Name = "ex_ControlRefreshRuntime"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private Const UI_NS As String = "urn:excelprototype:profiles"
' Два retained-реестра являются логической картой последнего render и не
' содержат копию данных Excel. g_ControlRegistry отвечает на простой вопрос
' "где сейчас расположен именованный control". g_LayoutRegistry хранит дерево
' layout-узлов: structural key, parent, orientation, fixed-size flag и bounds.
'
' Разделение принципиально для partial render: bounds самого контрола нужны для
' его очистки/повторной отрисовки, а layout-дерево — для распространения delta
' только по зависимой ветке. Полный сброс этих коллекций превратил бы partial
' render обратно в полный layout pass.
Private g_ControlRegistry As Object
Private g_LayoutRegistry As Object

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:ex_ControlRefreshRuntime.fn_Module_Dispose"
#End If
    On Error Resume Next
    Set g_ControlRegistry = Nothing
    Set g_LayoutRegistry = Nothing
    On Error GoTo 0
End Sub
' //
' // API
' //
Public Sub fn_ResetRegisteredControls()
    Set g_ControlRegistry = private_CreateDictionary()
    Set g_LayoutRegistry = private_CreateDictionary()
End Sub


Public Sub fn_ResetRegisteredControlsByWorksheet(ByVal sheetName As String)
    Dim key As Variant
    Dim keysToRemove As Collection
    Dim entry As Object

    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Then Exit Sub
    private_EnsureRegistry
    Set keysToRemove = New Collection
    For Each key In g_ControlRegistry.Keys
        Set entry = g_ControlRegistry(key)
        If Not entry Is Nothing Then
            If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), sheetName, VBA.vbTextCompare) = 0 Then keysToRemove.Add key
        End If
    Next key
    For Each key In keysToRemove
        g_ControlRegistry.Remove key
    Next key
    private_RemoveLayoutEntriesByWorksheet sheetName
End Sub


Public Sub fn_RegisterLayoutNodeRenderBounds( _
    ByVal layoutNode As Object, _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
)
    Dim layoutKey As String
    Dim parentKey As String
    Dim nodeKind As String
    Dim nodeName As String
    Dim orientation As String
    Dim entry As Object

    If layoutNode Is Nothing Then Exit Sub
    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Or rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    ' Ключ строится из положения узла в XML, а не из текущих координат Excel.
    ' Поэтому он остаётся стабильным после сдвига строк и позволяет заменить
    ' descriptor тем же ключом при следующем локальном render.
    layoutKey = private_BuildLayoutNodeKey(layoutNode)
    If VBA.Len(layoutKey) = 0 Then Exit Sub
    parentKey = private_BuildLayoutNodeKey(layoutNode.parentNode)
    nodeKind = VBA.LCase$(VBA.Trim$(VBA.CStr(layoutNode.baseName)))
    nodeName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(layoutNode, "name")))
    orientation = VBA.LCase$(VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(layoutNode, "orientation"))))
    If nodeKind = "stackpanel" And VBA.Len(orientation) = 0 Then orientation = "vertical"

    private_EnsureLayoutRegistry
    Set entry = private_CreateDictionary()
    entry("Key") = layoutKey
    entry("ParentKey") = parentKey
    entry("Kind") = nodeKind
    entry("Name") = nodeName
    entry("Orientation") = orientation
    entry("FixedRows") = private_HasPositiveLongAttr(layoutNode, "spanRows")
    entry("Sheet") = sheetName
    entry("RowStart") = VBA.CLng(rowStart)
    entry("ColStart") = VBA.CLng(colStart)
    entry("RowEnd") = VBA.CLng(rowEnd)
    entry("ColEnd") = VBA.CLng(colEnd)
    Set g_LayoutRegistry(private_BuildLayoutRegistryKey(sheetName, layoutKey)) = entry
End Sub


Public Sub fn_RegisterControlRenderBounds( _
    ByVal controlName As String, _
    ByVal controlType As String, _
    ByVal sheetName As String, _
    ByVal uiPath As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
)
    Dim key As String
    Dim entry As Object

    controlName = VBA.Trim$(controlName)
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    sheetName = VBA.Trim$(sheetName)
    uiPath = VBA.Trim$(uiPath)

    If VBA.Len(controlName) = 0 Then Exit Sub
    If VBA.Len(controlType) = 0 Then Exit Sub
    If VBA.Len(sheetName) = 0 Then Exit Sub
    If VBA.Len(uiPath) = 0 Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Then Exit Sub
    If rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    private_EnsureRegistry
    key = private_BuildRegistryKey(sheetName, controlName)

    Set entry = private_CreateDictionary()
    entry("Name") = controlName
    entry("Type") = controlType
    entry("Sheet") = sheetName
    entry("UiPath") = uiPath
    entry("RowStart") = VBA.CLng(rowStart)
    entry("ColStart") = VBA.CLng(colStart)
    entry("RowEnd") = VBA.CLng(rowEnd)
    entry("ColEnd") = VBA.CLng(colEnd)

    If g_ControlRegistry.Exists(key) Then g_ControlRegistry.Remove key
    g_ControlRegistry.Add key, entry
End Sub


Public Function fn_TryGetControlRenderBounds( _
    ByVal controlName As String, _
    ByVal sheetName As String, _
    ByRef outRowStart As Long, _
    ByRef outColStart As Long, _
    ByRef outRowEnd As Long, _
    ByRef outColEnd As Long _
) As Boolean
    Dim key As String
    Dim entry As Object

    outRowStart = 0
    outColStart = 0
    outRowEnd = 0
    outColEnd = 0
    controlName = VBA.Trim$(controlName)
    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(controlName) = 0 Or VBA.Len(sheetName) = 0 Then Exit Function
    key = private_BuildRegistryKey(sheetName, controlName)

    private_EnsureRegistry
    If Not g_ControlRegistry.Exists(key) Then Exit Function
    Set entry = g_ControlRegistry(key)
    If entry Is Nothing Then Exit Function
    If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function

    outRowStart = VBA.CLng(entry("RowStart"))
    outColStart = VBA.CLng(entry("ColStart"))
    outRowEnd = VBA.CLng(entry("RowEnd"))
    outColEnd = VBA.CLng(entry("ColEnd"))
    fn_TryGetControlRenderBounds = True
End Function


Public Function fn_TranslateRegisteredControlsBelow( _
    ByVal sheetName As String, _
    ByVal firstRow As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim key As Variant
    Dim entry As Object
    Dim rowStart As Long
    Dim rowEnd As Long

    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Or firstRow <= 0 Then Exit Function
    If rowDelta = 0 Then
        fn_TranslateRegisteredControlsBelow = True
        Exit Function
    End If

    private_EnsureRegistry
    For Each key In g_ControlRegistry.Keys
        Set entry = g_ControlRegistry(key)
        If entry Is Nothing Then GoTo ContinueEntry
        If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), sheetName, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry

        rowStart = VBA.CLng(entry("RowStart"))
        rowEnd = VBA.CLng(entry("RowEnd"))
        If rowStart >= firstRow Then
            entry("RowStart") = rowStart + rowDelta
            entry("RowEnd") = rowEnd + rowDelta
        ElseIf rowEnd >= firstRow Then
            ' Контейнер/контрол, пересекающий границу reflow, сохраняет начало,
            ' но должен охватить новую высоту перемещенного хвоста.
            entry("RowEnd") = rowEnd + rowDelta
        End If
ContinueEntry:
    Next key

    fn_TranslateRegisteredControlsBelow = True
End Function


Public Function fn_TranslateRegisteredControlsInRegion( _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim key As Variant
    Dim entry As Object

    private_EnsureRegistry
    For Each key In g_ControlRegistry.Keys
        Set entry = g_ControlRegistry(key)
        If private_EntryIsInsideRegion(entry, sheetName, rowStart, colStart, rowEnd, colEnd) Then
            entry("RowStart") = VBA.CLng(entry("RowStart")) + rowDelta
            entry("RowEnd") = VBA.CLng(entry("RowEnd")) + rowDelta
        End If
    Next key
    fn_TranslateRegisteredControlsInRegion = True
End Function


Public Function fn_TryBuildLayoutReflowPlan( _
    ByVal sheetName As String, _
    ByVal controlName As String, _
    ByVal newSpanRows As Long, _
    ByRef outPatches As Collection, _
    ByRef outAncestorUpdates As Collection _
) As Boolean
    Dim changedEntry As Object

    ' На этом этапе Excel ещё не изменяется. Метод строит только план физических
    ' переносов и новые bounds предков; применяет его obj_PageBase отдельно.
    Set outPatches = New Collection
    Set outAncestorUpdates = New Collection
    sheetName = VBA.Trim$(sheetName)
    controlName = VBA.Trim$(controlName)
    If VBA.Len(sheetName) = 0 Or VBA.Len(controlName) = 0 Or newSpanRows <= 0 Then Exit Function
    If Not private_TryFindLayoutEntryByControlName(sheetName, controlName, changedEntry) Then Exit Function

    fn_TryBuildLayoutReflowPlan = private_TryBuildLayoutReflowPlanFromEntry( _
        sheetName, changedEntry, newSpanRows, outPatches, outAncestorUpdates)
End Function


Public Function fn_TryBuildLayoutContainerReflowPlan( _
    ByVal sheetName As String, _
    ByVal containerName As String, _
    ByVal newSpanRows As Long, _
    ByRef outPatches As Collection, _
    ByRef outAncestorUpdates As Collection _
) As Boolean
    Dim changedEntry As Object

    ' Container-вариант нужен, когда меняется сразу группа потомков. Типичный
    ' случай — несколько Collapsed-контролов становятся Visible одновременно:
    ' у них не было собственных bounds, но у именованного родителя они есть.
    Set outPatches = New Collection
    Set outAncestorUpdates = New Collection
    sheetName = VBA.Trim$(sheetName)
    containerName = VBA.Trim$(containerName)
    If VBA.Len(sheetName) = 0 Or VBA.Len(containerName) = 0 Or newSpanRows <= 0 Then Exit Function
    If Not private_TryFindLayoutEntryByNodeName( _
        sheetName, "stackpanel", containerName, changedEntry) Then Exit Function

    fn_TryBuildLayoutContainerReflowPlan = private_TryBuildLayoutReflowPlanFromEntry( _
        sheetName, changedEntry, newSpanRows, outPatches, outAncestorUpdates)
End Function


Private Function private_TryBuildLayoutReflowPlanFromEntry( _
    ByVal sheetName As String, _
    ByVal changedEntry As Object, _
    ByVal newSpanRows As Long, _
    ByRef outPatches As Collection, _
    ByRef outAncestorUpdates As Collection _
) As Boolean
    Dim parentEntry As Object
    Dim children As Collection
    Dim childEntry As Object
    Dim patch As Object
    Dim updateEntry As Object
    Dim changedKey As String
    Dim parentKey As String
    Dim parentKind As String
    Dim orientation As String
    Dim oldChangedEnd As Long
    Dim newChangedEnd As Long
    Dim oldParentEnd As Long
    Dim newParentEnd As Long
    Dim maxEnd As Long
    Dim rowDelta As Long

    If changedEntry Is Nothing Or newSpanRows <= 0 Then Exit Function

    changedKey = VBA.CStr(changedEntry("Key"))
    oldChangedEnd = VBA.CLng(changedEntry("RowEnd"))
    newChangedEnd = VBA.CLng(changedEntry("RowStart")) + newSpanRows - 1

    Do
        parentKey = VBA.CStr(changedEntry("ParentKey"))
        If VBA.Len(parentKey) = 0 Then Exit Do
        ' <page> не имеет физических bounds и поэтому не регистрируется.
        ' Дойдя до него, propagation штатно завершается.
        If Not private_TryGetLayoutEntry(sheetName, parentKey, parentEntry) Then Exit Do
        If Not private_TryGetDirectLayoutChildren(sheetName, parentKey, children) Then Exit Function

        oldParentEnd = VBA.CLng(parentEntry("RowEnd"))
        parentKind = VBA.LCase$(VBA.CStr(parentEntry("Kind")))
        orientation = VBA.LCase$(VBA.CStr(parentEntry("Orientation")))

        If VBA.CBool(parentEntry("FixedRows")) Then
            ' Fixed-height parent является layout boundary: dynamic child не
            ' имеет права выйти за его allocated bounds и перекрыть следующий flow.
            If newChangedEnd > oldParentEnd Then Exit Function
            newParentEnd = oldParentEnd
            rowDelta = 0
            GoTo RecordParentUpdate
        End If

        If parentKind = "stackpanel" And orientation = "vertical" Then
            ' В vertical flow все следующие siblings действительно зависят от
            ' нижней границы changed node, поэтому переносим готовый хвост одним
            ' прямоугольным patch, не вызывая Render для каждого соседа.
            rowDelta = newChangedEnd - oldChangedEnd
            newParentEnd = oldParentEnd + rowDelta
            If rowDelta <> 0 And oldChangedEnd < oldParentEnd Then
                Set patch = private_CreateReflowPatch( _
                    oldChangedEnd + 1, VBA.CLng(parentEntry("ColStart")), _
                    oldParentEnd, VBA.CLng(parentEntry("ColEnd")), rowDelta)
                outPatches.Add patch
            End If
        Else
            ' Horizontal stack и grid не двигают соседей по вертикали.
            ' Меняется только max bottom родителя, если changed child стал выше.
            maxEnd = VBA.CLng(parentEntry("RowStart"))
            For Each childEntry In children
                If VBA.CStr(childEntry("Key")) = changedKey Then
                    If newChangedEnd > maxEnd Then maxEnd = newChangedEnd
                ElseIf VBA.CLng(childEntry("RowEnd")) > maxEnd Then
                    maxEnd = VBA.CLng(childEntry("RowEnd"))
                End If
            Next childEntry
            newParentEnd = maxEnd
            rowDelta = newParentEnd - oldParentEnd
        End If

RecordParentUpdate:
        Set updateEntry = private_CreateDictionary()
        updateEntry("Key") = VBA.CStr(parentEntry("Key"))
        updateEntry("RowStart") = VBA.CLng(parentEntry("RowStart"))
        updateEntry("ColStart") = VBA.CLng(parentEntry("ColStart"))
        updateEntry("OldRowEnd") = VBA.CLng(oldParentEnd)
        updateEntry("ColEnd") = VBA.CLng(parentEntry("ColEnd"))
        updateEntry("RowEnd") = VBA.CLng(newParentEnd)
        outAncestorUpdates.Add updateEntry

        changedKey = VBA.CStr(parentEntry("Key"))
        oldChangedEnd = oldParentEnd
        newChangedEnd = newParentEnd
        Set changedEntry = parentEntry
        If rowDelta = 0 Then Exit Do
    Loop

    private_TryBuildLayoutReflowPlanFromEntry = True
End Function


Public Function fn_CommitLayoutReflowPlan( _
    ByVal sheetName As String, _
    ByVal controlName As String, _
    ByVal newRowEnd As Long, _
    ByVal newColEnd As Long, _
    ByVal ancestorUpdates As Collection _
) As Boolean
    Dim entry As Object

    ' Commit вызывается после переноса и успешной отрисовки target. До этого
    ' расчёты обязаны опираться на старую геометрию реестра, иначе повторный
    ' reflow начнёт накапливать delta от уже "будущих" координат.
    If Not private_TryFindLayoutEntryByControlName(sheetName, controlName, entry) Then Exit Function
    fn_CommitLayoutReflowPlan = private_CommitLayoutReflowPlanForEntry( _
        sheetName, entry, newRowEnd, newColEnd, ancestorUpdates)
End Function


Public Function fn_CommitLayoutContainerReflowPlan( _
    ByVal sheetName As String, _
    ByVal containerName As String, _
    ByVal newRowEnd As Long, _
    ByVal newColEnd As Long, _
    ByVal ancestorUpdates As Collection _
) As Boolean
    Dim entry As Object

    If Not private_TryFindLayoutEntryByNodeName( _
        sheetName, "stackpanel", containerName, entry) Then Exit Function
    fn_CommitLayoutContainerReflowPlan = private_CommitLayoutReflowPlanForEntry( _
        sheetName, entry, newRowEnd, newColEnd, ancestorUpdates)
End Function


Private Function private_CommitLayoutReflowPlanForEntry( _
    ByVal sheetName As String, _
    ByVal entry As Object, _
    ByVal newRowEnd As Long, _
    ByVal newColEnd As Long, _
    ByVal ancestorUpdates As Collection _
) As Boolean
    Dim updateEntry As Object

    If entry Is Nothing Then Exit Function
    entry("RowEnd") = VBA.CLng(newRowEnd)
    entry("ColEnd") = VBA.CLng(newColEnd)

    If Not ancestorUpdates Is Nothing Then
        For Each updateEntry In ancestorUpdates
            If Not private_TryGetLayoutEntry(sheetName, VBA.CStr(updateEntry("Key")), entry) Then Exit Function
            entry("RowEnd") = VBA.CLng(updateEntry("RowEnd"))
        Next updateEntry
    End If
    private_CommitLayoutReflowPlanForEntry = True
End Function


Public Function fn_TranslateLayoutEntriesInRegion( _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim key As Variant
    Dim entry As Object

    private_EnsureLayoutRegistry
    For Each key In g_LayoutRegistry.Keys
        Set entry = g_LayoutRegistry(key)
        If private_EntryIsInsideRegion(entry, sheetName, rowStart, colStart, rowEnd, colEnd) Then
            entry("RowStart") = VBA.CLng(entry("RowStart")) + rowDelta
            entry("RowEnd") = VBA.CLng(entry("RowEnd")) + rowDelta
        End If
    Next key
    fn_TranslateLayoutEntriesInRegion = True
End Function


Public Function fn_TryRefreshStaticControl(ByVal controlName As String) As Boolean
    Dim key As String
    Dim entry As Object
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim pageRef As obj_IPage
    Dim pageBase As obj_PageBase
    Dim uiDoc As Object
    Dim controlNode As Object
    Dim escapedName As String
    Dim xPath As String
    Dim renderCtx As obj_LayoutRenderContext

    controlName = VBA.Trim$(controlName)
    If VBA.Len(controlName) = 0 Then Exit Function

    private_EnsureRegistry

    key = private_FindRegistryKey(controlName)
    If VBA.Len(key) = 0 Then Exit Function

    Set entry = g_ControlRegistry(key)
    If entry Is Nothing Then Exit Function
    If Not private_IsStaticControlType(VBA.CStr(entry("Type"))) Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VBA.CStr(entry("Sheet")))
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set wb = ws.Parent
    If wb Is Nothing Then Exit Function

    Set uiDoc = ex_XmlCore.fn_LoadDomByRelativePath( _
        wb, _
        VBA.CStr(entry("UiPath")), _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If uiDoc Is Nothing Then Exit Function

    escapedName = ex_XmlCore.fn_XPathLiteral(VBA.CStr(entry("Name")))
    xPath = "/p:page//p:control[@name=" & escapedName & "] | /p:uiDefinition/p:layout//p:control[@name=" & escapedName & "]"
    Set controlNode = uiDoc.selectSingleNode(xPath)
    If controlNode Is Nothing Then Exit Function

    If Not rt_PageManager.fn_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    Set pageBase = pageRef.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set renderCtx = New obj_LayoutRenderContext
    If Not renderCtx.Initialize(pageRef) Then Exit Function

    If Not ex_XmlLayoutEngine.fn_RenderNodeInBounds( _
        renderCtx:=renderCtx, _
        layoutNode:=controlNode, _
        rowStart:=VBA.CLng(entry("RowStart")), _
        colStart:=VBA.CLng(entry("ColStart")), _
        rowEnd:=VBA.CLng(entry("RowEnd")), _
        colEnd:=VBA.CLng(entry("ColEnd"))) Then Exit Function

    If Not pageBase.ApplyInlineRuns() Then Exit Function

    fn_TryRefreshStaticControl = True
End Function


Public Function fn_TryGetSheetMaxControlBounds( _
    ByVal sheetName As String, _
    ByRef outRowEnd As Long, _
    ByRef outColEnd As Long _
) As Boolean
    Dim key As Variant
    Dim entry As Object
    Dim entrySheetName As String
    Dim rowEnd As Long
    Dim colEnd As Long

    outRowEnd = 0
    outColEnd = 0
    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Then Exit Function

    private_EnsureRegistry

    For Each key In g_ControlRegistry.Keys
        Set entry = g_ControlRegistry(key)
        If entry Is Nothing Then GoTo ContinueEntry

        entrySheetName = VBA.Trim$(VBA.CStr(entry("Sheet")))
        If VBA.StrComp(entrySheetName, sheetName, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry

        rowEnd = VBA.CLng(entry("RowEnd"))
        colEnd = VBA.CLng(entry("ColEnd"))
        If rowEnd > outRowEnd Then outRowEnd = rowEnd
        If colEnd > outColEnd Then outColEnd = colEnd

ContinueEntry:
    Next key

    fn_TryGetSheetMaxControlBounds = (outRowEnd > 0 And outColEnd > 0)
End Function

' //
' // Internal
' //

Private Sub private_EnsureRegistry()
    If g_ControlRegistry Is Nothing Then
        Set g_ControlRegistry = private_CreateDictionary()
    End If
End Sub

Private Sub private_EnsureLayoutRegistry()
    If g_LayoutRegistry Is Nothing Then Set g_LayoutRegistry = private_CreateDictionary()
End Sub

Private Sub private_RemoveLayoutEntriesByWorksheet(ByVal sheetName As String)
    Dim key As Variant
    Dim keysToRemove As Collection
    Dim entry As Object

    private_EnsureLayoutRegistry
    Set keysToRemove = New Collection
    For Each key In g_LayoutRegistry.Keys
        Set entry = g_LayoutRegistry(key)
        If VBA.StrComp(VBA.CStr(entry("Sheet")), sheetName, VBA.vbTextCompare) = 0 Then keysToRemove.Add key
    Next key
    For Each key In keysToRemove
        g_LayoutRegistry.Remove key
    Next key
End Sub

Private Function private_BuildLayoutRegistryKey(ByVal sheetName As String, ByVal layoutKey As String) As String
    private_BuildLayoutRegistryKey = VBA.LCase$(VBA.Trim$(sheetName)) & "|" & VBA.LCase$(VBA.Trim$(layoutKey))
End Function

Private Function private_HasPositiveLongAttr(ByVal node As Object, ByVal attrName As String) As Boolean
    Dim rawText As String
    If node Is Nothing Then Exit Function
    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(node, attrName)))
    If VBA.Len(rawText) = 0 Or Not VBA.IsNumeric(rawText) Then Exit Function
    private_HasPositiveLongAttr = (VBA.CLng(rawText) > 0)
End Function

Private Function private_BuildLayoutNodeKey(ByVal layoutNode As Object) As String
    Dim currentNode As Object
    Dim sibling As Object
    Dim segmentText As String
    Dim resultText As String
    Dim siblingIndex As Long
    Dim nodeKind As String

    If layoutNode Is Nothing Then Exit Function
    Set currentNode = layoutNode
    Do While Not currentNode Is Nothing
        If currentNode.NodeType = 1 Then
            nodeKind = VBA.LCase$(VBA.Trim$(VBA.CStr(currentNode.baseName)))
            If nodeKind = "page" Or nodeKind = "grid" Or nodeKind = "stackpanel" Or _
               nodeKind = "control" Or nodeKind = "list" Or nodeKind = "itemcontrol" Then
                siblingIndex = 1
                Set sibling = currentNode.previousSibling
                Do While Not sibling Is Nothing
                    If sibling.NodeType = 1 Then
                        If VBA.LCase$(VBA.CStr(sibling.baseName)) = nodeKind Then siblingIndex = siblingIndex + 1
                    End If
                    Set sibling = sibling.previousSibling
                Loop
                segmentText = nodeKind & "[" & VBA.CStr(siblingIndex) & "]"
                If VBA.Len(resultText) = 0 Then
                    resultText = segmentText
                Else
                    resultText = segmentText & "/" & resultText
                End If
            End If
        End If
        Set currentNode = currentNode.parentNode
    Loop
    private_BuildLayoutNodeKey = resultText
End Function

Private Function private_TryFindLayoutEntryByControlName( _
    ByVal sheetName As String, _
    ByVal controlName As String, _
    ByRef outEntry As Object _
) As Boolean
    Dim key As Variant
    Dim entry As Object

    Set outEntry = Nothing
    private_EnsureLayoutRegistry
    For Each key In g_LayoutRegistry.Keys
        Set entry = g_LayoutRegistry(key)
        If VBA.StrComp(VBA.CStr(entry("Sheet")), sheetName, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        If VBA.StrComp(VBA.CStr(entry("Kind")), "control", VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        If VBA.StrComp(VBA.CStr(entry("Name")), controlName, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        Set outEntry = entry
        private_TryFindLayoutEntryByControlName = True
        Exit Function
ContinueEntry:
    Next key
End Function

Private Function private_TryFindLayoutEntryByNodeName( _
    ByVal sheetName As String, _
    ByVal nodeKind As String, _
    ByVal nodeName As String, _
    ByRef outEntry As Object _
) As Boolean
    Dim key As Variant
    Dim entry As Object

    Set outEntry = Nothing
    nodeKind = VBA.LCase$(VBA.Trim$(nodeKind))
    nodeName = VBA.Trim$(nodeName)
    If VBA.Len(nodeKind) = 0 Or VBA.Len(nodeName) = 0 Then Exit Function

    private_EnsureLayoutRegistry
    For Each key In g_LayoutRegistry.Keys
        Set entry = g_LayoutRegistry(key)
        If VBA.StrComp(VBA.CStr(entry("Sheet")), sheetName, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        If VBA.StrComp(VBA.CStr(entry("Kind")), nodeKind, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        If VBA.StrComp(VBA.CStr(entry("Name")), nodeName, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        Set outEntry = entry
        private_TryFindLayoutEntryByNodeName = True
        Exit Function
ContinueEntry:
    Next key
End Function

Private Function private_TryGetLayoutEntry( _
    ByVal sheetName As String, _
    ByVal layoutKey As String, _
    ByRef outEntry As Object _
) As Boolean
    Dim registryKey As String
    Set outEntry = Nothing
    private_EnsureLayoutRegistry
    registryKey = private_BuildLayoutRegistryKey(sheetName, layoutKey)
    If Not g_LayoutRegistry.Exists(registryKey) Then Exit Function
    Set outEntry = g_LayoutRegistry(registryKey)
    private_TryGetLayoutEntry = Not outEntry Is Nothing
End Function

Private Function private_TryGetDirectLayoutChildren( _
    ByVal sheetName As String, _
    ByVal parentKey As String, _
    ByRef outChildren As Collection _
) As Boolean
    Dim key As Variant
    Dim entry As Object
    Set outChildren = New Collection
    private_EnsureLayoutRegistry
    For Each key In g_LayoutRegistry.Keys
        Set entry = g_LayoutRegistry(key)
        If VBA.StrComp(VBA.CStr(entry("Sheet")), sheetName, VBA.vbTextCompare) = 0 And _
           VBA.StrComp(VBA.CStr(entry("ParentKey")), parentKey, VBA.vbTextCompare) = 0 Then outChildren.Add entry
    Next key
    private_TryGetDirectLayoutChildren = True
End Function

Private Function private_CreateReflowPatch( _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
) As Object
    Dim patch As Object
    Set patch = private_CreateDictionary()
    patch("RowStart") = rowStart
    patch("ColStart") = colStart
    patch("RowEnd") = rowEnd
    patch("ColEnd") = colEnd
    patch("RowDelta") = rowDelta
    Set private_CreateReflowPatch = patch
End Function

Private Function private_EntryIsInsideRegion( _
    ByVal entry As Object, _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
) As Boolean
    If entry Is Nothing Then Exit Function
    If VBA.StrComp(VBA.CStr(entry("Sheet")), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    private_EntryIsInsideRegion = _
        VBA.CLng(entry("RowStart")) >= rowStart And _
        VBA.CLng(entry("RowEnd")) <= rowEnd And _
        VBA.CLng(entry("ColStart")) >= colStart And _
        VBA.CLng(entry("ColEnd")) <= colEnd
End Function


Private Function private_CreateDictionary() As Object
    Set private_CreateDictionary = VBA.CreateObject("Scripting.Dictionary")
    private_CreateDictionary.CompareMode = 1
End Function


Private Function private_BuildRegistryKey(ByVal sheetName As String, ByVal controlName As String) As String
    private_BuildRegistryKey = VBA.LCase$(VBA.Trim$(sheetName)) & "|" & VBA.LCase$(VBA.Trim$(controlName))
End Function


Private Function private_FindRegistryKey(ByVal controlName As String) As String
    Dim key As Variant
    Dim entry As Object
    Dim activeSheetName As String
    Dim matchedKey As String
    Dim matchCount As Long

    controlName = VBA.LCase$(VBA.Trim$(controlName))
    If VBA.Len(controlName) = 0 Then Exit Function
    On Error Resume Next
    activeSheetName = VBA.Trim$(Application.ActiveSheet.Name)
    On Error GoTo 0

    For Each key In g_ControlRegistry.Keys
        Set entry = g_ControlRegistry(key)
        If entry Is Nothing Then GoTo ContinueEntry
        If VBA.LCase$(VBA.Trim$(VBA.CStr(entry("Name")))) <> controlName Then GoTo ContinueEntry
        If VBA.Len(activeSheetName) > 0 Then
            If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), activeSheetName, VBA.vbTextCompare) = 0 Then
                private_FindRegistryKey = VBA.CStr(key)
                Exit Function
            End If
        End If
        matchedKey = VBA.CStr(key)
        matchCount = matchCount + 1
ContinueEntry:
    Next key

    If matchCount = 1 Then private_FindRegistryKey = matchedKey
End Function


Private Function private_IsStaticControlType(ByVal controlType As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(controlType))
        Case "label", "banner", "button", "config", "select", "input"
            private_IsStaticControlType = True
    End Select
End Function
