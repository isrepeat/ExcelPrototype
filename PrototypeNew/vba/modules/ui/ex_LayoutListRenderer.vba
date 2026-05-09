Attribute VB_Name = "ex_LayoutListRenderer"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Renderer for <list> nodes.

Private Const UI_NS As String = "urn:excelprototype:profiles"

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_LayoutListRenderer.fn_Module_Dispose"
#End If
End Sub
' //
' // API
' //
Public Function fn_Render( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal layoutNode As Object, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pageBase As obj_PageBase
    Dim items As Collection
    Dim templateRoot As Object
    Dim listOrientation As String
    Dim tempDoc As Object
    Dim syntheticRoot As Object
    Dim itemValue As Variant
    Dim clonedNode As Object
    Dim itemIndex As Long

    If layoutNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list node is not specified."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(layoutNode.baseName)), "list", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: ex_LayoutListRenderer supports only <list> nodes."
#End If
        Exit Function
    End If
    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function
    Set pageBase = renderCtx.Page.GetPageBase()
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page base is not specified for list renderer."
#End If
        Exit Function
    End If

    ' itemsSource резолвится единым resolver-ом:
    ' - runtime source expression ({PageRuntimeSource/...}, {GlobalRuntimeSource/...})
    ' - или Binding, возвращающий Collection.
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource( _
        pageBase.RuntimeSources, _
        ex_XmlCore.fn_NodeAttrText(layoutNode, "itemsSource"), _
        items) Then Exit Function

    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list itemsSource resolved to Nothing."
#End If
        Exit Function
    End If
    If items.Count = 0 Then
        fn_Render = True
        Exit Function
    End If

    If Not private_TryResolveListTemplateRoot(layoutNode, templateRoot) Then Exit Function

    listOrientation = private_GetListOrientation(layoutNode)
    If VBA.Len(listOrientation) = 0 Then Exit Function

    Set tempDoc = ex_XmlCore.fn_CreateDom(UI_NS)
    private_CopyTemplatesToTempListDoc tempDoc, layoutNode.OwnerDocument

    Set syntheticRoot = tempDoc.createNode(1, "stackPanel", UI_NS)
    syntheticRoot.setAttribute "orientation", listOrientation

    itemIndex = 0
    For Each itemValue In items
        itemIndex = itemIndex + 1

        Set clonedNode = tempDoc.importNode(templateRoot, True)
        If clonedNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to clone list template node."
#End If
            Exit Function
        End If

        ' Для каждого итема создаем собственный dataContext и применяем Binding
        ' ко всем атрибутам внутри template.
        If Not private_ApplyListItemBindings(clonedNode, itemValue, renderCtx) Then Exit Function
        private_AppendSuffixToControlNames clonedNode, "_" & VBA.CStr(itemIndex)
        private_ApplyListItemValueToTemplate clonedNode, itemValue
        syntheticRoot.appendChild clonedNode
    Next itemValue

    fn_Render = ex_XmlLayoutEngine.fn_RenderContainerNodeInBounds( _
        renderCtx:=renderCtx, _
        containerNode:=syntheticRoot, _
        layoutRowStart:=rowStart, _
        layoutColStart:=colStart, _
        layoutRowEnd:=rowEnd, _
        layoutColEnd:=colEnd)
End Function


Public Function fn_TryMeasureContentSpan( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal listNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    Dim items As Collection
    Dim templateRoot As Object
    Dim itemRows As Long
    Dim itemCols As Long
    Dim orientation As String
    Dim itemValue As Variant
    Dim itemBindingSource As Object

    If listNode Is Nothing Then Exit Function
    If VBA.StrComp(VBA.LCase$(VBA.CStr(listNode.baseName)), "list", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: ex_LayoutListRenderer supports only <list> nodes."
#End If
        Exit Function
    End If

    If Not private_TryResolveItemsSourceForMeasure( _
        renderCtx, ex_XmlCore.fn_NodeAttrText(listNode, "itemsSource"), dataContext, items) Then Exit Function

    If items Is Nothing Or items.Count = 0 Then
        outSpanRows = 1
        outSpanColls = 1
        fn_TryMeasureContentSpan = True
        Exit Function
    End If

    If Not private_TryResolveListTemplateRoot(listNode, templateRoot) Then Exit Function

    orientation = private_GetListOrientation(listNode)
    If VBA.Len(orientation) = 0 Then Exit Function

    outSpanRows = 0
    outSpanColls = 0

    For Each itemValue In items
        Set itemBindingSource = Nothing
        If Not private_TryCreateListItemBindingSource(itemValue, itemBindingSource) Then Exit Function
        If Not ex_XmlLayoutEngine.fn_TryGetEffectiveNodeSpan(renderCtx, templateRoot, itemRows, itemCols, itemBindingSource) Then Exit Function

        If VBA.StrComp(orientation, "horizontal", VBA.vbBinaryCompare) = 0 Then
            If itemRows > outSpanRows Then outSpanRows = itemRows
            outSpanColls = outSpanColls + itemCols
        Else
            outSpanRows = outSpanRows + itemRows
            If itemCols > outSpanColls Then outSpanColls = itemCols
        End If
    Next itemValue

    If outSpanRows <= 0 Then outSpanRows = 1
    If outSpanColls <= 0 Then outSpanColls = 1
    fn_TryMeasureContentSpan = True
End Function

' //
' // Internal
' //
Private Sub private_CopyTemplatesToTempListDoc(ByVal targetDoc As Object, ByVal sourceDoc As Object)
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

    Set srcTemplateNodes = sourceDoc.selectNodes("/p:page/p:templates/p:template | /p:uiDefinition/p:templates/p:template")
    If srcTemplateNodes Is Nothing Then Exit Sub

    For Each srcTemplateNode In srcTemplateNodes
        templateName = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(srcTemplateNode, "name"))
        If VBA.Len(templateName) = 0 Then GoTo ContinueTemplate

        If Not targetTemplatesNode.selectSingleNode("p:template[@name=" & ex_XmlCore.fn_XPathLiteral(templateName) & "]") Is Nothing Then
            GoTo ContinueTemplate
        End If

        targetTemplatesNode.appendChild targetDoc.importNode(srcTemplateNode, True)

ContinueTemplate:
    Next srcTemplateNode
End Sub


Private Function private_TryResolveListTemplateRoot(ByVal listNode As Object, ByRef outTemplateRoot As Object) As Boolean
    Dim templateName As String
    Dim ownerDoc As Object
    Dim templateNode As Object
    Dim rootNodes As Object

    templateName = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(listNode, "itemsSourceTemplate"))
    If VBA.Len(templateName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list requires non-empty attribute 'itemsSourceTemplate'."
#End If
        Exit Function
    End If

    Set ownerDoc = listNode.OwnerDocument
    If ownerDoc Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to resolve owner document for list template."
#End If
        Exit Function
    End If

    Set templateNode = ownerDoc.selectSingleNode( _
        "/p:page/p:templates/p:template[@name=" & ex_XmlCore.fn_XPathLiteral(templateName) & "] | /p:uiDefinition/p:templates/p:template[@name=" & ex_XmlCore.fn_XPathLiteral(templateName) & "]")
    If templateNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list references missing template '" & templateName & "'."
#End If
        Exit Function
    End If

    Set rootNodes = templateNode.selectNodes("p:control | p:stackPanel | p:grid | p:list | p:itemControl")
    If rootNodes Is Nothing Or rootNodes.Length = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: template '" & templateName & "' has no visual root node."
#End If
        Exit Function
    End If
    If rootNodes.Length <> 1 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: template '" & templateName & "' must contain exactly one visual root node."
#End If
        Exit Function
    End If

    Set outTemplateRoot = rootNodes.Item(0)
    private_TryResolveListTemplateRoot = True
End Function


Private Function private_GetListOrientation(ByVal listNode As Object) As String
    private_GetListOrientation = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText(listNode, "orientation")))
    If VBA.Len(private_GetListOrientation) = 0 Then private_GetListOrientation = "vertical"

    If VBA.StrComp(private_GetListOrientation, "vertical", VBA.vbBinaryCompare) <> 0 And _
       VBA.StrComp(private_GetListOrientation, "horizontal", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list orientation must be 'vertical' or 'horizontal'."
#End If
        private_GetListOrientation = VBA.vbNullString
    End If
End Function


Private Sub private_AppendSuffixToControlNames(ByVal rootNode As Object, ByVal suffix As String)
    Dim childNode As Object
    Dim baseName As String
    Dim nodeName As String

    If rootNode Is Nothing Then Exit Sub
    If rootNode.NodeType <> 1 Then Exit Sub

    baseName = VBA.LCase$(VBA.CStr(rootNode.baseName))
    If VBA.StrComp(baseName, "control", VBA.vbBinaryCompare) = 0 Then
        nodeName = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(rootNode, "name"))
        If VBA.Len(nodeName) = 0 Then nodeName = "item"
        rootNode.setAttribute "name", nodeName & suffix
    End If

    For Each childNode In rootNode.ChildNodes
        private_AppendSuffixToControlNames childNode, suffix
    Next childNode
End Sub


Private Sub private_ApplyListItemValueToTemplate(ByVal templateRoot As Object, ByVal itemValue As Variant)
    Dim captionText As String
    Dim targetControl As Object
    Dim existingCaption As String

    If VBA.IsObject(itemValue) Then Exit Sub

    captionText = VBA.CStr(itemValue)
    Set targetControl = private_FindFirstControlNode(templateRoot)
    If targetControl Is Nothing Then Exit Sub

    existingCaption = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(targetControl, "caption"))
    If VBA.Len(existingCaption) > 0 Then Exit Sub

    targetControl.setAttribute "caption", captionText
End Sub


Private Function private_ApplyListItemBindings( _
    ByVal templateRoot As Object, _
    ByVal itemValue As Variant, _
    ByVal renderCtx As obj_LayoutRenderContext _
) As Boolean
    Dim dataContext As Object

    If Not private_TryCreateListItemBindingSource(itemValue, dataContext) Then Exit Function
    If Not private_ApplyNodeBindingsRecursive(templateRoot, dataContext, renderCtx) Then Exit Function

    private_ApplyListItemBindings = True
End Function


Private Function private_TryCreateListItemBindingSource(ByVal itemValue As Variant, ByRef outSource As Object) As Boolean
    Dim scalarSource As Object
    Dim scalarText As String

    If VBA.IsObject(itemValue) Then
        If itemValue Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: list item object is Nothing."
#End If
            Exit Function
        End If

        Set outSource = itemValue
        private_TryCreateListItemBindingSource = True
        Exit Function
    End If

    ' Скалярные элементы оборачиваем в dictionary, чтобы Path=Value/Path=Text работал
    ' так же, как для object-элементов списка.
    scalarText = VBA.CStr(itemValue)
    Set scalarSource = VBA.CreateObject("Scripting.Dictionary")
    scalarSource.CompareMode = 1
    scalarSource("Value") = scalarText
    scalarSource("Text") = scalarText

    Set outSource = scalarSource
    private_TryCreateListItemBindingSource = True
End Function


Private Function private_ApplyNodeBindingsRecursive( _
    ByVal rootNode As Object, _
    ByVal dataContext As Object, _
    ByVal renderCtx As obj_LayoutRenderContext _
) As Boolean
    Dim attrs As Object
    Dim attrNode As Object
    Dim attrName As String
    Dim rawText As String
    Dim resolvedValue As Variant
    Dim resolvedVisible As Boolean
    Dim childNode As Object
    Dim runtimeListSourceKey As String
    Dim runtimeObjectSourceKey As String
    Dim runtimeItems As Collection
    Dim resolvedObject As Object
    Dim rootNodeName As String

    If rootNode Is Nothing Then
        private_ApplyNodeBindingsRecursive = True
        Exit Function
    End If
    If rootNode.NodeType <> 1 Then
        private_ApplyNodeBindingsRecursive = True
        Exit Function
    End If

    Set attrs = rootNode.selectNodes("@*")
    If Not attrs Is Nothing Then
        For Each attrNode In attrs
            attrName = VBA.CStr(attrNode.nodeName)
            If VBA.LCase$(VBA.Left$(attrName, 5)) = "xmlns" Then GoTo ContinueAttr

            rawText = VBA.CStr(attrNode.Text)
            If VBA.InStr(1, rawText, "{Binding ", VBA.vbTextCompare) = 0 Then GoTo ContinueAttr

            If VBA.StrComp(VBA.LCase$(attrName), "visibility", VBA.vbBinaryCompare) = 0 Then
                If Not ex_BindingRuntime.fn_TryResolveVisibilityBinding(rawText, dataContext, resolvedVisible) Then Exit Function
                If resolvedVisible Then
                    rootNode.setAttribute attrName, "true"
                Else
                    rootNode.setAttribute attrName, "false"
                End If
                GoTo ContinueAttr
            End If

            ' Binding вычисляется от dataContext конкретного элемента списка.
            ' Объектные результаты не пишем в XML напрямую:
            ' регистрируем runtime-source key и подставляем ключ в атрибут.
            If Not ex_BindingRuntime.fn_TryResolveValueBinding(rawText, dataContext, resolvedValue) Then Exit Function

            If VBA.IsObject(resolvedValue) Then
                rootNodeName = VBA.LCase$(VBA.CStr(rootNode.baseName))

                If VBA.StrComp(VBA.LCase$(attrName), "itemssource", VBA.vbBinaryCompare) = 0 And _
                   (VBA.StrComp(rootNodeName, "list", VBA.vbBinaryCompare) = 0 Or _
                    VBA.StrComp(rootNodeName, "control", VBA.vbBinaryCompare) = 0) Then

                    If VBA.TypeName(resolvedValue) = "Collection" Then
                        Set runtimeItems = resolvedValue
                    ElseIf VBA.StrComp(rootNodeName, "control", VBA.vbBinaryCompare) = 0 Then
                        Set runtimeItems = New Collection
                        Set resolvedObject = resolvedValue
                        If resolvedObject Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
                            ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource binding resolved to Nothing object."
#End If
                            Exit Function
                        End If
                        runtimeItems.Add resolvedObject
                    Else
#If LOGGING_DEBUG_ENABLED Then
                        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list itemsSource binding must resolve to Collection."
#End If
                        Exit Function
                    End If

                    runtimeListSourceKey = private_RegisterRuntimeListItemsSourceKey(runtimeItems, renderCtx)
                    If VBA.Len(runtimeListSourceKey) = 0 Then Exit Function
                    rootNode.setAttribute attrName, runtimeListSourceKey
                ElseIf VBA.StrComp(VBA.LCase$(attrName), "objectsource", VBA.vbBinaryCompare) = 0 And _
                       VBA.StrComp(rootNodeName, "itemcontrol", VBA.vbBinaryCompare) = 0 Then

                    Set resolvedObject = resolvedValue
                    If resolvedObject Is Nothing Then
                        rootNode.setAttribute attrName, VBA.vbNullString
                    Else
                        runtimeObjectSourceKey = private_RegisterRuntimeObjectSourceKey(resolvedObject, renderCtx)
                        If VBA.Len(runtimeObjectSourceKey) = 0 Then Exit Function
                        rootNode.setAttribute attrName, runtimeObjectSourceKey
                    End If
                Else
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.fn_Diagnostic_LogError "PrototypeNew: template binding for attribute '" & attrName & "' must resolve to scalar value."
#End If
                    Exit Function
                End If
            Else
                rootNode.setAttribute attrName, VBA.CStr(resolvedValue)
            End If

ContinueAttr:
        Next attrNode
    End If

    For Each childNode In rootNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueChild
        If Not private_ApplyNodeBindingsRecursive(childNode, dataContext, renderCtx) Then Exit Function
ContinueChild:
    Next childNode

    private_ApplyNodeBindingsRecursive = True
End Function


Private Function private_RegisterRuntimeListItemsSourceKey( _
    ByVal items As Collection, _
    ByVal renderCtx As obj_LayoutRenderContext _
) As String
    Dim sourceKey As String

    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: runtime list items source is Nothing."
#End If
        Exit Function
    End If
    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    sourceKey = renderCtx.NextListRuntimeSourceKey()

    If Not renderCtx.Page.GetPageBase().RuntimeSources.SetItemsSource(sourceKey, items) Then Exit Function
    private_RegisterRuntimeListItemsSourceKey = sourceKey
End Function


Private Function private_RegisterRuntimeObjectSourceKey( _
    ByVal sourceObject As Object, _
    ByVal renderCtx As obj_LayoutRenderContext _
) As String
    Dim sourceKey As String

    If sourceObject Is Nothing Then Exit Function
    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    sourceKey = renderCtx.NextObjectRuntimeSourceKey()

    If Not renderCtx.Page.GetPageBase().RuntimeSources.SetObjectSource(sourceKey, sourceObject) Then Exit Function
    private_RegisterRuntimeObjectSourceKey = sourceKey
End Function


Private Function private_FindFirstControlNode(ByVal rootNode As Object) As Object
    Dim childNode As Object

    If rootNode Is Nothing Then Exit Function
    If rootNode.NodeType <> 1 Then Exit Function

    If VBA.StrComp(VBA.LCase$(VBA.CStr(rootNode.baseName)), "control", VBA.vbBinaryCompare) = 0 Then
        Set private_FindFirstControlNode = rootNode
        Exit Function
    End If

    For Each childNode In rootNode.ChildNodes
        Set private_FindFirstControlNode = private_FindFirstControlNode(childNode)
        If Not private_FindFirstControlNode Is Nothing Then Exit Function
    Next childNode
End Function


Private Function private_TryResolveItemsSourceForMeasure( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal rawItemsSource As String, _
    ByVal dataContext As Object, _
    ByRef outItems As Collection _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim sourceText As String
    Dim resolvedValue As Variant

    sourceText = VBA.Trim$(rawItemsSource)
    If VBA.Len(sourceText) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list itemsSource is required."
#End If
        Exit Function
    End If

    If private_IsBindingExpression(sourceText) Then
        If dataContext Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: list itemsSource binding requires item context during layout measurement."
#End If
            Exit Function
        End If

        ' На этапе measure binding тоже вычисляем, чтобы размер списка
        ' соответствовал реальным данным того же dataContext.
        If Not ex_BindingRuntime.fn_TryResolveValueBinding(sourceText, dataContext, resolvedValue) Then Exit Function

        If VBA.IsObject(resolvedValue) Then
            If VBA.TypeName(resolvedValue) <> "Collection" Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PrototypeNew: list itemsSource binding must resolve to Collection."
#End If
                Exit Function
            End If

            Set outItems = resolvedValue
            private_TryResolveItemsSourceForMeasure = True
            Exit Function
        End If

        sourceText = VBA.Trim$(VBA.CStr(resolvedValue))
        If VBA.Len(sourceText) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: list itemsSource binding resolved to empty value."
#End If
            Exit Function
        End If
    End If

    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list measure render context is not specified."
#End If
        Exit Function
    End If

    Set pageBase = renderCtx.Page.GetPageBase()
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: list measure page context is not specified in render context."
#End If
        Exit Function
    End If

    ' Небиндинговый текст source резолвим тем же путем, что и при render.
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, sourceText, outItems) Then Exit Function
    private_TryResolveItemsSourceForMeasure = True
End Function


Private Function private_IsBindingExpression(ByVal rawText As String) As Boolean
    Dim normalized As String

    normalized = VBA.Trim$(rawText)
    If VBA.Len(normalized) < 10 Then Exit Function
    If VBA.StrComp(VBA.Left$(normalized, 9), "{Binding ", VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(normalized, 1) <> "}" Then Exit Function

    private_IsBindingExpression = True
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
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    Set outWs = renderCtx.Worksheet
    If outWs Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified."
#End If
        Exit Function
    End If

    Set outWb = renderCtx.Workbook
    If outWb Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: workbook is not specified."
#End If
        Exit Function
    End If

    private_TryGetPageRenderContext = True
End Function
