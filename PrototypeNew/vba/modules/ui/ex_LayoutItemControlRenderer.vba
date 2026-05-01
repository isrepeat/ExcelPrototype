Attribute VB_Name = "ex_LayoutItemControlRenderer"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Renderer for <itemControl> nodes.

Private Const UI_NS As String = "urn:excelprototype:profiles"

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_LayoutItemControlRenderer.m_Module_Dispose"
#End If
End Sub
' //
' // API
' //
Public Function m_Render( _
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
    Dim sourceObject As Object
    Dim templateRoot As Object
    Dim tempDoc As Object
    Dim syntheticRoot As Object
    Dim clonedNode As Object
    Dim suffixValue As Long

    If layoutNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemControl node is not specified."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(layoutNode.baseName)), "itemcontrol", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: ex_LayoutItemControlRenderer supports only <itemControl> nodes."
#End If
        Exit Function
    End If
    If Not private_TryGetPageRenderContext(renderCtx, wb, ws) Then Exit Function
    Set pageBase = renderCtx.PageBase
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: page base is not specified for itemControl renderer."
#End If
        Exit Function
    End If

    ' objectSource для itemControl резолвится через общий resolver:
    ' - runtime source expression ({PageRuntimeSource/...}, {GlobalRuntimeSource/...})
    ' - либо Binding, который должен вернуть Object.
    If Not ex_RuntimeSourceResolver.m_TryResolveObjectSource( _
        pageBase.RuntimeSources, _
        ex_XmlCore.m_NodeAttrText(layoutNode, "objectSource"), _
        sourceObject, _
        True) Then Exit Function
    If sourceObject Is Nothing Then
        m_Render = True
        Exit Function
    End If

    If Not private_TryResolveItemControlTemplateRoot(layoutNode, templateRoot) Then Exit Function

    Set tempDoc = ex_XmlCore.m_CreateDom(UI_NS)
    private_CopyTemplatesToTempListDoc tempDoc, layoutNode.OwnerDocument

    Set syntheticRoot = tempDoc.createNode(1, "stackPanel", UI_NS)
    syntheticRoot.setAttribute "orientation", "vertical"

    Set clonedNode = tempDoc.importNode(templateRoot, True)
    If clonedNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: failed to clone itemControl template node."
#End If
        Exit Function
    End If

    If Not private_ApplyNodeBindingsRecursive(clonedNode, sourceObject, renderCtx) Then Exit Function
    suffixValue = renderCtx.NextObjectRenderSuffix()
    private_AppendSuffixToControlNames clonedNode, "_obj" & VBA.CStr(suffixValue)
    syntheticRoot.appendChild clonedNode

    m_Render = ex_XmlLayoutEngine.m_RenderContainerNodeInBounds( _
        renderCtx:=renderCtx, _
        containerNode:=syntheticRoot, _
        layoutRowStart:=rowStart, _
        layoutColStart:=colStart, _
        layoutRowEnd:=rowEnd, _
        layoutColEnd:=colEnd)
End Function


Public Function m_TryMeasureContentSpan( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal itemControlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanCols As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    Dim sourceObject As Object
    Dim templateRoot As Object

    If itemControlNode Is Nothing Then Exit Function
    If VBA.StrComp(VBA.LCase$(VBA.CStr(itemControlNode.baseName)), "itemcontrol", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: ex_LayoutItemControlRenderer supports only <itemControl> nodes."
#End If
        Exit Function
    End If

    If Not private_TryResolveObjectSourceForMeasure( _
        renderCtx, _
        ex_XmlCore.m_NodeAttrText(itemControlNode, "objectSource"), _
        dataContext, _
        sourceObject) Then Exit Function

    If sourceObject Is Nothing Then
        outSpanRows = 0
        outSpanCols = 0
        m_TryMeasureContentSpan = True
        Exit Function
    End If

    If Not private_TryResolveItemControlTemplateRoot(itemControlNode, templateRoot) Then Exit Function
    If Not ex_XmlLayoutEngine.m_TryGetEffectiveNodeSpan(renderCtx, templateRoot, outSpanRows, outSpanCols, sourceObject) Then Exit Function
    m_TryMeasureContentSpan = True
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
        templateName = VBA.Trim$(ex_XmlCore.m_NodeAttrText(srcTemplateNode, "name"))
        If VBA.Len(templateName) = 0 Then GoTo ContinueTemplate

        If Not targetTemplatesNode.selectSingleNode("p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]") Is Nothing Then
            GoTo ContinueTemplate
        End If

        targetTemplatesNode.appendChild targetDoc.importNode(srcTemplateNode, True)

ContinueTemplate:
    Next srcTemplateNode
End Sub


Private Function private_TryResolveItemControlTemplateRoot(ByVal itemControlNode As Object, ByRef outTemplateRoot As Object) As Boolean
    Dim templateName As String
    Dim ownerDoc As Object
    Dim templateNode As Object
    Dim rootNodes As Object

    templateName = VBA.Trim$(ex_XmlCore.m_NodeAttrText(itemControlNode, "objectSourceTemplate"))
    If VBA.Len(templateName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemControl requires non-empty attribute 'objectSourceTemplate'."
#End If
        Exit Function
    End If

    Set ownerDoc = itemControlNode.OwnerDocument
    If ownerDoc Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: failed to resolve owner document for itemControl template."
#End If
        Exit Function
    End If

    Set templateNode = ownerDoc.selectSingleNode( _
        "/p:page/p:templates/p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "] | /p:uiDefinition/p:templates/p:template[@name=" & ex_XmlCore.m_XPathLiteral(templateName) & "]")
    If templateNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemControl references missing template '" & templateName & "'."
#End If
        Exit Function
    End If

    Set rootNodes = templateNode.selectNodes("p:control | p:stackPanel | p:grid | p:list | p:itemControl")
    If rootNodes Is Nothing Or rootNodes.Length = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: template '" & templateName & "' has no visual root node."
#End If
        Exit Function
    End If
    If rootNodes.Length <> 1 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: template '" & templateName & "' must contain exactly one visual root node."
#End If
        Exit Function
    End If

    Set outTemplateRoot = rootNodes.Item(0)
    private_TryResolveItemControlTemplateRoot = True
End Function


Private Sub private_AppendSuffixToControlNames(ByVal rootNode As Object, ByVal suffix As String)
    Dim childNode As Object
    Dim baseName As String
    Dim nodeName As String

    If rootNode Is Nothing Then Exit Sub
    If rootNode.NodeType <> 1 Then Exit Sub

    baseName = VBA.LCase$(VBA.CStr(rootNode.baseName))
    If VBA.StrComp(baseName, "control", VBA.vbBinaryCompare) = 0 Then
        nodeName = VBA.Trim$(ex_XmlCore.m_NodeAttrText(rootNode, "name"))
        If VBA.Len(nodeName) = 0 Then nodeName = "item"
        rootNode.setAttribute "name", nodeName & suffix
    End If

    For Each childNode In rootNode.ChildNodes
        private_AppendSuffixToControlNames childNode, suffix
    Next childNode
End Sub


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
                If Not ex_BindingRuntime.m_TryResolveVisibilityBinding(rawText, dataContext, resolvedVisible) Then Exit Function
                If resolvedVisible Then
                    rootNode.setAttribute attrName, "true"
                Else
                    rootNode.setAttribute attrName, "false"
                End If
                GoTo ContinueAttr
            End If

            ' Здесь binding применяется к текущему dataContext итема.
            ' Для object-результата мы не оставляем сырой объект в XML-атрибуте,
            ' а регистрируем его в RuntimeSources и подставляем ключ.
            If Not ex_BindingRuntime.m_TryResolveValueBinding(rawText, dataContext, resolvedValue) Then Exit Function

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
                            ex_Core.m_Diagnostic_LogError "PrototypeNew: itemsSource binding resolved to Nothing object."
#End If
                            Exit Function
                        End If
                        runtimeItems.Add resolvedObject
                    Else
#If LOGGING_DEBUG_ENABLED Then
                        ex_Core.m_Diagnostic_LogError "PrototypeNew: list itemsSource binding must resolve to Collection."
#End If
                        Exit Function
                    End If

                    runtimeListSourceKey = private_RegisterRuntimeListItemsSourceKey(runtimeItems, renderCtx)
                    If VBA.Len(runtimeListSourceKey) = 0 Then Exit Function
                    ' После этого itemsSource проходит обычный путь через RuntimeSourceResolver.
                    rootNode.setAttribute attrName, runtimeListSourceKey
                ElseIf VBA.StrComp(VBA.LCase$(attrName), "objectsource", VBA.vbBinaryCompare) = 0 And _
                       VBA.StrComp(rootNodeName, "itemcontrol", VBA.vbBinaryCompare) = 0 Then

                    Set resolvedObject = resolvedValue
                    If resolvedObject Is Nothing Then
                        rootNode.setAttribute attrName, VBA.vbNullString
                    Else
                        runtimeObjectSourceKey = private_RegisterRuntimeObjectSourceKey(resolvedObject, renderCtx)
                        If VBA.Len(runtimeObjectSourceKey) = 0 Then Exit Function
                        ' objectSource также преобразуем в runtime key вместо прямой object-ссылки.
                        rootNode.setAttribute attrName, runtimeObjectSourceKey
                    End If
                Else
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.m_Diagnostic_LogError "PrototypeNew: template binding for attribute '" & attrName & "' must resolve to scalar value."
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
        ex_Core.m_Diagnostic_LogError "PrototypeNew: runtime list items source is Nothing."
#End If
        Exit Function
    End If
    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    sourceKey = renderCtx.NextListRuntimeSourceKey()

    If Not renderCtx.PageBase.RuntimeSources.SetItemsSource(sourceKey, items) Then Exit Function
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
        ex_Core.m_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    sourceKey = renderCtx.NextObjectRuntimeSourceKey()

    If Not renderCtx.PageBase.RuntimeSources.SetObjectSource(sourceKey, sourceObject) Then Exit Function
    private_RegisterRuntimeObjectSourceKey = sourceKey
End Function


Private Function private_TryResolveObjectSourceForMeasure( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal rawObjectSource As String, _
    ByVal dataContext As Object, _
    ByRef outObject As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim sourceText As String
    Dim resolvedValue As Variant

    sourceText = VBA.Trim$(rawObjectSource)
    If VBA.Len(sourceText) = 0 Then
        private_TryResolveObjectSourceForMeasure = True
        Exit Function
    End If

    If private_IsBindingExpression(sourceText) Then
        If dataContext Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: itemControl objectSource binding requires item context during layout measurement."
#End If
            Exit Function
        End If

        ' На этапе measure binding тоже должен быть вычислен, иначе span будет неточным.
        If Not ex_BindingRuntime.m_TryResolveValueBinding(sourceText, dataContext, resolvedValue) Then Exit Function

        If VBA.IsObject(resolvedValue) Then
            Set outObject = resolvedValue
            private_TryResolveObjectSourceForMeasure = True
            Exit Function
        End If

        sourceText = VBA.Trim$(VBA.CStr(resolvedValue))
    End If

    If VBA.Len(sourceText) = 0 Then
        private_TryResolveObjectSourceForMeasure = True
        Exit Function
    End If

    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemControl measure render context is not specified."
#End If
        Exit Function
    End If

    Set pageBase = renderCtx.PageBase
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemControl measure page context is not specified in render context."
#End If
        Exit Function
    End If

    ' Если это не binding-объект, резолвим source текст тем же путем, что и в render.
    If Not ex_RuntimeSourceResolver.m_TryResolveObjectSource(pageBase.RuntimeSources, sourceText, outObject, True) Then Exit Function
    private_TryResolveObjectSourceForMeasure = True
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
