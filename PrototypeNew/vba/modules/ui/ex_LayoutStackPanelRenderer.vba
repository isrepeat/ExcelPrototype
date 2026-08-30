Attribute VB_Name = "ex_LayoutStackPanelRenderer"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:ex_LayoutStackPanelRenderer.fn_Module_Dispose"
#End If
End Sub

' Renderer for <stackPanel> nodes.

' //
' // API
' //
Public Function fn_Render( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal layoutNode As Object, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    Dim stackDepth As Long
    Dim pageBase As obj_PageBase
    Dim containerName As String

    If layoutNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: stackPanel node is not specified."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(layoutNode.baseName)), "stackpanel", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: ex_LayoutStackPanelRenderer supports only <stackPanel> nodes."
#End If
        Exit Function
    End If

    If Not renderCtx Is Nothing Then
        stackDepth = private_GetStackPanelDepth(layoutNode)
        containerName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(layoutNode, "name")))
        ex_StylePipelineEngine.fn_RegisterLayoutBound _
            renderCtx.Worksheet, rowStart, colStart, rowEnd, colEnd, _
            "stackpanel", containerName, stackDepth

        If VBA.Len(containerName) > 0 Then
            ' Только именованный container доступен публичному partial-reflow API.
            ' Безымянные stackPanel по-прежнему участвуют в retained layout tree,
            ' но однозначно адресовать их из page/controller кода нельзя. Здесь
            ' сохраняются именно фактические bounds последнего render — они
            ' служат исходной областью очистки и точкой привязки нового subtree.
            Set pageBase = renderCtx.Page.GetPageBase()
            If pageBase Is Nothing Then Exit Function
            If Not pageBase.RegisterLayoutContainer( _
                containerName, _
                "stackpanel", _
                renderCtx.Worksheet.Name, _
                rowStart, _
                colStart, _
                rowEnd, _
                colEnd) Then Exit Function
        End If
    End If

    fn_Render = ex_XmlLayoutEngine.fn_RenderContainerNodeInBounds( _
        renderCtx:=renderCtx, _
        containerNode:=layoutNode, _
        layoutRowStart:=rowStart, _
        layoutColStart:=colStart, _
        layoutRowEnd:=rowEnd, _
        layoutColEnd:=colEnd, _
        dataContext:=dataContext)
End Function

Private Function private_GetStackPanelDepth(ByVal stackPanelNode As Object) As Long
    Dim currentNode As Object
    Dim baseName As String

    If stackPanelNode Is Nothing Then Exit Function

    On Error Resume Next
    Set currentNode = stackPanelNode.parentNode
    On Error GoTo 0

    Do While Not currentNode Is Nothing
        On Error Resume Next
        baseName = VBA.LCase$(VBA.Trim$(VBA.CStr(currentNode.baseName)))
        On Error GoTo 0

        If VBA.StrComp(baseName, "stackpanel", VBA.vbBinaryCompare) = 0 Then
            private_GetStackPanelDepth = private_GetStackPanelDepth + 1
        End If

        On Error Resume Next
        Set currentNode = currentNode.parentNode
        On Error GoTo 0
    Loop
End Function
