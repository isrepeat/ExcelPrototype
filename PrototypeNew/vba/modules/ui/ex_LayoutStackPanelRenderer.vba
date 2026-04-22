Attribute VB_Name = "ex_LayoutStackPanelRenderer"
Option Explicit

' Renderer for <stackPanel> nodes.

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
    If layoutNode Is Nothing Then
        VBA.MsgBox "PrototypeNew: stackPanel node is not specified.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(layoutNode.baseName)), "stackpanel", VBA.vbBinaryCompare) <> 0 Then
        VBA.MsgBox "PrototypeNew: ex_LayoutStackPanelRenderer supports only <stackPanel> nodes.", VBA.vbExclamation
        Exit Function
    End If

    m_Render = ex_XmlLayoutEngine.m_RenderContainerNodeInBounds( _
        renderCtx:=renderCtx, _
        containerNode:=layoutNode, _
        layoutRowStart:=rowStart, _
        layoutColStart:=colStart, _
        layoutRowEnd:=rowEnd, _
        layoutColEnd:=colEnd)
End Function
