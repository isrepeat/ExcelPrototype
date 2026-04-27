Attribute VB_Name = "ex_LayoutGridRenderer"
Option Explicit

' Renderer for <grid> nodes.

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
        ex_Core.m_Diagnostic_LogError "PrototypeNew: grid node is not specified."
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(layoutNode.baseName)), "grid", VBA.vbBinaryCompare) <> 0 Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: ex_LayoutGridRenderer supports only <grid> nodes."
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
