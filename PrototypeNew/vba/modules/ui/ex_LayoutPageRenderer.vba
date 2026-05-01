Attribute VB_Name = "ex_LayoutPageRenderer"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_LayoutPageRenderer.m_Module_Dispose"
#End If
End Sub

' Renderer for <page> nodes.

' //
' // API
' //
Public Function m_Render( _
    ByVal renderCtx As obj_LayoutRenderContext, _
    ByVal pageNode As Object _
) As Boolean
    Dim ws As Worksheet
    Dim childNode As Object
    Dim pageAnchorCell As String
    Dim nodeAnchorCell As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim spanRows As Long
    Dim spanColls As Long

    If Not private_TryGetPageWorksheet(renderCtx, ws) Then Exit Function
    If pageNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: page root node is not specified."
#End If
        Exit Function
    End If
    If pageNode.NodeType <> 1 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: page root node must be an element."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(pageNode.baseName)), "page", VBA.vbBinaryCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: expected <page> root node, got '" & VBA.CStr(pageNode.baseName) & "'."
#End If
        Exit Function
    End If

    pageAnchorCell = VBA.Trim$(ex_XmlCore.m_NodeAttrText(pageNode, "anchorCell"))
    If VBA.Len(pageAnchorCell) = 0 Then pageAnchorCell = "A1"

    For Each childNode In pageNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueNode

        Select Case VBA.LCase$(VBA.CStr(childNode.baseName))
            Case "styles", "templates"
                GoTo ContinueNode
        End Select

        If Not ex_XmlLayoutEngine.m_IsVisualLayoutNode(childNode) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PrototypeNew: unsupported node '" & VBA.CStr(childNode.baseName) & "' inside <page>."
#End If
            Exit Function
        End If

        nodeAnchorCell = VBA.Trim$(ex_XmlCore.m_NodeAttrText(childNode, "anchorCell"))
        If VBA.Len(nodeAnchorCell) = 0 Then nodeAnchorCell = pageAnchorCell

        If Not ex_XmlLayoutEngine.m_TryResolveNodeBoundsFromAnchor( _
            renderCtx:=renderCtx, _
            node:=childNode, _
            anchorCellAddr:=nodeAnchorCell, _
            outRow:=rowIndex, _
            outCol:=colIndex, _
            outSpanRows:=spanRows, _
            outSpanColls:=spanColls) Then Exit Function

        If spanRows <= 0 Or spanColls <= 0 Then GoTo ContinueNode

        If Not ex_XmlLayoutEngine.m_RenderNodeBySpan( _
            renderCtx:=renderCtx, _
            layoutNode:=childNode, _
            rowIndex:=rowIndex, _
            colIndex:=colIndex, _
            spanRows:=spanRows, _
            spanColls:=spanColls) Then Exit Function

ContinueNode:
    Next childNode

    m_Render = True
End Function

' //
' // Internal
' //
Private Function private_TryGetPageWorksheet(ByVal renderCtx As obj_LayoutRenderContext, ByRef outWorksheet As Worksheet) As Boolean
    Set outWorksheet = Nothing
    If renderCtx Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: render context is not specified."
#End If
        Exit Function
    End If

    Set outWorksheet = renderCtx.Worksheet
    If outWorksheet Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: worksheet is not specified."
#End If
        Exit Function
    End If

    private_TryGetPageWorksheet = True
End Function
