VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerControlVM"
Option Explicit
Implements obj_IControl

Private m_Base As obj_ControlBase
Private m_ControlName As String
Private m_ViewItem As obj_BannerViewItem
Private m_Layout As obj_ControlLayout
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim headerText As String
    Dim messageText As String
    Dim visibleRaw As String
    Dim isVisible As Boolean

    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_ViewItem = Nothing
    Set m_Base = Nothing

    Set m_Base = New obj_ControlBase
    If Not m_Base.Configure(page, controlNode, "Banner", "banner", m_ControlName) Then Exit Sub

    headerText = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "header"))
    messageText = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "message"))
    visibleRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "visible")))

    If VBA.Len(visibleRaw) = 0 Then
        isVisible = (VBA.Len(VBA.Trim$(headerText)) > 0 Or VBA.Len(VBA.Trim$(messageText)) > 0)
    Else
        isVisible = private_ParseBooleanText(visibleRaw)
    End If

    Set m_ViewItem = New obj_BannerViewItem
    m_ViewItem.Model.Header = headerText
    m_ViewItem.Model.Message = messageText
    m_ViewItem.Model.Visible = isVisible
    m_ViewItem.Presentation.EffectiveVisible = isVisible
    m_ViewItem.Presentation.PartName = "banner"

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromNode(controlNode, "Banner", m_ControlName, "style") Then Exit Sub
    m_ViewItem.Presentation.StyleName = m_Layout.StyleName

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "Banner: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set page = Nothing
    If Not m_Base Is Nothing Then Set page = m_Base.PageBase
    If page Is Nothing Then
        VBA.MsgBox "Banner: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_Layout.LayoutSheetName)
    If ws Is Nothing Then
        VBA.MsgBox "Banner: sheet '" & m_Layout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If m_ViewItem Is Nothing Then
        VBA.MsgBox "Banner: view item is not configured for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If Not m_ViewItem.Render(ws, m_Layout.RowStart, m_Layout.ColStart, m_Layout.RowEnd, m_Layout.ColEnd, m_ControlName) Then Exit Sub
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "header", "message", "visible"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //
' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_ParseBooleanText(ByVal rawText As String) As Boolean
    rawText = VBA.LCase$(VBA.Trim$(rawText))
    private_ParseBooleanText = (rawText = "1" Or rawText = "true" Or rawText = "yes")
End Function

Private Function private_GetWorksheetByName(ByVal page As obj_PageBase, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If page Is Nothing Then Exit Function
    Set ws = page.Worksheet
    If ws Is Nothing Then Exit Function

    sheetName = VBA.LCase$(VBA.Trim$(sheetName))
    If VBA.Len(sheetName) > 0 Then
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(ws.Name)), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    End If

    Set private_GetWorksheetByName = ws
End Function
