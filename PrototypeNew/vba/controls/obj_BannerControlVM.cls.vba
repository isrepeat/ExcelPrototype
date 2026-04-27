VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerControlVM"
Option Explicit
Implements obj_IControl

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_BannerViewItem As obj_BannerViewItem
Private m_ControlLayout As obj_ControlLayout
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
    Set m_ControlLayout = Nothing
    Set m_BannerViewItem = Nothing
    Set m_ControlBase = Nothing

    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Configure(page, controlNode, "Banner", "banner", m_ControlName) Then Exit Sub

    headerText = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "header"))
    messageText = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "message"))
    visibleRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "visible")))

    If VBA.Len(visibleRaw) = 0 Then
        isVisible = (VBA.Len(VBA.Trim$(headerText)) > 0 Or VBA.Len(VBA.Trim$(messageText)) > 0)
    Else
        isVisible = private_ParseBooleanText(visibleRaw)
    End If

    Set m_BannerViewItem = New obj_BannerViewItem
    m_BannerViewItem.Model.Header = headerText
    m_BannerViewItem.Model.Message = messageText
    m_BannerViewItem.Model.Visible = isVisible
    m_BannerViewItem.Presentation.EffectiveVisible = isVisible

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Banner", m_ControlName, "style") Then Exit Sub
    m_BannerViewItem.Presentation.StyleName = m_ControlLayout.StyleName

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim pageBase As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "Banner: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
        VBA.MsgBox "Banner: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
        VBA.MsgBox "Banner: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If m_BannerViewItem Is Nothing Then
        VBA.MsgBox "Banner: view item is not configured for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If Not m_BannerViewItem.Render(pageBase, m_ControlLayout.RowStart, m_ControlLayout.ColStart, m_ControlLayout.RowEnd, m_ControlLayout.ColEnd, m_ControlName) Then Exit Sub
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
