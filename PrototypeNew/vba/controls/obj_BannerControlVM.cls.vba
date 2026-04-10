VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerControlVM"
Option Explicit
Implements obj_IControl

Private m_ControlName As String
Private m_ViewItem As obj_BannerViewItem
Private m_Layout As obj_ControlLayout
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim headerText As String
    Dim messageText As String
    Dim visibleRaw As String
    Dim isVisible As Boolean

    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_ViewItem = Nothing

    If controlNode Is Nothing Then
        MsgBox "Banner: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "banner"

    headerText = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "header"))
    messageText = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "message"))
    visibleRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "visible")))

    If Len(visibleRaw) = 0 Then
        isVisible = (Len(Trim$(headerText)) > 0 Or Len(Trim$(messageText)) > 0)
    Else
        isVisible = mp_ParseBooleanText(visibleRaw)
    End If

    Set m_ViewItem = New obj_BannerViewItem
    m_ViewItem.Model.Header = headerText
    m_ViewItem.Model.Message = messageText
    m_ViewItem.Model.Visible = isVisible
    m_ViewItem.Presentation.EffectiveVisible = isVisible
    m_ViewItem.Presentation.PartName = "banner"

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.m_TryReadFromNode(controlNode, "Banner", m_ControlName, "style") Then Exit Sub
    m_ViewItem.Presentation.StyleName = m_Layout.StyleName

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet

    If Not m_IsConfigured Then
        MsgBox "Banner: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Banner: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_Layout.LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Banner: sheet '" & m_Layout.LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_ViewItem Is Nothing Then
        MsgBox "Banner: view item is not configured for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not m_ViewItem.m_Render(ws, m_Layout.RowStart, m_Layout.ColStart, m_Layout.RowEnd, m_Layout.ColEnd, m_ControlName) Then Exit Sub
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "header", "message", "visible"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // Internal
' //
Private Function mp_ParseBooleanText(ByVal rawText As String) As Boolean
    rawText = LCase$(Trim$(rawText))
    mp_ParseBooleanText = (rawText = "1" Or rawText = "true" Or rawText = "yes")
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
