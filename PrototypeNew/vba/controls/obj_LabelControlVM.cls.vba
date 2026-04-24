VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_LabelControlVM"
Option Explicit
Implements obj_IControl

Private m_Base As obj_ControlBase
Private m_ControlName As String
Private m_TextRaw As String
Private m_TextResolved As String
Private m_Layout As obj_ControlLayout
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim dataContext As Object

    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_Base = Nothing

    Set m_Base = New obj_ControlBase
    If Not m_Base.Configure(page, controlNode, "Label", "label", m_ControlName) Then Exit Sub

    m_TextRaw = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "text"))
    If VBA.Len(VBA.Trim$(m_TextRaw)) = 0 Then
        m_TextRaw = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "caption"))
    End If

    Set dataContext = m_Base.DataContext
    If dataContext Is Nothing Then Set dataContext = Me
    If Not ex_BindingRuntime.m_TryResolveTextBinding(m_TextRaw, dataContext, m_TextResolved) Then Exit Sub

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromNode(controlNode, "Label", m_ControlName, "style") Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "Label: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set page = Nothing
    If Not m_Base Is Nothing Then Set page = m_Base.PageBase
    If page Is Nothing Then
        VBA.MsgBox "Label: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_Layout.LayoutSheetName)
    If ws Is Nothing Then
        VBA.MsgBox "Label: sheet '" & m_Layout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_Layout.RowStart, m_Layout.ColStart), ws.Cells(m_Layout.RowEnd, m_Layout.ColEnd))
    On Error GoTo 0

    targetRange.Value2 = m_TextResolved
    targetRange.HorizontalAlignment = xlHAlignLeft
    targetRange.VerticalAlignment = xlVAlignCenter
    targetRange.WrapText = False
    If Not private_ApplyPresetStyle(targetRange, m_Layout.StyleName) Then Exit Sub
    Exit Sub

EH_RANGE:
    VBA.MsgBox "Label: failed to resolve target range for control '" & m_ControlName & "'.", VBA.vbExclamation
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "text", "caption"
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
Private Function private_ApplyPresetStyle(ByVal targetRange As Range, ByVal styleName As String) As Boolean
    If targetRange Is Nothing Then Exit Function

    Select Case VBA.LCase$(VBA.Trim$(styleName))
        Case VBA.vbNullString
            ' no-op

        Case "tablesection"
            targetRange.Interior.Color = VBA.RGB(23, 58, 94)
            targetRange.Font.Color = VBA.RGB(234, 246, 255)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(14, 34, 57)
            targetRange.Borders.Weight = xlThin

        Case "tableheadercell"
            targetRange.Interior.Color = VBA.RGB(43, 74, 107)
            targetRange.Font.Color = VBA.RGB(221, 238, 255)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(31, 54, 80)
            targetRange.Borders.Weight = xlThin

        Case "tabledatacell"
            targetRange.Interior.Color = VBA.RGB(58, 58, 58)
            targetRange.Font.Color = VBA.RGB(240, 240, 240)
            targetRange.Font.Bold = False
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(42, 42, 42)
            targetRange.Borders.Weight = xlThin

        Case "tablespacer"
            targetRange.Interior.Color = VBA.RGB(31, 31, 31)
            targetRange.Font.Color = VBA.RGB(31, 31, 31)
            targetRange.Font.Bold = False
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(31, 31, 31)
            targetRange.Borders.Weight = xlHairline

        Case Else
            VBA.MsgBox "Label: unsupported style '" & styleName & "' for control '" & m_ControlName & "'.", VBA.vbExclamation
            Exit Function
    End Select

    private_ApplyPresetStyle = True
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
