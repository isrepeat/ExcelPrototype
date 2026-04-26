VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerViewItem"
Option Explicit
Implements obj_IViewItem

Private Const INLINE_PART_BANNER As String = "banner"

Private m_Model As obj_Banner
Private m_Presentation As obj_ViewPresentation
Private m_HeaderInlinePart As obj_InlineTextPart
Private m_MessageInlinePart As obj_InlineTextPart

Private Sub Class_Initialize()
    Set m_Model = New obj_Banner
    Set m_Presentation = New obj_ViewPresentation
    Set m_HeaderInlinePart = New obj_InlineTextPart
    Set m_MessageInlinePart = New obj_InlineTextPart
End Sub

Public Property Get Model() As obj_Banner
    Set Model = m_Model
End Property

Public Property Set Model(ByVal value As obj_Banner)
    If value Is Nothing Then
        Set m_Model = New obj_Banner
    Else
        Set m_Model = value
    End If
End Property

Public Property Get Presentation() As obj_ViewPresentation
    Set Presentation = m_Presentation
End Property

Public Property Set Presentation(ByVal value As obj_ViewPresentation)
    If value Is Nothing Then
        Set m_Presentation = New obj_ViewPresentation
    Else
        Set m_Presentation = value
    End If
End Property

' //
' // Interface
' //
Private Function obj_IViewItem_Render( _
    ByVal page As obj_PageBase, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal viewName As String = "" _
) As Boolean
    obj_IViewItem_Render = Me.Render(page, rowStart, colStart, rowEnd, colEnd, viewName)
End Function

Private Function obj_IViewItem_IsVisible() As Boolean
    obj_IViewItem_IsVisible = Me.IsVisible()
End Function

' //
' // API
' //
Public Function Render( _
    ByVal page As obj_PageBase, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal viewName As String = "" _
) As Boolean
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim headerRange As Range
    Dim messageRange As Range
    Dim messageRowStart As Long
    Dim visibleNow As Boolean
    Dim headerTextResolved As String
    Dim messageTextResolved As String
    Dim inlineProfile As obj_InlineTextProfile

    If page Is Nothing Then
        VBA.MsgBox "BannerViewItem: page is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set ws = page.Worksheet
    If ws Is Nothing Then
        VBA.MsgBox "BannerViewItem: page worksheet is not specified.", VBA.vbExclamation
        Exit Function
    End If
    If rowStart <= 0 Or colStart <= 0 Then
        VBA.MsgBox "BannerViewItem: invalid render start row/column.", VBA.vbExclamation
        Exit Function
    End If
    If rowEnd < rowStart Or colEnd < colStart Then
        VBA.MsgBox "BannerViewItem: invalid render bounds.", VBA.vbExclamation
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    On Error GoTo 0

    ' Flow: берем единый профиль страницы по partName -> назначаем его частям ->
    ' resolve текста в plain text + runs.
    If Not page.TryGetInlineTextProfile(INLINE_PART_BANNER, inlineProfile) Then Exit Function
    Set m_HeaderInlinePart.InlineProfile = inlineProfile
    Set m_MessageInlinePart.InlineProfile = inlineProfile

    If Not m_HeaderInlinePart.Resolve(m_Model.Header) Then Exit Function
    If Not m_MessageInlinePart.Resolve(m_Model.Message) Then Exit Function
    headerTextResolved = m_HeaderInlinePart.ResolvedText
    messageTextResolved = m_MessageInlinePart.ResolvedText

    visibleNow = private_IsVisibleResolved()

    targetRange.UnMerge
    If Not visibleNow Then
        targetRange.ClearContents
        targetRange.Interior.Pattern = xlNone
        targetRange.Borders.LineStyle = xlNone
        Render = True
        Exit Function
    End If

    Set headerRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowStart, colEnd))
    headerRange.UnMerge
    headerRange.Merge
    headerRange.Value2 = headerTextResolved

    messageRowStart = rowStart + 1
    If messageRowStart > rowEnd Then messageRowStart = rowEnd
    Set messageRange = ws.Range(ws.Cells(messageRowStart, colStart), ws.Cells(rowEnd, colEnd))
    messageRange.UnMerge
    messageRange.Merge
    messageRange.Value2 = messageTextResolved

    targetRange.Interior.Color = VBA.RGB(45, 74, 104)
    targetRange.Borders.LineStyle = xlContinuous
    targetRange.Borders.Color = VBA.RGB(26, 43, 61)
    targetRange.Borders.Weight = xlThin

    headerRange.Font.Color = VBA.RGB(245, 251, 255)
    headerRange.Font.Bold = False
    headerRange.Font.Size = 11
    headerRange.HorizontalAlignment = xlHAlignLeft
    headerRange.VerticalAlignment = xlVAlignCenter
    headerRange.WrapText = False

    messageRange.Font.Color = VBA.RGB(226, 238, 248)
    messageRange.Font.Bold = False
    messageRange.Font.Size = 10
    messageRange.HorizontalAlignment = xlHAlignLeft
    messageRange.VerticalAlignment = xlVAlignTop
    messageRange.WrapText = True

    ' Регистрируем runs; фактическое применение централизованно делает PageBase.ApplyInlineRuns.
    If Not m_HeaderInlinePart.RegisterForRange(page, headerRange) Then Exit Function
    If Not m_MessageInlinePart.RegisterForRange(page, messageRange) Then Exit Function

    Render = True
    Exit Function

EH_RANGE:
    VBA.MsgBox "BannerViewItem: failed to resolve target range for view '" & viewName & "'.", VBA.vbExclamation
End Function

Public Function IsVisible() As Boolean
    IsVisible = private_IsVisibleResolved()
End Function

' //
' // Internal
' //
Private Function private_IsVisibleResolved() As Boolean
    If m_Presentation Is Nothing Then
        If m_Model Is Nothing Then
            private_IsVisibleResolved = False
        Else
            private_IsVisibleResolved = m_Model.Visible
        End If
        Exit Function
    End If

    If m_Model Is Nothing Then
        private_IsVisibleResolved = m_Presentation.EffectiveVisible
    Else
        private_IsVisibleResolved = (m_Model.Visible And m_Presentation.EffectiveVisible)
    End If
End Function
