VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerViewItem"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IViewItem

Private Const INLINE_PART_BANNER As String = "banner"

Private m_Banner As obj_Banner
Private m_ViewPresentation As obj_ViewPresentation
Private m_HeaderInlineTextPart As obj_InlineTextPart
Private m_MessageInlineTextPart As obj_InlineTextPart

Public Property Get Model() As obj_Banner
    Set Model = m_Banner
End Property

Public Property Get Presentation() As obj_ViewPresentation
    Set Presentation = m_ViewPresentation
End Property

Public Property Set Presentation(ByVal value As obj_ViewPresentation)
    If value Is Nothing Then
        Set m_ViewPresentation = New obj_ViewPresentation
    Else
        Set m_ViewPresentation = value
    End If
End Property

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_ViewPresentation = New obj_ViewPresentation
    Set m_HeaderInlineTextPart = New obj_InlineTextPart
    Set m_MessageInlineTextPart = New obj_InlineTextPart
    Call Me.Initialize(Nothing)
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

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
Public Function Initialize(ByVal value As obj_Banner) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    If value Is Nothing Then
        Set m_Banner = New obj_Banner
    Else
        Set m_Banner = value
    End If

    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Err.Clear
    Err.Clear
    Set m_Banner = Nothing
    Set m_ViewPresentation = Nothing
    Set m_HeaderInlineTextPart = Nothing
    Set m_MessageInlineTextPart = Nothing
    On Error GoTo 0
End Sub

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
    Dim inlineTextProfile As obj_InlineTextProfile

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "BannerViewItem: page is not specified."
#End If
        Exit Function
    End If

    Set ws = page.Worksheet
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "BannerViewItem: page worksheet is not specified."
#End If
        Exit Function
    End If
    If rowStart <= 0 Or colStart <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "BannerViewItem: invalid render start row/column."
#End If
        Exit Function
    End If
    If rowEnd < rowStart Or colEnd < colStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "BannerViewItem: invalid render bounds."
#End If
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    On Error GoTo 0

    ' Flow: берем единый профиль страницы по partName -> назначаем его частям ->
    ' resolve текста в plain text + runs.
    If Not page.TryGetInlineTextProfile(INLINE_PART_BANNER, inlineTextProfile) Then Exit Function
    Set m_HeaderInlineTextPart.InlineProfile = inlineTextProfile
    Set m_MessageInlineTextPart.InlineProfile = inlineTextProfile

    If Not m_HeaderInlineTextPart.Resolve(m_Banner.Header) Then Exit Function
    If Not m_MessageInlineTextPart.Resolve(m_Banner.Message) Then Exit Function
    headerTextResolved = m_HeaderInlineTextPart.ResolvedText
    messageTextResolved = m_MessageInlineTextPart.ResolvedText

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
    If Not m_HeaderInlineTextPart.RegisterForRange(page, headerRange) Then Exit Function
    If Not m_MessageInlineTextPart.RegisterForRange(page, messageRange) Then Exit Function

    Render = True
    Exit Function

EH_RANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "BannerViewItem: failed to resolve target range for view '" & viewName & "'."
#End If
End Function

Public Function IsVisible() As Boolean
    IsVisible = private_IsVisibleResolved()
End Function

' //
' // Internal
' //
Private Function private_IsVisibleResolved() As Boolean
    If m_ViewPresentation Is Nothing Then
        If m_Banner Is Nothing Then
            private_IsVisibleResolved = False
        Else
            private_IsVisibleResolved = m_Banner.Visible
        End If
        Exit Function
    End If

    If m_Banner Is Nothing Then
        private_IsVisibleResolved = m_ViewPresentation.EffectiveVisible
    Else
        private_IsVisibleResolved = (m_Banner.Visible And m_ViewPresentation.EffectiveVisible)
    End If
End Function

