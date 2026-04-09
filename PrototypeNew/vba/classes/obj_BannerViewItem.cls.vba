VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Model As obj_Banner
Private m_Presentation As obj_Presentation

Private Sub Class_Initialize()
    Set m_Model = New obj_Banner
    Set m_Presentation = New obj_Presentation
    m_Presentation.PartName = "banner"
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

Public Property Get Presentation() As obj_Presentation
    Set Presentation = m_Presentation
End Property

Public Property Set Presentation(ByVal value As obj_Presentation)
    If value Is Nothing Then
        Set m_Presentation = New obj_Presentation
    Else
        Set m_Presentation = value
    End If
End Property

' //
' // Interface
' //
Private Function obj_IViewItem_Render( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal viewName As String = "" _
) As Boolean
    obj_IViewItem_Render = m_Render(ws, rowStart, colStart, rowEnd, colEnd, viewName)
End Function

Private Function obj_IViewItem_IsVisible() As Boolean
    obj_IViewItem_IsVisible = m_IsVisible()
End Function

' //
' // API
' //
Public Function m_Render( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal viewName As String = "" _
) As Boolean
    Dim targetRange As Range
    Dim headerRange As Range
    Dim messageRange As Range
    Dim messageRowStart As Long
    Dim visibleNow As Boolean

    If ws Is Nothing Then
        MsgBox "BannerViewItem: worksheet is not specified.", vbExclamation
        Exit Function
    End If
    If rowStart <= 0 Or colStart <= 0 Then
        MsgBox "BannerViewItem: invalid render start row/column.", vbExclamation
        Exit Function
    End If
    If rowEnd < rowStart Or colEnd < colStart Then
        MsgBox "BannerViewItem: invalid render bounds.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    On Error GoTo 0

    visibleNow = mp_IsVisibleResolved()

    targetRange.UnMerge
    If Not visibleNow Then
        targetRange.ClearContents
        targetRange.Interior.Pattern = xlNone
        targetRange.Borders.LineStyle = xlNone
        m_Render = True
        Exit Function
    End If

    Set headerRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowStart, colEnd))
    headerRange.UnMerge
    headerRange.Merge
    headerRange.Value2 = m_Model.Header

    messageRowStart = rowStart + 1
    If messageRowStart > rowEnd Then messageRowStart = rowEnd
    Set messageRange = ws.Range(ws.Cells(messageRowStart, colStart), ws.Cells(rowEnd, colEnd))
    messageRange.UnMerge
    messageRange.Merge
    messageRange.Value2 = m_Model.Message

    targetRange.Interior.Color = RGB(45, 74, 104)
    targetRange.Borders.LineStyle = xlContinuous
    targetRange.Borders.Color = RGB(26, 43, 61)
    targetRange.Borders.Weight = xlThin

    headerRange.Font.Color = RGB(245, 251, 255)
    headerRange.Font.Bold = True
    headerRange.Font.Size = 11
    headerRange.HorizontalAlignment = xlHAlignLeft
    headerRange.VerticalAlignment = xlVAlignCenter
    headerRange.WrapText = False

    messageRange.Font.Color = RGB(226, 238, 248)
    messageRange.Font.Bold = False
    messageRange.Font.Size = 10
    messageRange.HorizontalAlignment = xlHAlignLeft
    messageRange.VerticalAlignment = xlVAlignTop
    messageRange.WrapText = True

    m_Render = True
    Exit Function

EH_RANGE:
    MsgBox "BannerViewItem: failed to resolve target range for view '" & viewName & "'.", vbExclamation
End Function

Public Function m_IsVisible() As Boolean
    m_IsVisible = mp_IsVisibleResolved()
End Function

' //
' // Internal
' //
Private Function mp_IsVisibleResolved() As Boolean
    If m_Presentation Is Nothing Then
        If m_Model Is Nothing Then
            mp_IsVisibleResolved = False
        Else
            mp_IsVisibleResolved = m_Model.Visible
        End If
        Exit Function
    End If

    If m_Model Is Nothing Then
        mp_IsVisibleResolved = m_Presentation.EffectiveVisible
    Else
        mp_IsVisibleResolved = (m_Model.Visible And m_Presentation.EffectiveVisible)
    End If
End Function