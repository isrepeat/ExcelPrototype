VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_RowViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Model As obj_Row
Private m_Presentation As obj_ViewPresentation
Private m_Banner As obj_BannerViewItem

Private Sub Class_Initialize()
    Set m_Model = New obj_Row
    Set m_Presentation = New obj_ViewPresentation
    Set m_Banner = Nothing
End Sub

Public Property Get Model() As obj_Row
    Set Model = m_Model
End Property

Public Property Set Model(ByVal value As obj_Row)
    If value Is Nothing Then
        Set m_Model = New obj_Row
    Else
        Set m_Model = value
    End If
End Property

Public Property Get Row() As obj_Row
    Set Row = m_Model
End Property

Public Property Set Row(ByVal value As obj_Row)
    Set Model = value
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

Public Property Get Banner() As obj_BannerViewItem
    Set Banner = m_Banner
End Property

Public Property Set Banner(ByVal value As obj_BannerViewItem)
    Set m_Banner = value
End Property

Public Property Get RowVisible() As Boolean
    RowVisible = Me.IsVisible()
End Property

Public Property Let RowVisible(ByVal value As Boolean)
    m_Presentation.EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get SpacerRowsAfter() As Long
    SpacerRowsAfter = m_Presentation.SpacerRowsAfter
End Property

Public Property Let SpacerRowsAfter(ByVal value As Long)
    m_Presentation.SpacerRowsAfter = VBA.CLng(value)
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
    VBA.MsgBox "RowViewItem: direct render is not supported.", VBA.vbExclamation
End Function

Private Function obj_IViewItem_IsVisible() As Boolean
    obj_IViewItem_IsVisible = Me.IsVisible()
End Function

' //
' // API
' //
Public Function IsVisible() As Boolean
    IsVisible = private_IsVisibleResolved()
End Function

Public Function HasBanner() As Boolean
    If m_Banner Is Nothing Then Exit Function
    HasBanner = m_Banner.IsVisible()
End Function

' //
' // Internal
' //
Private Function private_IsVisibleResolved() As Boolean
    If m_Presentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_Presentation.EffectiveVisible
End Function
