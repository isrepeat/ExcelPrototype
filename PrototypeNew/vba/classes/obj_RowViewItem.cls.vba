VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_RowViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Row As obj_Row
Private m_ViewPresentation As obj_ViewPresentation
Private m_BannerViewItem As obj_BannerViewItem

Private Sub Class_Initialize()
    Set m_ViewPresentation = New obj_ViewPresentation
    Set m_BannerViewItem = Nothing
    Call Me.Initialize(Nothing)
End Sub

Public Function Initialize(ByVal value As obj_Row) As Boolean
    If value Is Nothing Then
        Set m_Row = New obj_Row
    Else
        Set m_Row = value
    End If

    Initialize = True
End Function

Public Property Get Model() As obj_Row
    Set Model = m_Row
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

Public Property Get Banner() As obj_BannerViewItem
    Set Banner = m_BannerViewItem
End Property

Public Property Set Banner(ByVal value As obj_BannerViewItem)
    Set m_BannerViewItem = value
End Property

Public Property Get RowVisible() As Boolean
    RowVisible = Me.IsVisible()
End Property

Public Property Let RowVisible(ByVal value As Boolean)
    m_ViewPresentation.EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get SpacerRowsAfter() As Long
    SpacerRowsAfter = m_ViewPresentation.SpacerRowsAfter
End Property

Public Property Let SpacerRowsAfter(ByVal value As Long)
    m_ViewPresentation.SpacerRowsAfter = VBA.CLng(value)
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
    ex_Core.m_Diagnostic_LogError "RowViewItem: direct render is not supported."
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
    If m_BannerViewItem Is Nothing Then Exit Function
    HasBanner = m_BannerViewItem.IsVisible()
End Function

' //
' // Internal
' //
Private Function private_IsVisibleResolved() As Boolean
    If m_ViewPresentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_ViewPresentation.EffectiveVisible
End Function
