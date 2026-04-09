VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_RowViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Model As obj_Row
Private m_Presentation As obj_Presentation
Private m_Banner As obj_BannerViewItem

Private Sub Class_Initialize()
    Set m_Model = New obj_Row
    Set m_Presentation = New obj_Presentation
    Set m_Banner = Nothing
    m_Presentation.PartName = "row"
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

Public Property Get Banner() As obj_BannerViewItem
    Set Banner = m_Banner
End Property

Public Property Set Banner(ByVal value As obj_BannerViewItem)
    Set m_Banner = value
End Property

Public Property Get RowVisible() As Boolean
    RowVisible = m_IsVisible()
End Property

Public Property Let RowVisible(ByVal value As Boolean)
    m_Presentation.EffectiveVisible = CBool(value)
End Property

Public Property Get SpacerRowsAfter() As Long
    SpacerRowsAfter = m_Presentation.SpacerRowsAfter
End Property

Public Property Let SpacerRowsAfter(ByVal value As Long)
    m_Presentation.SpacerRowsAfter = CLng(value)
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
    MsgBox "RowViewItem: direct render is not supported.", vbExclamation
End Function

Private Function obj_IViewItem_IsVisible() As Boolean
    obj_IViewItem_IsVisible = m_IsVisible()
End Function

' //
' // API
' //
Public Function m_IsVisible() As Boolean
    m_IsVisible = mp_IsVisibleResolved()
End Function

Public Function m_HasBanner() As Boolean
    If m_Banner Is Nothing Then Exit Function
    m_HasBanner = m_Banner.m_IsVisible()
End Function

' //
' // Internal
' //
Private Function mp_IsVisibleResolved() As Boolean
    If m_Presentation Is Nothing Then
        mp_IsVisibleResolved = True
        Exit Function
    End If

    mp_IsVisibleResolved = m_Presentation.EffectiveVisible
End Function
