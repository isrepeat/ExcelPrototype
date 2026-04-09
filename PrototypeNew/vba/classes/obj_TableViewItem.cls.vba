VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Model As obj_TableDynamic
Private m_Presentation As obj_Presentation
Private m_Banner As obj_BannerViewItem
Private m_RowItems As Collection

Private Sub Class_Initialize()
    Set m_Model = New obj_TableDynamic
    Set m_Presentation = New obj_Presentation
    Set m_Banner = Nothing
    Set m_RowItems = New Collection
    m_Presentation.PartName = "table"
End Sub

Public Property Get Model() As obj_TableDynamic
    Set Model = m_Model
End Property

Public Property Set Model(ByVal value As obj_TableDynamic)
    If value Is Nothing Then
        Set m_Model = New obj_TableDynamic
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

Public Property Get Banner() As obj_BannerViewItem
    Set Banner = m_Banner
End Property

Public Property Set Banner(ByVal value As obj_BannerViewItem)
    Set m_Banner = value
End Property

Public Property Get RowItems() As Collection
    Set RowItems = m_RowItems
End Property

Public Property Set RowItems(ByVal value As Collection)
    If value Is Nothing Then
        Set m_RowItems = New Collection
    Else
        Set m_RowItems = value
    End If
End Property

Public Property Get ItemVisible() As Boolean
    ItemVisible = m_IsVisible()
End Property

Public Property Let ItemVisible(ByVal value As Boolean)
    m_Presentation.EffectiveVisible = CBool(value)
End Property

Public Property Get RowCount() As Long
    If m_Model Is Nothing Then Exit Property
    RowCount = m_Model.RowCount
End Property

Public Property Get ColumnCount() As Long
    If m_Model Is Nothing Then Exit Property
    ColumnCount = m_Model.ColumnCount
End Property

Public Property Get SectionTitle() As String
    If m_Model Is Nothing Then Exit Property
    SectionTitle = m_Model.SectionTitle
End Property

Public Property Let SectionTitle(ByVal value As String)
    If m_Model Is Nothing Then Set m_Model = New obj_TableDynamic
    m_Model.SectionTitle = CStr(value)
End Property

Public Property Get HeaderText() As String
    If m_Model Is Nothing Then Exit Property
    HeaderText = m_Model.HeaderText
End Property

Public Property Get Rows() As Collection
    If m_Model Is Nothing Then Exit Property
    Set Rows = m_Model.Rows
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
    MsgBox "TableViewItem: direct render is not supported.", vbExclamation
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
