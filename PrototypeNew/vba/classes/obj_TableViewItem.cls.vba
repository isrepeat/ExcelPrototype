VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Model As obj_TableDynamic
Private m_Presentation As obj_ViewPresentation
Private m_Banner As obj_BannerViewItem
Private m_RowItems As Collection

Private Sub Class_Initialize()
    Set m_Model = New obj_TableDynamic
    Set m_Presentation = New obj_ViewPresentation
    Set m_Banner = Nothing
    Set m_RowItems = New Collection
    Call private_TrySyncRowItemsFromModel()
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

    If Not private_TrySyncRowItemsFromModel() Then
        Set m_RowItems = New Collection
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
    ItemVisible = Me.IsVisible()
End Property

Public Property Let ItemVisible(ByVal value As Boolean)
    m_Presentation.EffectiveVisible = VBA.CBool(value)
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
    m_Model.SectionTitle = VBA.CStr(value)
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
    ByVal page As obj_PageBase, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal viewName As String = "" _
) As Boolean
    VBA.MsgBox "TableViewItem: direct render is not supported.", VBA.vbExclamation
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

Public Function TryResyncRowItemsFromModel() As Boolean
    TryResyncRowItemsFromModel = private_TrySyncRowItemsFromModel()
End Function

' //
' // Internal
' //
Private Function private_TrySyncRowItemsFromModel() As Boolean
    Dim syncedRows As Collection
    Dim sourceRows As Collection
    Dim rowRaw As Variant
    Dim rowModel As obj_Row
    Dim rowView As obj_RowViewItem

    Set syncedRows = New Collection
    If m_Model Is Nothing Then
        Set m_RowItems = syncedRows
        private_TrySyncRowItemsFromModel = True
        Exit Function
    End If

    Set sourceRows = m_Model.Rows
    If sourceRows Is Nothing Then
        Set m_RowItems = syncedRows
        private_TrySyncRowItemsFromModel = True
        Exit Function
    End If

    For Each rowRaw In sourceRows
        If Not VBA.IsObject(rowRaw) Then
            VBA.MsgBox "TableViewItem: model rows must contain obj_Row objects.", VBA.vbExclamation
            Exit Function
        End If
        If VBA.LCase$(VBA.TypeName(rowRaw)) <> "obj_row" Then
            VBA.MsgBox "TableViewItem: unsupported row type '" & VBA.TypeName(rowRaw) & "'. Expected obj_Row.", VBA.vbExclamation
            Exit Function
        End If

        Set rowModel = rowRaw
        Set rowView = New obj_RowViewItem
        Set rowView.Row = rowModel
        rowView.RowVisible = True
        syncedRows.Add rowView
    Next rowRaw

    Set m_RowItems = syncedRows
    private_TrySyncRowItemsFromModel = True
End Function

Private Function private_IsVisibleResolved() As Boolean
    If m_Presentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_Presentation.EffectiveVisible
End Function
