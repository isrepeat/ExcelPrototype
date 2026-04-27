VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableViewItem"
Option Explicit
Implements obj_IViewItem

Private m_TableDynamic As obj_TableDynamic
Private m_ViewPresentation As obj_ViewPresentation
Private m_BannerViewItem As obj_BannerViewItem
Private m_RowViewItems As list__obj_RowViewItem

Private Sub Class_Initialize()
    Set m_ViewPresentation = New obj_ViewPresentation
    Set m_BannerViewItem = Nothing
    Set m_RowViewItems = New list__obj_RowViewItem
    Call Me.Initialize(Nothing)
End Sub

Public Function Initialize(ByVal value As obj_TableDynamic) As Boolean
    If value Is Nothing Then
        Set m_TableDynamic = New obj_TableDynamic
    Else
        Set m_TableDynamic = value
    End If

    If Not private_TrySyncRowItemsFromModel() Then
        ex_Core.m_Diagnostic_LogInfo "warning: TableViewItem.Initialize: sync row view items from model failed. Fallback to empty list."
        Set m_RowViewItems = New list__obj_RowViewItem
    End If

    Initialize = True
End Function

Public Property Get Model() As obj_TableDynamic
    Set Model = m_TableDynamic
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

Public Property Get RowItems() As list__obj_RowViewItem
    Set RowItems = m_RowViewItems
End Property

Public Property Get ItemVisible() As Boolean
    ItemVisible = Me.IsVisible()
End Property

Public Property Let ItemVisible(ByVal value As Boolean)
    m_ViewPresentation.EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get RowCount() As Long
    If m_TableDynamic Is Nothing Then Exit Property
    RowCount = m_TableDynamic.RowCount
End Property

Public Property Get ColumnCount() As Long
    If m_TableDynamic Is Nothing Then Exit Property
    ColumnCount = m_TableDynamic.ColumnCount
End Property

Public Property Get SectionTitle() As String
    If m_TableDynamic Is Nothing Then Exit Property
    SectionTitle = m_TableDynamic.SectionTitle
End Property

Public Property Let SectionTitle(ByVal value As String)
    If m_TableDynamic Is Nothing Then Set m_TableDynamic = New obj_TableDynamic
    m_TableDynamic.SectionTitle = VBA.CStr(value)
End Property

Public Property Get HeaderText() As String
    If m_TableDynamic Is Nothing Then Exit Property
    HeaderText = m_TableDynamic.HeaderText
End Property

Public Property Get Rows() As list__obj_Row
    If m_TableDynamic Is Nothing Then Exit Property
    Set Rows = m_TableDynamic.Rows
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
    ex_Core.m_Diagnostic_LogError "TableViewItem: direct render is not supported."
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

Public Function TryResyncRowItemsFromModel() As Boolean
    TryResyncRowItemsFromModel = private_TrySyncRowItemsFromModel()
End Function

' //
' // Internal
' //
Private Function private_TrySyncRowItemsFromModel() As Boolean
    Dim sourceRow As obj_Row
    Dim sourceRows As list__obj_Row

    Dim rowViewItem As obj_RowViewItem
    Dim rowViewItems As list__obj_RowViewItem

    On Error GoTo EH_SYNC

    Set rowViewItems = New list__obj_RowViewItem
    If m_TableDynamic Is Nothing Then
        Set m_RowViewItems = rowViewItems
        private_TrySyncRowItemsFromModel = True
        Exit Function
    End If

    Set sourceRows = m_TableDynamic.Rows
    If sourceRows Is Nothing Then
        Set m_RowViewItems = rowViewItems
        private_TrySyncRowItemsFromModel = True
        Exit Function
    End If

    For Each sourceRow In sourceRows
        Set rowViewItem = New obj_RowViewItem
        If Not rowViewItem.Initialize(sourceRow) Then
            ex_Core.m_Diagnostic_LogError "TableViewItem.private_TrySyncRowItemsFromModel: failed to initialize obj_RowViewItem from source row."
            Exit Function
        End If
        rowViewItem.RowVisible = True
        rowViewItems.Add rowViewItem
    Next sourceRow

    Set m_RowViewItems = rowViewItems
    private_TrySyncRowItemsFromModel = True
    Exit Function

EH_SYNC:
    ex_Core.m_Diagnostic_LogError "TableViewItem.private_TrySyncRowItemsFromModel: unexpected error while syncing row view items: " & Err.Description
End Function

Private Function private_IsVisibleResolved() As Boolean
    If m_ViewPresentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_ViewPresentation.EffectiveVisible
End Function
