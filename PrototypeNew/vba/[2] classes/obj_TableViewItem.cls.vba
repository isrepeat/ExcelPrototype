VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableViewItem"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IViewItem

Private m_TableDynamic As obj_TableDynamic
Private m_ViewPresentation As obj_ViewPresentation
Private m_BannerViewItem As obj_BannerViewItem
Private m_RowViewItems As list__obj_RowViewItem

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

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_ViewPresentation = New obj_ViewPresentation
    Set m_BannerViewItem = Nothing
    Set m_RowViewItems = New list__obj_RowViewItem
    Call Me.Initialize(Nothing)
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "TableViewItem: direct render is not supported."
#End If
End Function

Private Function obj_IViewItem_IsVisible() As Boolean
    obj_IViewItem_IsVisible = Me.IsVisible()
End Function

' //
' // API
' //
Public Function Initialize(ByVal value As obj_TableDynamic) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    If value Is Nothing Then
        Set m_TableDynamic = New obj_TableDynamic
    Else
        Set m_TableDynamic = value
    End If

    If Not private_TrySyncRowItemsFromModel() Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "warning: TableViewItem.Initialize: sync row view items from model failed. Fallback to empty list."
#End If
        Set m_RowViewItems = New list__obj_RowViewItem
    End If

    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Err.Clear
    Set m_TableDynamic = Nothing
    Set m_ViewPresentation = Nothing
    Set m_BannerViewItem = Nothing
    Set m_RowViewItems = Nothing
    On Error GoTo 0
End Sub

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
    Dim rowIndex As Long

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

    For rowIndex = 1 To sourceRows.Count
        Set sourceRow = sourceRows.Item(rowIndex)
        If sourceRow Is Nothing Then GoTo ContinueSourceRow

        Set rowViewItem = New obj_RowViewItem
        If Not rowViewItem.Initialize(sourceRow) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableViewItem.private_TrySyncRowItemsFromModel: failed to initialize obj_RowViewItem from source row."
#End If
            Exit Function
        End If
        rowViewItem.RowVisible = True
        rowViewItems.Add rowViewItem
ContinueSourceRow:
    Next rowIndex

    Set m_RowViewItems = rowViewItems
    private_TrySyncRowItemsFromModel = True
    Exit Function

EH_SYNC:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "TableViewItem.private_TrySyncRowItemsFromModel: unexpected error while syncing row view items: " & Err.Description
#End If
End Function

Private Function private_IsVisibleResolved() As Boolean
    If m_ViewPresentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_ViewPresentation.EffectiveVisible
End Function

