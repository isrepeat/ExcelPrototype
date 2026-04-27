VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigTableViewItem"
Option Explicit
Implements obj_IViewItem

Private m_ConfigTable As obj_ConfigTable
Private m_ViewPresentation As obj_ViewPresentation
Private m_ConfigEntryViewItems As list__obj_ConfigEntryViewItem

Private Sub Class_Initialize()
    Set m_ViewPresentation = New obj_ViewPresentation
    Set m_ConfigEntryViewItems = New list__obj_ConfigEntryViewItem
    Call Me.Initialize(Nothing)
End Sub

Public Function Initialize(ByVal value As obj_ConfigTable) As Boolean
    If value Is Nothing Then
        Set m_ConfigTable = New obj_ConfigTable
    Else
        Set m_ConfigTable = value
    End If

    If Not private_TrySyncEntryItemsFromModel() Then
        Set m_ConfigEntryViewItems = New list__obj_ConfigEntryViewItem
    End If

    Initialize = True
End Function

Public Property Get Model() As obj_ConfigTable
    Set Model = m_ConfigTable
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

Public Property Get EntryItems() As list__obj_ConfigEntryViewItem
    Set EntryItems = m_ConfigEntryViewItems
End Property

Public Property Get ItemVisible() As Boolean
    ItemVisible = Me.IsVisible()
End Property

Public Property Let ItemVisible(ByVal value As Boolean)
    m_ViewPresentation.EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get Count() As Long
    If m_ConfigTable Is Nothing Then Exit Property
    Count = m_ConfigTable.Count
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
    VBA.MsgBox "ConfigTableViewItem: direct render is not supported.", VBA.vbExclamation
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

Public Function TryResyncEntryItemsFromModel() As Boolean
    TryResyncEntryItemsFromModel = private_TrySyncEntryItemsFromModel()
End Function

' //
' // Internal
' //
Private Function private_TrySyncEntryItemsFromModel() As Boolean
    Dim configEntryViewItems As list__obj_ConfigEntryViewItem
    Dim sourceConfigEntries As list__obj_ConfigEntry
    Dim sourceConfigEntry As obj_ConfigEntry
    Dim configEntryViewItem As obj_ConfigEntryViewItem

    Set configEntryViewItems = New list__obj_ConfigEntryViewItem
    If m_ConfigTable Is Nothing Then
        Set m_ConfigEntryViewItems = configEntryViewItems
        private_TrySyncEntryItemsFromModel = True
        Exit Function
    End If

    Set sourceConfigEntries = m_ConfigTable.Items
    If sourceConfigEntries Is Nothing Then
        Set m_ConfigEntryViewItems = configEntryViewItems
        private_TrySyncEntryItemsFromModel = True
        Exit Function
    End If

    For Each sourceConfigEntry In sourceConfigEntries
        Set configEntryViewItem = New obj_ConfigEntryViewItem
        If Not configEntryViewItem.Initialize(sourceConfigEntry) Then Exit Function
        configEntryViewItems.Add configEntryViewItem
    Next sourceConfigEntry

    Set m_ConfigEntryViewItems = configEntryViewItems
    private_TrySyncEntryItemsFromModel = True
End Function

Private Function private_IsVisibleResolved() As Boolean
    If m_ViewPresentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_ViewPresentation.EffectiveVisible
End Function
