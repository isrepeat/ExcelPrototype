VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigTableViewItem"
Option Explicit
Implements obj_IViewItem

Private m_Model As obj_ConfigTable
Private m_Presentation As obj_ViewPresentation
Private m_EntryItems As list__obj_ConfigEntryViewItem

Private Sub Class_Initialize()
    Set m_Model = New obj_ConfigTable
    Set m_Presentation = New obj_ViewPresentation
    Set m_EntryItems = New list__obj_ConfigEntryViewItem
    Call private_TrySyncEntryItemsFromModel()
End Sub

Public Property Get Model() As obj_ConfigTable
    Set Model = m_Model
End Property

Public Property Set Model(ByVal value As obj_ConfigTable)
    If value Is Nothing Then
        Set m_Model = New obj_ConfigTable
    Else
        Set m_Model = value
    End If

    If Not private_TrySyncEntryItemsFromModel() Then
        Set m_EntryItems = New list__obj_ConfigEntryViewItem
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

Public Property Get EntryItems() As list__obj_ConfigEntryViewItem
    Set EntryItems = m_EntryItems
End Property

Public Property Set EntryItems(ByVal value As list__obj_ConfigEntryViewItem)
    If value Is Nothing Then
        Set m_EntryItems = New list__obj_ConfigEntryViewItem
    Else
        Set m_EntryItems = value
    End If
End Property

Public Property Get ItemVisible() As Boolean
    ItemVisible = Me.IsVisible()
End Property

Public Property Let ItemVisible(ByVal value As Boolean)
    m_Presentation.EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get Count() As Long
    If m_Model Is Nothing Then Exit Property
    Count = m_Model.Count
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
    Dim syncedEntries As list__obj_ConfigEntryViewItem
    Dim sourceItems As list__obj_ConfigEntry
    Dim entryModel As obj_ConfigEntry
    Dim entryView As obj_ConfigEntryViewItem

    Set syncedEntries = New list__obj_ConfigEntryViewItem
    If m_Model Is Nothing Then
        Set m_EntryItems = syncedEntries
        private_TrySyncEntryItemsFromModel = True
        Exit Function
    End If

    Set sourceItems = m_Model.Items
    If sourceItems Is Nothing Then
        Set m_EntryItems = syncedEntries
        private_TrySyncEntryItemsFromModel = True
        Exit Function
    End If

    For Each entryModel In sourceItems
        Set entryView = New obj_ConfigEntryViewItem
        Set entryView.Model = entryModel
        syncedEntries.Add entryView
    Next entryModel

    Set m_EntryItems = syncedEntries
    private_TrySyncEntryItemsFromModel = True
End Function

Private Function private_IsVisibleResolved() As Boolean
    If m_Presentation Is Nothing Then
        private_IsVisibleResolved = True
        Exit Function
    End If

    private_IsVisibleResolved = m_Presentation.EffectiveVisible
End Function
