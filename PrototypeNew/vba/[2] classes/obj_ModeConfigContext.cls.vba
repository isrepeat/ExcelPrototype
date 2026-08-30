VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ModeConfigContext"
Option Explicit

Private m_ContextId As String
Private m_ModeId As String
Private m_ProfileId As String
Private m_ConfigTable As obj_ConfigTable
Private m_Revision As Long
Private m_ParentPage As obj_IPage
Private m_OwnerController As obj_PageMainCtrl

' Контекст принадлежит конкретной паре mode/profile. Открытая страница хранит
' ссылку на него и поэтому не зависит от текущего выбора DevConfig на Main.
Public Function Initialize( _
    ByVal contextId As String, _
    ByVal modeId As String, _
    ByVal profileId As String, _
    ByVal configTable As obj_ConfigTable, _
    Optional ByVal parentPage As obj_IPage = Nothing, _
    Optional ByVal ownerController As obj_PageMainCtrl = Nothing _
) As Boolean
    contextId = VBA.Trim$(contextId)
    modeId = VBA.Trim$(modeId)
    profileId = VBA.Trim$(profileId)
    If VBA.Len(contextId) = 0 Or VBA.Len(modeId) = 0 Or VBA.Len(profileId) = 0 Then Exit Function
    If configTable Is Nothing Then Exit Function

    m_ContextId = contextId
    m_ModeId = modeId
    m_ProfileId = profileId
    Set m_ConfigTable = configTable
    Set m_ParentPage = parentPage
    Set m_OwnerController = ownerController
    m_Revision = 1
    Initialize = True
End Function

Public Function ReplaceConfigTable(ByVal configTable As obj_ConfigTable) As Boolean
    If configTable Is Nothing Then Exit Function
    ' Revision меняется только вместе с данными: потребители используют его
    ' для точечной пересборки parser/provider/exporter-кэшей.
    If private_AreConfigTablesEqual(m_ConfigTable, configTable) Then
        ReplaceConfigTable = True
        Exit Function
    End If
    Set m_ConfigTable = configTable
    If m_Revision < 2147483647 Then
        m_Revision = m_Revision + 1
    Else
        m_Revision = 1
    End If
    ReplaceConfigTable = True
End Function

Public Function RefreshFromOwnerIfActive() As Boolean
    ' Main обновляет таблицу из несохранённого DevConfig только когда на нём
    ' всё ещё выбрана та же пара mode/profile.
    If m_OwnerController Is Nothing Then
        RefreshFromOwnerIfActive = True
        Exit Function
    End If
    RefreshFromOwnerIfActive = m_OwnerController.TryRefreshModeConfigContext(Me)
End Function

Public Function BindOwner( _
    ByVal parentPage As obj_IPage, _
    ByVal ownerController As obj_PageMainCtrl _
) As Boolean
    If parentPage Is Nothing Then Exit Function
    If ownerController Is Nothing Then Exit Function
    Set m_ParentPage = parentPage
    Set m_OwnerController = ownerController
    BindOwner = True
End Function

Public Property Get ContextId() As String
    ContextId = m_ContextId
End Property

Public Property Get ModeId() As String
    ModeId = m_ModeId
End Property

Public Property Get ProfileId() As String
    ProfileId = m_ProfileId
End Property

Public Property Get ConfigTable() As obj_ConfigTable
    Set ConfigTable = m_ConfigTable
End Property

Public Property Get Revision() As Long
    Revision = m_Revision
End Property

Public Property Get ParentPage() As obj_IPage
    Set ParentPage = m_ParentPage
End Property

Private Function private_AreConfigTablesEqual( _
    ByVal leftTable As obj_ConfigTable, _
    ByVal rightTable As obj_ConfigTable _
) As Boolean
    Dim leftItems As list__obj_ConfigEntry
    Dim rightItems As list__obj_ConfigEntry
    Dim leftEntry As obj_ConfigEntry
    Dim rightEntry As obj_ConfigEntry
    Dim itemIndex As Long

    If leftTable Is Nothing Or rightTable Is Nothing Then Exit Function
    Set leftItems = leftTable.Items
    Set rightItems = rightTable.Items
    If leftItems Is Nothing Or rightItems Is Nothing Then Exit Function
    If leftItems.Count <> rightItems.Count Then Exit Function

    For itemIndex = 1 To leftItems.Count
        Set leftEntry = leftItems.Item(itemIndex)
        Set rightEntry = rightItems.Item(itemIndex)
        If leftEntry Is Nothing Or rightEntry Is Nothing Then Exit Function
        If VBA.StrComp(leftEntry.Attr, rightEntry.Attr, VBA.vbBinaryCompare) <> 0 Then Exit Function
        If VBA.StrComp(leftEntry.Key, rightEntry.Key, VBA.vbBinaryCompare) <> 0 Then Exit Function
        If VBA.StrComp(leftEntry.Value, rightEntry.Value, VBA.vbBinaryCompare) <> 0 Then Exit Function
    Next itemIndex

    private_AreConfigTablesEqual = True
End Function
