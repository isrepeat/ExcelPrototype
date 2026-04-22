VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectControlVMStatic"
Attribute VB_PredeclaredId = True
Option Explicit

Private Const STATE_NS As String = "urn:excelprototype:select-state:v1"
Private Const ROOT_NODE As String = "selectStates"
Private Const ENTRY_NODE As String = "entry"
Private Const KEY_ATTR As String = "key"
Private Const VALUE_ATTR As String = "selectedId"

' //
' // API
' //
Public Function TryGetSelectedId(ByVal selectKey As String, ByRef outSelectedId As String) As Boolean
    selectKey = VBA.LCase$(VBA.Trim$(selectKey))
    If VBA.Len(selectKey) = 0 Then
        VBA.MsgBox "SelectControlVMStatic: select key is empty.", VBA.vbExclamation
        Exit Function
    End If

    TryGetSelectedId = ex_XmlKeyValueStateStore.m_TryGetValue( _
        namespaceUri:=STATE_NS, _
        rootNodeName:=ROOT_NODE, _
        entryNodeName:=ENTRY_NODE, _
        keyAttrName:=KEY_ATTR, _
        keyValue:=selectKey, _
        valueAttrName:=VALUE_ATTR, _
        outValue:=outSelectedId)
End Function

Public Function SetSelectedId(ByVal selectKey As String, ByVal selectedId As String) As Boolean
    selectKey = VBA.LCase$(VBA.Trim$(selectKey))
    selectedId = VBA.Trim$(selectedId)

    If VBA.Len(selectKey) = 0 Then
        VBA.MsgBox "SelectControlVMStatic: select key is empty.", VBA.vbExclamation
        Exit Function
    End If

    SetSelectedId = ex_XmlKeyValueStateStore.m_SetValue( _
        namespaceUri:=STATE_NS, _
        rootNodeName:=ROOT_NODE, _
        entryNodeName:=ENTRY_NODE, _
        keyAttrName:=KEY_ATTR, _
        keyValue:=selectKey, _
        valueAttrName:=VALUE_ATTR, _
        valueText:=selectedId)
End Function

