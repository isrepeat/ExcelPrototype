VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectControlVMStatic"
Attribute VB_PredeclaredId = True
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private Const STATE_NS As String = "urn:excelprototype:select-state:v1"
Private Const ROOT_NODE As String = "selectStates"
Private Const ENTRY_NODE As String = "entry"
Private Const KEY_ATTR As String = "key"
Private Const VALUE_ATTR As String = "selectedId"

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
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
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    On Error GoTo 0
End Sub

Public Function TryGetSelectedId(ByVal selectKey As String, ByRef outSelectedId As String) As Boolean
    selectKey = VBA.LCase$(VBA.Trim$(selectKey))
    If VBA.Len(selectKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectControlVMStatic: select key is empty."
#End If
        Exit Function
    End If

    TryGetSelectedId = ex_XmlKeyValueStateStore.fn_TryGetValue( _
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
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectControlVMStatic: select key is empty."
#End If
        Exit Function
    End If

    SetSelectedId = ex_XmlKeyValueStateStore.fn_SetValue( _
        namespaceUri:=STATE_NS, _
        rootNodeName:=ROOT_NODE, _
        entryNodeName:=ENTRY_NODE, _
        keyAttrName:=KEY_ATTR, _
        keyValue:=selectKey, _
        valueAttrName:=VALUE_ATTR, _
        valueText:=selectedId)
End Function

