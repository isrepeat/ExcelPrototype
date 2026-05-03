VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SIP_StaticList"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Implements obj_ISelectItemsSourceProvider

Private m_ProviderKey As String
Private m_Items As Collection
Private m_Stamp As String

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_ISelectItemsSourceProvider_GetProviderKey() As String
    obj_ISelectItemsSourceProvider_GetProviderKey = m_ProviderKey
End Function

Private Function obj_ISelectItemsSourceProvider_TryGetCurrentStamp(ByRef outStamp As String) As Boolean
    ' Для статического списка stamp фиксируется в Initialize().
    ' Пока items не переинициализировали — всегда cache-hit.
    outStamp = m_Stamp
    obj_ISelectItemsSourceProvider_TryGetCurrentStamp = (VBA.Len(VBA.Trim$(outStamp)) > 0)
End Function

Private Function obj_ISelectItemsSourceProvider_TryBuildItems(ByRef outItems As Collection) As Boolean
    obj_ISelectItemsSourceProvider_TryBuildItems = private_TryCloneItems(m_Items, outItems)
End Function

' //
' // API
' //
Public Function Initialize(ByVal providerKey As String, ByVal items As Collection) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    providerKey = VBA.LCase$(VBA.Trim$(providerKey))

    If VBA.Len(providerKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: static select source provider key is empty."
#End If
        Exit Function
    End If
    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: static select source items are not specified for key '" & providerKey & "'."
#End If
        Exit Function
    End If

    If Not private_TryCloneItems(items, m_Items) Then Exit Function
    m_ProviderKey = providerKey
    m_Stamp = private_BuildStampByItems(m_Items)
    If VBA.Len(VBA.Trim$(m_Stamp)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: failed to build static source stamp for key '" & providerKey & "'."
#End If
        Exit Function
    End If

    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    m_ProviderKey = VBA.vbNullString
    m_Stamp = VBA.vbNullString
    Set m_Items = Nothing
    On Error GoTo 0
End Sub

Public Function Configure(ByVal providerKey As String, ByVal items As Collection) As Boolean
    ' Backward-compatible wrapper.
    Configure = Initialize(providerKey, items)
End Function

' //
' // Internal
' //
Private Function private_TryCloneItems(ByVal sourceItems As Collection, ByRef outItems As Collection) As Boolean
    Dim rawItem As Variant

    Set outItems = Nothing
    If sourceItems Is Nothing Then Exit Function

    Set outItems = New Collection
    For Each rawItem In sourceItems
        If VBA.IsObject(rawItem) Then
            outItems.Add rawItem
        Else
            outItems.Add rawItem
        End If
    Next rawItem

    private_TryCloneItems = True
End Function

Private Function private_BuildStampByItems(ByVal sourceItems As Collection) As String
    Dim rawItem As Variant
    Dim itemSignature As String
    Dim itemsSignature As String
    Dim itemCount As Long

    If sourceItems Is Nothing Then Exit Function

    ' Формируем deterministic signature из элементов списка.
    For Each rawItem In sourceItems
        itemCount = itemCount + 1

        If VBA.IsObject(rawItem) Then
            itemSignature = VBA.TypeName(rawItem)
            On Error Resume Next
            itemSignature = itemSignature & "|" & VBA.CStr(CallByName(rawItem, "Id", VbGet))
            itemSignature = itemSignature & "|" & VBA.CStr(CallByName(rawItem, "Caption", VbGet))
            On Error GoTo 0
        Else
            itemSignature = VBA.CStr(rawItem)
        End If

        itemsSignature = itemsSignature & "#" & itemSignature
    Next rawItem

    private_BuildStampByItems = "static|" & VBA.CStr(itemCount) & "|" & itemsSignature
End Function
