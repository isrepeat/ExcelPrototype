Attribute VB_Name = "ex_SelectItemsSourceProviders"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_Manager As obj_SIPCacheManager

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_SelectItemsSourceProviders.m_Module_Dispose"
#End If
    On Error Resume Next
    If Not m_Manager Is Nothing Then m_Manager.Dispose
    Set m_Manager = Nothing
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function m_RegisterProvider( _
    ByVal provider As obj_ISelectItemsSourceProvider, _
    Optional ByVal replaceExisting As Boolean = True _
) As Boolean
    ' Тонкий facade над cache manager:
    ' actions/controls работают с одним модулем, без прямой зависимости от класса manager.
    If Not private_TryEnsureManager() Then Exit Function
    m_RegisterProvider = m_Manager.RegisterProvider(provider, replaceExisting)
End Function

Public Function m_HasProvider(ByVal providerKey As String) As Boolean
    If Not private_TryEnsureManager() Then Exit Function
    m_HasProvider = m_Manager.HasProvider(providerKey)
End Function

Public Function m_UnregisterProvider(ByVal providerKey As String) As Boolean
    If Not private_TryEnsureManager() Then Exit Function
    m_UnregisterProvider = m_Manager.UnregisterProvider(providerKey)
End Function

Public Function m_ResetProviderCache(ByVal providerKey As String) As Boolean
    If Not private_TryEnsureManager() Then Exit Function
    m_ResetProviderCache = m_Manager.ResetCacheByProviderKey(providerKey)
End Function

Public Sub m_ResetAllProviderCaches()
    If Not private_TryEnsureManager() Then Exit Sub
    m_Manager.ResetAllCaches
End Sub

Public Function m_TryResolveItemsByProviderKey( _
    ByVal providerKey As String, _
    ByRef outItems As Collection, _
    Optional ByRef outUsedCache As Boolean = False, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    If Not private_TryEnsureManager() Then Exit Function
    m_TryResolveItemsByProviderKey = m_Manager.TryResolveItemsByProviderKey(providerKey, outItems, outUsedCache, allowMissing)
End Function

' //
' // Internal
' //
Private Function private_TryEnsureManager() As Boolean
    If Not m_Manager Is Nothing Then
        private_TryEnsureManager = True
        Exit Function
    End If

    ' Singleton-подобный runtime manager для текущего запуска Excel.
    ' Создается лениво при первом обращении.
    Set m_Manager = New obj_SIPCacheManager
    If m_Manager Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "SelectItemsSourceProviders: failed to create cache manager."
#End If
        Exit Function
    End If
    If Not m_Manager.Initialize Then
        Set m_Manager = Nothing
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "SelectItemsSourceProviders: failed to initialize cache manager."
#End If
        Exit Function
    End If

    private_TryEnsureManager = True
End Function
