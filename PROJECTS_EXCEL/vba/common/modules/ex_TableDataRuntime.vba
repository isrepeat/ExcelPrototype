Option Explicit

Private m_lookupServices As Collection

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    fn_ResetCaches
    Set m_lookupServices = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Register(ByVal lookupService As obj_LookupService)
    If m_lookupServices Is Nothing Then Set m_lookupServices = New Collection
    m_lookupServices.Add lookupService
End Sub

Public Sub fn_Unregister(ByVal lookupService As obj_LookupService)
    Dim index As Long
    Dim registeredService As obj_LookupService

    If m_lookupServices Is Nothing Then Exit Sub
    For index = m_lookupServices.Count To 1 Step -1
        Set registeredService = m_lookupServices(index)
        If registeredService Is lookupService Then m_lookupServices.Remove index
    Next index
End Sub

Public Sub fn_ResetCaches()
    Dim lookupService As obj_LookupService

    If m_lookupServices Is Nothing Then Exit Sub
    For Each lookupService In m_lookupServices
        lookupService.InvalidateCache
    Next lookupService
End Sub

Public Function fn_TryQueryLoaded( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByRef available As Boolean, _
    ByRef result As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim lookupService As obj_LookupService

    available = False
    Set result = Nothing
    fn_TryQueryLoaded = True
    If m_lookupServices Is Nothing Then Exit Function
    For Each lookupService In m_lookupServices
        If Not lookupService.TryQueryLoaded(source, query, available, result, diagnostic) Then
            fn_TryQueryLoaded = False
            Exit Function
        End If
        If available Then Exit Function
    Next lookupService
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------