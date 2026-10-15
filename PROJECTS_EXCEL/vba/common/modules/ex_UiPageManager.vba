Attribute VB_Name = "ex_UiPageManager"
Option Explicit

#Const ENABLE_LOGGING_VERBOSE = False

Private m_activePage As obj_IPage

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_MANAGER_DISPOSE_STARTED"
#End If
    ex_UiRuntime.fn_Module_Dispose
    If Not m_activePage Is Nothing Then
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "UI_ACTIVE_PAGE_DISPOSE_STARTED"
#End If
        m_activePage.Dispose
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "UI_ACTIVE_PAGE_DISPOSE_COMPLETED"
#End If
    End If
    Set m_activePage = Nothing
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_MANAGER_DISPOSE_COMPLETED"
#End If
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_ShowPage(ByVal pageId As String, ByVal profileId As String) As Boolean
    Dim nextPage As obj_IPage

    Set nextPage = ex_UiPageFactory.fn_Create(pageId)
    If nextPage Is Nothing Then Exit Function
    If Not nextPage.Initialize(profileId) Then
        nextPage.Dispose
        Exit Function
    End If

    ex_UiRuntime.fn_Module_Dispose
    If Not m_activePage Is Nothing Then m_activePage.Dispose
    Set m_activePage = nextPage
    fn_ShowPage = m_activePage.Render()
End Function

Public Function fn_HandleCellChange(ByVal target As Range) As Boolean
    Dim bindingHandled As Boolean
    Dim pageHandled As Boolean

    If target Is Nothing Then Exit Function
    bindingHandled = ex_UiBindings.fn_HandleCellChange(target)
    If Not m_activePage Is Nothing Then pageHandled = m_activePage.HandleCellChange(target)
    fn_HandleCellChange = bindingHandled Or pageHandled
End Function

Public Function fn_RenderPage() As Boolean
    If m_activePage Is Nothing Then Exit Function
    fn_RenderPage = m_activePage.Render()
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------