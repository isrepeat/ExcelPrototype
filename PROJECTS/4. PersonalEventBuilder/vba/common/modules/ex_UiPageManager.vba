Attribute VB_Name = "ex_UiPageManager"
Option Explicit

Private m_activePage As obj_IPage

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    If Not m_activePage Is Nothing Then m_activePage.Dispose
    Set m_activePage = Nothing
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

    If Not m_activePage Is Nothing Then m_activePage.Dispose
    Set m_activePage = nextPage
    fn_ShowPage = m_activePage.Render()
End Function

Public Function fn_HandleCellChange(ByVal target As Range) As Boolean
    If m_activePage Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function
    fn_HandleCellChange = m_activePage.HandleCellChange(target)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------