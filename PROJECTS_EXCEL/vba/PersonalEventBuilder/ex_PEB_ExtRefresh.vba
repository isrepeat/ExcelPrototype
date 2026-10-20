Attribute VB_Name = "ex_PEB_ExtRefresh"
Option Explicit

Private m_owner As obj_PEB_FormExtensions
Private m_scheduled As Boolean
Private m_at As Date

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    If m_scheduled Then
        Application.OnTime m_at, private_Callback(), Schedule:=False
    End If
    m_scheduled = False
    Set m_owner = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Request(ByVal owner As obj_PEB_FormExtensions) As Boolean
    On Error GoTo EH
    If owner Is Nothing Then Exit Function
    Set m_owner = owner
    If Not m_scheduled Then
        m_at = Now
        Application.OnTime m_at, private_Callback()
        m_scheduled = True
    End If
    ex_Core.fn_Diagnostic_WriteLog "EXTENSION_REFRESH_SCHEDULED"
    ex_Core.fn_Diagnostic_Flush True
    fn_Request = True
    Exit Function
EH:
    ex_Core.fn_Diagnostic_WriteError "Extensions.Schedule", Err.Number, Err.Source, Err.Description
End Function

Public Sub fn_Cancel(ByVal owner As obj_PEB_FormExtensions)
    If m_owner Is owner Then fn_Module_Dispose
End Sub

Public Sub fn_Render()
    Dim owner As obj_PEB_FormExtensions
    Dim runtimeContext As Object
    Dim diagnostic As String
    Dim entered As Boolean

    On Error GoTo EH
    Set owner = m_owner
    Set m_owner = Nothing
    m_scheduled = False
    If owner Is Nothing Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    entered = True
    ex_Core.fn_Diagnostic_WriteLog "EXTENSION_REFRESH_STARTED"
    ex_Core.fn_Diagnostic_Flush True
    If Not owner.TryRefresh(diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    ex_Core.fn_Diagnostic_WriteLog "EXTENSION_REFRESH_COMPLETED"
Cleanup:
    If entered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_Flush True
    Exit Sub
EH:
    diagnostic = Err.Description
    ex_Core.fn_Diagnostic_WriteError "Extensions.Refresh", Err.Number, Err.Source, diagnostic
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, ex_PEB_Export.fn_Config("extensions.errorTitle")
    Resume Cleanup
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Callback() As String
    private_Callback = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_PEB_ExtRefresh.fn_Render"
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------