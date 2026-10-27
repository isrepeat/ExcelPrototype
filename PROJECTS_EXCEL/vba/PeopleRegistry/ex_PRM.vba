Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_PR_Registry.fn_Module_Dispose
    ex_AppHotkeys.fn_Module_Dispose
    ex_UiPageManager.fn_Module_Dispose
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Config(ByVal suffix As String) As String
    Dim key As String

    key = "PeopleRegistry::" & suffix
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, fn_Config) Or Len(fn_Config) = 0 Then
        Err.Raise 5, "ex_PRM.fn_Config", "Required configuration key not found: " & key
    End If
End Function

Public Function fn_CreatePage(ByVal pageId As String) As obj_IPage
    If StrComp(pageId, "PeopleRegistry", vbTextCompare) <> 0 Then
        Err.Raise 5, "ex_PRM.fn_CreatePage", "Unknown page: " & pageId
    End If
    Select Case fn_Config("ui.pageClass")
        Case "obj_PRM_PgMain"
            Set fn_CreatePage = New obj_PRM_PgMain
        Case Else
            Err.Raise 5, "ex_PRM.fn_CreatePage", "Unknown configured page class."
    End Select
End Function

Public Function fn_Initialize() As Boolean
    Dim profileId As String

    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then Exit Function
    If StrComp(profileId, "PeopleRegistry", vbTextCompare) <> 0 Then
        Err.Raise 5, "ex_PRM.fn_Initialize", "Unexpected workbook profile: " & profileId
    End If
    If Not ex_UiPageManager.fn_ShowPage(profileId, profileId) Then Exit Function
    ex_AppCommands.fn_RegisterHotkeys
    If Application.ActiveWorkbook Is ThisWorkbook Then ex_AppHotkeys.fn_Activate
    fn_Initialize = True
End Function

' --------------------------------------
' } // namespace API
' --------------------------------------