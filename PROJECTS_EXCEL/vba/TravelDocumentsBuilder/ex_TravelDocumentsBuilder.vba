Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_UiPageManager.fn_Module_Dispose
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_CreatePage(ByVal pageId As String) As obj_IPage
    If VBA.StrComp(VBA.Trim$(pageId), "TravelDocumentsBuilder", VBA.vbTextCompare) <> 0 Then
        VBA.Err.Raise VBA.vbObjectError + 2230, "ex_TravelDocumentsBuilder.fn_CreatePage", _
            "The page is not registered in this workbook: " & pageId
    End If
    Set fn_CreatePage = New obj_TDB_PgMain
End Function

Public Function fn_Initialize() As Boolean
    Dim profileId As String

    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then Exit Function
    fn_Initialize = ex_UiPageManager.fn_ShowPage("TravelDocumentsBuilder", profileId)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------