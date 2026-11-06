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
Public Sub fn_MarkFallbackDates( _
    ByVal table As ListObject, _
    ByVal dateColumn As Long, _
    ByVal notesColumn As Long, _
    ByVal marker As String _
)
    Dim noteCell As Range
    Dim dates As Range
    Dim rowIndex As Long

    If table.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    Set dates = table.ListColumns(dateColumn).DataBodyRange
    For Each noteCell In table.ListColumns(notesColumn).DataBodyRange.Cells
        rowIndex = rowIndex + 1
        If VBA.Left$(VBA.CStr(noteCell.Value2), VBA.Len(marker)) = marker Then
            dates.Cells(rowIndex, 1).Font.Color = VBA.RGB(255, 217, 102)
        End If
    Next noteCell
End Sub

Public Function fn_CreatePage(ByVal pageId As String) As obj_IPage
    If VBA.StrComp(VBA.Trim$(pageId), "PersonnelAtDisposalDaysCalculation", VBA.vbTextCompare) <> 0 Then
        VBA.Err.Raise VBA.vbObjectError + 2211, "ex_PADC.fn_CreatePage", _
            "The page is not registered in this workbook: " & pageId
    End If
    Set fn_CreatePage = New obj_PADC_PgMain
End Function

Public Function fn_Initialize() As Boolean
    Dim profileId As String

    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then
        Exit Function
    End If
    If VBA.StrComp(profileId, "PersonnelAtDisposalDaysCalculation", VBA.vbTextCompare) <> 0 Then
        Exit Function
    End If
    If Not ex_UiPageManager.fn_ShowPage("PersonnelAtDisposalDaysCalculation", profileId) Then
        Exit Function
    End If

    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_COMPLETED | Profile=" & profileId
    fn_Initialize = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------