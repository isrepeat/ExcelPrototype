Attribute VB_Name = "ex_TDB_Candidates"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_IsBrowse(ByVal text As String) As Boolean
    fn_IsBrowse = (ex_TableQuery.fn_NormalizeText(text) = ex_TravelDocumentsBuilder.fn_Config("candidates.browse.token"))
End Function

Public Function fn_BrowseLimit(ByVal text As String) As Long
    If fn_IsBrowse(text) Then
        fn_BrowseLimit = VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("candidates.browse.limit"))
        If fn_BrowseLimit < 1 Then
            VBA.Err.Raise 5, "Candidates.Browse", ex_TravelDocumentsBuilder.fn_Config("candidates.browse.text.InvalidLimit")
        End If
    End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------