Option Explicit

Private activeLookupCellAddress As String

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_RegisterRoutes() As Boolean
    Dim inputSheetName As String
    Dim lookupCellAddresses As Collection
    Dim candidateRangeAddress As String
    Dim maxCandidateCount As Long
    Dim lookupCellAddress As Variant

    If Not private_Candidates_TryGetConfig( _
        inputSheetName, lookupCellAddresses, candidateRangeAddress, _
        maxCandidateCount) Then Exit Function
    For Each lookupCellAddress In lookupCellAddresses
        If Not ex_CellChangeRouter.fn_RegisterChangeRoute( _
            inputSheetName, VBA.CStr(lookupCellAddress), _
            "ex_PersonnelCandidates.fn_OnInputChanged") Then Exit Function
    Next lookupCellAddress
    If Not ex_CellChangeRouter.fn_RegisterSelectionRoute( _
        inputSheetName, candidateRangeAddress, _
        "ex_PersonnelCandidates.fn_OnCandidateSelected") Then Exit Function
    activeLookupCellAddress = VBA.vbNullString
    ex_Helpers.LogDebug "Personnel candidate routes registered | Sheet=" & _
        inputSheetName & " | LookupCells=" & _
        VBA.CStr(lookupCellAddresses.Count) & " | CandidateRange=" & _
        candidateRangeAddress
    fn_RegisterRoutes = True
End Function

Public Sub fn_OnInputChanged( _
    ByVal changedSheet As Object, _
    ByVal target As Range _
)
    Dim inputSheetName As String
    Dim lookupCellAddresses As Collection
    Dim candidateRangeAddress As String
    Dim maxCandidateCount As Long
    Dim sourceSheet As Worksheet

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    Set sourceSheet = changedSheet
    If Not private_Candidates_TryGetConfig( _
        inputSheetName, lookupCellAddresses, candidateRangeAddress, _
        maxCandidateCount) Then Exit Sub
    If VBA.StrComp(sourceSheet.Name, inputSheetName, _
            VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "Personnel candidate callback was received from unexpected " & _
            "sheet '" & sourceSheet.Name & "'.", VBA.vbExclamation, _
            "Document Generation"
        Exit Sub
    End If
    If target.CountLarge <> 1 Then
        VBA.MsgBox "Personnel candidate search requires exactly one changed cell.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If

    activeLookupCellAddress = target.Address(False, False)
    ex_Helpers.LogDebug "Personnel candidate search started | Sheet=" & _
        sourceSheet.Name & " | Cell=" & activeLookupCellAddress
    private_Candidates_Refresh sourceSheet, target, candidateRangeAddress, _
        maxCandidateCount
End Sub

Public Sub fn_OnCandidateSelected( _
    ByVal changedSheet As Object, _
    ByVal target As Range _
)
    Dim inputSheetName As String
    Dim lookupCellAddresses As Collection
    Dim candidateRangeAddress As String
    Dim maxCandidateCount As Long
    Dim sourceSheet As Worksheet

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    If target.CountLarge <> 1 Then Exit Sub
    Set sourceSheet = changedSheet
    If Not private_Candidates_TryGetConfig( _
        inputSheetName, lookupCellAddresses, candidateRangeAddress, _
        maxCandidateCount) Then Exit Sub
    If VBA.StrComp(sourceSheet.Name, inputSheetName, _
            VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "Personnel candidate selection was received from unexpected " & _
            "sheet '" & sourceSheet.Name & "'.", VBA.vbExclamation, _
            "Document Generation"
        Exit Sub
    End If
    If VBA.Len(ex_Helpers.private_Text_Normalize( _
        VBA.CStr(target.Value2))) = 0 Then Exit Sub
    If VBA.Len(activeLookupCellAddress) = 0 Then
        VBA.MsgBox "Select a personnel lookup field before choosing a candidate.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If

    private_Candidates_Accept sourceSheet, target, candidateRangeAddress
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Candidates {
' --------------------------------------
Private Function private_Candidates_TryGetConfig( _
    ByRef outInputSheetName As String, _
    ByRef outLookupCellAddresses As Collection, _
    ByRef outCandidateRangeAddress As String, _
    ByRef outMaxCandidateCount As Long _
) As Boolean
    If Not ex_VacationTicketGeneration.fn_TryGetPersonnelCandidatesConfig( _
        outInputSheetName, outLookupCellAddresses, outCandidateRangeAddress, _
        outMaxCandidateCount) Then Exit Function
    If outLookupCellAddresses Is Nothing Or _
       outLookupCellAddresses.Count = 0 Then
        VBA.MsgBox "Vacation personnel lookup cells are not configured.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If VBA.Len(outCandidateRangeAddress) = 0 Or outMaxCandidateCount <= 0 Then
        VBA.MsgBox "Vacation candidate range or maximum count is not configured.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    private_Candidates_TryGetConfig = True
End Function

Private Sub private_Candidates_Refresh( _
    ByVal sourceSheet As Worksheet, _
    ByVal lookupCell As Range, _
    ByVal candidateRangeAddress As String, _
    ByVal maxCandidateCount As Long _
)
    Dim inputText As String
    Dim candidates As Collection
    Dim candidateItem As Variant
    Dim candidateRange As Range
    Dim outputRow As Range

    Set candidateRange = sourceSheet.Range(candidateRangeAddress)
    private_Candidates_Clear candidateRange
    inputText = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(lookupCell.Value2))
    If VBA.Len(inputText) = 0 Then
        ex_Helpers.LogDebug "Personnel candidate search cleared | Sheet=" & _
            sourceSheet.Name & " | Cell=" & activeLookupCellAddress
        activeLookupCellAddress = VBA.vbNullString
        Exit Sub
    End If
    If candidateRange.Columns.Count <> 2 Then
        VBA.MsgBox "Vacation candidate output range must contain exactly two columns: " & _
            "FIO and IPN.", VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    If maxCandidateCount > candidateRange.Rows.Count Then
        VBA.MsgBox "Vacation candidate limit exceeds the configured output rows.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    If Not ex_PersonnelData.ex_TryFindPersonCandidates( _
        inputText, maxCandidateCount, candidates) Then Exit Sub

    ex_Helpers.LogDebug "Personnel candidate search completed | Sheet=" & _
        sourceSheet.Name & " | Cell=" & activeLookupCellAddress & _
        " | Count=" & VBA.CStr(candidates.Count)

    Set outputRow = candidateRange.Rows(1)
    candidateRange.Columns(2).NumberFormat = "@"
    For Each candidateItem In candidates
        If Not VBA.IsArray(candidateItem) Then
            VBA.MsgBox "SHPO candidate has invalid data format.", _
                VBA.vbExclamation, "Document Generation"
            Exit Sub
        End If
        outputRow.Cells(1, 1).Value = VBA.CStr(candidateItem(0))
        outputRow.Cells(1, 2).Value = VBA.CStr(candidateItem(1))
        Set outputRow = outputRow.Offset(1, 0)
    Next candidateItem
End Sub

Private Sub private_Candidates_Accept( _
    ByVal sourceSheet As Worksheet, _
    ByVal candidateCell As Range, _
    ByVal candidateRangeAddress As String _
)
    Dim ipnText As String
    Dim eventsWereEnabled As Boolean
    Dim candidateRange As Range

    Set candidateRange = sourceSheet.Range(candidateRangeAddress)
    If candidateRange.Columns.Count <> 2 Then
        VBA.MsgBox "Vacation candidate output range must contain exactly two columns: " & _
            "FIO and IPN.", VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    ipnText = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Cells(candidateCell.Row, _
            candidateRange.Columns(2).Column).Value2))
    If Not ex_Helpers.private_Text_IsIpn(ipnText) Then
        VBA.MsgBox "Candidate has invalid IPN in " & _
            candidateCell.Address(False, False) & ".", VBA.vbExclamation, _
            "Document Generation"
        Exit Sub
    End If

    eventsWereEnabled = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    sourceSheet.Range(activeLookupCellAddress).Value = ipnText
    private_Candidates_Clear sourceSheet.Range(candidateRangeAddress)
    ex_Helpers.LogDebug "Personnel candidate accepted | Sheet=" & _
        sourceSheet.Name & " | TargetCell=" & activeLookupCellAddress & _
        " | CandidateRow=" & VBA.CStr(candidateCell.Row)
    activeLookupCellAddress = VBA.vbNullString
    Application.EnableEvents = eventsWereEnabled
    Exit Sub
EH:
    Application.EnableEvents = eventsWereEnabled
    ex_Helpers.LogError "Failed to accept personnel candidate | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to accept personnel candidate: " & _
        VBA.Err.Description, VBA.vbExclamation, "Document Generation"
End Sub

Private Sub private_Candidates_Clear(ByVal candidateRange As Range)
    candidateRange.ClearContents
End Sub
' --------------------------------------
' } // namespace Candidates
' --------------------------------------