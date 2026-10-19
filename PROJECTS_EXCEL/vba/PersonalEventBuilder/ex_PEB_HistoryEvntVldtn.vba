Attribute VB_Name = "ex_PEB_HistoryEvntVldtn"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Validate( _
    ByVal current As Object, _
    ByVal previous As Object, _
    ByRef diagnostic As String, _
    ByVal historyIndex As Object _
) As Boolean
    Dim snapshot As Object
    Dim state As Object
    Dim statuses As Object
    Dim eventPlan As Object
    Dim anchor As Object
    Dim ordered As New Collection
    Dim stored As Object
    Dim personId As String
    Dim enabled As String
    Dim boundaryDate As Double
    Dim boundaryOrder As Double
    Dim eventDate As Double
    Dim startedAt As Double
    Dim stageAt As Double

    On Error GoTo EH
    startedAt = VBA.Timer
    diagnostic = VBA.vbNullString
    enabled = ex_PEB_Export.fn_Config("validation.Movement.enabled")
    If enabled = "false" Or current("Status") = "Cancelled" Then
        fn_Validate = True
        GoTo Cleanup
    End If
    If enabled <> "true" Then
        ex_PEB_MovementLinks.fn_Fail "InvalidValidationEnabledSetting", enabled
    End If
    If Not ex_PEB_MovementRegistry.fn_TryEvent(current("EventId"), eventPlan) Then
        fn_Validate = True
        GoTo Cleanup
    End If
    personId = current("PersonId")
    Set statuses = ex_PEB_MovementRegistry.fn_Statuses()
    Set state = VBA.CreateObject("Scripting.Dictionary")
    state("IsOpen") = False
    state("Type") = VBA.vbNullString
    stageAt = VBA.Timer
    Set snapshot = ex_PEB_MovementSource.fn_ReadLatest(personId)
    If Not snapshot Is Nothing Then
        state("Type") = ex_PEB_Export.fn_Text(snapshot("EventName"))
        boundaryDate = ex_PEB_MovementSource.fn_OrderDate(snapshot("ArrivalDate"))
        state("IsOpen") = (VBA.Len(ex_PEB_Export.fn_Text(snapshot("ArrivalDate"))) = 0)
        If state("IsOpen") Then
            boundaryDate = ex_PEB_MovementSource.fn_OrderDate(snapshot("DepartureDate"))
        End If
    End If
    ex_Core.fn_Diagnostic_WritePerf "Movement.SaveValidation.ReadLatest", stageAt
    stageAt = VBA.Timer
    If historyIndex("ByPerson").Exists(personId) Then
        ' Locate the represented state boundary; no hashes or target links are read.
        For Each stored In historyIndex("ByPerson")(personId)
            If ex_PEB_MovementSource.fn_OrderDate(stored("EventDate")) = boundaryDate Then
                If ex_PEB_MovementRegistry.fn_TryEvent(stored("EventId"), eventPlan) Then
                    If private_IsBoundary(state, eventPlan, statuses) Then
                        If anchor Is Nothing Or VBA.Val(stored("RowOrder")) >= boundaryOrder Then
                            Set anchor = stored
                            boundaryOrder = VBA.Val(stored("RowOrder"))
                        End If
                    End If
                End If
            End If
        Next stored
        If Not anchor Is Nothing Then
            If anchor("Status") = "Cancelled" Or anchor("RowId") = current("RowId") Then
                Set eventPlan = ex_PEB_MovementRegistry.fn_Event(anchor("EventId"))
                state("IsOpen") = (eventPlan("operation") <> "Departure")
                If state("IsOpen") Then
                    state("Type") = statuses(eventPlan("sourceStatus"))
                End If
            End If
        End If
        For Each stored In historyIndex("ByPerson")(personId)
            If stored("RowId") <> current("RowId") And stored("Status") <> "Cancelled" Then
                eventDate = ex_PEB_MovementSource.fn_OrderDate(stored("EventDate"))
                If eventDate > boundaryDate Or (eventDate = boundaryDate And VBA.Val(stored("RowOrder")) > boundaryOrder) Then
                    If private_Before(stored, current) Then
                        If ex_PEB_MovementRegistry.fn_TryEvent(stored("EventId"), eventPlan) Then
                            private_Insert ordered, stored
                        End If
                    End If
                End If
            End If
        Next stored
    End If
    For Each stored In ordered
        Set eventPlan = ex_PEB_MovementRegistry.fn_Event(stored("EventId"))
        private_Apply state, eventPlan, statuses
    Next stored
    Set stored = current
    Set eventPlan = ex_PEB_MovementRegistry.fn_Event(current("EventId"))
    private_Apply state, eventPlan, statuses
    ex_Core.fn_Diagnostic_WritePerf "Movement.SaveValidation.Rules", stageAt
    fn_Validate = True
Cleanup:
    ex_Core.fn_Diagnostic_WritePerf "Movement.SaveValidation.Total", startedAt
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    If Not stored Is Nothing Then
        diagnostic = VBA.Replace$(VBA.Replace$(ex_PEB_Export.fn_Config("text.MovementSaveValidationFailed"), _
            "{id}", stored("RowId")), "{details}", diagnostic)
    End If
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_IsBoundary( _
    ByVal state As Object, _
    ByVal eventPlan As Object, _
    ByVal statuses As Object _
) As Boolean
    Dim statusId As String

    If state("IsOpen") Then
        statusId = eventPlan("targetStatus")
    ElseIf eventPlan("operation") = "Arrival" Then
        statusId = eventPlan("sourceStatus")
    End If
    If VBA.Len(statusId) > 0 Then
        private_IsBoundary = (VBA.StrComp(state("Type"), statuses(statusId), VBA.vbTextCompare) = 0)
    End If
End Function

Private Sub private_Apply( _
    ByVal state As Object, _
    ByVal eventPlan As Object, _
    ByVal statuses As Object _
)
    Dim expected As String
    Dim message As String

    If eventPlan("operation") = "Departure" Then
        If state("IsOpen") Then
            message = ex_PEB_Export.fn_Config("text.SimpleDepartureWhileOpen")
            Err.Raise 5, , VBA.Replace$(message, "{actual}", state("Type"))
        End If
    Else
        If Not state("IsOpen") Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.SimpleArrivalWithoutDeparture")
        End If
        expected = statuses(eventPlan("sourceStatus"))
        If VBA.StrComp(state("Type"), expected, VBA.vbTextCompare) <> 0 Then
            message = ex_PEB_Export.fn_Config("text.SimpleArrivalTypeMismatch")
            message = VBA.Replace$(message, "{expected}", expected)
            Err.Raise 5, , VBA.Replace$(message, "{actual}", state("Type"))
        End If
    End If
    state("IsOpen") = (eventPlan("operation") <> "Arrival")
    If state("IsOpen") Then
        state("Type") = statuses(eventPlan("targetStatus"))
    End If
End Sub

Private Function private_Before( _
    ByVal first As Object, _
    ByVal second As Object _
) As Boolean
    Dim firstDate As Double
    Dim secondDate As Double

    firstDate = ex_PEB_MovementSource.fn_OrderDate(first("EventDate"))
    secondDate = ex_PEB_MovementSource.fn_OrderDate(second("EventDate"))
    private_Before = (firstDate < secondDate Or _
        (firstDate = secondDate And VBA.Val(first("RowOrder")) < VBA.Val(second("RowOrder"))))
End Function

Private Sub private_Insert( _
    ByVal ordered As Collection, _
    ByVal stored As Object _
)
    Dim index As Long

    For index = 1 To ordered.Count
        If private_Before(stored, ordered(index)) Then
            ordered.Add stored, Before:=index
            Exit Sub
        End If
    Next index
    ordered.Add stored
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------