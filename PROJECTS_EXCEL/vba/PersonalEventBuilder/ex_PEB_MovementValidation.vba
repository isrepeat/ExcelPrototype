Attribute VB_Name = "ex_PEB_MovementValidation"
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
Public Sub fn_Insert( _
    ByVal ordered As Collection, _
    ByVal stored As Object _
)
    Dim other As Object
    Dim index As Long
    Dim eventDate As Date
    Dim otherDate As Date
    Dim sequence As Long

    eventDate = ex_PEB_History.fn_Date(stored("EventDate"))
    sequence = VBA.CLng(stored("RowOrder"))
    If sequence < 1 Then
        ex_PEB_MovementLinks.fn_Fail "InvalidRowOrder", stored("RowId")
    End If
    For index = 1 To ordered.Count
        Set other = ordered(index)
        otherDate = ex_PEB_History.fn_Date(other("EventDate"))
        If eventDate = otherDate And sequence = VBA.CLng(other("RowOrder")) And _
                stored("PersonId") = other("PersonId") Then
            ex_PEB_MovementLinks.fn_Fail "AmbiguousSameDayOrder", stored("RowId")
        End If
        If eventDate < otherDate Or (eventDate = otherDate And sequence < VBA.CLng(other("RowOrder"))) Then
            ordered.Add stored, Before:=index
            Exit Sub
        End If
    Next index
    ordered.Add stored
End Sub

Public Function fn_Apply( _
    ByVal periods As Collection, _
    ByVal stored As Object, _
    ByVal retainedClosings As Object _
) As Long
    Dim plan As Object
    Dim statuses As Object
    Dim period As Object
    Dim selected As Object
    Dim eventDate As Date
    Dim expected As String
    Dim actual As String
    Dim diagnostic As String
    Dim count As Long
    Dim openCount As Long
    Dim matches As Boolean
    Dim index As Long
    Dim selectedIndex As Long

    Set plan = ex_PEB_MovementRegistry.fn_Event(stored("EventId"))
    Set statuses = ex_PEB_MovementRegistry.fn_Statuses()
    eventDate = ex_PEB_History.fn_Date(stored("EventDate"))
    If VBA.Len(plan("sourceStatus")) > 0 Then
        expected = statuses(plan("sourceStatus"))
    End If
    For Each period In periods
        index = index + 1
        If private_Text(period("PersonId")) = stored("PersonId") Then
            If VBA.Len(private_Text(period("ArrivalDate"))) = 0 Then
                If private_Date(period("DepartureDate")) <= eventDate Then
                    openCount = openCount + 1
                End If
                actual = actual & private_Text(period("EventName")) & " (" & private_Text(period("DepartureDate")) & "); "
                matches = (private_Text(period("EventName")) = expected)
                If VBA.Len(plan("sourceDestination")) > 0 Then
                    matches = matches And (private_Text(period("Destination")) = plan("sourceDestination"))
                End If
                If matches And private_Date(period("DepartureDate")) <= eventDate Then
                    count = count + 1
                    Set selected = period
                    selectedIndex = index
                End If
            End If
        End If
    Next period
    If plan("operation") = "Departure" Then
        If openCount > 0 Then
            ex_PEB_MovementLinks.fn_Fail "RepeatDeparture", stored("RowId")
        End If
    Else
        If count <> 1 Then
            If VBA.Len(actual) = 0 Then
                actual = ex_PEB_Export.fn_Config("text.MovementNoOpenAbsence")
            End If
            diagnostic = ex_PEB_Export.fn_Config("text.MovementOpenAbsenceMismatch")
            diagnostic = VBA.Replace$(diagnostic, "{person}", stored("PersonId"))
            diagnostic = VBA.Replace$(diagnostic, "{expected}", expected)
            diagnostic = VBA.Replace$(diagnostic, "{date}", stored("EventDate"))
            diagnostic = VBA.Replace$(diagnostic, "{actual}", actual)
            If count > 1 Then
                ex_PEB_MovementLinks.fn_Fail "AmbiguousOpenAbsence", diagnostic
            End If
            Err.Raise 5, , diagnostic
        End If
        selected("ArrivalDate") = eventDate
        selected("ArrivalOrder") = stored("AcceptedOrderNumber")
    End If
    If plan("operation") <> "Arrival" Then
        Set period = VBA.CreateObject("Scripting.Dictionary")
        period("PersonId") = stored("PersonId")
        period("EventName") = statuses(plan("targetStatus"))
        period("DepartureDate") = eventDate
        period("DepartureOrder") = stored("AcceptedOrderNumber")
        period("ArrivalDate") = Empty
        period("ArrivalOrder") = Empty
        If retainedClosings.Exists(stored("RowId")) Then
            period("ArrivalDate") = retainedClosings(stored("RowId"))("ArrivalDate")
            period("ArrivalOrder") = retainedClosings(stored("RowId"))("ArrivalOrder")
        End If
        period("Destination") = plan("destinationValue")
        If VBA.Len(plan("destinationValue")) = 0 Then
            period("Destination") = stored(plan("destinationField"))
        End If
        periods.Add period
    End If
    fn_Apply = selectedIndex
End Function

Public Sub fn_CheckPeriods( _
    ByVal periods As Collection _
)
    Dim first As Object
    Dim second As Object
    Dim startAt As Double
    Dim endAt As Double
    Dim otherEnd As Double
    Dim index As Long
    Dim other As Long

    For index = 1 To periods.Count
        Set first = periods(index)
        startAt = private_Date(first("DepartureDate"))
        endAt = private_End(first)
        If endAt < startAt Then
            ex_PEB_MovementLinks.fn_Fail "InvalidPeriod", private_Text(first("PersonId"))
        End If
        For other = index + 1 To periods.Count
            Set second = periods(other)
            If private_Text(first("PersonId")) = private_Text(second("PersonId")) Then
                otherEnd = private_End(second)
                If startAt < otherEnd And private_Date(second("DepartureDate")) < endAt Then
                    ex_PEB_MovementLinks.fn_Fail "OverlappingPeriods", private_Text(first("PersonId"))
                End If
            End If
        Next other
    Next index
End Sub

' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_End(ByVal period As Object) As Double
    private_End = 2958466#
    If VBA.Len(private_Text(period("ArrivalDate"))) > 0 Then
        private_End = private_Date(period("ArrivalDate"))
    End If
End Function

Private Function private_Date(ByVal value As Variant) As Date
    If VBA.VarType(value) <> VBA.vbString And VBA.IsNumeric(value) Then
        If VBA.CDbl(value) < 61 Or VBA.CDbl(value) > 2958465# Then
            ex_PEB_MovementLinks.fn_Fail "InvalidMovementDate", VBA.CStr(value)
        End If
        private_Date = VBA.CDate(value)
    Else
        private_Date = ex_PEB_History.fn_Date(value)
    End If
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        ex_PEB_MovementLinks.fn_Fail "InvalidSnapshotValue", VBA.CStr(value)
    End If
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
        private_Text = VBA.Trim$(VBA.CStr(value))
    End If
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------