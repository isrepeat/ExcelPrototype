Option Explicit

Private Const CFG_TBTICKETS As String = "tbTickets"
Private Const CFG_TBTICKETS_COLUMN_IPN As String = "tbTickets::column.ipn"
Private Const CFG_TBTICKETS_COLUMN_OUT_ORDER As String = "tbTickets::column.out_order"
Private Const CFG_TBTICKETS_COLUMN_EVENT As String = "tbTickets::column.event"
Private Const CFG_TBTICKETS_EVENT_DONATION As String = "tbTickets::event.donation"
Private Const CFG_TBTICKETS_COLUMN_DOCUMENT As String = "tbTickets::column.document"
Private Const CFG_TBTICKETS_COLUMN_RANK As String = "tbTickets::column.rank"
Private Const CFG_TBTICKETS_COLUMN_FIO As String = "tbTickets::column.fio"
Private Const CFG_TBTICKETS_COLUMN_POSITION As String = "tbTickets::column.position"
Private Const CFG_TBTICKETS_COLUMN_OUT_FOOD As String = "tbTickets::column.out_food"
Private Const CFG_TBTICKETS_COLUMN_OUT_DATE As String = "tbTickets::column.out_date"
Private Const CFG_TBTICKETS_COLUMN_DURATION As String = "tbTickets::column.duration"
Private Const CFG_TBTICKETS_COLUMN_ROAD As String = "tbTickets::column.road"
Private Const CFG_TBTICKETS_COLUMN_TVO_FIO As String = "tbTickets::column.tvo_fio"
Private Const CFG_TBTICKETS_COLUMN_TVO_IPN As String = "tbTickets::column.tvo_ipn"
Private Const CFG_TBTICKETS_COLUMN_TVO_POSITION As String = "tbTickets::column.tvo_position"
Private Const CFG_TBTICKETS_COLUMN_STATUS As String = "tbTickets::column.status"
Private Const CFG_TBTICKETS_COLUMN_ARRIVAL_PLAN As String = "tbTickets::column.arrival_plan"
Private Const CFG_COMMON_POSITION_ROZP_PREFIX As String = "common::position.rozp.prefix"
Private Const CFG_COMMON_POSITION_ROZP_SHORT As String = "common::position.rozp.short"
Private Const CFG_COMMON_POSITION_SPIS_PREFIX As String = "common::position.spis.prefix"
Private Const CFG_COMMON_POSITION_SPIS_SHORT As String = "common::position.spis.short"

Private Function private_ConfigText(ByVal configKey As String) As String
    If Not ex_Config.fn_TryGetText(configKey, private_ConfigText) Then
        Err.Raise VBA.vbObjectError + 1201, "ex_Tickets", "Configuration value is required: " & configKey
    End If
End Function

' --------------------------------------
' namespace API {
' --------------------------------------
' Finds the single open ticket registry table.
Public Function ex_TryGetOpenTable(ByRef outTicketsTable As ListObject) As Boolean
    ex_TryGetOpenTable = ex_Document.ex_TryFindOpenTable( _
        private_ConfigText(CFG_TBTICKETS), outTicketsTable)
End Function

' Returns a short position value for registry columns.
Public Function ex_GetRegistryPosition(ByVal positionCode As String) As String
    ex_GetRegistryPosition = private_Position_ToRegistryValue(positionCode)
End Function

' Finds a ticket row by IPN, order number, and event type.
' Empty eventText means the main row, except a separate donation row.
Public Function ex_TryFindVacationRow( _
    ByVal ticketsTable As ListObject, _
    ByVal ipnText As String, _
    ByVal orderNo As String, _
    ByVal eventText As String, _
    ByRef outTicketRow As ListRow _
) As Boolean
    Dim ticketRow As ListRow
    Dim ipnColumnIndex As Long, orderColumnIndex As Long, eventColumnIndex As Long
    Dim existingIpnText As String, existingOrderNo As String, existingEventText As String
    Dim isEventMatch As Boolean

    On Error GoTo EH
    Set outTicketRow = Nothing
    ipnColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_IPN)).Index
    orderColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_OUT_ORDER)).Index
    eventColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_EVENT)).Index
    For Each ticketRow In ticketsTable.ListRows
        existingIpnText = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, ipnColumnIndex).Text))
        If VBA.StrComp(existingIpnText, ipnText, VBA.vbTextCompare) = 0 Then
            existingOrderNo = ex_Helpers.private_Text_Normalize( _
                VBA.CStr(ticketRow.Range.Cells(1, orderColumnIndex).Text))
            existingEventText = ex_Helpers.private_Text_Normalize( _
                VBA.CStr(ticketRow.Range.Cells(1, eventColumnIndex).Text))
            If VBA.Len(eventText) = 0 Then
                isEventMatch = (VBA.StrComp(existingEventText, _
                    private_ConfigText(CFG_TBTICKETS_EVENT_DONATION), VBA.vbTextCompare) <> 0)
            Else
                isEventMatch = (VBA.StrComp(existingEventText, _
                    eventText, VBA.vbTextCompare) = 0)
            End If
            If VBA.StrComp(existingOrderNo, orderNo, VBA.vbTextCompare) = 0 And _
               isEventMatch Then
                If Not outTicketRow Is Nothing Then
                    ex_Helpers.ex_ShowErrorMessage "Multiple vacation tickets were found " & _
                        "for IPN '" & ipnText & "', order '" & orderNo & _
                        "' and event '" & eventText & "'.", _
                        VBA.vbExclamation, "Document Generation"
                    Exit Function
                End If
                Set outTicketRow = ticketRow
            End If
        End If
    Next ticketRow
    ex_TryFindVacationRow = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to find vacation ticket | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to find vacation ticket: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Finds a separate donation row by IPN, ticket number, and the "(?)" prefix.
Public Function ex_TryFindDonationRow( _
    ByVal ticketsTable As ListObject, _
    ByVal ipnText As String, _
    ByVal ticketNo As String, _
    ByRef outTicketRow As ListRow _
) As Boolean
    Dim ticketRow As ListRow
    Dim ipnColumnIndex As Long, documentColumnIndex As Long
    Dim eventColumnIndex As Long, orderColumnIndex As Long
    Dim existingIpnText As String, existingTicketNo As String
    Dim existingEventText As String, existingOrderNo As String

    On Error GoTo EH
    Set outTicketRow = Nothing
    ipnColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_IPN)).Index
    documentColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_DOCUMENT)).Index
    eventColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_EVENT)).Index
    orderColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_OUT_ORDER)).Index
    For Each ticketRow In ticketsTable.ListRows
        existingIpnText = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, ipnColumnIndex).Text))
        existingTicketNo = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, documentColumnIndex).Text))
        existingEventText = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, eventColumnIndex).Text))
        existingOrderNo = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, orderColumnIndex).Text))
        If VBA.StrComp(existingIpnText, ipnText, VBA.vbTextCompare) = 0 And _
           VBA.StrComp(existingTicketNo, ticketNo, VBA.vbTextCompare) = 0 And _
           VBA.StrComp(existingEventText, private_ConfigText(CFG_TBTICKETS_EVENT_DONATION), _
               VBA.vbTextCompare) = 0 And _
           (VBA.Len(existingOrderNo) = 0 Or _
            VBA.Left$(existingOrderNo, 4) = "(?) ") Then
            If Not outTicketRow Is Nothing Then
                ex_Helpers.ex_ShowErrorMessage "Multiple donation rows were found " & _
                    "for IPN '" & ipnText & "' and ticket '" & ticketNo & "'.", _
                    VBA.vbExclamation, "Document Generation"
                Exit Function
            End If
            Set outTicketRow = ticketRow
        End If
    Next ticketRow
    ex_TryFindDonationRow = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to find donation ticket row | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to find donation ticket row: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Returns the saved ticket number for update mode.
Public Function ex_TryReadTicketNo( _
    ByVal ticketsTable As ListObject, _
    ByVal ticketRow As ListRow, _
    ByRef outTicketNo As String _
) As Boolean
    Dim documentColumnIndex As Long

    On Error GoTo EH
    outTicketNo = VBA.vbNullString
    documentColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_DOCUMENT)).Index
    outTicketNo = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(ticketRow.Range.Cells(1, documentColumnIndex).Text))
    If VBA.Len(outTicketNo) = 0 Then
        ex_Helpers.ex_ShowErrorMessage "Existing vacation ticket has no ticket number.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    ex_TryReadTicketNo = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to read existing ticket number | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to read existing ticket number: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Removes the donation row when extra days are no longer set.
Public Function ex_TryDeleteVacationRow( _
    ByVal ticketRow As ListRow _
) As Boolean
    On Error GoTo EH
    ticketRow.Delete
    ex_TryDeleteVacationRow = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to delete vacation ticket row | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to delete vacation ticket row: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Builds the next ticket number for an order from existing tbTickets rows.
Public Function ex_TryBuildNextTicketNo( _
    ByVal ticketsTable As ListObject, _
    ByVal orderNo As String, _
    ByVal orderDate As Date, _
    ByRef outTicketNo As String _
) As Boolean
    Dim ticketRow As ListRow
    Dim orderColumnIndex As Long, documentColumnIndex As Long
    Dim normalizedOrderNo As String, existingOrderNo As String
    Dim existingTicketNo As String, ticketPrefix As String, sequenceText As String
    Dim greatestSequence As Long, ticketSequence As Long

    On Error GoTo EH
    outTicketNo = VBA.vbNullString
    normalizedOrderNo = ex_Helpers.private_Text_Normalize(orderNo)
    If Not ex_Helpers.private_Text_IsDigits(normalizedOrderNo) Then
        ex_Helpers.LogError "Resolved order number is not numeric: " & orderNo
        ex_Helpers.ex_ShowErrorMessage "Resolved order number must contain digits only: " & _
            orderNo, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    orderColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_OUT_ORDER)).Index
    documentColumnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_DOCUMENT)).Index
    ticketPrefix = VBA.CStr(VBA.Year(orderDate)) & "/" & normalizedOrderNo & "/"
    For Each ticketRow In ticketsTable.ListRows
        existingOrderNo = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, orderColumnIndex).Text))
        If VBA.StrComp(existingOrderNo, normalizedOrderNo, VBA.vbTextCompare) = 0 Then
            existingTicketNo = ex_Helpers.private_Text_Normalize( _
                VBA.CStr(ticketRow.Range.Cells(1, documentColumnIndex).Text))
            If VBA.Left$(existingTicketNo, VBA.Len(ticketPrefix)) <> ticketPrefix Then
                ex_Helpers.LogError "Ticket number has invalid prefix | Value=" & _
                    existingTicketNo & " | Expected=" & ticketPrefix
                ex_Helpers.ex_ShowErrorMessage "Ticket number '" & existingTicketNo & _
                    "' does not match order '" & normalizedOrderNo & "'.", _
                    VBA.vbExclamation, "Document Generation"
                Exit Function
            End If
            sequenceText = VBA.Mid$(existingTicketNo, VBA.Len(ticketPrefix) + 1)
            If Not ex_Helpers.private_Text_IsDigits(sequenceText) Then
                ex_Helpers.LogError "Ticket number has invalid sequence: " & existingTicketNo
                ex_Helpers.ex_ShowErrorMessage "Ticket number has an invalid sequence: " & _
                    existingTicketNo, VBA.vbExclamation, "Document Generation"
                Exit Function
            End If
            ticketSequence = VBA.CLng(sequenceText)
            If ticketSequence > greatestSequence Then greatestSequence = ticketSequence
        End If
    Next ticketRow

    If greatestSequence = 2147483647 Then
        ex_Helpers.LogError "Ticket sequence limit reached | Prefix=" & ticketPrefix
        ex_Helpers.ex_ShowErrorMessage "Ticket number sequence limit was reached for order '" & _
            normalizedOrderNo & "'.", VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outTicketNo = ticketPrefix & VBA.CStr(greatestSequence + 1)
    ex_Helpers.LogDebug "Next vacation ticket number: " & outTicketNo
    ex_TryBuildNextTicketNo = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to build next ticket number | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to build the next ticket number: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Adds a new tbTickets row or updates the found row.
Public Function ex_TrySaveVacationRow( _
    ByVal ticketsTable As ListObject, _
    ByRef ioTicketRow As ListRow, _
    ByVal rankText As String, _
    ByVal fioText As String, _
    ByVal ipnText As String, _
    ByVal positionCode As String, _
    ByVal eventText As String, _
    ByVal orderNo As String, _
    ByVal foodDepartureDate As Date, _
    ByVal departureDate As Date, _
    ByVal vacationDays As Long, _
    ByVal roadDays As Long, _
    ByVal ticketNo As String, _
    ByVal tvoFioText As String, _
    ByVal tvoIpnText As String, _
    ByVal tvoPositionCode As String, _
    ByVal statusText As String _
) As Boolean
    Dim tableValues As Object
    Dim valueKey As Variant
    Dim columnIndex As Long
    Dim arrivalPlanFormula As String

    On Error GoTo EH
    Set tableValues = VBA.CreateObject("Scripting.Dictionary")
    tableValues.CompareMode = VBA.vbBinaryCompare
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_RANK), rankText
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_FIO), fioText
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_IPN), ipnText
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_POSITION), private_Position_ToRegistryValue(positionCode)
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_EVENT), eventText
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_OUT_ORDER), orderNo
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_OUT_FOOD), foodDepartureDate
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_OUT_DATE), departureDate
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_DURATION), private_Value_ZeroToBlank(vacationDays)
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_ROAD), private_Value_ZeroToBlank(roadDays)
    ' The arrival-plan field is calculated by the tbTickets table formula.
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_DOCUMENT), ticketNo
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_TVO_FIO), tvoFioText
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_TVO_IPN), tvoIpnText
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_TVO_POSITION), _
        private_Position_ToRegistryValue(tvoPositionCode)
    tableValues.Add private_ConfigText(CFG_TBTICKETS_COLUMN_STATUS), statusText

    ' Check all headers first to avoid adding an empty row for a schema error.
    For Each valueKey In tableValues.Keys
        columnIndex = ticketsTable.ListColumns(VBA.CStr(valueKey)).Index
    Next valueKey
    If Not private_Tickets_TryGetCalculatedFormula( _
        ticketsTable, private_ConfigText(CFG_TBTICKETS_COLUMN_ARRIVAL_PLAN), arrivalPlanFormula) Then Exit Function
    If ioTicketRow Is Nothing Then Set ioTicketRow = ticketsTable.ListRows.Add
    For Each valueKey In tableValues.Keys
        columnIndex = ticketsTable.ListColumns(VBA.CStr(valueKey)).Index
        ioTicketRow.Range.Cells(1, columnIndex).Value = tableValues(valueKey)
    Next valueKey
    columnIndex = ticketsTable.ListColumns(private_ConfigText(CFG_TBTICKETS_COLUMN_ARRIVAL_PLAN)).Index
    ioTicketRow.Range.Cells(1, columnIndex).Formula = arrivalPlanFormula
    ex_TrySaveVacationRow = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to save vacation ticket row | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to save vacation ticket row: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' Reads the saved calculated-column formula from any tbTickets row.
Private Function private_Tickets_TryGetCalculatedFormula( _
    ByVal ticketsTable As ListObject, _
    ByVal columnName As String, _
    ByRef outFormula As String _
) As Boolean
    Dim ticketRow As ListRow
    Dim columnIndex As Long
    Dim formulaCell As Range

    On Error GoTo EH
    outFormula = VBA.vbNullString
    columnIndex = ticketsTable.ListColumns(columnName).Index
    For Each ticketRow In ticketsTable.ListRows
        Set formulaCell = ticketRow.Range.Cells(1, columnIndex)
        If formulaCell.HasFormula Then
            outFormula = VBA.CStr(formulaCell.Formula)
            Exit For
        End If
    Next ticketRow
    If VBA.Len(outFormula) = 0 Then
        ex_Helpers.ex_ShowErrorMessage "Table tbTickets has no saved formula in " & _
            "column '" & columnName & "'. Add the formula to one table row and retry.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    private_Tickets_TryGetCalculatedFormula = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to read calculated column formula | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to read the formula in column '" & _
        columnName & "': " & Err.Description, VBA.vbExclamation, _
        "Document Generation"
End Function

' Normalizes a special position code for a short registry value.
Private Function private_Position_ToRegistryValue( _
    ByVal positionCode As String _
) As String
    Dim normalizedCode As String

    normalizedCode = VBA.UCase$(ex_Helpers.private_Text_Normalize(positionCode))
    normalizedCode = VBA.Replace$(normalizedCode, " ", VBA.vbNullString)
    normalizedCode = VBA.Replace$(normalizedCode, ex_Helpers.fn_FromCodePoints("1040"), "A")
    normalizedCode = VBA.Replace$(normalizedCode, ex_Helpers.fn_FromCodePoints("1042"), "B")
    If VBA.Left$(normalizedCode, VBA.Len(private_ConfigText(CFG_COMMON_POSITION_ROZP_PREFIX))) = _
        private_ConfigText(CFG_COMMON_POSITION_ROZP_PREFIX) Then
        private_Position_ToRegistryValue = private_ConfigText(CFG_COMMON_POSITION_ROZP_SHORT)
    ElseIf VBA.Left$(normalizedCode, VBA.Len(private_ConfigText(CFG_COMMON_POSITION_SPIS_PREFIX))) = _
        private_ConfigText(CFG_COMMON_POSITION_SPIS_PREFIX) Then
        private_Position_ToRegistryValue = private_ConfigText(CFG_COMMON_POSITION_SPIS_SHORT)
    Else
        private_Position_ToRegistryValue = positionCode
    End If
End Function

' Does not write a technical zero value to a registry cell.
Private Function private_Value_ZeroToBlank(ByVal valueData As Variant) As Variant
    If VBA.IsNumeric(valueData) Then
        If VBA.CDbl(valueData) = 0 Then
            private_Value_ZeroToBlank = VBA.vbNullString
            Exit Function
        End If
    End If
    private_Value_ZeroToBlank = valueData
End Function