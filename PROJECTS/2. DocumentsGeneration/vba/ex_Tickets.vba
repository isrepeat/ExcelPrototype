Option Explicit

Private Const TICKETS_TABLE_NAME As String = "tbTickets"
Private Const TICKETS_COL_RANK As String = "Звання"
Private Const TICKETS_COL_FIO As String = "ПІБ"
Private Const TICKETS_COL_IPN As String = "ІПН"
Private Const TICKETS_COL_POSITION As String = "Посада"
Private Const TICKETS_COL_EVENT As String = "Подія"
Private Const TICKETS_COL_OUT_ORDER As String = "Вибуття.Наказ"
Private Const TICKETS_COL_OUT_FOOD As String = "Вибуття.Продовольче"
Private Const TICKETS_COL_OUT_DATE As String = "Вибуття"
Private Const TICKETS_COL_DURATION As String = "Вибуття.Термін"
Private Const TICKETS_COL_ROAD As String = "Вибуття.Дорога"
Private Const TICKETS_COL_ARRIVAL_PLAN As String = "Прибуття.План"
Private Const TICKETS_COL_DOCUMENT As String = "Супровідний документ"
Private Const TICKETS_COL_TVO_FIO As String = "ТВО.ПІБ"
Private Const TICKETS_COL_TVO_IPN As String = "ТВО.ІПН"
Private Const TICKETS_COL_TVO_POSITION As String = "ТВО.Посада"
Private Const TICKETS_COL_STATUS As String = "Статус"

Private Const POSITION_PREFIX_ROZP As String = "A1A"
Private Const POSITION_PREFIX_SPIS As String = "A1B"
Private Const POSITION_VALUE_ROZP As String = "РОЗП"
Private Const POSITION_VALUE_SPIS As String = "СПИС"

' --------------------------------------
' namespace API {
' --------------------------------------
' Находит единственную открытую таблицу реестра билетов.
Public Function ex_TryGetOpenTable(ByRef outTicketsTable As ListObject) As Boolean
    ex_TryGetOpenTable = ex_Document.ex_TryFindOpenTable( _
        TICKETS_TABLE_NAME, outTicketsTable)
End Function

' Возвращает краткое представление кода должности для колонок реестра.
Public Function ex_GetRegistryPosition(ByVal positionCode As String) As String
    ex_GetRegistryPosition = private_Position_ToRegistryValue(positionCode)
End Function

' Находит строку билета по ИПН и номеру приказа; отсутствие строки не является ошибкой.
Public Function ex_TryFindVacationRow( _
    ByVal ticketsTable As ListObject, _
    ByVal ipnText As String, _
    ByVal orderNo As String, _
    ByRef outTicketRow As ListRow _
) As Boolean
    Dim ticketRow As ListRow
    Dim ipnColumnIndex As Long, orderColumnIndex As Long
    Dim existingIpnText As String, existingOrderNo As String

    On Error GoTo EH
    Set outTicketRow = Nothing
    ipnColumnIndex = ticketsTable.ListColumns(TICKETS_COL_IPN).Index
    orderColumnIndex = ticketsTable.ListColumns(TICKETS_COL_OUT_ORDER).Index
    For Each ticketRow In ticketsTable.ListRows
        existingIpnText = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, ipnColumnIndex).Text))
        If VBA.StrComp(existingIpnText, ipnText, VBA.vbTextCompare) = 0 Then
            existingOrderNo = ex_Helpers.private_Text_Normalize( _
                VBA.CStr(ticketRow.Range.Cells(1, orderColumnIndex).Text))
            If VBA.StrComp(existingOrderNo, orderNo, VBA.vbTextCompare) = 0 Then
                If Not outTicketRow Is Nothing Then
                    ex_Helpers.ex_ShowErrorMessage "Multiple vacation tickets were found " & _
                        "for IPN '" & ipnText & "' and order '" & orderNo & "'.", _
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

' Возвращает сохранённый номер билета для режима обновления.
Public Function ex_TryReadTicketNo( _
    ByVal ticketsTable As ListObject, _
    ByVal ticketRow As ListRow, _
    ByRef outTicketNo As String _
) As Boolean
    Dim documentColumnIndex As Long

    On Error GoTo EH
    outTicketNo = VBA.vbNullString
    documentColumnIndex = ticketsTable.ListColumns(TICKETS_COL_DOCUMENT).Index
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

' Формирует следующий номер билета для приказа по уже внесённым строкам tbTickets.
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

    orderColumnIndex = ticketsTable.ListColumns(TICKETS_COL_OUT_ORDER).Index
    documentColumnIndex = ticketsTable.ListColumns(TICKETS_COL_DOCUMENT).Index
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

' Добавляет новую либо обновляет найденную строку tbTickets.
Public Function ex_TrySaveVacationRow( _
    ByVal ticketsTable As ListObject, _
    ByRef ioTicketRow As ListRow, _
    ByVal rankText As String, _
    ByVal fioText As String, _
    ByVal ipnText As String, _
    ByVal positionCode As String, _
    ByVal eventText As String, _
    ByVal orderNo As String, _
    ByVal orderDate As Date, _
    ByVal departureDate As Date, _
    ByVal vacationDays As Long, _
    ByVal roadDays As Long, _
    ByVal dateTo As Date, _
    ByVal ticketNo As String, _
    ByVal tvoFioText As String, _
    ByVal tvoIpnText As String, _
    ByVal tvoPositionCode As String, _
    ByVal statusText As String _
) As Boolean
    Dim tableValues As Object
    Dim valueKey As Variant
    Dim columnIndex As Long

    On Error GoTo EH
    Set tableValues = VBA.CreateObject("Scripting.Dictionary")
    tableValues.CompareMode = VBA.vbBinaryCompare
    tableValues.Add TICKETS_COL_RANK, rankText
    tableValues.Add TICKETS_COL_FIO, fioText
    tableValues.Add TICKETS_COL_IPN, ipnText
    tableValues.Add TICKETS_COL_POSITION, private_Position_ToRegistryValue(positionCode)
    tableValues.Add TICKETS_COL_EVENT, eventText
    tableValues.Add TICKETS_COL_OUT_ORDER, orderNo
    tableValues.Add TICKETS_COL_OUT_FOOD, orderDate
    tableValues.Add TICKETS_COL_OUT_DATE, departureDate
    tableValues.Add TICKETS_COL_DURATION, private_Value_ZeroToBlank(vacationDays)
    tableValues.Add TICKETS_COL_ROAD, private_Value_ZeroToBlank(roadDays)
    tableValues.Add TICKETS_COL_ARRIVAL_PLAN, dateTo
    tableValues.Add TICKETS_COL_DOCUMENT, ticketNo
    tableValues.Add TICKETS_COL_TVO_FIO, tvoFioText
    tableValues.Add TICKETS_COL_TVO_IPN, tvoIpnText
    tableValues.Add TICKETS_COL_TVO_POSITION, _
        private_Position_ToRegistryValue(tvoPositionCode)
    tableValues.Add TICKETS_COL_STATUS, statusText

    ' Сначала проверяем все заголовки, чтобы не создать пустую строку при ошибке схемы.
    For Each valueKey In tableValues.Keys
        columnIndex = ticketsTable.ListColumns(VBA.CStr(valueKey)).Index
    Next valueKey
    If ioTicketRow Is Nothing Then Set ioTicketRow = ticketsTable.ListRows.Add
    For Each valueKey In tableValues.Keys
        columnIndex = ticketsTable.ListColumns(VBA.CStr(valueKey)).Index
        ioTicketRow.Range.Cells(1, columnIndex).Value = tableValues(valueKey)
    Next valueKey
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

' Нормализует специальный код должности для краткого представления в реестре.
Private Function private_Position_ToRegistryValue( _
    ByVal positionCode As String _
) As String
    Dim normalizedCode As String

    normalizedCode = VBA.UCase$(ex_Helpers.private_Text_Normalize(positionCode))
    normalizedCode = VBA.Replace$(normalizedCode, " ", VBA.vbNullString)
    normalizedCode = VBA.Replace$(normalizedCode, "А", "A")
    normalizedCode = VBA.Replace$(normalizedCode, "В", "B")
    If VBA.Left$(normalizedCode, VBA.Len(POSITION_PREFIX_ROZP)) = _
        POSITION_PREFIX_ROZP Then
        private_Position_ToRegistryValue = POSITION_VALUE_ROZP
    ElseIf VBA.Left$(normalizedCode, VBA.Len(POSITION_PREFIX_SPIS)) = _
        POSITION_PREFIX_SPIS Then
        private_Position_ToRegistryValue = POSITION_VALUE_SPIS
    Else
        private_Position_ToRegistryValue = positionCode
    End If
End Function

' Не выводит техническое нулевое значение в ячейку реестра.
Private Function private_Value_ZeroToBlank(ByVal valueData As Variant) As Variant
    If VBA.IsNumeric(valueData) Then
        If VBA.CDbl(valueData) = 0 Then
            private_Value_ZeroToBlank = VBA.vbNullString
            Exit Function
        End If
    End If
    private_Value_ZeroToBlank = valueData
End Function