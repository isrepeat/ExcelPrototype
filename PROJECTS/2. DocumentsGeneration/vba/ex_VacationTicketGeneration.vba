Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

Private Const INPUT_SHEET_NAME As String = "Відпустки"

' Стабильные aliases полей формы. Адреса инкапсулированы в Input mapper-е.
Private Const INPUT_ALIAS_PERSON_LOOKUP As String = "PersonLookup"
Private Const INPUT_ALIAS_ORDER_REFERENCE As String = "OrderReference"
Private Const INPUT_ALIAS_VACATION_KIND As String = "VacationKind"
Private Const INPUT_ALIAS_VACATION_PLACE As String = "VacationPlace"
Private Const INPUT_ALIAS_VACATION_DAYS As String = "VacationDays"
Private Const INPUT_ALIAS_ROAD_DAYS As String = "RoadDays"
Private Const INPUT_ALIAS_DONATION_DAYS As String = "DonationDays"
Private Const INPUT_ALIAS_TVO_LOOKUP As String = "TvoLookup"
Private Const INPUT_ALIAS_TEMPLATE_PATH As String = "TemplatePath"
Private Const INPUT_ALIAS_OUTPUT_FOLDER_PATH As String = "OutputFolderPath"

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

' Канонические типы отпусков и их текст для Word-шаблона.
Private Const VACATION_KIND_ANNUAL As String = "Щорічна відпустка"
Private Const VACATION_KIND_DONATION As String = "Відпочинок за донацію крові"
Private Const VACATION_KIND_FAMILY As String = "Відпустка за сімейними обставинами"
Private Const VACATION_KIND_TREATMENT As String = "Відпустка для лікування"
Private Const VACATION_KIND_MATERNITY As String = _
    "Відпустка у зв'язку з вагітністю та пологами"
Private Const VACATION_KIND_CHILDCARE As String = "Відпустка по догляду за дитиною"
Private Const VACATION_TEXT_ANNUAL As String = _
    "у частину щорічної основної відпустки"
Private Const VACATION_TEXT_DONATION As String = _
    "у відпочинок за донацію крові"
Private Const VACATION_TEXT_FAMILY As String = _
    "у відпустку за сімейними обставинами"
Private Const VACATION_TEXT_TREATMENT As String = _
    "у відпустку для лікування"
Private Const VACATION_TEXT_MATERNITY As String = _
    "у відпустку у зв'язку з вагітністю та пологами"
Private Const VACATION_TEXT_CHILDCARE As String = _
    "у відпустку по догляду за дитиною"

' Форматы дат для Word-шаблона. Текст «до 08:00 год.» находится в шаблоне.
Private Const DATE_FORMAT_PATTERN As String = """{dd}"" {month} {yyyy} р."

' Стабильный шаблон имени созданного документа.
Private Const DOCUMENT_NAME_PATTERN As String = "Відпускний квиток {FIO} ({IPN})"

' Aliases контекста имени генерируемого документа.
Private Const GENERATED_CONTEXT_ALIAS_FIO As String = "FIO"
Private Const GENERATED_CONTEXT_ALIAS_IPN As String = "IPN"

Private inputCellMap As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_VacationTicketGeneration_Create()
    Dim personLookup As String, tvoLookup As String
    Dim tvoIpnText As String, tvoFioText As String
    Dim orderReference As String
    Dim vacationKind As String, vacationKindText As String
    Dim vacationPlace As String
    Dim templatePath As String, outputFolderPath As String
    Dim ipnText As String, fioText As String
    Dim rankText As String, orderNo As String
    Dim personPositionCode As String, tvoPositionCode As String
    Dim orderDate As Date, dateFrom As Date, dateTo As Date, dateArrival As Date
    Dim vacationDays As Long, roadDays As Long, donationDays As Long
    Dim ticketNo As String, ticketDateText As String
    Dim dateFromText As String, dateToText As String
    Dim dateArrivalText As String, personalLine As String, personalInitials As String
    Dim placeholderNames As Variant, placeholderValues As Variant
    Dim documentPath As String, documentNameValues As Object
    Dim tableValues As Object
    Dim ticketsTable As ListObject
    Dim performanceStart As Single
    Dim personnelSessionStarted As Boolean

    On Error GoTo EH
    private_Initialize
    ex_Helpers.ClearLog
    performanceStart = VBA.Timer
    private_Performance_LogCheckpoint performanceStart, "Start"
    ex_Helpers.LogDebug "Vacation ticket generation started"
    Call ex_Document.ex_LogWorkbookContext(INPUT_SHEET_NAME, inputCellMap)

    personLookup = private_Input_ReadRequired(INPUT_ALIAS_PERSON_LOOKUP, "ПІБ або ІПН")
    orderReference = private_Input_ReadRequired(INPUT_ALIAS_ORDER_REFERENCE, "номер або дату наказу")
    vacationKind = private_Input_ReadRequired(INPUT_ALIAS_VACATION_KIND, "вид відпустки")
    vacationPlace = private_Input_ReadRequired(INPUT_ALIAS_VACATION_PLACE, "місце відпустки")
    tvoLookup = private_Input_ReadRequired(INPUT_ALIAS_TVO_LOOKUP, "ПІБ або ІПН ТВО")
    templatePath = private_Input_ReadRequired(INPUT_ALIAS_TEMPLATE_PATH, "шлях до шаблону")
    outputFolderPath = private_Input_ReadRequired( _
        INPUT_ALIAS_OUTPUT_FOLDER_PATH, "шлях до папки результатів")
    If VBA.Len(personLookup) = 0 Or VBA.Len(orderReference) = 0 Or _
        VBA.Len(vacationKind) = 0 Or VBA.Len(vacationPlace) = 0 Or _
        VBA.Len(tvoLookup) = 0 Or _
        VBA.Len(templatePath) = 0 Or VBA.Len(outputFolderPath) = 0 Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Input read"

    ex_Helpers.LogDebug "Vacation ticket person lookup: " & personLookup
    ex_Helpers.LogDebug "Vacation ticket order reference: " & orderReference

    If Not private_Input_TryReadNonNegativeDays( _
        INPUT_ALIAS_VACATION_DAYS, "термін вибуття", vacationDays) Then GoTo CleanExit
    If Not private_Input_TryReadOptionalNonNegativeDays( _
        INPUT_ALIAS_ROAD_DAYS, "додаткові дні на дорогу", roadDays) Then GoTo CleanExit
    If Not private_Input_TryReadOptionalNonNegativeDays( _
        INPUT_ALIAS_DONATION_DAYS, "додаткові дні на донацію", donationDays) Then GoTo CleanExit
    If Not private_Vacation_TryMapKind(vacationKind, vacationKindText) Then GoTo CleanExit

    If Not ex_PersonnelData.ex_TryBeginSession() Then GoTo CleanExit
    personnelSessionStarted = True
    private_Performance_LogCheckpoint performanceStart, "SHPO session opened"
    If Not ex_PersonnelData.ex_TryResolveIpn(personLookup, ipnText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveIpn(tvoLookup, tvoIpnText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveFioNominative( _
        tvoIpnText, tvoFioText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolvePositionCode( _
        tvoIpnText, tvoPositionCode) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveFioNominative(ipnText, fioText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveRankNominative(ipnText, rankText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolvePositionCode( _
        ipnText, personPositionCode) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveOrderReference( _
        orderReference, orderNo, orderDate) Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Personnel and order data resolved"
    If Not ex_Document.ex_TryFindOpenTable( _
        TICKETS_TABLE_NAME, ticketsTable) Then GoTo CleanExit
    If Not private_Tickets_TryValidateNoDuplicatePerson( _
        ticketsTable, ipnText, orderDate) Then GoTo CleanExit
    If Not private_Tickets_TryBuildNextTicketNo( _
        ticketsTable, orderNo, orderDate, ticketNo) Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Registry validated and ticket number assigned"

    ' Даты рассчитываются по утверждённому правилу отпуска.
    dateFrom = VBA.DateAdd("d", 1, orderDate)
    dateTo = VBA.DateAdd("d", vacationDays + roadDays + donationDays, dateFrom)
    dateArrival = VBA.DateAdd("d", 1, dateTo)
    ex_Helpers.LogDebug "Vacation period | From=" & VBA.CStr(dateFrom) & _
        " | To=" & VBA.CStr(dateTo) & " | Arrival=" & VBA.CStr(dateArrival)
    If Not ex_Helpers.private_Date_TryFormat( _
        orderDate, DATE_FORMAT_PATTERN, ticketDateText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        dateFrom, DATE_FORMAT_PATTERN, dateFromText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        dateTo, DATE_FORMAT_PATTERN, dateToText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        dateArrival, DATE_FORMAT_PATTERN, dateArrivalText) Then GoTo CleanExit

    personalLine = rankText & " " & fioText
    ' Краткая запись используется в строке о возвращении из отпуска.
    personalInitials = private_Person_BuildInitials(rankText, fioText)
    If VBA.Len(personalInitials) = 0 Then GoTo CleanExit

    ' Имена должны точно совпадать с плейсхолдерами Word-шаблона.
    placeholderNames = Array( _
        "TicketDate", "TicketNum", "PersonalLine", "VacationKind", _
        "VacationPlace", "VacationDuration", "DateFrom", "DateTo", _
        "PersonalInitials", "DateArrival")
    placeholderValues = Array( _
        ticketDateText, ticketNo, personalLine, vacationKindText, vacationPlace, _
        VBA.CStr(vacationDays) & " календарних днів", dateFromText, _
        dateToText, personalInitials, dateArrivalText)

    Set documentNameValues = VBA.CreateObject("Scripting.Dictionary")
    documentNameValues.CompareMode = VBA.vbBinaryCompare
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_FIO, fioText
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_IPN, ipnText
    private_Performance_LogCheckpoint performanceStart, "Word generation started"
    If Not ex_Document.ex_TryGenerateWordDocument( _
        templatePath, DOCUMENT_NAME_PATTERN, documentNameValues, _
        placeholderNames, placeholderValues, documentPath, _
        outputFolderPath) Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Word generation completed"

    Set tableValues = private_Tickets_BuildRowValues( _
        rankText, fioText, ipnText, personPositionCode, vacationKind, orderNo, _
        orderDate, vacationDays, roadDays, dateTo, ticketNo, _
        tvoFioText, tvoIpnText, tvoPositionCode)
    If Not ex_Document.ex_TryAppendTableRow( _
        TICKETS_TABLE_NAME, tableValues) Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Registry row appended"
    ex_Helpers.WriteLog "DOCUMENT: " & documentPath
    ex_Helpers.LogDebug "Vacation ticket generation completed"
    ex_Helpers.ex_ShowStatusBarMessage _
        "Vacation ticket generated: " & documentPath
    GoTo CleanExit

CleanExit:
    If personnelSessionStarted Then ex_PersonnelData.ex_EndSession
    Exit Sub
EH:
    ex_Helpers.LogError "Vacation ticket generation failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Vacation ticket generation failed: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Performance {
' --------------------------------------
' Временная диагностика длительности этапов генерации. Удалить после замеров.
Private Sub private_Performance_LogCheckpoint( _
    ByVal startTime As Single, _
    ByVal checkpointName As String _
)
    Dim elapsedSeconds As Single

    elapsedSeconds = VBA.Timer - startTime
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400!
    ex_Helpers.WriteLog "PERF | ElapsedMs=" & _
        VBA.Format$(elapsedSeconds * 1000!, "0") & " | Checkpoint=" & checkpointName
End Sub
' --------------------------------------
' } // namespace Performance
' --------------------------------------

' Инициализирует module-level состояние перед каждым запуском генерации.
Private Sub private_Initialize()
    Set inputCellMap = VBA.CreateObject("Scripting.Dictionary")
    inputCellMap.CompareMode = VBA.vbBinaryCompare

    ' Адреса значений соответствуют строкам конфигурационной таблицы на листе.
    inputCellMap.Add INPUT_ALIAS_PERSON_LOOKUP, "C4"
    inputCellMap.Add INPUT_ALIAS_ORDER_REFERENCE, "C5"
    inputCellMap.Add INPUT_ALIAS_VACATION_KIND, "C6"
    inputCellMap.Add INPUT_ALIAS_VACATION_PLACE, "C7"
    inputCellMap.Add INPUT_ALIAS_VACATION_DAYS, "C8"
    inputCellMap.Add INPUT_ALIAS_ROAD_DAYS, "C9"
    inputCellMap.Add INPUT_ALIAS_DONATION_DAYS, "C10"
    inputCellMap.Add INPUT_ALIAS_TVO_LOOKUP, "C11"
    inputCellMap.Add INPUT_ALIAS_TEMPLATE_PATH, "C12"
    inputCellMap.Add INPUT_ALIAS_OUTPUT_FOLDER_PATH, "C13"
End Sub

' --------------------------------------
' namespace Input {
' --------------------------------------
Private Function private_Input_ReadRequired( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String _
) As String
    If Not ex_Document.ex_TryReadRequired( _
        INPUT_SHEET_NAME, inputCellMap, fieldAlias, fieldCaption, _
        private_Input_ReadRequired) Then Exit Function
End Function

Private Function private_Input_TryReadNonNegativeDays( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByRef outDays As Long _
) As Boolean
    private_Input_TryReadNonNegativeDays = _
        ex_Document.ex_TryReadNonNegativeDays( _
            INPUT_SHEET_NAME, inputCellMap, fieldAlias, fieldCaption, _
            False, outDays)
End Function

Private Function private_Input_TryReadOptionalNonNegativeDays( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByRef outDays As Long _
) As Boolean
    private_Input_TryReadOptionalNonNegativeDays = _
        ex_Document.ex_TryReadNonNegativeDays( _
            INPUT_SHEET_NAME, inputCellMap, fieldAlias, fieldCaption, _
            True, outDays)
End Function

' --------------------------------------
' } // namespace Input
' --------------------------------------

' --------------------------------------
' namespace Person {
' --------------------------------------
Private Function private_Person_BuildInitials( _
    ByVal rankText As String, _
    ByVal fioText As String _
) As String
    Dim nameParts As Variant

    nameParts = VBA.Split(ex_Helpers.private_Text_Normalize(fioText), " ")
    If UBound(nameParts) < 2 Then
        ex_Helpers.ex_ShowErrorMessage "FIO must contain surname, name and patronymic: " & fioText, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    private_Person_BuildInitials = rankText & " " & nameParts(0) & _
        VBA.ChrW$(160) & VBA.Left$(nameParts(1), 1) & "." & _
        VBA.Left$(nameParts(2), 1) & "."
End Function
' --------------------------------------
' } // namespace Person
' --------------------------------------

' --------------------------------------
' namespace Vacation {
' --------------------------------------
Private Function private_Vacation_TryMapKind( _
    ByVal vacationKind As String, _
    ByRef outVacationText As String _
) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(vacationKind))
        Case VBA.LCase$(VACATION_KIND_ANNUAL)
            outVacationText = VACATION_TEXT_ANNUAL
        Case VBA.LCase$(VACATION_KIND_DONATION)
            outVacationText = VACATION_TEXT_DONATION
        Case VBA.LCase$(VACATION_KIND_FAMILY)
            outVacationText = VACATION_TEXT_FAMILY
        Case VBA.LCase$(VACATION_KIND_TREATMENT)
            outVacationText = VACATION_TEXT_TREATMENT
        Case VBA.LCase$(VACATION_KIND_MATERNITY)
            outVacationText = VACATION_TEXT_MATERNITY
        Case VBA.LCase$(VACATION_KIND_CHILDCARE)
            outVacationText = VACATION_TEXT_CHILDCARE
        Case Else
            ex_Helpers.LogError "Unsupported vacation kind: " & vacationKind
            ex_Helpers.ex_ShowErrorMessage "Unsupported vacation kind: " & vacationKind, _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Vacation_TryMapKind = True
End Function
' --------------------------------------
' } // namespace Vacation
' --------------------------------------

' --------------------------------------
' namespace Tickets {
' --------------------------------------
' Не допускает повторного оформления отпуска одному человеку на дату приказа.
Private Function private_Tickets_TryValidateNoDuplicatePerson( _
    ByVal ticketsTable As ListObject, _
    ByVal ipnText As String, _
    ByVal orderDate As Date _
) As Boolean
    Dim ticketRow As ListRow
    Dim ipnColumnIndex As Long, departureDateColumnIndex As Long
    Dim existingIpnText As String, existingDateText As String
    Dim existingOrderDate As Date

    On Error GoTo EH
    ipnColumnIndex = ticketsTable.ListColumns(TICKETS_COL_IPN).Index
    departureDateColumnIndex = ticketsTable.ListColumns(TICKETS_COL_OUT_DATE).Index
    For Each ticketRow In ticketsTable.ListRows
        existingIpnText = ex_Helpers.private_Text_Normalize( _
            VBA.CStr(ticketRow.Range.Cells(1, ipnColumnIndex).Text))
        If VBA.StrComp(existingIpnText, ipnText, VBA.vbTextCompare) = 0 Then
            existingDateText = ex_Helpers.private_Text_Normalize( _
                VBA.CStr(ticketRow.Range.Cells(1, departureDateColumnIndex).Text))
            If Not ex_Helpers.private_Date_TryParse( _
                existingDateText, existingOrderDate) Then
                ex_Helpers.LogError "Existing ticket has invalid departure date | " & _
                    "IPN=" & ipnText & " | Value=" & existingDateText
                ex_Helpers.ex_ShowErrorMessage "Existing ticket for IPN '" & ipnText & _
                    "' has an invalid departure date: " & existingDateText, _
                    VBA.vbExclamation, "Document Generation"
                Exit Function
            End If
            If VBA.DateValue(existingOrderDate) = VBA.DateValue(orderDate) Then
                ex_Helpers.LogError "Duplicate vacation ticket | IPN=" & ipnText & _
                    " | OrderDate=" & VBA.CStr(orderDate)
                ex_Helpers.ex_ShowErrorMessage "A vacation ticket already exists for " & _
                    "this person on " & VBA.Format$(orderDate, "dd.mm.yyyy") & ".", _
                    VBA.vbExclamation, "Document Generation"
                Exit Function
            End If
        End If
    Next ticketRow
    private_Tickets_TryValidateNoDuplicatePerson = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to validate duplicate vacation ticket | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to validate duplicate vacation ticket: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Формирует следующий номер билета для приказа по уже внесённым строкам tbTickets.
Private Function private_Tickets_TryBuildNextTicketNo( _
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
    private_Tickets_TryBuildNextTicketNo = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to build next ticket number | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to build the next ticket number: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function
' --------------------------------------
' } // namespace Tickets
' --------------------------------------

' Собирает значения новой строки реестра tbTickets по именам его колонок.
Private Function private_Tickets_BuildRowValues( _
    ByVal rankText As String, _
    ByVal fioText As String, _
    ByVal ipnText As String, _
    ByVal positionCode As String, _
    ByVal eventText As String, _
    ByVal orderNo As String, _
    ByVal orderDate As Date, _
    ByVal vacationDays As Long, _
    ByVal roadDays As Long, _
    ByVal dateTo As Date, _
    ByVal ticketNo As String, _
    ByVal tvoFioText As String, _
    ByVal tvoIpnText As String, _
    ByVal tvoPositionCode As String _
) As Object
    Dim tableValues As Object

    Set tableValues = VBA.CreateObject("Scripting.Dictionary")
    tableValues.CompareMode = VBA.vbBinaryCompare
    tableValues.Add TICKETS_COL_RANK, rankText
    tableValues.Add TICKETS_COL_FIO, fioText
    tableValues.Add TICKETS_COL_IPN, ipnText
    tableValues.Add TICKETS_COL_POSITION, positionCode
    tableValues.Add TICKETS_COL_EVENT, eventText
    tableValues.Add TICKETS_COL_OUT_ORDER, orderNo
    tableValues.Add TICKETS_COL_OUT_FOOD, orderDate
    tableValues.Add TICKETS_COL_OUT_DATE, orderDate
    tableValues.Add TICKETS_COL_DURATION, vacationDays
    tableValues.Add TICKETS_COL_ROAD, roadDays
    tableValues.Add TICKETS_COL_ARRIVAL_PLAN, dateTo
    tableValues.Add TICKETS_COL_DOCUMENT, ticketNo
    tableValues.Add TICKETS_COL_TVO_FIO, tvoFioText
    tableValues.Add TICKETS_COL_TVO_IPN, tvoIpnText
    tableValues.Add TICKETS_COL_TVO_POSITION, tvoPositionCode
    Set private_Tickets_BuildRowValues = tableValues
End Function