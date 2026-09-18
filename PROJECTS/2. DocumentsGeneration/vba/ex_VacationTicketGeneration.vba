Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = ".log"
Private Const INPUT_SHEET_NAME As String = "Відпустки"
Private Const INPUT_MESSAGE_CELL_ADDRESS As String = "E7"

' Стабильные aliases полей формы. Адреса инкапсулированы в Input mapper-е.
Private Const INPUT_ALIAS_PERSON_LOOKUP As String = "PersonLookup"
Private Const INPUT_ALIAS_ORDER_REFERENCE As String = "OrderReference"
Private Const INPUT_ALIAS_DEPARTURE_DATE As String = "DepartureDate"
Private Const INPUT_ALIAS_VACATION_KIND As String = "VacationKind"
Private Const INPUT_ALIAS_VACATION_PLACE As String = "VacationPlace"
Private Const INPUT_ALIAS_VACATION_ABROAD As String = "VacationAbroad"
Private Const INPUT_ALIAS_VACATION_DAYS As String = "VacationDays"
Private Const INPUT_ALIAS_ROAD_DAYS As String = "RoadDays"
Private Const INPUT_ALIAS_DONATION_DAYS As String = "DonationDays"
Private Const INPUT_ALIAS_TVO_LOOKUP As String = "TvoLookup"
Private Const INPUT_ALIAS_STATUS As String = "Status"
Private Const INPUT_ALIAS_TEMPLATE_PATH As String = "TemplatePath"
Private Const INPUT_ALIAS_OUTPUT_FOLDER_PATH As String = "OutputFolderPath"
Private Const PERSONNEL_CANDIDATES_START_CELL_ADDRESS As String = "J4"
Private Const PERSONNEL_CANDIDATES_MAX_COUNT As Long = 20
Private Const PERSONNEL_CANDIDATES_TABLE_REF As String = "[АЛФ$A1:J12000]"
Private Const PERSONNEL_CANDIDATES_FIO_FIELD As String = "ПІБ"
Private Const PERSONNEL_CANDIDATES_IPN_FIELD As String = "ІПН"
Private Const PERSONNEL_CANDIDATES_FONT_NAME As String = "Times New Roman"
Private Const PERSONNEL_CANDIDATES_FONT_SIZE As Long = 11
' -4108 соответствует Excel-константе xlCenter.
Private Const PERSONNEL_CANDIDATES_HORIZONTAL_ALIGNMENT As Long = -4108
Private Const PERSONNEL_CANDIDATES_VERTICAL_ALIGNMENT As Long = -4108

' Канонические типы отпусков и отдельные тексты для Word и реестра Квитки.
Private Const VACATION_KIND_ANNUAL As String = "Щорічна відпустка"
Private Const VACATION_KIND_ANNUAL_WORD_TEXT As String = "у частину щорічної основної відпустки"
Private Const VACATION_KIND_ANNUAL_TICKETS_TEXT As String = VACATION_KIND_ANNUAL

Private Const VACATION_KIND_FAMILY As String = "Відпустка за сімейними обставинами"
Private Const VACATION_KIND_FAMILY_WORD_TEXT As String = "у відпустку за сімейними обставинами"
Private Const VACATION_KIND_FAMILY_TICKETS_TEXT As String = VACATION_KIND_FAMILY

Private Const VACATION_KIND_TREATMENT As String = "Відпустка для лікування після поранення (контузії, травми або каліцтва)"
Private Const VACATION_KIND_TREATMENT_WORD_TEXT As String = "у відпустку для лікування після поранення (контузії, травми або каліцтва)"
Private Const VACATION_KIND_TREATMENT_TICKETS_TEXT As String = "Відпустка для лікування"

Private Const VACATION_KIND_MATERNITY As String = "Відпустка у зв'язку з вагітністю та пологами"
Private Const VACATION_KIND_MATERNITY_WORD_TEXT As String = "у відпустку у зв'язку з вагітністю та пологами"
Private Const VACATION_KIND_MATERNITY_TICKETS_TEXT As String = VACATION_KIND_MATERNITY

Private Const VACATION_KIND_CHILDCARE As String = "Відпустка по догляду за дитиною"
Private Const VACATION_KIND_CHILDCARE_WORD_TEXT As String = "у відпустку по догляду за дитиною"
Private Const VACATION_KIND_CHILDCARE_TICKETS_TEXT As String = VACATION_KIND_CHILDCARE

Private Const VACATION_KIND_DONATION_TICKETS_TEXT As String = "Відпочинок за донацію крові"

' Допустимые значения признака выезда за границу в форме.
Private Const VACATION_ABROAD_YES As String = "Так"
Private Const VACATION_ABROAD_NO As String = "Ні"
Private Const VACATION_ABROAD_TEXT_YES As String = "Дозволено виїзд за кордон"

' Допустимые статусы отпускного билета и их представление в реестре.
Private Const VACATION_STATUS_ACTIVE As String = "Активна"
Private Const VACATION_STATUS_CANCELLED As String = "Скасовано"
Private Const VACATION_STATUS_TICKETS_CANCELLED As String = "СКАСОВАНО"

' Форматы дат для Word-шаблона. Текст «до 08:00 год.» находится в шаблоне.
Private Const DATE_FORMAT_PATTERN As String = """{dd}"" {month} {yyyy} р."

' Стабильный шаблон имени созданного документа.
Private Const DOCUMENT_NAME_PATTERN_ACTIVE As String = _
    "В.к. {TicketNo} {FIO} ({IPN})"
Private Const DOCUMENT_NAME_PATTERN_CANCELLED As String = _
    "В.к. {TicketNo} (СКАСОВАНО) {FIO} ({IPN})"

' Aliases контекста имени генерируемого документа.
Private Const GENERATED_CONTEXT_ALIAS_FIO As String = "FIO"
Private Const GENERATED_CONTEXT_ALIAS_IPN As String = "IPN"
Private Const GENERATED_CONTEXT_ALIAS_TICKET_NO As String = "TicketNo"

Private inputCellMap As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_VacationTicketGeneration_Create()
    private_Generate False
End Sub

' Обновляет существующий билет, найденный по ИПН и номеру приказа.
Public Sub fn_VacationTicketGeneration_Update()
    private_Generate True
End Sub

' Инициализирует зависимости формы, нужные до запуска генерации документа.
Public Function fn_TryInitializeUiRuntime() As Boolean
    fn_TryInitializeUiRuntime = ex_Helpers.ex_TryConfigureLogFileSuffix( _
        LOG_FILE_SUFFIX)
End Function

' Возвращает конфигурацию универсального поиска кандидатов формы отпуска.
Public Function fn_TryGetCandidatesConfig( _
    ByRef outCandidatesConfig As Object _
) As Boolean
    Dim lookupCellAddresses As Collection
    Dim columns As Collection
    Dim columnConfig As Object
    Dim styleConfig As Object

    Set outCandidatesConfig = Nothing
    private_Initialize
    If inputCellMap Is Nothing Then
        VBA.MsgBox "Vacation input cell map is not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not inputCellMap.Exists(INPUT_ALIAS_PERSON_LOOKUP) Or _
       Not inputCellMap.Exists(INPUT_ALIAS_TVO_LOOKUP) Then
        VBA.MsgBox "Vacation personnel lookup cells are not configured.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    Set lookupCellAddresses = New Collection
    lookupCellAddresses.Add VBA.CStr(inputCellMap( _
        INPUT_ALIAS_PERSON_LOOKUP))
    lookupCellAddresses.Add VBA.CStr(inputCellMap( _
        INPUT_ALIAS_TVO_LOOKUP))
    ' Порядок элементов Columns — порядок колонок от CandidateStartCellAddress.
    ' SourceIndex — индекс значения в массиве, возвращённом query callback.
    Set columns = New Collection
    Set columnConfig = VBA.CreateObject("Scripting.Dictionary")
    columnConfig.CompareMode = VBA.vbBinaryCompare
    columnConfig.Add "SourceIndex", 0
    columnConfig.Add "NumberFormat", "General"
    columns.Add columnConfig
    Set columnConfig = VBA.CreateObject("Scripting.Dictionary")
    columnConfig.CompareMode = VBA.vbBinaryCompare
    columnConfig.Add "SourceIndex", 1
    columnConfig.Add "NumberFormat", "@"
    columns.Add columnConfig
    Set styleConfig = VBA.CreateObject("Scripting.Dictionary")
    styleConfig.CompareMode = VBA.vbBinaryCompare
    styleConfig.Add "FontColor", VBA.RGB(255, 255, 255)
    styleConfig.Add "FillColor", VBA.RGB(0, 96, 32)
    styleConfig.Add "SelectedFillColor", VBA.RGB(112, 0, 56)
    styleConfig.Add "FontName", PERSONNEL_CANDIDATES_FONT_NAME
    styleConfig.Add "FontSize", PERSONNEL_CANDIDATES_FONT_SIZE
    styleConfig.Add "HorizontalAlignment", _
        PERSONNEL_CANDIDATES_HORIZONTAL_ALIGNMENT
    styleConfig.Add "VerticalAlignment", _
        PERSONNEL_CANDIDATES_VERTICAL_ALIGNMENT
    styleConfig.Add "WrapText", True
    Set outCandidatesConfig = VBA.CreateObject("Scripting.Dictionary")
    outCandidatesConfig.CompareMode = VBA.vbBinaryCompare
    outCandidatesConfig.Add "InputSheetName", INPUT_SHEET_NAME
    outCandidatesConfig.Add "LookupCellAddresses", lookupCellAddresses
    outCandidatesConfig.Add "CandidateStartCellAddress", PERSONNEL_CANDIDATES_START_CELL_ADDRESS
    outCandidatesConfig.Add "MaxCandidateCount", PERSONNEL_CANDIDATES_MAX_COUNT
    outCandidatesConfig.Add "QueryCallbackName", "ex_VacationTicketGeneration.fn_TryFindPersonnelCandidates"
    outCandidatesConfig.Add "SelectedValueIndex", 0
    outCandidatesConfig.Add "Columns", columns
    outCandidatesConfig.Add "Style", styleConfig
    fn_TryGetCandidatesConfig = True
End Function

' Выполняет предметный SQL-запрос формы отпуска и возвращает пары ПІБ/ІПН.
Public Function fn_TryFindPersonnelCandidates( _
    ByVal searchText As String, _
    ByVal maxCandidateCount As Long _
) As Collection
    Dim normalizedSearchText As String
    Dim sqlText As String
    Dim candidates As Collection
    Dim candidateFieldNames As Collection

    normalizedSearchText = ex_Helpers.private_Text_Normalize(searchText)
    If VBA.Len(normalizedSearchText) = 0 Then
        Set candidates = New Collection
        Set fn_TryFindPersonnelCandidates = candidates
        Exit Function
    End If
    If maxCandidateCount <= 0 Then
        VBA.MsgBox "Maximum candidate count must be greater than zero.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    sqlText = "SELECT TOP " & VBA.CStr(maxCandidateCount) & " [" & _
        PERSONNEL_CANDIDATES_FIO_FIELD & "], [" & _
        PERSONNEL_CANDIDATES_IPN_FIELD & "] FROM " & _
        PERSONNEL_CANDIDATES_TABLE_REF & " WHERE UCASE(TRIM(CSTR(IIF(ISNULL([" & _
        PERSONNEL_CANDIDATES_FIO_FIELD & "]), '', [" & _
        PERSONNEL_CANDIDATES_FIO_FIELD & "])))) LIKE '%" & _
        ex_ExternalTables.ex_EscapeSql(VBA.UCase$(normalizedSearchText)) & _
        "%' OR UCASE(TRIM(CSTR(IIF(ISNULL([" & _
        PERSONNEL_CANDIDATES_IPN_FIELD & "]), '', [" & _
        PERSONNEL_CANDIDATES_IPN_FIELD & "])))) LIKE '%" & _
        ex_ExternalTables.ex_EscapeSql(VBA.UCase$(normalizedSearchText)) & _
        "%' ORDER BY [" & PERSONNEL_CANDIDATES_FIO_FIELD & "]"
    ex_Helpers.LogDebug "Vacation candidate query SQL: " & sqlText
    Set candidateFieldNames = New Collection
    candidateFieldNames.Add PERSONNEL_CANDIDATES_FIO_FIELD
    candidateFieldNames.Add PERSONNEL_CANDIDATES_IPN_FIELD
    If Not ex_PersonnelData.ex_TryExecuteShpoCandidateSql( _
        sqlText, candidateFieldNames, candidates) Then Exit Function
    Set fn_TryFindPersonnelCandidates = candidates
End Function

Private Sub private_Generate(ByVal isUpdateMode As Boolean)
    Dim personLookup As String, tvoLookup As String
    Dim tvoIpnText As String, tvoFioText As String
    Dim orderReference As String
    Dim vacationKind As String, vacationKindText As String
    Dim vacationRegistryText As String
    Dim vacationAbroad As String, vacationAbroadText As String
    Dim vacationStatus As String, ticketsStatusText As String
    Dim vacationPlace As String
    Dim templatePath As String, outputFolderPath As String
    Dim ipnText As String, fioText As String
    Dim rankText As String, orderNo As String, donationOrderNo As String
    Dim personPositionCode As String, tvoPositionCode As String
    Dim orderDate As Date, departureDate As Date
    Dim mainReturnDate As Date, finalReturnDate As Date, dateArrival As Date
    Dim vacationDays As Long, roadDays As Long, donationDays As Long
    Dim ticketNo As String, ticketDateText As String
    Dim dateFromText As String, dateToText As String
    Dim dateArrivalText As String, personalLine As String, personalInitials As String
    Dim placeholderNames As Variant, placeholderValues As Variant
    Dim documentPath As String, matchedDocumentPath As String
    Dim overwriteDocumentPath As String
    Dim documentNamePattern As String
    Dim documentNameValues As Object
    Dim ticketsTable As ListObject, ticketRow As ListRow, donationTicketRow As ListRow
    Dim performanceStart As Single
    Dim personnelSessionStarted As Boolean
    Dim overwriteExistingDocument As Boolean
    Dim archiveExistingDocument As Boolean
    Dim restoreMissingDocument As Boolean
    Dim mainTicketRowCreated As Boolean
    Dim donationTicketRowCreated As Boolean
    Dim registryRowsSaved As Boolean
    Dim operationCompleted As Boolean

    On Error GoTo EH
    If Not ex_Helpers.ex_TryConfigureLogFileSuffix(LOG_FILE_SUFFIX) Then Exit Sub
    private_Initialize

    If Not ex_Helpers.ex_TryConfigureMessageTarget( _
        INPUT_SHEET_NAME, INPUT_MESSAGE_CELL_ADDRESS) Then Exit Sub
    ex_Helpers.ClearLog
    performanceStart = VBA.Timer
    private_Performance_LogCheckpoint performanceStart, "Start"
    ex_Helpers.LogDebug "Vacation ticket generation started | UpdateMode=" & _
        VBA.CStr(isUpdateMode)
    Call ex_Document.ex_LogWorkbookContext(INPUT_SHEET_NAME, inputCellMap)

    personLookup = private_Input_ReadRequired(INPUT_ALIAS_PERSON_LOOKUP, "ПІБ або ІПН")
    vacationKind = private_Input_ReadRequired(INPUT_ALIAS_VACATION_KIND, "вид відпустки")
    vacationPlace = private_Input_ReadRequired(INPUT_ALIAS_VACATION_PLACE, "місце відпустки")
    vacationAbroad = private_Input_ReadRequired( _
        INPUT_ALIAS_VACATION_ABROAD, "відпустка за кордон")
    orderReference = private_Input_ReadRequired(INPUT_ALIAS_ORDER_REFERENCE, "номер або дату наказу")
    If Not private_Input_TryReadDate( _
        INPUT_ALIAS_DEPARTURE_DATE, "дату вибуття", departureDate) Then GoTo CleanExit
    If Not private_Input_TryReadOptional( _
        INPUT_ALIAS_TVO_LOOKUP, tvoLookup) Then GoTo CleanExit
    vacationStatus = private_Input_ReadRequired(INPUT_ALIAS_STATUS, "статус")
    templatePath = private_Input_ReadRequired(INPUT_ALIAS_TEMPLATE_PATH, "шлях до шаблону")
    outputFolderPath = private_Input_ReadRequired( _
        INPUT_ALIAS_OUTPUT_FOLDER_PATH, "шлях до папки результатів")
    If VBA.Len(personLookup) = 0 Or VBA.Len(orderReference) = 0 Or _
        VBA.Len(vacationKind) = 0 Or VBA.Len(vacationPlace) = 0 Or _
        VBA.Len(vacationAbroad) = 0 Or _
        VBA.Len(vacationStatus) = 0 Or _
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
    If Not private_Vacation_TryMapKind( _
        vacationKind, vacationKindText, vacationRegistryText) Then GoTo CleanExit
    If Not private_Vacation_TryMapAbroad( _
        vacationAbroad, vacationAbroadText) Then GoTo CleanExit
    If Not private_Vacation_TryMapStatus( _
        vacationStatus, ticketsStatusText) Then GoTo CleanExit
    If VBA.StrComp(ticketsStatusText, VACATION_STATUS_TICKETS_CANCELLED, _
            VBA.vbTextCompare) = 0 Then
        documentNamePattern = DOCUMENT_NAME_PATTERN_CANCELLED
    Else
        documentNamePattern = DOCUMENT_NAME_PATTERN_ACTIVE
    End If

    If Not ex_PersonnelData.ex_TryBeginSession() Then GoTo CleanExit
    personnelSessionStarted = True
    private_Performance_LogCheckpoint performanceStart, "SHPO session opened"
    If Not ex_PersonnelData.ex_TryResolveIpn(personLookup, ipnText) Then GoTo CleanExit
    ' ТВО не обязателен: пустое поле оставляет колонки ТВО реестра пустыми.
    If VBA.Len(tvoLookup) > 0 Then
        If Not ex_PersonnelData.ex_TryResolveIpn(tvoLookup, tvoIpnText) Then GoTo CleanExit
        If Not ex_PersonnelData.ex_TryResolveFioNominative( _
            tvoIpnText, tvoFioText) Then GoTo CleanExit
        If Not ex_PersonnelData.ex_TryResolvePositionCode( _
            tvoIpnText, tvoPositionCode) Then GoTo CleanExit
    End If
    If Not ex_PersonnelData.ex_TryResolveFioNominative(ipnText, fioText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveRankNominative(ipnText, rankText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolvePositionCode( _
        ipnText, personPositionCode) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveOrderReference( _
        orderReference, orderNo, orderDate) Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Personnel and order data resolved"
    If Not ex_Tickets.ex_TryGetOpenTable(ticketsTable) Then GoTo CleanExit
    If Not ex_Tickets.ex_TryFindVacationRow( _
        ticketsTable, ipnText, orderNo, VBA.vbNullString, ticketRow) Then GoTo CleanExit
    If isUpdateMode Then
        If ticketRow Is Nothing Then
            ex_Helpers.ex_ShowErrorMessage "Vacation ticket was not found for IPN '" & _
                ipnText & "' and order '" & orderNo & "'.", _
                VBA.vbExclamation, "Document Generation"
            GoTo CleanExit
        End If
        If Not ex_Tickets.ex_TryReadTicketNo( _
            ticketsTable, ticketRow, ticketNo) Then GoTo CleanExit
        If Not ex_Helpers.private_Path_TryFindVacationTicketDocument( _
            outputFolderPath, ticketNo, ipnText, matchedDocumentPath) Then GoTo CleanExit
        private_Performance_LogCheckpoint performanceStart, "Registry row found for update"
    Else
        If Not ticketRow Is Nothing Then
            If Not ex_Tickets.ex_TryReadTicketNo( _
                ticketsTable, ticketRow, ticketNo) Then GoTo CleanExit
            If Not ex_Helpers.private_Path_TryFindVacationTicketDocument( _
                outputFolderPath, ticketNo, ipnText, matchedDocumentPath, False) Then GoTo CleanExit
            If VBA.Len(matchedDocumentPath) > 0 Then
                ex_Helpers.ex_ShowErrorMessage "A vacation ticket already exists for " & _
                    "this person and order '" & orderNo & "'. " & _
                    "Use Update data instead.", VBA.vbExclamation, "Document Generation"
                GoTo CleanExit
            End If
            If VBA.MsgBox("A vacation ticket record already exists in tbTickets, " & _
                "but its Word file was not found." & VBA.vbCrLf & VBA.vbCrLf & _
                "Create the Word file from the current form data?", _
                VBA.vbYesNo + VBA.vbQuestion, "Document Generation") <> VBA.vbYes Then
                GoTo CleanExit
            End If
            restoreMissingDocument = True
        Else
            If Not ex_Tickets.ex_TryBuildNextTicketNo( _
                ticketsTable, orderNo, orderDate, ticketNo) Then GoTo CleanExit
            private_Performance_LogCheckpoint performanceStart, "New ticket number assigned"
        End If
    End If
    If Not ex_Tickets.ex_TryFindDonationRow( _
        ticketsTable, ipnText, ticketNo, donationTicketRow) Then GoTo CleanExit

    ' Даты рассчитываются от указанной пользователем даты выбытия.
    mainReturnDate = VBA.DateAdd("d", vacationDays + roadDays, departureDate)
    finalReturnDate = VBA.DateAdd("d", donationDays, mainReturnDate)
    dateArrival = VBA.DateAdd("d", 1, finalReturnDate)
    If donationDays > 0 Then
        If Not ex_PersonnelData.ex_TryFindOrderNoByDate( _
            mainReturnDate, donationOrderNo) Then GoTo CleanExit
        If VBA.Len(donationOrderNo) = 0 Then donationOrderNo = "NNN"
        donationOrderNo = "(?) " & donationOrderNo
    End If
    ex_Helpers.LogDebug "Vacation period | From=" & VBA.CStr(departureDate) & _
        " | MainReturn=" & VBA.CStr(mainReturnDate) & _
        " | FinalReturn=" & VBA.CStr(finalReturnDate) & _
        " | Arrival=" & VBA.CStr(dateArrival)
    If Not ex_Helpers.private_Date_TryFormat( _
        orderDate, DATE_FORMAT_PATTERN, ticketDateText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        departureDate, DATE_FORMAT_PATTERN, dateFromText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        finalReturnDate, DATE_FORMAT_PATTERN, dateToText) Then GoTo CleanExit
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
        "PersonalInitials", "DateArrival", "VacationAbroadText", _
        "FooterSpacer")
    placeholderValues = Array( _
        ticketDateText, ticketNo, personalLine, vacationKindText, vacationPlace, _
        private_Vacation_BuildWordDuration(vacationDays, donationDays, roadDays), dateFromText, _
        dateToText, personalInitials, dateArrivalText, vacationAbroadText, _
        VBA.vbCr)

    Set documentNameValues = VBA.CreateObject("Scripting.Dictionary")
    documentNameValues.CompareMode = VBA.vbBinaryCompare
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_FIO, fioText
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_IPN, ipnText
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_TICKET_NO, ticketNo
    If Not isUpdateMode Then
        If Not ex_Helpers.private_Path_TryFindVacationTicketDocument( _
            outputFolderPath, ticketNo, ipnText, matchedDocumentPath, False) Then GoTo CleanExit
        If VBA.Len(matchedDocumentPath) > 0 Then
            If VBA.MsgBox("A document already exists for this vacation ticket:" & _
                VBA.vbCrLf & matchedDocumentPath & VBA.vbCrLf & VBA.vbCrLf & _
                "Replace it with the new document?", _
                VBA.vbYesNo + VBA.vbQuestion, "Document Generation") = VBA.vbYes Then
                overwriteExistingDocument = True
            Else
                archiveExistingDocument = True
            End If
        End If
        If overwriteExistingDocument Then _
            overwriteDocumentPath = matchedDocumentPath
    Else
        overwriteDocumentPath = matchedDocumentPath
    End If
    If Not ex_Helpers.private_Path_TryEnsureDocumentReadable( _
        templatePath) Then GoTo CleanExit
    If Not ex_Helpers.private_Path_TryProbeTemplateCopy( _
        templatePath, outputFolderPath) Then GoTo CleanExit
    If VBA.Len(matchedDocumentPath) > 0 Then
        If Not ex_Helpers.private_Path_TryEnsureDocumentWritable( _
            matchedDocumentPath) Then GoTo CleanExit
    End If
    mainTicketRowCreated = (ticketRow Is Nothing)
    registryRowsSaved = True
    If Not ex_Tickets.ex_TrySaveVacationRow( _
        ticketsTable, ticketRow, rankText, fioText, ipnText, personPositionCode, _
        vacationRegistryText, orderNo, orderDate, departureDate, vacationDays, roadDays, _
        ticketNo, tvoFioText, tvoIpnText, tvoPositionCode, ticketsStatusText) Then GoTo CleanExit
    If donationDays > 0 Then
        donationTicketRowCreated = (donationTicketRow Is Nothing)
        If Not ex_Tickets.ex_TrySaveVacationRow( _
            ticketsTable, donationTicketRow, rankText, fioText, ipnText, personPositionCode, _
            VACATION_KIND_DONATION_TICKETS_TEXT, donationOrderNo, finalReturnDate, _
            mainReturnDate, donationDays, 0, ticketNo, tvoFioText, tvoIpnText, _
            tvoPositionCode, ticketsStatusText) Then GoTo CleanExit
    End If
    private_Performance_LogCheckpoint performanceStart, "Registry row saved"
    If archiveExistingDocument Then
        If Not ex_Helpers.private_Path_TryArchiveDocument( _
            matchedDocumentPath) Then GoTo CleanExit
    End If
    private_Performance_LogCheckpoint performanceStart, "Word generation started"
    If Not ex_Document.ex_TryGenerateWordDocument( _
        templatePath, documentNamePattern, documentNameValues, _
        placeholderNames, placeholderValues, documentPath, _
        outputFolderPath, overwriteDocumentPath, _
        Not isUpdateMode Or restoreMissingDocument, _
        isUpdateMode Or overwriteExistingDocument) Then GoTo CleanExit
    private_Performance_LogCheckpoint performanceStart, "Word generation completed"
    If donationDays = 0 And Not donationTicketRow Is Nothing Then
        If Not ex_Tickets.ex_TryDeleteVacationRow( _
            donationTicketRow) Then GoTo CleanExit
    End If

    ex_Helpers.WriteLog "DOCUMENT: " & documentPath
    ex_Helpers.LogDebug "Vacation ticket generation completed"
    If isUpdateMode Then
        ex_Helpers.ex_ShowStatusMessage "Vacation ticket updated: " & documentPath
    Else
        ex_Helpers.ex_ShowStatusMessage "Vacation ticket generated: " & documentPath
    End If
    operationCompleted = True
    GoTo CleanExit

CleanExit:
    If registryRowsSaved And Not operationCompleted Then
        private_TryRollbackCreatedTicketRows ticketRow, mainTicketRowCreated, _
            donationTicketRow, donationTicketRowCreated
    End If
    If personnelSessionStarted Then ex_PersonnelData.ex_EndSession
    ex_Helpers.ex_ClearMessageTarget
    Exit Sub
EH:
    ex_Helpers.LogError "Vacation ticket generation failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Vacation ticket generation failed: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Sub

' Удаляет только строки, добавленные в текущем запуске после неудачной генерации Word.
Private Sub private_TryRollbackCreatedTicketRows( _
    ByVal ticketRow As ListRow, _
    ByVal mainTicketRowCreated As Boolean, _
    ByVal donationTicketRow As ListRow, _
    ByVal donationTicketRowCreated As Boolean _
)
    On Error GoTo EH
    If donationTicketRowCreated Then donationTicketRow.Delete
    If mainTicketRowCreated Then ticketRow.Delete
    ex_Helpers.LogDebug "Created ticket rows were rolled back after a failed operation"
    Exit Sub
EH:
    ex_Helpers.LogError "Failed to roll back created ticket rows | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Word generation failed and the newly created " & _
        "tbTickets rows could not be removed: " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
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
    inputCellMap.Add INPUT_ALIAS_VACATION_KIND, "C5"
    inputCellMap.Add INPUT_ALIAS_VACATION_PLACE, "C6"
    inputCellMap.Add INPUT_ALIAS_VACATION_ABROAD, "C7"
    inputCellMap.Add INPUT_ALIAS_ORDER_REFERENCE, "C8"
    inputCellMap.Add INPUT_ALIAS_DEPARTURE_DATE, "C9"
    inputCellMap.Add INPUT_ALIAS_VACATION_DAYS, "C10"
    inputCellMap.Add INPUT_ALIAS_ROAD_DAYS, "C11"
    inputCellMap.Add INPUT_ALIAS_DONATION_DAYS, "C12"
    inputCellMap.Add INPUT_ALIAS_TVO_LOOKUP, "C13"
    inputCellMap.Add INPUT_ALIAS_STATUS, "C14"
    inputCellMap.Add INPUT_ALIAS_TEMPLATE_PATH, "C15"
    inputCellMap.Add INPUT_ALIAS_OUTPUT_FOLDER_PATH, "C16"
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

Private Function private_Input_TryReadOptional( _
    ByVal fieldAlias As String, _
    ByRef outValue As String _
) As Boolean
    private_Input_TryReadOptional = ex_Document.ex_TryReadOptional( _
        INPUT_SHEET_NAME, inputCellMap, fieldAlias, outValue)
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

Private Function private_Input_TryReadDate( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByRef outDate As Date _
) As Boolean
    Dim dateText As String

    If Not ex_Document.ex_TryReadRequired( _
        INPUT_SHEET_NAME, inputCellMap, fieldAlias, fieldCaption, dateText) Then Exit Function
    If ex_Helpers.private_Date_TryParse(dateText, outDate) Then
        private_Input_TryReadDate = True
        Exit Function
    End If
    ex_Helpers.LogError "Invalid required date | Field=" & fieldAlias & _
        " | Value=" & dateText
    ex_Helpers.ex_ShowErrorMessage "Enter a valid date for " & fieldCaption & ".", _
        VBA.vbExclamation, "Document Generation"
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
Private Function private_Vacation_BuildWordDuration( _
    ByVal vacationDays As Long, _
    ByVal donationDays As Long, _
    ByVal roadDays As Long _
) As String
    Dim durationText As String

    durationText = VBA.CStr(vacationDays) & " " & _
        ex_Helpers.ex_GetUkrainianCountForm( _
            vacationDays, "календарний день", "календарні дні", _
            "календарних днів")
    If donationDays > 0 Then
        durationText = durationText & " та " & VBA.CStr(donationDays) & " " & _
            ex_Helpers.ex_GetUkrainianCountForm( _
                donationDays, "додатковий день", "додаткові дні", _
                "додаткових днів") & _
            " відпочинку за донацію донорської крові та/або компонентів крові"
    End If
    If roadDays > 0 Then
        durationText = durationText & " та " & VBA.CStr(roadDays) & " " & _
            ex_Helpers.ex_GetUkrainianCountForm( _
                roadDays, "добу", "доби", "діб") & " на дорогу"
    End If
    private_Vacation_BuildWordDuration = durationText
End Function

Private Function private_Vacation_TryMapKind( _
    ByVal vacationKind As String, _
    ByRef outWordText As String, _
    ByRef outRegistryText As String _
) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(vacationKind))
        Case VBA.LCase$(VACATION_KIND_ANNUAL)
            outWordText = VACATION_KIND_ANNUAL_WORD_TEXT
            outRegistryText = VACATION_KIND_ANNUAL_TICKETS_TEXT
        Case VBA.LCase$(VACATION_KIND_FAMILY)
            outWordText = VACATION_KIND_FAMILY_WORD_TEXT
            outRegistryText = VACATION_KIND_FAMILY_TICKETS_TEXT
        Case VBA.LCase$(VACATION_KIND_TREATMENT)
            outWordText = VACATION_KIND_TREATMENT_WORD_TEXT
            outRegistryText = VACATION_KIND_TREATMENT_TICKETS_TEXT
        Case VBA.LCase$(VACATION_KIND_MATERNITY)
            outWordText = VACATION_KIND_MATERNITY_WORD_TEXT
            outRegistryText = VACATION_KIND_MATERNITY_TICKETS_TEXT
        Case VBA.LCase$(VACATION_KIND_CHILDCARE)
            outWordText = VACATION_KIND_CHILDCARE_WORD_TEXT
            outRegistryText = VACATION_KIND_CHILDCARE_TICKETS_TEXT
        Case Else
            ex_Helpers.LogError "Unsupported vacation kind: " & vacationKind
            ex_Helpers.ex_ShowErrorMessage "Unsupported vacation kind: " & vacationKind, _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Vacation_TryMapKind = True
End Function

Private Function private_Vacation_TryMapAbroad( _
    ByVal vacationAbroad As String, _
    ByRef outVacationAbroadText As String _
) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(vacationAbroad))
        Case VBA.LCase$(VACATION_ABROAD_YES)
            outVacationAbroadText = VACATION_ABROAD_TEXT_YES
        Case VBA.LCase$(VACATION_ABROAD_NO)
            ' Сохраняем отдельную строку шаблона при отсутствии разрешения.
            outVacationAbroadText = VBA.vbCr
        Case Else
            ex_Helpers.LogError "Unsupported vacation abroad value: " & vacationAbroad
            ex_Helpers.ex_ShowErrorMessage _
                "Vacation abroad value must be 'Так' or 'Ні': " & vacationAbroad, _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Vacation_TryMapAbroad = True
End Function

Private Function private_Vacation_TryMapStatus( _
    ByVal vacationStatus As String, _
    ByRef outTicketsStatusText As String _
) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(vacationStatus))
        Case VBA.LCase$(VACATION_STATUS_ACTIVE)
            ' Активный билет не требует отметки в реестре.
            outTicketsStatusText = VBA.vbNullString
        Case VBA.LCase$(VACATION_STATUS_CANCELLED)
            outTicketsStatusText = VACATION_STATUS_TICKETS_CANCELLED
        Case Else
            ex_Helpers.LogError "Unsupported vacation status: " & vacationStatus
            ex_Helpers.ex_ShowErrorMessage _
                "Vacation status must be 'Активна' or 'Скасовано': " & vacationStatus, _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Vacation_TryMapStatus = True
End Function
' --------------------------------------
' } // namespace Vacation
' --------------------------------------