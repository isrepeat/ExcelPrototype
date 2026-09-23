Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = ".log"
Private Const CFG_WS_VACATION As String = "wsVacation"
Private Const CFG_MESSAGE_TARGET As String = "wsVacation::message_target"

' Stable form field aliases. Addresses stay inside the input mapper.
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
Private Const CFG_CANDIDATES_START As String = "wsVacation::candidates.start"
Private Const CFG_CANDIDATES_MAX_COUNT As String = "wsVacation::candidates.max_count"
Private Const CFG_PERSONNEL_ALF_RANGE As String = "rangePersonnelAlf"
Private Const CFG_PERSONNEL_ALF_FIO As String = "rangePersonnelAlf::column.fio"
Private Const CFG_PERSONNEL_ALF_IPN As String = "rangePersonnelAlf::column.ipn"
Private Const CFG_CANDIDATES_PERSON_LOOKUP_TITLE As String = "wsVacation::candidates.person_lookup.title"
Private Const CFG_CANDIDATES_TVO_LOOKUP_TITLE As String = "wsVacation::candidates.tvo_lookup.title"
Private Const CFG_CANDIDATES_TABLE_TITLE_PREFIX As String = "wsVacation::candidates.table_title_prefix"
Private Const CFG_CANDIDATES_HIDE_COMMAND As String = "wsVacation::candidates.hide_command"

' Vacation kind values are read from the workbook config.
Private Const CFG_KIND_ANNUAL_INPUT As String = "wsVacation::kind.annual.input"
Private Const CFG_KIND_ANNUAL_WORD As String = "wsVacation::kind.annual.word"
Private Const CFG_KIND_ANNUAL_TICKETS As String = "wsVacation::kind.annual.tickets"
Private Const CFG_KIND_FAMILY_INPUT As String = "wsVacation::kind.family.input"
Private Const CFG_KIND_FAMILY_WORD As String = "wsVacation::kind.family.word"
Private Const CFG_KIND_FAMILY_TICKETS As String = "wsVacation::kind.family.tickets"
Private Const CFG_KIND_TREATMENT_INPUT As String = "wsVacation::kind.treatment.input"
Private Const CFG_KIND_TREATMENT_WORD As String = "wsVacation::kind.treatment.word"
Private Const CFG_KIND_TREATMENT_TICKETS As String = "wsVacation::kind.treatment.tickets"
Private Const CFG_KIND_MATERNITY_INPUT As String = "wsVacation::kind.maternity.input"
Private Const CFG_KIND_MATERNITY_WORD As String = "wsVacation::kind.maternity.word"
Private Const CFG_KIND_MATERNITY_TICKETS As String = "wsVacation::kind.maternity.tickets"
Private Const CFG_KIND_CHILDCARE_INPUT As String = "wsVacation::kind.childcare.input"
Private Const CFG_KIND_CHILDCARE_WORD As String = "wsVacation::kind.childcare.word"
Private Const CFG_KIND_CHILDCARE_TICKETS As String = "wsVacation::kind.childcare.tickets"
Private Const CFG_KIND_DONATION_TICKETS As String = "tbTickets::event.donation"

Private Const CFG_ABROAD_YES As String = "wsVacation::abroad.yes"
Private Const CFG_ABROAD_NO As String = "wsVacation::abroad.no"
Private Const CFG_ABROAD_YES_WORD As String = "wsVacation::abroad.yes.word"
Private Const CFG_STATUS_ACTIVE As String = "wsVacation::status.active"
Private Const CFG_STATUS_CANCELLED As String = "wsVacation::status.cancelled"
Private Const CFG_STATUS_CANCELLED_TICKETS As String = "wsVacation::status.cancelled.tickets"
Private Const CFG_DATE_WORD As String = "wsVacation::date.word"
Private Const CFG_DOCUMENT_NAME_ACTIVE As String = "wsVacation::document_name.active"
Private Const CFG_DOCUMENT_NAME_CANCELLED As String = "wsVacation::document_name.cancelled"
Private Const CFG_DURATION_CONNECTOR As String = "wsVacation::duration.connector"
Private Const CFG_DURATION_VACATION_ONE As String = "wsVacation::duration.vacation.one"
Private Const CFG_DURATION_VACATION_FEW As String = "wsVacation::duration.vacation.few"
Private Const CFG_DURATION_VACATION_MANY As String = "wsVacation::duration.vacation.many"
Private Const CFG_DURATION_DONATION_ONE As String = "wsVacation::duration.donation.one"
Private Const CFG_DURATION_DONATION_FEW As String = "wsVacation::duration.donation.few"
Private Const CFG_DURATION_DONATION_MANY As String = "wsVacation::duration.donation.many"
Private Const CFG_DURATION_DONATION_SUFFIX As String = "wsVacation::duration.donation.suffix"
Private Const CFG_DURATION_ROAD_ONE As String = "wsVacation::duration.road.one"
Private Const CFG_DURATION_ROAD_FEW As String = "wsVacation::duration.road.few"
Private Const CFG_DURATION_ROAD_MANY As String = "wsVacation::duration.road.many"
Private Const CFG_DURATION_ROAD_SUFFIX As String = "wsVacation::duration.road.suffix"

' Context aliases for the generated document name.
Private Const GENERATED_CONTEXT_ALIAS_FIO As String = "FIO"
Private Const GENERATED_CONTEXT_ALIAS_IPN As String = "IPN"
Private Const GENERATED_CONTEXT_ALIAS_TICKET_NO As String = "TicketNo"

Private inputCellMap As Object
Private vacationKindAnnualInput As String
Private vacationKindAnnualWord As String
Private vacationKindAnnualTickets As String
Private vacationKindFamilyInput As String
Private vacationKindFamilyWord As String
Private vacationKindFamilyTickets As String
Private vacationKindTreatmentInput As String
Private vacationKindTreatmentWord As String
Private vacationKindTreatmentTickets As String
Private vacationKindMaternityInput As String
Private vacationKindMaternityWord As String
Private vacationKindMaternityTickets As String
Private vacationKindChildcareInput As String
Private vacationKindChildcareWord As String
Private vacationKindChildcareTickets As String
Private vacationKindDonationTickets As String
Private vacationSheetName As String
Private personnelAlfRange As String
Private personnelAlfFioColumn As String
Private personnelAlfIpnColumn As String
Private candidatesStartCellAddress As String
Private candidatesMaxCount As Long
Private candidatesPersonLookupTitle As String
Private candidatesTvoLookupTitle As String
Private candidatesTableTitlePrefix As String
Private candidatesHideCommand As String
Private vacationAbroadYes As String
Private vacationAbroadNo As String
Private vacationAbroadYesWord As String
Private vacationStatusActive As String
Private vacationStatusCancelled As String
Private vacationStatusCancelledTickets As String
Private dateWordPattern As String
Private documentNamePatternActive As String
Private documentNamePatternCancelled As String
Private durationConnector As String
Private durationVacationOne As String
Private durationVacationFew As String
Private durationVacationMany As String
Private durationDonationOne As String
Private durationDonationFew As String
Private durationDonationMany As String
Private durationDonationSuffix As String
Private durationRoadOne As String
Private durationRoadFew As String
Private durationRoadMany As String
Private durationRoadSuffix As String

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_VacationTicketGeneration_Create()
    private_Generate False
End Sub

' Updates an existing ticket found by IPN and order number.
Public Sub fn_VacationTicketGeneration_Update()
    private_Generate True
End Sub

' Initializes form dependencies before document generation.
Public Function fn_TryInitializeUiRuntime() As Boolean
    fn_TryInitializeUiRuntime = ex_Helpers.ex_TryConfigureLogFileSuffix( _
        LOG_FILE_SUFFIX)
End Function

' Returns common candidate-search configuration for the vacation form.
Public Function fn_TryGetCandidatesConfig( _
    ByRef outCandidatesConfig As Object _
) As Boolean
    Dim lookupCellAddresses As Collection
    Dim lookupFieldTitles As Object
    Dim columns As Collection
    Dim columnConfig As Object
    Dim styles As Object
    Dim titleStyle As Object

    Set outCandidatesConfig = Nothing
    If Not private_Initialize() Then Exit Function
    If inputCellMap Is Nothing Then
        Call ex_Helpers.ex_ShowMessageBox( _
            "Vacation input cell map is not initialized.", VBA.vbExclamation, _
            "Document Generation")
        Exit Function
    End If
    If Not inputCellMap.Exists(INPUT_ALIAS_PERSON_LOOKUP) Or _
       Not inputCellMap.Exists(INPUT_ALIAS_TVO_LOOKUP) Then
        Call ex_Helpers.ex_ShowMessageBox( _
            "Vacation personnel lookup cells are not configured.", VBA.vbExclamation, _
            "Document Generation")
        Exit Function
    End If

    Set lookupCellAddresses = New Collection
    lookupCellAddresses.Add VBA.CStr(inputCellMap( _
        INPUT_ALIAS_PERSON_LOOKUP))
    lookupCellAddresses.Add VBA.CStr(inputCellMap( _
        INPUT_ALIAS_TVO_LOOKUP))
    Set lookupFieldTitles = VBA.CreateObject("Scripting.Dictionary")
    lookupFieldTitles.CompareMode = VBA.vbBinaryCompare
    lookupFieldTitles.Add VBA.CStr(inputCellMap(INPUT_ALIAS_PERSON_LOOKUP)), _
        candidatesPersonLookupTitle
    lookupFieldTitles.Add VBA.CStr(inputCellMap(INPUT_ALIAS_TVO_LOOKUP)), _
        candidatesTvoLookupTitle
    ' CandidateStartCellAddress points to the left cell of the command row.
    ' Headers and data start one and two rows below it.
    ' SourceIndex is the value index in the query-callback result array.
    Set columns = New Collection
    Set columnConfig = VBA.CreateObject("Scripting.Dictionary")
    columnConfig.CompareMode = VBA.vbBinaryCompare
    columnConfig.Add "SourceIndex", 0
    columnConfig.Add "Header", personnelAlfFioColumn
    columnConfig.Add "NumberFormat", "General"
    columns.Add columnConfig
    Set columnConfig = VBA.CreateObject("Scripting.Dictionary")
    columnConfig.CompareMode = VBA.vbBinaryCompare
    columnConfig.Add "SourceIndex", 1
    columnConfig.Add "Header", personnelAlfIpnColumn
    columnConfig.Add "NumberFormat", "@"
    columns.Add columnConfig
    Set styles = VBA.CreateObject("Scripting.Dictionary")
    styles.CompareMode = VBA.vbBinaryCompare
    styles.Add "Candidate", private_CandidatesConfig_CreateCellStyle( _
        "Times New Roman", 16, VBA.RGB(255, 255, 255), VBA.RGB(0, 96, 32), _
        -4108, -4108, True)
    styles.Add "Chrome", private_CandidatesConfig_CreateCellStyle( _
        "Times New Roman", 16, VBA.RGB(255, 255, 255), VBA.RGB(0, 0, 0), _
        -4108, -4108, True)
    styles.Add "Command", private_CandidatesConfig_CreateCellStyle( _
        "Times New Roman", 16, VBA.RGB(244, 176, 132), VBA.RGB(64, 64, 64), _
        -4108, -4108, True)
    Set titleStyle = private_CandidatesConfig_CreateCellStyle( _
        "Times New Roman", 16, VBA.RGB(166, 166, 166), VBA.RGB(0, 0, 0), _
        -4108, -4108, True)
    titleStyle.Remove "FillColor"
    styles.Add "Title", titleStyle
    styles.Add "Selected", private_CandidatesConfig_CreateCellStyle( _
        "Times New Roman", 16, VBA.RGB(255, 255, 255), VBA.RGB(112, 0, 56), _
        -4108, -4108, True)
    Set outCandidatesConfig = VBA.CreateObject("Scripting.Dictionary")
    outCandidatesConfig.CompareMode = VBA.vbBinaryCompare
    outCandidatesConfig.Add "InputSheetName", vacationSheetName
    outCandidatesConfig.Add "LookupCellAddresses", lookupCellAddresses
    outCandidatesConfig.Add "LookupFieldTitles", lookupFieldTitles
    outCandidatesConfig.Add "CandidateStartCellAddress", candidatesStartCellAddress
    outCandidatesConfig.Add "TableTitlePrefix", candidatesTableTitlePrefix
    outCandidatesConfig.Add "HideCommandText", candidatesHideCommand
    outCandidatesConfig.Add "MaxCandidateCount", candidatesMaxCount
    outCandidatesConfig.Add "QueryCallbackName", "ex_VacationTicketGeneration.fn_TryFindPersonnelCandidates"
    outCandidatesConfig.Add "SelectedValueIndex", 0
    outCandidatesConfig.Add "Columns", columns
    outCandidatesConfig.Add "Styles", styles
    outCandidatesConfig.Add "TableStyleName", "Candidate"
    outCandidatesConfig.Add "ChromeStyleName", "Chrome"
    outCandidatesConfig.Add "CommandStyleName", "Command"
    outCandidatesConfig.Add "TitleStyleName", "Title"
    outCandidatesConfig.Add "SelectedStyleName", "Selected"
    fn_TryGetCandidatesConfig = True
End Function

' Runs the vacation-form SQL query and returns FIO/IPN pairs.
Public Function fn_TryFindPersonnelCandidates( _
    ByVal searchText As String, _
    ByVal maxCandidateCount As Long _
) As Collection
    Dim normalizedSearchText As String
    Dim sqlText As String
    Dim candidates As Collection
    Dim candidateFieldNames As Collection

    If Not private_Initialize() Then Exit Function
    normalizedSearchText = ex_Helpers.private_Text_Normalize(searchText)
    If VBA.Len(normalizedSearchText) = 0 Then
        Set candidates = New Collection
        Set fn_TryFindPersonnelCandidates = candidates
        Exit Function
    End If
    If maxCandidateCount <= 0 Then
        Call ex_Helpers.ex_ShowMessageBox( _
            "Maximum candidate count must be greater than zero.", VBA.vbExclamation, _
            "Document Generation")
        Exit Function
    End If
    sqlText = "SELECT TOP " & VBA.CStr(maxCandidateCount) & " [" & _
        personnelAlfFioColumn & "], [" & _
        personnelAlfIpnColumn & "] FROM " & _
        personnelAlfRange & " WHERE UCASE(TRIM(CSTR(IIF(ISNULL([" & _
        personnelAlfFioColumn & "]), '', [" & _
        personnelAlfFioColumn & "])))) LIKE '%" & _
        ex_ExternalTables.ex_EscapeSql(VBA.UCase$(normalizedSearchText)) & _
        "%' OR UCASE(TRIM(CSTR(IIF(ISNULL([" & _
        personnelAlfIpnColumn & "]), '', [" & _
        personnelAlfIpnColumn & "])))) LIKE '%" & _
        ex_ExternalTables.ex_EscapeSql(VBA.UCase$(normalizedSearchText)) & _
        "%' ORDER BY [" & personnelAlfFioColumn & "]"
    ex_Helpers.LogDebug "Vacation candidate query SQL: " & sqlText
    Set candidateFieldNames = New Collection
    candidateFieldNames.Add personnelAlfFioColumn
    candidateFieldNames.Add personnelAlfIpnColumn
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
    If Not private_Initialize() Then Exit Sub

    If Not ex_Helpers.ex_TryConfigureMessageTarget( _
        private_ConfigText(CFG_WS_VACATION), private_ConfigText(CFG_MESSAGE_TARGET)) Then Exit Sub
    ex_Helpers.ClearLog
    performanceStart = VBA.Timer
    private_Performance_LogCheckpoint performanceStart, "Start"
    ex_Helpers.LogDebug "Vacation ticket generation started | UpdateMode=" & _
        VBA.CStr(isUpdateMode)
    Call ex_Document.ex_LogWorkbookContext(private_ConfigText(CFG_WS_VACATION), inputCellMap)

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
    If VBA.StrComp(ticketsStatusText, vacationStatusCancelledTickets, _
            VBA.vbTextCompare) = 0 Then
        documentNamePattern = documentNamePatternCancelled
    Else
        documentNamePattern = documentNamePatternActive
    End If

    If Not ex_PersonnelData.ex_TryBeginSession() Then GoTo CleanExit
    personnelSessionStarted = True
    private_Performance_LogCheckpoint performanceStart, "SHPO session opened"
    If Not ex_PersonnelData.ex_TryResolveIpn(personLookup, ipnText) Then GoTo CleanExit
    ' TVO is optional. An empty field keeps TVO registry columns empty.
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
            If ex_Helpers.ex_ShowMessageBox( _
                "A vacation ticket record already exists in tbTickets, " & _
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

    ' Dates are calculated from the departure date entered by the user.
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
        orderDate, dateWordPattern, ticketDateText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        departureDate, dateWordPattern, dateFromText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        finalReturnDate, dateWordPattern, dateToText) Then GoTo CleanExit
    If Not ex_Helpers.private_Date_TryFormat( _
        dateArrival, dateWordPattern, dateArrivalText) Then GoTo CleanExit

    personalLine = rankText & " " & fioText
    ' A short name is used in the return-from-vacation line.
    personalInitials = private_Person_BuildInitials(rankText, fioText)
    If VBA.Len(personalInitials) = 0 Then GoTo CleanExit

    ' Names must exactly match Word template placeholders.
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
            If ex_Helpers.ex_ShowMessageBox( _
                "A document already exists for this vacation ticket:" & _
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
        ticketNo, tvoFioText, tvoIpnText, tvoPositionCode, ticketsStatusText, _
        vacationPlace) Then GoTo CleanExit
    If donationDays > 0 Then
        donationTicketRowCreated = (donationTicketRow Is Nothing)
        If Not ex_Tickets.ex_TrySaveVacationRow( _
            ticketsTable, donationTicketRow, rankText, fioText, ipnText, personPositionCode, _
            vacationKindDonationTickets, donationOrderNo, finalReturnDate, _
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

' Removes only rows added in this run after failed Word generation.
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
' namespace CandidatesConfig {
' --------------------------------------
' Creates a cell style in the common contract used by the candidate module.
Private Function private_CandidatesConfig_CreateCellStyle( _
    ByVal fontName As String, _
    ByVal fontSize As Double, _
    ByVal fontColor As Long, _
    ByVal fillColor As Long, _
    ByVal horizontalAlignment As Long, _
    ByVal verticalAlignment As Long, _
    ByVal wrapText As Boolean _
) As Object
    Dim cellStyle As Object

    Set cellStyle = VBA.CreateObject("Scripting.Dictionary")
    cellStyle.CompareMode = VBA.vbBinaryCompare
    cellStyle.Add "FontName", fontName
    cellStyle.Add "FontSize", fontSize
    cellStyle.Add "FontColor", fontColor
    cellStyle.Add "FillColor", fillColor
    cellStyle.Add "HorizontalAlignment", horizontalAlignment
    cellStyle.Add "VerticalAlignment", verticalAlignment
    cellStyle.Add "WrapText", wrapText
    Set private_CandidatesConfig_CreateCellStyle = cellStyle
End Function
' --------------------------------------
' } // namespace CandidatesConfig
' --------------------------------------

' --------------------------------------
' namespace Performance {
' --------------------------------------
' Temporary timing diagnostics for generation steps. Remove after measurement.
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

' Initializes module-level state before each generation run.
Private Function private_Initialize() As Boolean
    Set inputCellMap = VBA.CreateObject("Scripting.Dictionary")
    inputCellMap.CompareMode = VBA.vbBinaryCompare

    If Not private_Input_AddConfigCell(INPUT_ALIAS_PERSON_LOOKUP, "wsVacation::input.person_lookup") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_VACATION_KIND, "wsVacation::input.kind") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_VACATION_PLACE, "wsVacation::input.place") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_VACATION_ABROAD, "wsVacation::input.abroad") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_ORDER_REFERENCE, "wsVacation::input.order_reference") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_DEPARTURE_DATE, "wsVacation::input.departure_date") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_VACATION_DAYS, "wsVacation::input.days") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_ROAD_DAYS, "wsVacation::input.road_days") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_DONATION_DAYS, "wsVacation::input.donation_days") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_TVO_LOOKUP, "wsVacation::input.tvo_lookup") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_STATUS, "wsVacation::input.status") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_TEMPLATE_PATH, "wsVacation::input.template_path") Then Exit Function
    If Not private_Input_AddConfigCell(INPUT_ALIAS_OUTPUT_FOLDER_PATH, "wsVacation::input.output_folder_path") Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_ANNUAL_INPUT, vacationKindAnnualInput) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_ANNUAL_WORD, vacationKindAnnualWord) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_ANNUAL_TICKETS, vacationKindAnnualTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_FAMILY_INPUT, vacationKindFamilyInput) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_FAMILY_WORD, vacationKindFamilyWord) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_FAMILY_TICKETS, vacationKindFamilyTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_TREATMENT_INPUT, vacationKindTreatmentInput) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_TREATMENT_WORD, vacationKindTreatmentWord) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_TREATMENT_TICKETS, vacationKindTreatmentTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_MATERNITY_INPUT, vacationKindMaternityInput) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_MATERNITY_WORD, vacationKindMaternityWord) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_MATERNITY_TICKETS, vacationKindMaternityTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_CHILDCARE_INPUT, vacationKindChildcareInput) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_CHILDCARE_WORD, vacationKindChildcareWord) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_CHILDCARE_TICKETS, vacationKindChildcareTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_KIND_DONATION_TICKETS, vacationKindDonationTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_ABROAD_YES, vacationAbroadYes) Then Exit Function
    If Not private_LoadConfigText(CFG_ABROAD_NO, vacationAbroadNo) Then Exit Function
    If Not private_LoadConfigText(CFG_ABROAD_YES_WORD, vacationAbroadYesWord) Then Exit Function
    If Not private_LoadConfigText(CFG_STATUS_ACTIVE, vacationStatusActive) Then Exit Function
    If Not private_LoadConfigText(CFG_STATUS_CANCELLED, vacationStatusCancelled) Then Exit Function
    If Not private_LoadConfigText(CFG_STATUS_CANCELLED_TICKETS, vacationStatusCancelledTickets) Then Exit Function
    If Not private_LoadConfigText(CFG_DATE_WORD, dateWordPattern) Then Exit Function
    If Not private_LoadConfigText(CFG_DOCUMENT_NAME_ACTIVE, documentNamePatternActive) Then Exit Function
    If Not private_LoadConfigText(CFG_DOCUMENT_NAME_CANCELLED, documentNamePatternCancelled) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_CONNECTOR, durationConnector) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_VACATION_ONE, durationVacationOne) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_VACATION_FEW, durationVacationFew) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_VACATION_MANY, durationVacationMany) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_DONATION_ONE, durationDonationOne) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_DONATION_FEW, durationDonationFew) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_DONATION_MANY, durationDonationMany) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_DONATION_SUFFIX, durationDonationSuffix) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_ROAD_ONE, durationRoadOne) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_ROAD_FEW, durationRoadFew) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_ROAD_MANY, durationRoadMany) Then Exit Function
    If Not private_LoadConfigText(CFG_DURATION_ROAD_SUFFIX, durationRoadSuffix) Then Exit Function
    If Not private_LoadConfigText(CFG_WS_VACATION, vacationSheetName) Then Exit Function
    If Not private_LoadConfigText(CFG_PERSONNEL_ALF_RANGE, personnelAlfRange) Then Exit Function
    If Not private_LoadConfigText(CFG_PERSONNEL_ALF_FIO, personnelAlfFioColumn) Then Exit Function
    If Not private_LoadConfigText(CFG_PERSONNEL_ALF_IPN, personnelAlfIpnColumn) Then Exit Function
    If Not private_LoadConfigText(CFG_CANDIDATES_START, candidatesStartCellAddress) Then Exit Function
    If Not private_LoadConfigLong(CFG_CANDIDATES_MAX_COUNT, candidatesMaxCount) Then Exit Function
    If Not private_LoadConfigText(CFG_CANDIDATES_PERSON_LOOKUP_TITLE, candidatesPersonLookupTitle) Then Exit Function
    If Not private_LoadConfigText(CFG_CANDIDATES_TVO_LOOKUP_TITLE, candidatesTvoLookupTitle) Then Exit Function
    If Not private_LoadConfigText(CFG_CANDIDATES_TABLE_TITLE_PREFIX, candidatesTableTitlePrefix) Then Exit Function
    If Not private_LoadConfigText(CFG_CANDIDATES_HIDE_COMMAND, candidatesHideCommand) Then Exit Function
    private_Initialize = True
End Function

Private Function private_Input_AddConfigCell(ByVal fieldAlias As String, ByVal configKey As String) As Boolean
    Dim cellAddress As String
    If Not ex_Config.fn_TryGetText(configKey, cellAddress) Then Exit Function
    inputCellMap.Add fieldAlias, cellAddress
    private_Input_AddConfigCell = True
End Function

Private Function private_ConfigText(ByVal configKey As String) As String
    Dim valueText As String
    If ex_Config.fn_TryGetText(configKey, valueText) Then private_ConfigText = valueText
End Function

Private Function private_LoadConfigText(ByVal configKey As String, ByRef outValue As String) As Boolean
    outValue = VBA.vbNullString
    private_LoadConfigText = ex_Config.fn_TryGetText(configKey, outValue)
End Function

Private Function private_LoadConfigLong(ByVal configKey As String, ByRef outValue As Long) As Boolean
    outValue = 0
    private_LoadConfigLong = ex_Config.fn_TryGetLong(configKey, outValue)
End Function

' --------------------------------------
' namespace Input {
' --------------------------------------
Private Function private_Input_ReadRequired( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String _
) As String
    If Not ex_Document.ex_TryReadRequired( _
        private_ConfigText(CFG_WS_VACATION), inputCellMap, fieldAlias, fieldCaption, _
        private_Input_ReadRequired) Then Exit Function
End Function

Private Function private_Input_TryReadOptional( _
    ByVal fieldAlias As String, _
    ByRef outValue As String _
) As Boolean
    private_Input_TryReadOptional = ex_Document.ex_TryReadOptional( _
        private_ConfigText(CFG_WS_VACATION), inputCellMap, fieldAlias, outValue)
End Function

Private Function private_Input_TryReadNonNegativeDays( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByRef outDays As Long _
) As Boolean
    private_Input_TryReadNonNegativeDays = _
        ex_Document.ex_TryReadNonNegativeDays( _
            private_ConfigText(CFG_WS_VACATION), inputCellMap, fieldAlias, fieldCaption, _
            False, outDays)
End Function

Private Function private_Input_TryReadDate( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String, _
    ByRef outDate As Date _
) As Boolean
    Dim dateText As String

    If Not ex_Document.ex_TryReadRequired( _
        private_ConfigText(CFG_WS_VACATION), inputCellMap, fieldAlias, fieldCaption, dateText) Then Exit Function
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
            private_ConfigText(CFG_WS_VACATION), inputCellMap, fieldAlias, fieldCaption, _
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
            vacationDays, durationVacationOne, durationVacationFew, _
            durationVacationMany)
    If donationDays > 0 Then
        durationText = durationText & " " & durationConnector & " " & _
            VBA.CStr(donationDays) & " " & _
            ex_Helpers.ex_GetUkrainianCountForm( _
                donationDays, durationDonationOne, durationDonationFew, _
                durationDonationMany) & durationDonationSuffix
    End If
    If roadDays > 0 Then
        durationText = durationText & " " & durationConnector & " " & _
            VBA.CStr(roadDays) & " " & _
            ex_Helpers.ex_GetUkrainianCountForm( _
                roadDays, durationRoadOne, durationRoadFew, _
                durationRoadMany) & durationRoadSuffix
    End If
    private_Vacation_BuildWordDuration = durationText
End Function

Private Function private_Vacation_TryMapKind( _
    ByVal vacationKind As String, _
    ByRef outWordText As String, _
    ByRef outRegistryText As String _
) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(vacationKind))
        Case VBA.LCase$(vacationKindAnnualInput)
            outWordText = vacationKindAnnualWord
            outRegistryText = vacationKindAnnualTickets
        Case VBA.LCase$(vacationKindFamilyInput)
            outWordText = vacationKindFamilyWord
            outRegistryText = vacationKindFamilyTickets
        Case VBA.LCase$(vacationKindTreatmentInput)
            outWordText = vacationKindTreatmentWord
            outRegistryText = vacationKindTreatmentTickets
        Case VBA.LCase$(vacationKindMaternityInput)
            outWordText = vacationKindMaternityWord
            outRegistryText = vacationKindMaternityTickets
        Case VBA.LCase$(vacationKindChildcareInput)
            outWordText = vacationKindChildcareWord
            outRegistryText = vacationKindChildcareTickets
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
        Case VBA.LCase$(vacationAbroadYes)
            outVacationAbroadText = vacationAbroadYesWord
        Case VBA.LCase$(vacationAbroadNo)
            ' Keep a separate template line when there is no permission.
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
        Case VBA.LCase$(vacationStatusActive)
            ' An active ticket does not need a registry mark.
            outTicketsStatusText = VBA.vbNullString
        Case VBA.LCase$(vacationStatusCancelled)
            outTicketsStatusText = vacationStatusCancelledTickets
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