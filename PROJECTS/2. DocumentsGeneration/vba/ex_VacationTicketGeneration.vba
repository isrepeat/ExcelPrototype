Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

Private Const INPUT_SHEET_NAME As String = "Відпустки"

' Стабильные aliases полей формы. Адреса инкапсулированы в Input mapper-е.
Private Const INPUT_ALIAS_PERSON_LOOKUP As String = "PersonLookup"
Private Const INPUT_ALIAS_ORDER_REFERENCE As String = "OrderReference"
Private Const INPUT_ALIAS_TICKET_NO As String = "TicketNo"
Private Const INPUT_ALIAS_VACATION_KIND As String = "VacationKind"
Private Const INPUT_ALIAS_VACATION_PLACE As String = "VacationPlace"
Private Const INPUT_ALIAS_VACATION_DAYS As String = "VacationDays"
Private Const INPUT_ALIAS_ROAD_DAYS As String = "RoadDays"
Private Const INPUT_ALIAS_DONATION_DAYS As String = "DonationDays"
Private Const INPUT_ALIAS_TVO_LOOKUP As String = "TvoLookup"
Private Const INPUT_ALIAS_TEMPLATE_PATH As String = "TemplatePath"
Private Const INPUT_ALIAS_OUTPUT_FOLDER_PATH As String = "OutputFolderPath"

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
    Dim rawTicketNo As String, vacationKind As String, vacationKindText As String
    Dim vacationPlace As String
    Dim templatePath As String, outputFolderPath As String
    Dim ipnText As String, fioText As String
    Dim rankText As String, orderNo As String
    Dim orderDate As Date, dateFrom As Date, dateTo As Date, dateArrival As Date
    Dim vacationDays As Long, roadDays As Long, donationDays As Long
    Dim ticketNo As String, ticketDateText As String
    Dim dateFromText As String, dateToText As String
    Dim dateArrivalText As String, personalLine As String, personalInitials As String
    Dim placeholderNames As Variant, placeholderValues As Variant
    Dim documentPath As String, documentNameValues As Object

    On Error GoTo EH
    private_Initialize
    ex_Helpers.ClearLog
    ex_Helpers.LogDebug "Vacation ticket generation started"
    Call ex_Document.ex_LogWorkbookContext(INPUT_SHEET_NAME, inputCellMap)

    personLookup = private_Input_ReadRequired(INPUT_ALIAS_PERSON_LOOKUP, "ПІБ або ІПН")
    orderReference = private_Input_ReadRequired(INPUT_ALIAS_ORDER_REFERENCE, "номер або дату наказу")
    rawTicketNo = private_Input_ReadRequired(INPUT_ALIAS_TICKET_NO, "номер квитка")
    vacationKind = private_Input_ReadRequired(INPUT_ALIAS_VACATION_KIND, "вид відпустки")
    vacationPlace = private_Input_ReadRequired(INPUT_ALIAS_VACATION_PLACE, "місце відпустки")
    tvoLookup = private_Input_ReadRequired(INPUT_ALIAS_TVO_LOOKUP, "ПІБ або ІПН ТВО")
    templatePath = private_Input_ReadRequired(INPUT_ALIAS_TEMPLATE_PATH, "шлях до шаблону")
    outputFolderPath = private_Input_ReadRequired( _
        INPUT_ALIAS_OUTPUT_FOLDER_PATH, "шлях до папки результатів")
    If VBA.Len(personLookup) = 0 Or VBA.Len(orderReference) = 0 Or _
        VBA.Len(rawTicketNo) = 0 Or _
        VBA.Len(vacationKind) = 0 Or VBA.Len(vacationPlace) = 0 Or _
        VBA.Len(tvoLookup) = 0 Or _
        VBA.Len(templatePath) = 0 Or VBA.Len(outputFolderPath) = 0 Then Exit Sub

    ex_Helpers.LogDebug "Vacation ticket person lookup: " & personLookup
    ex_Helpers.LogDebug "Vacation ticket order reference: " & orderReference

    If Not private_Input_TryReadNonNegativeDays( _
        INPUT_ALIAS_VACATION_DAYS, "термін вибуття", vacationDays) Then Exit Sub
    If Not private_Input_TryReadOptionalNonNegativeDays( _
        INPUT_ALIAS_ROAD_DAYS, "додаткові дні на дорогу", roadDays) Then Exit Sub
    If Not private_Input_TryReadOptionalNonNegativeDays( _
        INPUT_ALIAS_DONATION_DAYS, "додаткові дні на донацію", donationDays) Then Exit Sub
    If Not private_Vacation_TryMapKind(vacationKind, vacationKindText) Then Exit Sub

    If Not ex_PersonnelData.ex_TryResolveIpn(personLookup, ipnText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveIpn(tvoLookup, tvoIpnText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveFioNominative( _
        tvoIpnText, tvoFioText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveFioNominative(ipnText, fioText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveRankNominative(ipnText, rankText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveOrderReference( _
        orderReference, orderNo, orderDate) Then Exit Sub
    If Not ex_PersonnelData.ex_TryBuildTicketNo( _
        rawTicketNo, orderNo, orderDate, ticketNo) Then Exit Sub

    ' Даты рассчитываются по утверждённому правилу отпуска.
    dateFrom = VBA.DateAdd("d", 1, orderDate)
    dateTo = VBA.DateAdd("d", vacationDays + roadDays + donationDays, dateFrom)
    dateArrival = VBA.DateAdd("d", 1, dateTo)
    ex_Helpers.LogDebug "Vacation period | From=" & VBA.CStr(dateFrom) & _
        " | To=" & VBA.CStr(dateTo) & " | Arrival=" & VBA.CStr(dateArrival)
    If Not ex_Helpers.private_Date_TryFormat( _
        orderDate, DATE_FORMAT_PATTERN, ticketDateText) Then Exit Sub
    If Not ex_Helpers.private_Date_TryFormat( _
        dateFrom, DATE_FORMAT_PATTERN, dateFromText) Then Exit Sub
    If Not ex_Helpers.private_Date_TryFormat( _
        dateTo, DATE_FORMAT_PATTERN, dateToText) Then Exit Sub
    If Not ex_Helpers.private_Date_TryFormat( _
        dateArrival, DATE_FORMAT_PATTERN, dateArrivalText) Then Exit Sub

    personalLine = rankText & " " & fioText
    ' Краткая запись используется в строке о возвращении из отпуска.
    personalInitials = private_Person_BuildInitials(rankText, fioText)
    If VBA.Len(personalInitials) = 0 Then Exit Sub

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
    If Not ex_Document.ex_TryGenerateWordDocument( _
        templatePath, DOCUMENT_NAME_PATTERN, documentNameValues, _
        placeholderNames, placeholderValues, documentPath, _
        outputFolderPath) Then Exit Sub
    ex_Helpers.WriteLog "DOCUMENT: " & documentPath
    ex_Helpers.LogDebug "Vacation ticket generation completed"
    Exit Sub
EH:
    ex_Helpers.LogError "Vacation ticket generation failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    VBA.MsgBox "Vacation ticket generation failed: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' Инициализирует module-level состояние перед каждым запуском генерации.
Private Sub private_Initialize()
    Set inputCellMap = VBA.CreateObject("Scripting.Dictionary")
    inputCellMap.CompareMode = VBA.vbBinaryCompare

    ' Адреса значений соответствуют строкам конфигурационной таблицы на листе.
    inputCellMap.Add INPUT_ALIAS_PERSON_LOOKUP, "C4"
    inputCellMap.Add INPUT_ALIAS_ORDER_REFERENCE, "C5"
    inputCellMap.Add INPUT_ALIAS_TICKET_NO, "C6"
    inputCellMap.Add INPUT_ALIAS_VACATION_KIND, "C7"
    inputCellMap.Add INPUT_ALIAS_VACATION_PLACE, "C8"
    inputCellMap.Add INPUT_ALIAS_VACATION_DAYS, "C9"
    inputCellMap.Add INPUT_ALIAS_ROAD_DAYS, "C10"
    inputCellMap.Add INPUT_ALIAS_DONATION_DAYS, "C11"
    inputCellMap.Add INPUT_ALIAS_TVO_LOOKUP, "C12"
    inputCellMap.Add INPUT_ALIAS_TEMPLATE_PATH, "C13"
    inputCellMap.Add INPUT_ALIAS_OUTPUT_FOLDER_PATH, "C14"
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
        VBA.MsgBox "FIO must contain surname, name and patronymic: " & fioText, _
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
            VBA.MsgBox "Unsupported vacation kind: " & vacationKind, _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Vacation_TryMapKind = True
End Function
' --------------------------------------
' } // namespace Vacation
' --------------------------------------