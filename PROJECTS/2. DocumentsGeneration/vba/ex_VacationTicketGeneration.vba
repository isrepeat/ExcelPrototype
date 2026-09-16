Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

Private Const INPUT_SHEET_NAME As String = "Відпустки"

' Стабильные aliases полей формы. Адреса инкапсулированы в Input mapper-е.
Private Const INPUT_ALIAS_PERSON_LOOKUP As String = "PersonLookup"
Private Const INPUT_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const INPUT_ALIAS_ORDER_REFERENCE As String = "OrderReference"
Private Const INPUT_ALIAS_TICKET_NO As String = "TicketNo"
Private Const INPUT_ALIAS_VACATION_KIND As String = "VacationKind"
Private Const INPUT_ALIAS_VACATION_PLACE As String = "VacationPlace"
Private Const INPUT_ALIAS_VACATION_DAYS As String = "VacationDays"
Private Const INPUT_ALIAS_ROAD_DAYS As String = "RoadDays"
Private Const INPUT_ALIAS_DONATION_DAYS As String = "DonationDays"
Private Const INPUT_ALIAS_TVO_LOOKUP As String = "TvoLookup"
Private Const INPUT_ALIAS_TEMPLATE_PATH As String = "TemplatePath"

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
    Dim personLookup As String, positionCode As String, orderReference As String
    Dim rawTicketNo As String, vacationKind As String, vacationPlace As String
    Dim templatePath As String, ipnText As String, fioText As String
    Dim rankText As String, positionText As String, orderNo As String
    Dim orderDate As Date, dateFrom As Date, dateTo As Date, dateArrival As Date
    Dim vacationDays As Long, roadDays As Long, donationDays As Long
    Dim ticketNo As String, dateFromText As String, dateToText As String
    Dim dateArrivalText As String, personalLine As String, personalInitials As String
    Dim placeholderNames As Variant, placeholderValues As Variant
    Dim documentPath As String, documentNameValues As Object

    On Error GoTo EH
    private_Initialize
    ex_Helpers.ClearLog
    ex_Helpers.LogDebug "Vacation ticket generation started"
    ex_Document.ex_LogWorkbookContext(INPUT_SHEET_NAME, inputCellMap)

    personLookup = private_Input_ReadRequired(INPUT_ALIAS_PERSON_LOOKUP, "ПІБ або ІПН")
    positionCode = private_Input_ReadRequired(INPUT_ALIAS_POSITION_CODE, "код посади")
    orderReference = private_Input_ReadRequired(INPUT_ALIAS_ORDER_REFERENCE, "номер або дату наказу")
    rawTicketNo = private_Input_ReadRequired(INPUT_ALIAS_TICKET_NO, "номер квитка")
    vacationKind = private_Input_ReadRequired(INPUT_ALIAS_VACATION_KIND, "вид відпустки")
    vacationPlace = private_Input_ReadRequired(INPUT_ALIAS_VACATION_PLACE, "місце відпустки")
    templatePath = private_Input_ReadRequired(INPUT_ALIAS_TEMPLATE_PATH, "шлях до шаблону")
    If VBA.Len(personLookup) = 0 Or VBA.Len(positionCode) = 0 Or _
        VBA.Len(orderReference) = 0 Or VBA.Len(rawTicketNo) = 0 Or _
        VBA.Len(vacationKind) = 0 Or VBA.Len(vacationPlace) = 0 Or _
        VBA.Len(templatePath) = 0 Then Exit Sub

    ex_Helpers.LogDebug "Vacation ticket person lookup: " & personLookup
    ex_Helpers.LogDebug "Vacation ticket position code: " & positionCode
    ex_Helpers.LogDebug "Vacation ticket order reference: " & orderReference

    If Not private_Input_TryReadNonNegativeDays( _
        INPUT_ALIAS_VACATION_DAYS, "термін вибуття", vacationDays) Then Exit Sub
    If Not private_Input_TryReadOptionalNonNegativeDays( _
        INPUT_ALIAS_ROAD_DAYS, "додаткові дні на дорогу", roadDays) Then Exit Sub
    If Not private_Input_TryReadOptionalNonNegativeDays( _
        INPUT_ALIAS_DONATION_DAYS, "додаткові дні на донацію", donationDays) Then Exit Sub

    If Not ex_PersonnelData.ex_TryResolveIpn(personLookup, ipnText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveFioNominative(ipnText, fioText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolveRankNominative(ipnText, rankText) Then Exit Sub
    If Not ex_PersonnelData.ex_TryResolvePositionGenitive( _
        positionCode, rankText, positionText) Then Exit Sub
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
        "PersonalInitials", "DateArrival", "Position")
    placeholderValues = Array( _
        dateFromText, ticketNo, personalLine, vacationKind, vacationPlace, _
        VBA.CStr(vacationDays) & " календарних днів", dateFromText, _
        dateToText, personalInitials, dateArrivalText, positionText)

    Set documentNameValues = VBA.CreateObject("Scripting.Dictionary")
    documentNameValues.CompareMode = VBA.vbBinaryCompare
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_FIO, fioText
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_IPN, ipnText
    If Not ex_Document.ex_TryGenerateWordDocument( _
        templatePath, DOCUMENT_NAME_PATTERN, documentNameValues, _
        placeholderNames, placeholderValues, documentPath) Then Exit Sub
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
    inputCellMap.Add INPUT_ALIAS_POSITION_CODE, "C5"
    inputCellMap.Add INPUT_ALIAS_ORDER_REFERENCE, "C6"
    inputCellMap.Add INPUT_ALIAS_TICKET_NO, "C7"
    inputCellMap.Add INPUT_ALIAS_VACATION_KIND, "C8"
    inputCellMap.Add INPUT_ALIAS_VACATION_PLACE, "C9"
    inputCellMap.Add INPUT_ALIAS_VACATION_DAYS, "C10"
    inputCellMap.Add INPUT_ALIAS_ROAD_DAYS, "C11"
    inputCellMap.Add INPUT_ALIAS_DONATION_DAYS, "C12"
    inputCellMap.Add INPUT_ALIAS_TVO_LOOKUP, "C13"
    inputCellMap.Add INPUT_ALIAS_TEMPLATE_PATH, "C14"
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
    private_Person_BuildInitials = rankText & " " & nameParts(0) & " " & _
        VBA.Left$(nameParts(1), 1) & "." & VBA.Left$(nameParts(2), 1) & "."
End Function
' --------------------------------------
' } // namespace Person
' --------------------------------------