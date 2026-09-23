Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = ".log"
Private Const CFG_WS As String = "wsTravelCertificate"
Private Const CFG_MESSAGE_TARGET As String = "wsTravelCertificate::message_target"
Private Const CFG_DOCUMENT_NAME As String = "wsTravelCertificate::document_name"
Private Const CFG_TICKETS_EVENT As String = "wsTravelCertificate::tickets.event"
Private Const CFG_TICKETS_STATUS As String = "wsTravelCertificate::tickets.status"
Private Const CFG_PERSONNEL_ALF_RANGE As String = "rangePersonnelAlf"
Private Const CFG_PERSONNEL_ALF_FIO As String = "rangePersonnelAlf::column.fio"
Private Const CFG_PERSONNEL_ALF_IPN As String = "rangePersonnelAlf::column.ipn"
Private Const CFG_CANDIDATES_START As String = "wsTravelCertificate::candidates.start"
Private Const CFG_CANDIDATES_MAX_COUNT As String = "wsTravelCertificate::candidates.max_count"
Private Const CFG_CANDIDATES_TITLE As String = "wsTravelCertificate::candidates.title"
Private Const CFG_CANDIDATES_TABLE_TITLE As String = "wsTravelCertificate::candidates.table_title"
Private Const CFG_CANDIDATES_HIDE As String = "wsTravelCertificate::candidates.hide"

Private Const INPUT_PERSON_LOOKUP As String = "PersonLookup"
Private Const INPUT_ORDER_REFERENCE As String = "OrderReference"
Private Const INPUT_TRAVEL_PLACE As String = "TravelPlace"
Private Const INPUT_MILITARY_PLACE As String = "MilitaryPlace"
Private Const INPUT_DATE_FROM As String = "DateFrom"
Private Const INPUT_DURATION_DAYS As String = "DurationDays"
Private Const INPUT_WEAPON_LINE As String = "WeaponLine"
Private Const INPUT_TEMPLATE_PATH As String = "TemplatePath"
Private Const INPUT_OUTPUT_FOLDER_PATH As String = "OutputFolderPath"

Private inputCellMap As Object
Private sheetName As String
Private documentNamePattern As String
Private ticketsEventText As String
Private ticketsStatusText As String
Private personnelAlfRange As String, personnelAlfFioColumn As String, personnelAlfIpnColumn As String
Private candidatesStartCellAddress As String, candidatesTitle As String, candidatesHideCommand As String
Private candidatesTableTitle As String
Private candidatesMaxCount As Long

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_TravelCertificateGeneration_Create()
    private_Generate False
End Sub

Public Sub fn_TravelCertificateGeneration_Update()
    private_Generate True
End Sub

Private Sub private_Generate(ByVal isUpdateMode As Boolean)
    Dim personLookup As String, positionCode As String, orderReference As String
    Dim travelPlace As String, militaryPlace As String
    Dim weaponLine As String, templatePath As String, outputFolderPath As String
    Dim ipnText As String, fioText As String, fioDative As String
    Dim rankText As String, rankDative As String, positionText As String
    Dim documentPath As String
    Dim orderNo As String, ticketNo As String
    Dim orderDate As Date, dateFrom As Date
    Dim durationDays As Long
    Dim ticketsTable As ListObject, ticketRow As ListRow
    Dim matchedDocumentPath As String
    Dim placeholderNames As Variant, placeholderValues As Variant
    Dim documentNameValues As Object
    Dim sessionStarted As Boolean

    On Error GoTo EH
    If Not fn_TryInitializeUiRuntime() Then Exit Sub
    If Not private_Initialize() Then Exit Sub
    If Not ex_Helpers.ex_TryConfigureMessageTarget( _
        sheetName, private_ConfigText(CFG_MESSAGE_TARGET)) Then Exit Sub
    ex_Helpers.ClearLog
    ex_Helpers.LogDebug "Travel certificate generation started"
    ex_Document.ex_LogWorkbookContext sheetName, inputCellMap

    personLookup = private_ReadRequired(INPUT_PERSON_LOOKUP, "FIO or IPN")
    orderReference = private_ReadRequired(INPUT_ORDER_REFERENCE, "order number or date")
    travelPlace = private_ReadRequired(INPUT_TRAVEL_PLACE, "destination")
    militaryPlace = private_ReadRequired(INPUT_MILITARY_PLACE, "military unit")
    weaponLine = private_ReadOptional(INPUT_WEAPON_LINE)
    templatePath = private_ReadRequired(INPUT_TEMPLATE_PATH, "Word template path")
    outputFolderPath = private_ReadRequired(INPUT_OUTPUT_FOLDER_PATH, "results folder path")
    If Not private_TryReadDate(INPUT_DATE_FROM, "departure date", dateFrom) Then GoTo CleanExit
    If Not private_TryReadDurationDays(durationDays) Then GoTo CleanExit
    If VBA.Len(personLookup) = 0 Or VBA.Len(orderReference) = 0 Or _
        VBA.Len(travelPlace) = 0 Or _
        VBA.Len(militaryPlace) = 0 Or VBA.Len(templatePath) = 0 Or _
        VBA.Len(outputFolderPath) = 0 Then GoTo CleanExit

    If Not ex_PersonnelData.ex_TryBeginSession() Then GoTo CleanExit
    sessionStarted = True
    If Not ex_PersonnelData.ex_TryResolveIpn(personLookup, ipnText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveFioNominative(ipnText, fioText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveFioDative(ipnText, fioDative) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveRankNominative(ipnText, rankText) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolveRankDative(ipnText, rankDative) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolvePositionCode(ipnText, positionCode) Then GoTo CleanExit
    If Not ex_PersonnelData.ex_TryResolvePositionDative(positionCode, rankText, positionText) Then GoTo CleanExit
    positionText = private_Text_LowercaseFirstLetter(positionText)
    If Not ex_PersonnelData.ex_TryResolveOrderReference(orderReference, orderNo, orderDate) Then GoTo CleanExit
    If Not ex_Tickets.ex_TryGetOpenTable(ticketsTable) Then GoTo CleanExit
    If Not ex_Tickets.ex_TryFindVacationRow( _
        ticketsTable, ipnText, orderNo, VBA.vbNullString, ticketRow) Then GoTo CleanExit
    If isUpdateMode Then
        If ticketRow Is Nothing Then
            ex_Helpers.ex_ShowErrorMessage "Travel certificate record was not found for the selected person and order.", _
                VBA.vbExclamation, "Document Generation"
            GoTo CleanExit
        End If
        If Not ex_Tickets.ex_TryReadTicketNo( _
            ticketsTable, ticketRow, ticketNo) Then GoTo CleanExit
        If Not ex_Helpers.private_Path_TryFindVacationTicketDocument( _
            outputFolderPath, ticketNo, ipnText, matchedDocumentPath) Then GoTo CleanExit
    ElseIf ticketRow Is Nothing Then
        If Not ex_Tickets.ex_TryBuildNextTicketNo( _
            ticketsTable, orderNo, orderDate, ticketNo) Then GoTo CleanExit
    Else
        ex_Helpers.ex_ShowErrorMessage "Travel certificate record already exists. Use Update data.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    placeholderNames = Array( _
        "TicketNum", "PersonalLine", "Position", "WeaponLine", "FooterSpacer")
    placeholderValues = Array(ticketNo, rankDative & " " & fioDative, _
        positionText, weaponLine, VBA.vbCr)
    Set documentNameValues = VBA.CreateObject("Scripting.Dictionary")
    documentNameValues.CompareMode = VBA.vbBinaryCompare
    documentNameValues.Add "FIO", fioText
    documentNameValues.Add "IPN", ipnText
    documentNameValues.Add "TicketNo", ticketNo
    If Not ex_Document.ex_TryGenerateWordDocument( _
        templatePath, documentNamePattern, documentNameValues, placeholderNames, _
        placeholderValues, documentPath, outputFolderPath, matchedDocumentPath, _
        Not isUpdateMode, isUpdateMode) Then GoTo CleanExit
    If Not ex_Tickets.ex_TrySaveVacationRow( _
        ticketsTable, ticketRow, rankText, fioText, ipnText, positionCode, _
        ticketsEventText, orderNo, dateFrom, dateFrom, durationDays, 0, ticketNo, _
        VBA.vbNullString, VBA.vbNullString, VBA.vbNullString, ticketsStatusText, _
        travelPlace & "; " & militaryPlace) Then GoTo CleanExit
    ex_Helpers.WriteLog "DOCUMENT: " & documentPath
    ex_Helpers.ex_ShowStatusMessage private_Message_Generated() & documentPath

CleanExit:
    If sessionStarted Then ex_PersonnelData.ex_EndSession
    ex_Helpers.ex_ClearMessageTarget
    Exit Sub
EH:
    ex_Helpers.LogError "Travel certificate generation failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Travel certificate generation failed: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Sub

' Initializes shared UI dependencies before form interaction.
Public Function fn_TryInitializeUiRuntime() As Boolean
    fn_TryInitializeUiRuntime = ex_Helpers.ex_TryConfigureLogFileSuffix(LOG_FILE_SUFFIX)
End Function

' Returns the standard candidate-search contract for the input controller.
Public Function fn_TryGetCandidatesConfig(ByRef outCandidatesConfig As Object) As Boolean
    Dim lookups As Collection, titles As Object, columns As Collection, styles As Object
    Dim column As Object, titleStyle As Object
    Set outCandidatesConfig = Nothing
    If Not private_Initialize() Then Exit Function
    Set lookups = New Collection
    lookups.Add inputCellMap(INPUT_PERSON_LOOKUP)
    Set titles = VBA.CreateObject("Scripting.Dictionary")
    titles.CompareMode = VBA.vbBinaryCompare
    titles.Add inputCellMap(INPUT_PERSON_LOOKUP), candidatesTitle
    Set columns = New Collection
    Set column = private_CreateColumn(0, personnelAlfFioColumn, "General"): columns.Add column
    Set column = private_CreateColumn(1, personnelAlfIpnColumn, "@"): columns.Add column
    Set styles = private_CreateCandidateStyles()
    Set outCandidatesConfig = VBA.CreateObject("Scripting.Dictionary")
    outCandidatesConfig.CompareMode = VBA.vbBinaryCompare
    outCandidatesConfig.Add "InputSheetName", sheetName
    outCandidatesConfig.Add "LookupCellAddresses", lookups
    outCandidatesConfig.Add "LookupFieldTitles", titles
    outCandidatesConfig.Add "CandidateStartCellAddress", candidatesStartCellAddress
    outCandidatesConfig.Add "TableTitlePrefix", candidatesTableTitle
    outCandidatesConfig.Add "HideCommandText", candidatesHideCommand
    outCandidatesConfig.Add "MaxCandidateCount", candidatesMaxCount
    outCandidatesConfig.Add "QueryCallbackName", "ex_TravelCertificateGeneration.fn_TryFindPersonnelCandidates"
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

' Searches personnel by FIO or IPN for the form picker.
Public Function fn_TryFindPersonnelCandidates(ByVal searchText As String, ByVal maxCount As Long) As Collection
    Dim search As String, sql As String, fields As Collection, result As Collection
    If Not private_Initialize() Then Exit Function
    Set result = New Collection
    search = ex_Helpers.private_Text_Normalize(searchText)
    If VBA.Len(search) = 0 Then Set fn_TryFindPersonnelCandidates = result: Exit Function
    If maxCount <= 0 Then
        ex_Helpers.ex_ShowErrorMessage "Maximum candidate count must be greater than zero.", VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    sql = "SELECT TOP " & maxCount & " [" & personnelAlfFioColumn & "], [" & personnelAlfIpnColumn & "] FROM " & personnelAlfRange & " WHERE UCASE(CSTR([" & personnelAlfFioColumn & "])) LIKE '%" & ex_ExternalTables.ex_EscapeSql(VBA.UCase$(search)) & "%' OR UCASE(CSTR([" & personnelAlfIpnColumn & "])) LIKE '%" & ex_ExternalTables.ex_EscapeSql(VBA.UCase$(search)) & "%' ORDER BY [" & personnelAlfFioColumn & "]"
    Set fields = New Collection
    fields.Add personnelAlfFioColumn
    fields.Add personnelAlfIpnColumn
    If Not ex_PersonnelData.ex_TryExecuteShpoCandidateSql(sql, fields, result) Then Exit Function
    Set fn_TryFindPersonnelCandidates = result
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Initialize() As Boolean
    Set inputCellMap = VBA.CreateObject("Scripting.Dictionary")
    inputCellMap.CompareMode = VBA.vbBinaryCompare
    If Not private_AddInputCell(INPUT_PERSON_LOOKUP, "wsTravelCertificate::input.person_lookup") Then Exit Function
    If Not private_AddInputCell(INPUT_ORDER_REFERENCE, "wsTravelCertificate::input.order_reference") Then Exit Function
    If Not private_AddInputCell(INPUT_TRAVEL_PLACE, "wsTravelCertificate::input.travel_place") Then Exit Function
    If Not private_AddInputCell(INPUT_MILITARY_PLACE, "wsTravelCertificate::input.military_place") Then Exit Function
    If Not private_AddInputCell(INPUT_DATE_FROM, "wsTravelCertificate::input.date_from") Then Exit Function
    If Not private_AddInputCell(INPUT_DURATION_DAYS, "wsTravelCertificate::input.duration_days") Then Exit Function
    If Not private_AddInputCell(INPUT_WEAPON_LINE, "wsTravelCertificate::input.weapon_line") Then Exit Function
    If Not private_AddInputCell(INPUT_TEMPLATE_PATH, "wsTravelCertificate::input.template_path") Then Exit Function
    If Not private_AddInputCell(INPUT_OUTPUT_FOLDER_PATH, "wsTravelCertificate::input.output_folder_path") Then Exit Function
    If Not private_LoadText(CFG_WS, sheetName) Then Exit Function
    If Not private_LoadText(CFG_DOCUMENT_NAME, documentNamePattern) Then Exit Function
    If Not private_LoadText(CFG_TICKETS_EVENT, ticketsEventText) Then Exit Function
    If Not private_LoadText(CFG_TICKETS_STATUS, ticketsStatusText) Then Exit Function
    If Not private_LoadText(CFG_PERSONNEL_ALF_RANGE, personnelAlfRange) Then Exit Function
    If Not private_LoadText(CFG_PERSONNEL_ALF_FIO, personnelAlfFioColumn) Then Exit Function
    If Not private_LoadText(CFG_PERSONNEL_ALF_IPN, personnelAlfIpnColumn) Then Exit Function
    If Not private_LoadText(CFG_CANDIDATES_START, candidatesStartCellAddress) Then Exit Function
    If Not private_LoadLong(CFG_CANDIDATES_MAX_COUNT, candidatesMaxCount) Then Exit Function
    If Not private_LoadText(CFG_CANDIDATES_TITLE, candidatesTitle) Then Exit Function
    If Not private_LoadText(CFG_CANDIDATES_TABLE_TITLE, candidatesTableTitle) Then Exit Function
    If Not private_LoadText(CFG_CANDIDATES_HIDE, candidatesHideCommand) Then Exit Function
    private_Initialize = True
End Function

Private Function private_AddInputCell(ByVal aliasName As String, ByVal configKey As String) As Boolean
    Dim cellAddress As String
    If Not ex_Config.fn_TryGetText(configKey, cellAddress) Then Exit Function
    inputCellMap.Add aliasName, cellAddress
    private_AddInputCell = True
End Function

Private Function private_ConfigText(ByVal configKey As String) As String
    ex_Config.fn_TryGetText configKey, private_ConfigText
End Function

Private Function private_LoadText(ByVal configKey As String, ByRef outValue As String) As Boolean
    outValue = VBA.vbNullString
    private_LoadText = ex_Config.fn_TryGetText(configKey, outValue)
End Function

Private Function private_LoadLong(ByVal configKey As String, ByRef outValue As Long) As Boolean
    outValue = 0
    private_LoadLong = ex_Config.fn_TryGetLong(configKey, outValue)
End Function

' --------------------------------------
' namespace Input {
' --------------------------------------
Private Function private_ReadRequired(ByVal aliasName As String, ByVal caption As String) As String
    ex_Document.ex_TryReadRequired sheetName, inputCellMap, aliasName, caption, private_ReadRequired
End Function

Private Function private_ReadOptional(ByVal aliasName As String) As String
    ex_Document.ex_TryReadOptional sheetName, inputCellMap, aliasName, private_ReadOptional
End Function

Private Function private_TryReadDate( _
    ByVal aliasName As String, ByVal caption As String, ByRef outDate As Date _
) As Boolean
    Dim dateText As String
    dateText = private_ReadRequired(aliasName, caption)
    If VBA.Len(dateText) = 0 Then Exit Function
    If ex_Helpers.private_Date_TryParse(dateText, outDate) Then
        private_TryReadDate = True
    Else
        ex_Helpers.ex_ShowErrorMessage "Enter a valid date for " & caption & ".", _
            VBA.vbExclamation, "Document Generation"
    End If
End Function

Private Function private_TryReadDurationDays(ByRef outDays As Long) As Boolean
    If Not ex_Document.ex_TryReadNonNegativeDays( _
        sheetName, inputCellMap, INPUT_DURATION_DAYS, "travel duration", _
        True, outDays) Then Exit Function
    ' A blank duration represents a trip until a separate order is issued.
    private_TryReadDurationDays = True
End Function

' --------------------------------------
' } // namespace Input
' --------------------------------------

Private Function private_CreateColumn(ByVal sourceIndex As Long, ByVal headerText As String, ByVal numberFormat As String) As Object
    Set private_CreateColumn = VBA.CreateObject("Scripting.Dictionary")
    private_CreateColumn.CompareMode = VBA.vbBinaryCompare
    private_CreateColumn.Add "SourceIndex", sourceIndex
    private_CreateColumn.Add "Header", headerText
    private_CreateColumn.Add "NumberFormat", numberFormat
End Function

Private Function private_CreateCandidateStyles() As Object
    Dim titleStyle As Object
    Set private_CreateCandidateStyles = VBA.CreateObject("Scripting.Dictionary")
    private_CreateCandidateStyles.CompareMode = VBA.vbBinaryCompare
    private_CreateCandidateStyles.Add "Candidate", private_CreateStyle(VBA.RGB(255, 255, 255), VBA.RGB(0, 96, 32))
    private_CreateCandidateStyles.Add "Chrome", private_CreateStyle(VBA.RGB(255, 255, 255), VBA.RGB(0, 0, 0))
    private_CreateCandidateStyles.Add "Command", private_CreateStyle(VBA.RGB(244, 176, 132), VBA.RGB(64, 64, 64))
    Set titleStyle = private_CreateStyle(VBA.RGB(166, 166, 166), VBA.RGB(0, 0, 0))
    titleStyle.Remove "FillColor"
    private_CreateCandidateStyles.Add "Title", titleStyle
    private_CreateCandidateStyles.Add "Selected", private_CreateStyle(VBA.RGB(255, 255, 255), VBA.RGB(112, 0, 56))
End Function

Private Function private_CreateStyle(ByVal fontColor As Long, ByVal fillColor As Long) As Object
    Set private_CreateStyle = VBA.CreateObject("Scripting.Dictionary")
    private_CreateStyle.CompareMode = VBA.vbBinaryCompare
    private_CreateStyle.Add "FontName", "Times New Roman"
    private_CreateStyle.Add "FontSize", 16
    private_CreateStyle.Add "FontColor", fontColor
    private_CreateStyle.Add "FillColor", fillColor
    private_CreateStyle.Add "HorizontalAlignment", -4108
    private_CreateStyle.Add "VerticalAlignment", -4108
    private_CreateStyle.Add "WrapText", True
End Function

Private Function private_Message_Generated() As String
    private_Message_Generated = ex_Helpers.fn_FromCodePoints( _
        "1055,1086,1089,1074,1110,1076,1095,1077,1085,1085,1103," & _
        "32,1087,1088,1086,32,1074,1110,1076,1088,1103,1076,1078," & _
        "1077,1085,1085,1103,32,1089,1092,1086,1088,1084,1086,1074," & _
        "1072,1085,1086,58,32")
End Function

Private Function private_Text_LowercaseFirstLetter(ByVal valueText As String) As String
    If VBA.Len(valueText) = 0 Then Exit Function
    private_Text_LowercaseFirstLetter = VBA.LCase$(VBA.Left$(valueText, 1)) & _
        VBA.Mid$(valueText, 2)
End Function