Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

' Стабильные aliases полей формы. Адреса инкапсулированы в Input mapper-е.
Private Const INPUT_ALIAS_PERSON_LOOKUP As String = "PersonLookup"
Private Const INPUT_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const INPUT_ALIAS_ORDER_REFERENCE As String = "OrderReference"
Private Const INPUT_ALIAS_TICKET_NO As String = "TicketNo"
Private Const INPUT_ALIAS_TO_MILITARY_NUM As String = "ToMilitaryNum"
Private Const INPUT_ALIAS_MSG_FROM_MILITARY_NUM As String = "MsgFromMilitaryNum"
Private Const INPUT_ALIAS_MSG_POSITION As String = "MsgPosition"
Private Const INPUT_ALIAS_MSG_NUM As String = "MsgNum"
Private Const INPUT_ALIAS_MSG_DATE As String = "MsgDate"
Private Const INPUT_ALIAS_DATE_FROM As String = "DateFrom"
Private Const INPUT_ALIAS_DATE_TO As String = "DateTo"
Private Const INPUT_ALIAS_TEMPLATE_PATH As String = "TemplatePath"

Private Const INPUT_SHEET_NAME As String = "Відрядження"
Private Const SHPO_RELATIVE_PATH As String = "ШПО.xlsx"
Private Const ORDERS_RELATIVE_PATH As String = "Накази.xlsx"
Private Const ALF_TABLE_REF As String = "[АЛФ$A1:J12000]"
Private Const OS_TABLE_REF As String = "[ОС$A1:AB12000]"
Private Const RANKS_TABLE_REF As String = "[Звання$A1:E12000]"
Private Const POSITIONS_TABLE_REF As String = "[Посади$A1:E12000]"
Private Const ORDERS_2025_TABLE_REF As String = "[Накази$A2:B12000]"
Private Const ORDERS_2026_TABLE_REF As String = "[Накази$D2:E12000]"
Private Const POSITION_ROZP_PREFIX As String = "A1A"
Private Const POSITION_ROZP_TEXT_PREFIX As String = "у розпорядженні командира військової частини "
Private Const POSITION_ROZP_OFFICER_UNIT As String = "А3369"
Private Const POSITION_ROZP_OTHER_UNIT As String = "А7383"

Private Const AD_OPEN_STATIC As Long = 3
Private Const AD_LOCK_READ_ONLY As Long = 1

' Aliases генерируемого контекста, могут использоваться для formatter-ов.
Private Const GENERATED_CONTEXT_ALIAS_FIO As String = "FIO"
Private Const GENERATED_CONTEXT_ALIAS_IPN As String = "IPN"
Private Const GENERATED_CONTEXT_ALIAS_ORDER_NO As String = "OrderNo"
Private Const GENERATED_CONTEXT_ALIAS_ORDER_MIRRORED_DATE As String = "OrderDateMirrored"
Private Const GENERATED_CONTEXT_ALIAS_TICKET_NO As String = "TicketNo"

Private Const ORDER_DATE_FORMAT_PATTERN As String = "{dd}.{mm}.{yyyy}"
Private Const ORDER_DATE_MIRRORED_FORMAT_PATTERN As String = "{yyyy}.{mm}.{dd}"
Private Const TICKET_DATE_FORMAT_PATTERN As String = """{dd}"" {month} {yyyy} р."
Private Const GENERATED_DOCUMENT_NAME_PATTERN As String = "{FIO} ({IPN}) ({OrderDateMirrored})"

Private inputCellMap As Object

' Инициализирует module-level состояние перед каждым запуском генерации.
Private Sub Initialize()
    If Not inputCellMap Is Nothing Then
        inputCellMap.RemoveAll
        Set inputCellMap = Nothing
    End If

    Set inputCellMap = VBA.CreateObject("Scripting.Dictionary")
    inputCellMap.CompareMode = VBA.vbBinaryCompare
    inputCellMap.Add INPUT_ALIAS_PERSON_LOOKUP, "C4"
    inputCellMap.Add INPUT_ALIAS_POSITION_CODE, "C5"
    inputCellMap.Add INPUT_ALIAS_ORDER_REFERENCE, "C6"
    inputCellMap.Add INPUT_ALIAS_TICKET_NO, "C7"
    inputCellMap.Add INPUT_ALIAS_TO_MILITARY_NUM, "C8"
    inputCellMap.Add INPUT_ALIAS_MSG_FROM_MILITARY_NUM, "C9"
    inputCellMap.Add INPUT_ALIAS_MSG_POSITION, "C10"
    inputCellMap.Add INPUT_ALIAS_MSG_NUM, "C11"
    inputCellMap.Add INPUT_ALIAS_MSG_DATE, "C12"
    inputCellMap.Add INPUT_ALIAS_DATE_FROM, "C13"
    inputCellMap.Add INPUT_ALIAS_DATE_TO, "C14"
    inputCellMap.Add INPUT_ALIAS_TEMPLATE_PATH, "C15"
End Sub

' --------------------------------------
' namespace API {
' --------------------------------------
' Назначить этот макрос кнопке "Створити".
Public Sub fn_DocumentsGeneration_Create()
    Dim personLookup As String
    Dim shpoPath As String
    Dim connection As Object
    Dim ipnText As String
    Dim fioDefault As String
    Dim fioDative As String
    Dim rankText As String
    Dim rankDative As String
    Dim positionCode As String
    Dim positionText As String
    Dim templatePath As String
    Dim documentPath As String
    Dim orderReference As String
    Dim rawTicketNo As String
    Dim toMilitaryNum As String
    Dim msgFromMilitaryNum As String
    Dim msgPosition As String
    Dim msgNum As String
    Dim msgDate As String
    Dim dateFromText As String
    Dim dateToText As String
    Dim orderNo As String
    Dim orderDate As Date
    Dim orderDateText As String
    Dim orderDateMirroredText As String
    Dim ticketNo As String
    Dim ticketDateText As String
    Dim placeholderNames As Variant
    Dim placeholderValues As Variant
    Dim documentNameValues As Object
    Dim documentName As String

    On Error GoTo EH

    Initialize

    ClearLog
    LogDebug "Generation started"
    private_Log_WorkbookContext
    personLookup = private_Input_ReadPersonLookup()
    If VBA.Len(personLookup) = 0 Then Exit Sub

    positionCode = private_Input_ReadPositionCode()
    If VBA.Len(positionCode) = 0 Then Exit Sub

    orderReference = private_Input_ReadRequiredValue( _
        INPUT_ALIAS_ORDER_REFERENCE, "order number or date")
    If VBA.Len(orderReference) = 0 Then Exit Sub

    rawTicketNo = private_Input_ReadRequiredValue( _
        INPUT_ALIAS_TICKET_NO, "ticket number")
    If VBA.Len(rawTicketNo) = 0 Then Exit Sub

    toMilitaryNum = private_Input_ReadRequiredValue( _
        INPUT_ALIAS_TO_MILITARY_NUM, "destination military unit")
    If VBA.Len(toMilitaryNum) = 0 Then Exit Sub

    msgFromMilitaryNum = private_Input_ReadRequiredValue( _
        INPUT_ALIAS_MSG_FROM_MILITARY_NUM, "message military unit")
    If VBA.Len(msgFromMilitaryNum) = 0 Then Exit Sub

    msgPosition = private_Input_ReadRequiredValue( _
        INPUT_ALIAS_MSG_POSITION, "message position")
    If VBA.Len(msgPosition) = 0 Then Exit Sub

    msgNum = private_Input_ReadRequiredValue(INPUT_ALIAS_MSG_NUM, "message number")
    If VBA.Len(msgNum) = 0 Then Exit Sub

    msgDate = private_Input_ReadRequiredValue(INPUT_ALIAS_MSG_DATE, "message date")
    If VBA.Len(msgDate) = 0 Then Exit Sub

    dateFromText = private_Input_ReadOptionalValue(INPUT_ALIAS_DATE_FROM)
    dateToText = private_Input_ReadOptionalValue(INPUT_ALIAS_DATE_TO)
    templatePath = private_Input_ReadTemplatePath()
    If VBA.Len(templatePath) = 0 Then Exit Sub

    LogDebug "Person lookup: " & personLookup
    LogDebug "Person lookup Unicode: " & private_Text_ToUnicodeDebug(personLookup)
    LogDebug "Position code: " & positionCode
    LogDebug "Template input: " & templatePath

    If Not private_Order_TryResolveReference( _
        orderReference, orderNo, orderDate) Then Exit Sub

    If Not private_Order_TryBuildTicketNo( _
        rawTicketNo, orderNo, orderDate, ticketNo) Then Exit Sub

    If Not private_Date_TryFormat( _
        orderDate, ORDER_DATE_FORMAT_PATTERN, orderDateText) Then Exit Sub

    If Not private_Date_TryFormat( _
        orderDate, ORDER_DATE_MIRRORED_FORMAT_PATTERN, _
        orderDateMirroredText) Then Exit Sub

    If Not private_Date_TryFormat( _
        orderDate, TICKET_DATE_FORMAT_PATTERN, ticketDateText) Then Exit Sub

    If VBA.Len(dateFromText) = 0 Then
        dateFromText = orderDateText
        LogDebug "Start date is empty; order date fallback applied: " & _
            dateFromText
    End If

    If VBA.Len(dateToText) = 0 Then
        ' В шаблоне слово "до" уже находится перед плейсхолдером DateTo.
        dateToText = "окремого розпорядження"
        LogDebug "End date is empty; fallback applied: " & _
            "до окремого розпорядження"
    End If

    shpoPath = ThisWorkbook.Path & Application.PathSeparator & SHPO_RELATIVE_PATH
    If VBA.Len(VBA.Dir$(shpoPath)) = 0 Then
        LogError "SHPO file was not found: " & shpoPath
        VBA.MsgBox "SHPO file was not found: " & shpoPath, _
            VBA.vbExclamation, "Document Generation"

        Exit Sub
    End If

    If Not private_Shpo_TryOpenConnection(shpoPath, connection) Then Exit Sub
    LogDebug "SHPO opened: " & shpoPath

    If private_Text_IsIpn(personLookup) Then
        ipnText = personLookup
    ElseIf Not private_Shpo_TryLookup( _
        connection, ALF_TABLE_REF, "ПІБ", personLookup, "ІПН", ipnText) Then
        GoTo CleanExit
    End If

    If Not private_Shpo_TryLookup( _
        connection, ALF_TABLE_REF, "ІПН", ipnText, _
        "ПІБ", fioDefault) Then GoTo CleanExit

    If Not private_Shpo_TryLookup( _
        connection, ALF_TABLE_REF, "ІПН", ipnText, _
        "Давальний", fioDative) Then GoTo CleanExit

    If Not private_Shpo_TryLookup( _
        connection, OS_TABLE_REF, "ІПН", ipnText, _
        "Військове звання фактично", rankText) Then GoTo CleanExit

    If Not private_Shpo_TryLookup( _
        connection, RANKS_TABLE_REF, "Звання", rankText, _
        "Давальний", rankDative) Then GoTo CleanExit

    If Not private_Shpo_TryResolvePosition( _
        connection, positionCode, rankText, positionText) Then GoTo CleanExit

    LogDebug "Resolved IPN: " & ipnText
    LogDebug "Resolved FIO: " & fioDefault
    LogDebug "Resolved rank: " & rankText
    LogDebug "Resolved position: " & positionText
    WriteLog "RESULT: " & rankDative & " " & fioDative & " | " & positionText

    LogDebug "Resolved order | Number=" & orderNo & _
        " | Date=" & orderDateText & " | Ticket=" & ticketNo

    placeholderNames = Array( _
        "PersonalLine", "Position", "ToMilitaryNum", "DateFrom", "DateTo", _
        "OrderDate", "OrderNum", "MsgPosition", "MsgFromMilitaryNum", _
        "MsgNum", "MsgDate", "TicketNum", "TicketDate")

    placeholderValues = Array( _
        rankDative & " " & fioDative, positionText, toMilitaryNum, _
        dateFromText, dateToText, orderDateText, orderNo, msgPosition, _
        msgFromMilitaryNum, msgNum, msgDate, ticketNo, ticketDateText)

    Set documentNameValues = VBA.CreateObject("Scripting.Dictionary")
    documentNameValues.CompareMode = VBA.vbBinaryCompare
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_FIO, fioDefault
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_IPN, ipnText
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_ORDER_NO, orderNo
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_ORDER_MIRRORED_DATE, _
        orderDateMirroredText
    documentNameValues.Add GENERATED_CONTEXT_ALIAS_TICKET_NO, ticketNo

    If Not private_Text_TryFormat( _
        GENERATED_DOCUMENT_NAME_PATTERN, documentNameValues, _
        documentName) Then Exit Sub

    LogDebug "Generated document name: " & documentName
    If Not private_Word_TryGenerateDocument( _
        templatePath, documentName, placeholderNames, placeholderValues, _
        documentPath) Then GoTo CleanExit

    WriteLog "DOCUMENT: " & documentPath

CleanExit:
    On Error Resume Next
    If Not connection Is Nothing Then connection.Close
    Set connection = Nothing
    On Error GoTo 0
    Exit Sub

EH:
    LogError "Generation failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description

    VBA.MsgBox "Document generation failed: [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "Document Generation"

    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Input {
' --------------------------------------
Private Function private_Input_ReadPersonLookup() As String
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        LogError "Input sheet was not found | ExpectedNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & _
            " | WorksheetCount=" & VBA.CStr(ThisWorkbook.Worksheets.Count)
        VBA.MsgBox "Input sheet was not found. Check the document generation log.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    cellAddress = private_Input_GetCellAddress(INPUT_ALIAS_PERSON_LOOKUP)
    If VBA.Len(cellAddress) = 0 Then Exit Function

    private_Input_ReadPersonLookup = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadPersonLookup) = 0 Then
        LogError "Person lookup cell is empty | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
            cellAddress
        VBA.MsgBox "Enter FIO or IPN in cell " & _
            cellAddress & ".", _
            VBA.vbExclamation, "Document Generation"
    End If

End Function

Private Function private_Input_ReadPositionCode() As String
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    cellAddress = private_Input_GetCellAddress(INPUT_ALIAS_POSITION_CODE)
    If VBA.Len(cellAddress) = 0 Then Exit Function

    private_Input_ReadPositionCode = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadPositionCode) = 0 Then
        LogError "Position code cell is empty | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
            cellAddress
        VBA.MsgBox "Enter the position code in cell " & _
            cellAddress & ".", _
            VBA.vbExclamation, "Document Generation"
    End If

End Function

Private Function private_Input_ReadTemplatePath() As String
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    cellAddress = private_Input_GetCellAddress(INPUT_ALIAS_TEMPLATE_PATH)
    If VBA.Len(cellAddress) = 0 Then Exit Function

    private_Input_ReadTemplatePath = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadTemplatePath) = 0 Then
        LogError "Template path cell is empty | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
            cellAddress
        VBA.MsgBox "Enter the Word template path in cell " & _
            cellAddress & ".", _
            VBA.vbExclamation, "Document Generation"
    End If

End Function

Private Function private_Input_ReadRequiredValue( _
    ByVal fieldAlias As String, _
    ByVal fieldCaption As String _
) As String
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    cellAddress = private_Input_GetCellAddress(fieldAlias)
    If VBA.Len(cellAddress) = 0 Then Exit Function

    private_Input_ReadRequiredValue = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadRequiredValue) = 0 Then
        LogError "Required input is empty | Field=" & fieldCaption & _
            " | Cell=" & cellAddress
        VBA.MsgBox "Enter " & fieldCaption & " in cell " & cellAddress & ".", _
            VBA.vbExclamation, "Document Generation"
    End If

End Function

Private Function private_Input_ReadOptionalValue( _
    ByVal fieldAlias As String _
) As String
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    cellAddress = private_Input_GetCellAddress(fieldAlias)
    If VBA.Len(cellAddress) = 0 Then Exit Function

    private_Input_ReadOptionalValue = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
End Function

Private Function private_Input_GetCellAddress(ByVal fieldAlias As String) As String
    If inputCellMap Is Nothing Then
        LogError "Input cell mapper is not initialized"
        VBA.MsgBox "Input cell mapper is not initialized.", _
            VBA.vbExclamation, "Document Generation"

        Exit Function
    End If

    If Not inputCellMap.Exists(fieldAlias) Then
        LogError "Input field alias is not mapped: " & fieldAlias
        VBA.MsgBox "Input field alias is not mapped: " & fieldAlias, _
            VBA.vbExclamation, "Document Generation"

        Exit Function
    End If

    private_Input_GetCellAddress = VBA.CStr(inputCellMap(fieldAlias))
End Function
' --------------------------------------
' } // namespace Input
' --------------------------------------

' --------------------------------------
' namespace Shpo {
' --------------------------------------
Private Function private_Shpo_TryOpenConnection( _
    ByVal shpoPath As String, _
    ByRef outConnection As Object _
) As Boolean
    private_Shpo_TryOpenConnection = private_Ado_TryOpenConnection( _
        shpoPath, "SHPO", outConnection)
End Function

Private Function private_Shpo_TryResolvePosition( _
    ByVal connection As Object, _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    Dim normalizedCode As String

    normalizedCode = VBA.UCase$(private_Text_Normalize(positionCode))
    If private_Position_IsRozpCode(normalizedCode) Then
        If private_Rank_IsOfficer(rankText) Then
            outPositionText = POSITION_ROZP_TEXT_PREFIX & _
                POSITION_ROZP_OFFICER_UNIT
        Else
            outPositionText = POSITION_ROZP_TEXT_PREFIX & _
                POSITION_ROZP_OTHER_UNIT
        End If
        private_Shpo_TryResolvePosition = True
        Exit Function
    End If

    private_Shpo_TryResolvePosition = private_Shpo_TryLookup( _
        connection, POSITIONS_TABLE_REF, "Код", normalizedCode, _
        "Родовий", outPositionText)
End Function

Private Function private_Shpo_TryLookup( _
    ByVal connection As Object, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal resultHeader As String, _
    ByRef outValue As String _
) As Boolean
    Dim recordset As Object
    Dim sqlText As String

    On Error GoTo EH
    outValue = VBA.vbNullString
    sqlText = "SELECT TOP 2 [" & resultHeader & "] FROM " & tableRef & _
        " WHERE UCASE(TRIM(CSTR(IIF(ISNULL([" & keyHeader & _
        "]), '', [" & keyHeader & "])))) = '" & _
        private_Sql_EscapeLiteral(VBA.UCase$(private_Text_Normalize(keyValue))) & "'"
    LogDebug "Lookup SQL: " & sqlText

    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY

    If recordset.EOF Then
        LogError tableRef & ": value was not found | Key=" & keyHeader & _
            " | Value=" & keyValue
        VBA.MsgBox "SHPO " & tableRef & ": value '" & keyValue & _
            "' was not found in column '" & keyHeader & "'.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    outValue = private_Recordset_ReadText(recordset, resultHeader)
    recordset.MoveNext
    If Not recordset.EOF Then
        LogError tableRef & ": multiple rows found | Key=" & keyHeader & _
            " | Value=" & keyValue
        VBA.MsgBox "SHPO " & tableRef & ": multiple rows were found for '" & _
            keyValue & "'.", VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    If VBA.Len(outValue) = 0 Then
        LogError tableRef & ": result is empty | Column=" & resultHeader & _
            " | Key=" & keyValue
        VBA.MsgBox "SHPO " & tableRef & ": column '" & resultHeader & _
            "' is empty for '" & keyValue & "'.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If

    private_Shpo_TryLookup = True

CleanExit:
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    Set recordset = Nothing
    On Error GoTo 0
    Exit Function
EH:
    LogError "Failed to read " & tableRef & " | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    VBA.MsgBox "Failed to read SHPO " & tableRef & ": [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function
' --------------------------------------
' } // namespace Shpo
' --------------------------------------

' --------------------------------------
' namespace Order {
' --------------------------------------
Private Function private_Order_TryResolveReference( _
    ByVal orderInput As String, _
    ByRef outOrderNo As String, _
    ByRef outOrderDate As Date _
) As Boolean
    Dim ordersPath As String
    Dim ordersTableRef As String
    Dim connection As Object
    Dim recordset As Object
    Dim sqlText As String
    Dim inputIsDate As Boolean
    Dim inputDate As Date
    Dim orderYear As Long
    Dim candidateNo As String
    Dim candidateDate As Date
    Dim matchCount As Long

    On Error GoTo EH
    outOrderNo = VBA.vbNullString
    outOrderDate = 0
    inputIsDate = private_Date_LooksLikeFullDate(orderInput)
    If inputIsDate Then
        If Not private_Date_TryParse(orderInput, inputDate) Then
            VBA.MsgBox "Order date must be valid and use dd.mm.yyyy format.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
        orderYear = VBA.Year(inputDate)
    Else
        orderYear = VBA.Year(VBA.Date)
        orderInput = private_Order_NormalizeNumber(orderInput)
        If VBA.Len(orderInput) = 0 Then
            LogError "Order number is empty after normalization"
            VBA.MsgBox "Enter a valid order number.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
    End If

    If Not private_Order_TryGetTableRef(orderYear, ordersTableRef) Then Exit Function
    ordersPath = private_Path_ResolveFromWorkbook(ORDERS_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(ordersPath)) = 0 Then
        LogError "Orders workbook was not found: " & ordersPath
        VBA.MsgBox "Orders workbook was not found: " & ordersPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not private_Ado_TryOpenConnection( _
        ordersPath, "Orders", connection) Then Exit Function

    sqlText = "SELECT [Номер наказу], [Дата наказу] FROM " & ordersTableRef
    LogDebug "Order lookup SQL: " & sqlText
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY
    Do While Not recordset.EOF
        candidateNo = private_Order_NormalizeNumber( _
            private_Recordset_ReadText(recordset, "Номер наказу"))
        candidateDate = 0
        If private_Date_TryReadRecordsetDate( _
            recordset.Fields("Дата наказу").Value, candidateDate) Then
            If (inputIsDate And VBA.DateValue(candidateDate) = _
                    VBA.DateValue(inputDate)) Or _
               (Not inputIsDate And VBA.StrComp(candidateNo, orderInput, _
                    VBA.vbTextCompare) = 0) Then
                matchCount = matchCount + 1
                If matchCount = 1 Then
                    outOrderNo = candidateNo
                    outOrderDate = candidateDate
                End If
            End If
        End If
        recordset.MoveNext
    Loop

    If matchCount = 0 Then
        LogError "Order was not found | Input=" & orderInput & _
            " | Year=" & VBA.CStr(orderYear)
        VBA.MsgBox "Order was not found for value '" & orderInput & _
            "' in the " & VBA.CStr(orderYear) & " orders table.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    If matchCount > 1 Then
        LogError "Order reference is ambiguous | Input=" & orderInput & _
            " | Matches=" & VBA.CStr(matchCount)
        VBA.MsgBox "Multiple orders were found for value '" & orderInput & "'.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If

    private_Order_TryResolveReference = True
CleanExit:
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    If Not connection Is Nothing Then connection.Close
    Set recordset = Nothing
    Set connection = Nothing
    On Error GoTo 0
    Exit Function
EH:
    LogError "Order lookup failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    VBA.MsgBox "Order lookup failed: [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

Private Function private_Order_TryGetTableRef( _
    ByVal orderYear As Long, _
    ByRef outTableRef As String _
) As Boolean
    Select Case orderYear
        Case 2025: outTableRef = ORDERS_2025_TABLE_REF
        Case 2026: outTableRef = ORDERS_2026_TABLE_REF
        Case Else
            LogError "Orders table is not configured for year " & _
                VBA.CStr(orderYear)
            VBA.MsgBox "Orders table is not configured for year " & _
                VBA.CStr(orderYear) & ".", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Order_TryGetTableRef = True
End Function

Private Function private_Order_TryBuildTicketNo( _
    ByVal rawTicketNo As String, _
    ByVal orderNo As String, _
    ByVal orderDate As Date, _
    ByRef outTicketNo As String _
) As Boolean
    rawTicketNo = private_Text_Normalize(rawTicketNo)
    orderNo = private_Order_NormalizeNumber(orderNo)
    If Not private_Text_IsDigits(rawTicketNo) Or _
        Not private_Text_IsDigits(orderNo) Then
        LogError "Ticket number requires numeric order and ticket values"
        VBA.MsgBox "Order number and ticket number must contain digits only.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outTicketNo = VBA.CStr(VBA.Year(orderDate)) & "/" & _
        orderNo & "/" & rawTicketNo
    private_Order_TryBuildTicketNo = True
End Function

Private Function private_Order_NormalizeNumber( _
    ByVal orderNo As String _
) As String
    orderNo = private_Text_Normalize(orderNo)
    Do While VBA.Len(orderNo) > 1 And VBA.Left$(orderNo, 1) = "0"
        orderNo = VBA.Mid$(orderNo, 2)
    Loop
    private_Order_NormalizeNumber = orderNo
End Function
' --------------------------------------
' } // namespace Order
' --------------------------------------

' --------------------------------------
' namespace Word {
' --------------------------------------
Private Function private_Word_TryGenerateDocument( _
    ByVal templatePathInput As String, _
    ByVal documentName As String, _
    ByVal placeholderNames As Variant, _
    ByVal placeholderValues As Variant, _
    ByRef outDocumentPath As String _
) As Boolean
    Dim templatePath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim ownsWordApp As Boolean
    Dim placeholderIndex As Long

    On Error GoTo EH
    outDocumentPath = VBA.vbNullString
    templatePath = private_Path_ResolveFromWorkbook(templatePathInput)
    If VBA.Len(VBA.Dir$(templatePath)) = 0 Then
        LogError "Word template was not found: " & templatePath
        VBA.MsgBox "Word template was not found: " & templatePath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    outDocumentPath = private_Path_BuildGeneratedDocumentPath( _
        templatePath, documentName)
    If VBA.Len(outDocumentPath) = 0 Then Exit Function
    VBA.FileCopy templatePath, outDocumentPath
    LogDebug "Word template copied | Source=" & templatePath & _
        " | Target=" & outDocumentPath

    On Error Resume Next
    Set wordApp = VBA.GetObject(, "Word.Application")
    On Error GoTo EH
    If wordApp Is Nothing Then
        Set wordApp = VBA.CreateObject("Word.Application")
        ownsWordApp = True
    End If

    Set wordDoc = wordApp.Documents.Open(outDocumentPath)
    For placeholderIndex = LBound(placeholderNames) To UBound(placeholderNames)
        If Not private_Word_TryReplacePlaceholder( _
            wordDoc, VBA.CStr(placeholderNames(placeholderIndex)), _
            VBA.CStr(placeholderValues(placeholderIndex))) Then GoTo CleanFail
    Next placeholderIndex

    wordDoc.Save
    wordApp.Visible = True
    wordDoc.Activate
    private_Word_TryGenerateDocument = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If ownsWordApp And Not wordApp Is Nothing Then wordApp.Quit
    If VBA.Len(outDocumentPath) > 0 Then
        If VBA.Len(VBA.Dir$(outDocumentPath)) > 0 Then VBA.Kill outDocumentPath
    End If
    On Error GoTo 0
    Exit Function
EH:
    LogError "Word generation failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    VBA.MsgBox "Word generation failed: [" & VBA.CStr(Err.Number) & _
        "] " & Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanFail
End Function

Private Function private_Word_TryReplacePlaceholder( _
    ByVal wordDoc As Object, _
    ByVal placeholderName As String, _
    ByVal replacementText As String _
) As Boolean
    Dim markerText As String
    Dim markerRange As Object

    markerText = "<" & placeholderName & "></" & placeholderName & ">"
    Set markerRange = wordDoc.Content.Duplicate
    markerRange.Find.ClearFormatting
    markerRange.Find.Text = markerText
    markerRange.Find.Forward = True
    markerRange.Find.Wrap = 0
    markerRange.Find.MatchWildcards = False
    If Not markerRange.Find.Execute Then
        LogError "Required Word placeholder was not found: " & markerText
        VBA.MsgBox "Required Word placeholder was not found: " & markerText, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    markerRange.Text = replacementText
    LogDebug "Word placeholder replaced | Name=" & placeholderName & _
        " | ValueUnicode=" & private_Text_ToUnicodeDebug(replacementText)
    private_Word_TryReplacePlaceholder = True
End Function
' --------------------------------------
' } // namespace Word
' --------------------------------------

' --------------------------------------
' namespace Date {
' --------------------------------------
Private Function private_Date_TryFormat( _
    ByVal dateValue As Variant, _
    ByVal formatPattern As String, _
    ByRef outText As String _
) As Boolean
    Dim parsedDate As Date

    outText = VBA.vbNullString
    If Not private_Date_TryParse(dateValue, parsedDate) Then Exit Function
    If VBA.Len(formatPattern) = 0 Then
        LogError "Date format pattern is empty"
        Exit Function
    End If

    outText = formatPattern
    ' Сначала заменяем длинные токены, чтобы короткие не затрагивали их части.
    outText = VBA.Replace$(outText, "{month}", _
        private_Date_GetUaMonthGenitive(VBA.Month(parsedDate)))
    outText = VBA.Replace$(outText, "{yyyy}", _
        VBA.Format$(VBA.Year(parsedDate), "0000"))
    outText = VBA.Replace$(outText, "{yy}", _
        VBA.Right$(VBA.Format$(VBA.Year(parsedDate), "0000"), 2))
    outText = VBA.Replace$(outText, "{dd}", _
        VBA.Format$(VBA.Day(parsedDate), "00"))
    outText = VBA.Replace$(outText, "{d}", _
        VBA.CStr(VBA.Day(parsedDate)))
    outText = VBA.Replace$(outText, "{mm}", _
        VBA.Format$(VBA.Month(parsedDate), "00"))
    outText = VBA.Replace$(outText, "{m}", _
        VBA.CStr(VBA.Month(parsedDate)))
    ' В строковом формате из ячейки или внешнего конфига \" означает кавычку.
    outText = VBA.Replace$(outText, "\""", """")
    If VBA.InStr(1, outText, "{", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, outText, "}", VBA.vbBinaryCompare) > 0 Then
        LogError "Date format contains an unsupported token: " & _
            private_Text_ToUnicodeDebug(formatPattern)
        outText = VBA.vbNullString
        Exit Function
    End If

    private_Date_TryFormat = True
End Function

Private Function private_Date_TryParse( _
    ByVal dateValue As Variant, _
    ByRef outDate As Date _
) As Boolean
    Dim dateText As String
    Dim dateParts As Variant
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim parsedDate As Date

    On Error GoTo InvalidDate
    outDate = 0
    If VBA.IsDate(dateValue) And VBA.VarType(dateValue) <> VBA.vbString Then
        outDate = VBA.CDate(dateValue)
        private_Date_TryParse = True
        Exit Function
    End If

    dateText = VBA.Trim$(VBA.CStr(dateValue))
    dateText = VBA.Replace$(dateText, "/", ".")
    dateText = VBA.Replace$(dateText, "-", ".")
    dateParts = VBA.Split(dateText, ".")
    If UBound(dateParts) - LBound(dateParts) <> 2 Then GoTo InvalidDate
    If Not VBA.IsNumeric(dateParts(0)) Or _
        Not VBA.IsNumeric(dateParts(1)) Or _
        Not VBA.IsNumeric(dateParts(2)) Then GoTo InvalidDate

    dayValue = VBA.CLng(dateParts(0))
    monthValue = VBA.CLng(dateParts(1))
    yearValue = VBA.CLng(dateParts(2))
    parsedDate = VBA.DateSerial(yearValue, monthValue, dayValue)
    ' DateSerial нормализует 31.02, поэтому сверяем компоненты после парсинга.
    If VBA.Day(parsedDate) <> dayValue Or _
        VBA.Month(parsedDate) <> monthValue Or _
        VBA.Year(parsedDate) <> yearValue Then GoTo InvalidDate

    outDate = parsedDate
    private_Date_TryParse = True
    Exit Function

InvalidDate:
    LogError "Invalid date value: " & private_Text_ToUnicodeDebug( _
        VBA.CStr(dateValue))
End Function

Private Function private_Date_LooksLikeFullDate( _
    ByVal valueText As String _
) As Boolean
    private_Date_LooksLikeFullDate = ( _
        VBA.InStr(1, valueText, ".", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, valueText, "/", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, valueText, "-", VBA.vbBinaryCompare) > 0)
End Function

Private Function private_Date_TryReadRecordsetDate( _
    ByVal rawValue As Variant, _
    ByRef outDate As Date _
) As Boolean
    outDate = 0
    If VBA.IsNull(rawValue) Or VBA.IsEmpty(rawValue) Then Exit Function
    If VBA.IsDate(rawValue) Then
        outDate = VBA.CDate(rawValue)
        private_Date_TryReadRecordsetDate = True
        Exit Function
    End If
    private_Date_TryReadRecordsetDate = private_Date_TryParse( _
        rawValue, outDate)
End Function

Private Function private_Date_GetUaMonthGenitive( _
    ByVal monthNumber As Long _
) As String
    Select Case monthNumber
        Case 1: private_Date_GetUaMonthGenitive = "січня"
        Case 2: private_Date_GetUaMonthGenitive = "лютого"
        Case 3: private_Date_GetUaMonthGenitive = "березня"
        Case 4: private_Date_GetUaMonthGenitive = "квітня"
        Case 5: private_Date_GetUaMonthGenitive = "травня"
        Case 6: private_Date_GetUaMonthGenitive = "червня"
        Case 7: private_Date_GetUaMonthGenitive = "липня"
        Case 8: private_Date_GetUaMonthGenitive = "серпня"
        Case 9: private_Date_GetUaMonthGenitive = "вересня"
        Case 10: private_Date_GetUaMonthGenitive = "жовтня"
        Case 11: private_Date_GetUaMonthGenitive = "листопада"
        Case 12: private_Date_GetUaMonthGenitive = "грудня"
    End Select
End Function
' --------------------------------------
' } // namespace Date
' --------------------------------------

' --------------------------------------
' namespace Helpers {
' --------------------------------------
Private Function private_Text_IsIpn(ByVal valueText As String) As Boolean
    private_Text_IsIpn = (VBA.Len(valueText) > 0 And _
        Not valueText Like "*[!0-9]*")
End Function

Private Function private_Text_IsDigits(ByVal valueText As String) As Boolean
    valueText = private_Text_Normalize(valueText)
    private_Text_IsDigits = (VBA.Len(valueText) > 0 And _
        Not valueText Like "*[!0-9]*")
End Function

' Форматирует строку по именованным токенам: "{FIO} - {OrderNo}".
' Литеральные фигурные скобки задаются как "{{" и "}}".
Private Function private_Text_TryFormat( _
    ByVal formatPattern As String, _
    ByVal formatValues As Object, _
    ByRef outText As String _
) As Boolean
    Const OPEN_BRACE_SENTINEL As String = "<<__FORMAT_OPEN_BRACE__>>"
    Const CLOSE_BRACE_SENTINEL As String = "<<__FORMAT_CLOSE_BRACE__>>"
    Dim tokenKey As Variant
    Dim tokenName As String
    Dim tokenValue As String
    Dim validationText As String

    outText = VBA.vbNullString
    If VBA.Len(formatPattern) = 0 Then
        LogError "String format pattern is empty"
        Exit Function
    End If

    If formatValues Is Nothing Then
        LogError "String formatter value map is not initialized"
        Exit Function
    End If

    outText = VBA.Replace$(formatPattern, "{{", OPEN_BRACE_SENTINEL)
    outText = VBA.Replace$(outText, "}}", CLOSE_BRACE_SENTINEL)
    validationText = outText
    For Each tokenKey In formatValues.Keys
        tokenName = VBA.Trim$(VBA.CStr(tokenKey))
        If VBA.Len(tokenName) = 0 Then
            LogError "String formatter contains an empty token name"
            outText = VBA.vbNullString
            Exit Function
        End If

        If VBA.IsNull(formatValues(tokenKey)) Or _
            VBA.IsEmpty(formatValues(tokenKey)) Then
            tokenValue = VBA.vbNullString
        Else
            tokenValue = VBA.CStr(formatValues(tokenKey))
        End If

        outText = VBA.Replace$(outText, _
            "{" & tokenName & "}", tokenValue, 1, -1, VBA.vbBinaryCompare)
        validationText = VBA.Replace$(validationText, _
            "{" & tokenName & "}", VBA.vbNullString, _
            1, -1, VBA.vbBinaryCompare)
    Next tokenKey

    If VBA.InStr(1, validationText, "{", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, validationText, "}", VBA.vbBinaryCompare) > 0 Then
        LogError "String format contains an unknown or malformed token: " & _
            private_Text_ToUnicodeDebug(formatPattern)
        outText = VBA.vbNullString
        Exit Function
    End If

    outText = VBA.Replace$(outText, OPEN_BRACE_SENTINEL, "{")
    outText = VBA.Replace$(outText, CLOSE_BRACE_SENTINEL, "}")
    private_Text_TryFormat = True
End Function

Private Function private_Ado_TryOpenConnection( _
    ByVal workbookPath As String, _
    ByVal sourceCaption As String, _
    ByRef outConnection As Object _
) As Boolean
    On Error GoTo EH
    Set outConnection = VBA.CreateObject("ADODB.Connection")
    outConnection.Open _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & workbookPath & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    private_Ado_TryOpenConnection = True
    Exit Function
EH:
    LogError "Failed to open " & sourceCaption & " | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    VBA.MsgBox "Failed to open " & sourceCaption & ": [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Path_ResolveFromWorkbook( _
    ByVal pathText As String _
) As String
    pathText = private_Text_Normalize(pathText)
    If VBA.Len(pathText) >= 2 And VBA.Mid$(pathText, 2, 1) = ":" Then
        private_Path_ResolveFromWorkbook = pathText
    ElseIf VBA.Left$(pathText, 2) = "\\" Then
        private_Path_ResolveFromWorkbook = pathText
    Else
        private_Path_ResolveFromWorkbook = ThisWorkbook.Path & _
            Application.PathSeparator & pathText
    End If
End Function

Private Function private_Path_BuildGeneratedDocumentPath( _
    ByVal templatePath As String, _
    ByVal documentName As String _
) As String
    Dim folderPath As String
    Dim fileNameBase As String
    Dim dotPosition As Long
    Dim slashPosition As Long
    Dim extensionText As String
    Dim candidatePath As String
    Dim copyIndex As Long

    dotPosition = VBA.InStrRev(templatePath, ".")
    If dotPosition > 0 Then
        extensionText = VBA.Mid$(templatePath, dotPosition)
    Else
        extensionText = ".docx"
    End If

    slashPosition = VBA.InStrRev(templatePath, Application.PathSeparator)
    If slashPosition > 0 Then _
        folderPath = VBA.Left$(templatePath, slashPosition)
    fileNameBase = private_Path_SanitizeFileName(documentName)
    If VBA.Len(fileNameBase) = 0 Then
        LogError "Generated document file name is empty after FIO sanitization"
        VBA.MsgBox "Generated document file name is empty.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    candidatePath = folderPath & fileNameBase & extensionText
    copyIndex = 2
    Do While VBA.Len(VBA.Dir$(candidatePath)) > 0
        candidatePath = folderPath & fileNameBase & " (" & _
            VBA.CStr(copyIndex) & ")" & extensionText
        copyIndex = copyIndex + 1
    Loop
    private_Path_BuildGeneratedDocumentPath = candidatePath
End Function

Private Function private_Path_SanitizeFileName( _
    ByVal fileNameText As String _
) As String
    Dim invalidChar As Variant

    fileNameText = private_Text_Normalize(fileNameText)
    For Each invalidChar In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        fileNameText = VBA.Replace$(fileNameText, VBA.CStr(invalidChar), "_")
    Next invalidChar
    Do While VBA.Len(fileNameText) > 0 And _
        (VBA.Right$(fileNameText, 1) = "." Or _
         VBA.Right$(fileNameText, 1) = " ")
        fileNameText = VBA.Left$(fileNameText, VBA.Len(fileNameText) - 1)
    Loop
    private_Path_SanitizeFileName = fileNameText
End Function

Private Function private_Position_IsRozpCode( _
    ByVal positionCode As String _
) As Boolean
    positionCode = VBA.Replace$(positionCode, " ", VBA.vbNullString)
    positionCode = VBA.Replace$(positionCode, "А", "A")
    private_Position_IsRozpCode = ( _
        VBA.Left$(positionCode, VBA.Len(POSITION_ROZP_PREFIX)) = _
        POSITION_ROZP_PREFIX)
End Function

Private Function private_Rank_IsOfficer(ByVal rankText As String) As Boolean
    Select Case VBA.LCase$(private_Text_Normalize(rankText))
        Case "молодший лейтенант", "лейтенант", "старший лейтенант", _
             "капітан", "майор", "підполковник", "полковник"
            private_Rank_IsOfficer = True
    End Select
End Function

Private Function private_Text_Normalize(ByVal valueText As String) As String
    valueText = VBA.Replace$(valueText, VBA.ChrW$(160), " ")
    valueText = VBA.Trim$(valueText)
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace$(valueText, "  ", " ")
    Loop
    private_Text_Normalize = valueText
End Function

Private Function private_Text_ToUnicodeDebug(ByVal valueText As String) As String
    Dim charIndex As Long
    Dim charCode As Long
    Dim charText As String

    For charIndex = 1 To VBA.Len(valueText)
        charText = VBA.Mid$(valueText, charIndex, 1)
        charCode = VBA.AscW(charText)
        If charCode < 0 Then charCode = charCode + 65536
        If charCode >= 32 And charCode <= 126 Then
            private_Text_ToUnicodeDebug = _
                private_Text_ToUnicodeDebug & charText
        Else
            private_Text_ToUnicodeDebug = _
                private_Text_ToUnicodeDebug & "\u" & _
                VBA.Right$("0000" & VBA.Hex$(charCode), 4)
        End If
    Next charIndex
End Function

Private Function private_Sql_EscapeLiteral(ByVal valueText As String) As String
    private_Sql_EscapeLiteral = VBA.Replace$(valueText, "'", "''")
End Function

Private Function private_Recordset_ReadText( _
    ByVal recordset As Object, _
    ByVal fieldName As String _
) As String
    If VBA.IsNull(recordset.Fields(fieldName).Value) Then Exit Function
    private_Recordset_ReadText = private_Text_Normalize( _
        VBA.CStr(recordset.Fields(fieldName).Value))
End Function
' --------------------------------------
' } // namespace Helpers
' --------------------------------------

' --------------------------------------
' namespace Logging {
' --------------------------------------
Private Sub ClearLog()
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Output As #fileNumber
    Close #fileNumber
#End If
End Sub

Private Sub private_Log_WorkbookContext()
#If ENABLE_LOGGING Then
    Dim worksheetIndex As Long
    Dim worksheetObj As Worksheet
    Dim activeSheetText As String
    Dim personCellAddress As String
    Dim positionCellAddress As String
    Dim orderCellAddress As String
    Dim ticketCellAddress As String
    Dim dateFromCellAddress As String
    Dim dateToCellAddress As String
    Dim templateCellAddress As String

    On Error Resume Next
    activeSheetText = Application.ActiveSheet.Name
    On Error GoTo 0

    LogDebug "Workbook path: " & ThisWorkbook.FullName
    LogDebug "Worksheet count: " & VBA.CStr(ThisWorkbook.Worksheets.Count)
    LogDebug "Active sheet Unicode: " & _
        private_Text_ToUnicodeDebug(activeSheetText)

    For worksheetIndex = 1 To ThisWorkbook.Worksheets.Count
        Set worksheetObj = ThisWorkbook.Worksheets(worksheetIndex)
        LogDebug "Worksheet | Index=" & VBA.CStr(worksheetIndex) & _
            " | CodeName=" & worksheetObj.CodeName & _
            " | NameUnicode=" & private_Text_ToUnicodeDebug(worksheetObj.Name)
    Next worksheetIndex

    Set worksheetObj = Nothing
    On Error Resume Next
    Set worksheetObj = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    On Error GoTo 0
    If Not worksheetObj Is Nothing Then
        personCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_PERSON_LOOKUP)
        positionCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_POSITION_CODE)
        orderCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_ORDER_REFERENCE)
        ticketCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_TICKET_NO)
        dateFromCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_DATE_FROM)
        dateToCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_DATE_TO)
        templateCellAddress = private_Input_GetCellAddress(INPUT_ALIAS_TEMPLATE_PATH)

        LogDebug "Input binding | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & _
            " | PersonCell=" & personCellAddress & _
            " | PositionCell=" & positionCellAddress & _
            " | OrderCell=" & orderCellAddress & _
            " | TicketCell=" & ticketCellAddress & _
            " | DateFromCell=" & dateFromCellAddress & _
            " | DateToCell=" & dateToCellAddress & _
            " | TemplateCell=" & templateCellAddress

        LogDebug "Input raw values Unicode | Person=" & _
            private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(personCellAddress).Text)) & _
            " | Position=" & private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(positionCellAddress).Text)) & _
            " | Template=" & private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(templateCellAddress).Text))
    Else
        LogError "Input binding failed | ExpectedNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME)
    End If
#End If
End Sub

Private Sub LogError(ByVal messageText As String)
#If ENABLE_LOGGING Then
    WriteLog "ERROR: " & messageText
#End If
End Sub

Private Sub LogDebug(ByVal messageText As String)
#If ENABLE_LOGGING Then
#If ENABLE_DEBUG_LOGGING Then
    WriteLog "DEBUG: " & messageText
#End If
#End If
End Sub

Private Sub WriteLog(ByVal messageText As String)
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Append As #fileNumber
    Print #fileNumber, VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & messageText
    Close #fileNumber
#End If
End Sub

Private Function GetLogFilePath() As String
    Dim workbookName As String
    Dim baseName As String
    Dim dotPosition As Long

    workbookName = ThisWorkbook.Name
    dotPosition = VBA.InStrRev(workbookName, ".")
    If dotPosition > 0 Then
        baseName = VBA.Left$(workbookName, dotPosition - 1)
    Else
        baseName = workbookName
    End If
    GetLogFilePath = ThisWorkbook.Path & Application.PathSeparator & _
        baseName & LOG_FILE_SUFFIX
End Function
' --------------------------------------
' } // namespace Logging
' --------------------------------------
