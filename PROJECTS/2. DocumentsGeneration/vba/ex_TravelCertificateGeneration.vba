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
Private Const ORDERS_RELATIVE_PATH As String = "Накази.xlsx"
Private Const ORDERS_2025_TABLE_REF As String = "[Накази$A2:B12000]"
Private Const ORDERS_2026_TABLE_REF As String = "[Накази$D2:E12000]"

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

' --------------------------------------
' namespace API {
' --------------------------------------
' Назначить этот макрос кнопке "Створити".
Public Sub fn_DocumentsGeneration_Create()
    Dim personLookup As String
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

    private_Initialize

    ex_Helpers.ClearLog
    ex_Helpers.LogDebug "Generation started"
    private_Diagnoctics_LogWorkbookContext
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

    ex_Helpers.LogDebug "Person lookup: " & personLookup
    ex_Helpers.LogDebug "Person lookup Unicode: " & ex_Helpers.private_Text_ToUnicodeDebug(personLookup)
    ex_Helpers.LogDebug "Position code: " & positionCode
    ex_Helpers.LogDebug "Template input: " & templatePath

    If Not private_Order_TryResolveReference( _
        orderReference, orderNo, orderDate) Then Exit Sub

    If Not private_Order_TryBuildTicketNo( _
        rawTicketNo, orderNo, orderDate, ticketNo) Then Exit Sub

    If Not ex_Helpers.private_Date_TryFormat( _
        orderDate, ORDER_DATE_FORMAT_PATTERN, orderDateText) Then Exit Sub

    If Not ex_Helpers.private_Date_TryFormat( _
        orderDate, ORDER_DATE_MIRRORED_FORMAT_PATTERN, _
        orderDateMirroredText) Then Exit Sub

    If Not ex_Helpers.private_Date_TryFormat( _
        orderDate, TICKET_DATE_FORMAT_PATTERN, ticketDateText) Then Exit Sub

    If VBA.Len(dateFromText) = 0 Then
        dateFromText = orderDateText
        ex_Helpers.LogDebug "Start date is empty; order date fallback applied: " & _
            dateFromText
    End If

    If VBA.Len(dateToText) = 0 Then
        ' В шаблоне слово "до" уже находится перед плейсхолдером DateTo.
        dateToText = "окремого розпорядження"
        ex_Helpers.LogDebug "End date is empty; fallback applied: " & _
            "до окремого розпорядження"
    End If

    If Not ex_PersonnelData.ex_PersonnelData_TryResolveIpn( _
        personLookup, ipnText) Then Exit Sub
    If Not ex_PersonnelData.ex_PersonnelData_TryResolveFioNominative( _
        ipnText, fioDefault) Then Exit Sub
    If Not ex_PersonnelData.ex_PersonnelData_TryResolveFioDative( _
        ipnText, fioDative) Then Exit Sub
    If Not ex_PersonnelData.ex_PersonnelData_TryResolveRankNominative( _
        ipnText, rankText) Then Exit Sub
    If Not ex_PersonnelData.ex_PersonnelData_TryResolveRankDative( _
        ipnText, rankDative) Then Exit Sub
    If Not ex_PersonnelData.ex_PersonnelData_TryResolvePositionGenitive( _
        positionCode, rankText, positionText) Then Exit Sub

    ex_Helpers.LogDebug "Resolved IPN: " & ipnText
    ex_Helpers.LogDebug "Resolved FIO: " & fioDefault
    ex_Helpers.LogDebug "Resolved rank: " & rankText
    ex_Helpers.LogDebug "Resolved position: " & positionText
    ex_Helpers.WriteLog "RESULT: " & rankDative & " " & fioDative & " | " & positionText

    ex_Helpers.LogDebug "Resolved order | Number=" & orderNo & _
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

    If Not ex_Helpers.private_Text_TryFormat( _
        GENERATED_DOCUMENT_NAME_PATTERN, documentNameValues, _
        documentName) Then Exit Sub

    ex_Helpers.LogDebug "Generated document name: " & documentName
    If Not ex_Helpers.private_Word_TryGenerateDocument( _
        templatePath, documentName, placeholderNames, placeholderValues, _
        documentPath) Then GoTo CleanExit

    ex_Helpers.WriteLog "DOCUMENT: " & documentPath

CleanExit:
    Exit Sub

EH:
    ex_Helpers.LogError "Generation failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description

    VBA.MsgBox "Document generation failed: [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "Document Generation"

    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' Инициализирует module-level состояние перед каждым запуском генерации.
Private Sub private_Initialize()
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
' namespace Input {
' --------------------------------------
Private Function private_Input_ReadPersonLookup() As String
    Dim sourceSheet As Worksheet
    Dim cellAddress As String

    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        ex_Helpers.LogError "Input sheet was not found | ExpectedNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & _
            " | WorksheetCount=" & VBA.CStr(ThisWorkbook.Worksheets.Count)
        VBA.MsgBox "Input sheet was not found. Check the document generation log.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    cellAddress = private_Input_GetCellAddress(INPUT_ALIAS_PERSON_LOOKUP)
    If VBA.Len(cellAddress) = 0 Then Exit Function

    private_Input_ReadPersonLookup = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadPersonLookup) = 0 Then
        ex_Helpers.LogError "Person lookup cell is empty | SheetNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
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

    private_Input_ReadPositionCode = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadPositionCode) = 0 Then
        ex_Helpers.LogError "Position code cell is empty | SheetNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
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

    private_Input_ReadTemplatePath = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadTemplatePath) = 0 Then
        ex_Helpers.LogError "Template path cell is empty | SheetNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
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

    private_Input_ReadRequiredValue = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
    If VBA.Len(private_Input_ReadRequiredValue) = 0 Then
        ex_Helpers.LogError "Required input is empty | Field=" & fieldCaption & _
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

    private_Input_ReadOptionalValue = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(cellAddress).Text))
End Function

Private Function private_Input_GetCellAddress(ByVal fieldAlias As String) As String
    If inputCellMap Is Nothing Then
        ex_Helpers.LogError "Input cell mapper is not initialized"
        VBA.MsgBox "Input cell mapper is not initialized.", _
            VBA.vbExclamation, "Document Generation"

        Exit Function
    End If

    If Not inputCellMap.Exists(fieldAlias) Then
        ex_Helpers.LogError "Input field alias is not mapped: " & fieldAlias
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
    inputIsDate = ex_Helpers.private_Date_LooksLikeFullDate(orderInput)
    If inputIsDate Then
        If Not ex_Helpers.private_Date_TryParse(orderInput, inputDate) Then
            VBA.MsgBox "Order date must be valid and use dd.mm.yyyy format.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
        orderYear = VBA.Year(inputDate)
    Else
        orderYear = VBA.Year(VBA.Date)
        orderInput = private_Order_NormalizeNumber(orderInput)
        If VBA.Len(orderInput) = 0 Then
            ex_Helpers.LogError "Order number is empty after normalization"
            VBA.MsgBox "Enter a valid order number.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
    End If

    If Not private_Order_TryGetTableRef(orderYear, ordersTableRef) Then Exit Function
    ordersPath = ex_Helpers.private_Path_ResolveFromWorkbook(ORDERS_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(ordersPath)) = 0 Then
        ex_Helpers.LogError "Orders workbook was not found: " & ordersPath
        VBA.MsgBox "Orders workbook was not found: " & ordersPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_ExternalTables.ex_ExternalTables_TryOpenConnection( _
        ordersPath, "Orders", connection) Then Exit Function

    sqlText = "SELECT [Номер наказу], [Дата наказу] FROM " & ordersTableRef
    ex_Helpers.LogDebug "Order lookup SQL: " & sqlText
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY
    Do While Not recordset.EOF
        candidateNo = private_Order_NormalizeNumber( _
            ex_ExternalTables.ex_ExternalTables_ReadText(recordset, "Номер наказу"))
        candidateDate = 0
        If ex_Helpers.private_Date_TryReadRecordsetDate( _
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
        ex_Helpers.LogError "Order was not found | Input=" & orderInput & _
            " | Year=" & VBA.CStr(orderYear)
        VBA.MsgBox "Order was not found for value '" & orderInput & _
            "' in the " & VBA.CStr(orderYear) & " orders table.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    If matchCount > 1 Then
        ex_Helpers.LogError "Order reference is ambiguous | Input=" & orderInput & _
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
    ex_Helpers.LogError "Order lookup failed | Number=" & VBA.CStr(Err.Number) & _
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
            ex_Helpers.LogError "Orders table is not configured for year " & _
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
    rawTicketNo = ex_Helpers.private_Text_Normalize(rawTicketNo)
    orderNo = private_Order_NormalizeNumber(orderNo)
    If Not ex_Helpers.private_Text_IsDigits(rawTicketNo) Or _
        Not ex_Helpers.private_Text_IsDigits(orderNo) Then
        ex_Helpers.LogError "Ticket number requires numeric order and ticket values"
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
    orderNo = ex_Helpers.private_Text_Normalize(orderNo)
    Do While VBA.Len(orderNo) > 1 And VBA.Left$(orderNo, 1) = "0"
        orderNo = VBA.Mid$(orderNo, 2)
    Loop
    private_Order_NormalizeNumber = orderNo
End Function
' --------------------------------------
' } // namespace Order
' --------------------------------------

' --------------------------------------
' namespace Diagnostics {
' --------------------------------------
Private Sub private_Diagnoctics_LogWorkbookContext()
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

    ex_Helpers.LogDebug "Workbook path: " & ThisWorkbook.FullName
    ex_Helpers.LogDebug "Worksheet count: " & VBA.CStr(ThisWorkbook.Worksheets.Count)
    ex_Helpers.LogDebug "Active sheet Unicode: " & _
        ex_Helpers.private_Text_ToUnicodeDebug(activeSheetText)

    For worksheetIndex = 1 To ThisWorkbook.Worksheets.Count
        Set worksheetObj = ThisWorkbook.Worksheets(worksheetIndex)
        ex_Helpers.LogDebug "Worksheet | Index=" & VBA.CStr(worksheetIndex) & _
            " | CodeName=" & worksheetObj.CodeName & _
            " | NameUnicode=" & ex_Helpers.private_Text_ToUnicodeDebug(worksheetObj.Name)
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

        ex_Helpers.LogDebug "Input binding | SheetNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & _
            " | PersonCell=" & personCellAddress & _
            " | PositionCell=" & positionCellAddress & _
            " | OrderCell=" & orderCellAddress & _
            " | TicketCell=" & ticketCellAddress & _
            " | DateFromCell=" & dateFromCellAddress & _
            " | DateToCell=" & dateToCellAddress & _
            " | TemplateCell=" & templateCellAddress

        ex_Helpers.LogDebug "Input raw values Unicode | Person=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(personCellAddress).Text)) & _
            " | Position=" & ex_Helpers.private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(positionCellAddress).Text)) & _
            " | Template=" & ex_Helpers.private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(templateCellAddress).Text))
    Else
        ex_Helpers.LogError "Input binding failed | ExpectedNameUnicode=" & _
            ex_Helpers.private_Text_ToUnicodeDebug(INPUT_SHEET_NAME)
    End If
#End If
End Sub
' --------------------------------------
' } // namespace Diagnostics
' --------------------------------------