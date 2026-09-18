Option Explicit

Private Const SHPO_RELATIVE_PATH As String = "Dependencies\ШПО.xlsx"
Private Const ORDERS_RELATIVE_PATH As String = "Dependencies\Накази.xlsx"
Private Const ORDERS_2025_TABLE_REF As String = "[Накази$A2:B12000]"
Private Const ORDERS_2026_TABLE_REF As String = "[Накази$D2:E12000]"
Private Const AD_OPEN_STATIC As Long = 3
Private Const AD_LOCK_READ_ONLY As Long = 1

Private Const ALF_TABLE_REF As String = "[АЛФ$A1:J12000]"
Private Const OS_TABLE_REF As String = "[ОС$A1:AB12000]"
Private Const RANKS_TABLE_REF As String = "[Звання$A1:E12000]"
Private Const POSITIONS_TABLE_REF As String = "[Посади$A1:E12000]"
Private Const POSITION_ROZP_PREFIX As String = "A1A"
Private Const POSITION_ROZP_TEXT_PREFIX As String = "у розпорядженні командира військової частини "
Private Const POSITION_ROZP_OFFICER_UNIT As String = "А3369"
Private Const POSITION_ROZP_OTHER_UNIT As String = "А7383"

Private shpoSessionConnection As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function ex_TryBeginSession() As Boolean
    Dim shpoPath As String

    On Error GoTo EH
    If Not shpoSessionConnection Is Nothing Then
        ex_TryBeginSession = True
        Exit Function
    End If
    shpoPath = ex_Helpers.private_Path_ResolveFromWorkbook( _
        SHPO_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(shpoPath)) = 0 Then
        ex_Helpers.LogError "SHPO file was not found: " & shpoPath
        ex_Helpers.ex_ShowErrorMessage "SHPO file was not found: " & shpoPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_ExternalTables.ex_TryOpenConnection( _
        shpoPath, "SHPO", shpoSessionConnection) Then Exit Function
    ex_Helpers.LogDebug "SHPO session connection opened"
    ex_TryBeginSession = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to open SHPO session | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to open SHPO session: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
    ex_EndSession
End Function

Public Sub ex_EndSession()
    On Error Resume Next
    If Not shpoSessionConnection Is Nothing Then shpoSessionConnection.Close
    Set shpoSessionConnection = Nothing
    On Error GoTo 0
End Sub

Public Function ex_TryResolveIpn( _
    ByVal personLookup As String, _
    ByRef outIpnText As String _
) As Boolean
    outIpnText = VBA.vbNullString
    If ex_Helpers.private_Text_IsIpn(personLookup) Then
        outIpnText = personLookup
        ex_TryResolveIpn = True
        Exit Function
    End If

    ex_TryResolveIpn = private_TryLookupShpoValue( _
        ALF_TABLE_REF, "ПІБ", personLookup, "ІПН", outIpnText)
End Function

' Возвращает кандидатов по части ПІБ или ІПН. Каждый элемент коллекции —
' массив из двух значений: ПІБ (индекс 0) и ІПН (индекс 1).
Public Function ex_TryFindPersonCandidates( _
    ByVal searchText As String, _
    ByVal maxCandidateCount As Long, _
    ByRef outCandidates As Collection _
) As Boolean
    Dim shpoPath As String
    Dim connection As Object, recordset As Object
    Dim usesSessionConnection As Boolean
    Dim normalizedSearchText As String
    Dim sqlText As String
    Dim fioText As String, ipnText As String

    On Error GoTo EH
    Set outCandidates = New Collection
    normalizedSearchText = ex_Helpers.private_Text_Normalize(searchText)
    If VBA.Len(normalizedSearchText) = 0 Then
        ex_TryFindPersonCandidates = True
        Exit Function
    End If
    If maxCandidateCount <= 0 Then
        ex_Helpers.ex_ShowErrorMessage "Maximum candidate count must be greater than zero.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    shpoPath = ex_Helpers.private_Path_ResolveFromWorkbook( _
        SHPO_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(shpoPath)) = 0 Then
        ex_Helpers.ex_ShowErrorMessage "SHPO file was not found: " & shpoPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not shpoSessionConnection Is Nothing Then
        Set connection = shpoSessionConnection
        usesSessionConnection = True
    ElseIf Not ex_ExternalTables.ex_TryOpenConnection( _
        shpoPath, "SHPO", connection) Then
        GoTo CleanExit
    End If

    sqlText = "SELECT TOP " & VBA.CStr(maxCandidateCount) & _
        " [ПІБ], [ІПН] FROM " & ALF_TABLE_REF & _
        " WHERE UCASE(TRIM(CSTR(IIF(ISNULL([ПІБ]), '', [ПІБ])))) LIKE '%" & _
        ex_ExternalTables.ex_EscapeSql(VBA.UCase$(normalizedSearchText)) & _
        "%' OR UCASE(TRIM(CSTR(IIF(ISNULL([ІПН]), '', [ІПН])))) LIKE '%" & _
        ex_ExternalTables.ex_EscapeSql(VBA.UCase$(normalizedSearchText)) & _
        "%' ORDER BY [ПІБ]"
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY
    Do While Not recordset.EOF
        fioText = ex_ExternalTables.ex_ReadText(recordset, "ПІБ")
        ipnText = ex_ExternalTables.ex_ReadText(recordset, "ІПН")
        If VBA.Len(fioText) > 0 And VBA.Len(ipnText) > 0 Then
            outCandidates.Add VBA.Array(fioText, ipnText)
        End If
        recordset.MoveNext
    Loop
    ex_TryFindPersonCandidates = True

CleanExit:
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    If Not usesSessionConnection And Not connection Is Nothing Then connection.Close
    Set recordset = Nothing
    Set connection = Nothing
    On Error GoTo 0
    Exit Function
EH:
    ex_Helpers.LogError "SHPO candidate lookup failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to search SHPO candidates: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

' Совместимый API для ex_VacationPersonAutocomplete. Возвращает только ПІБ,
' потому что этот модуль записывает выбранную строку непосредственно в C4.
Public Function ex_TryFindFioCandidates( _
    ByVal fioPart As String, _
    ByRef outCandidates As Collection _
) As Boolean
    Const MAX_CANDIDATES As Long = 50

    Dim recordset As Object
    Dim sqlText As String
    Dim normalizedPart As String
    Dim fioText As String

    On Error GoTo EH
    Set outCandidates = New Collection
    normalizedPart = ex_Helpers.private_Text_Normalize(fioPart)
    If VBA.Len(normalizedPart) < 2 Then
        ex_Helpers.ex_ShowErrorMessage "Enter at least two FIO characters to search SHPO.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_TryBeginSession() Then Exit Function

    sqlText = "SELECT TOP " & VBA.CStr(MAX_CANDIDATES) & " [ПІБ] FROM " & _
        ALF_TABLE_REF & " WHERE UCASE(TRIM(CSTR(IIF(ISNULL([ПІБ]), '', [ПІБ])))) " & _
        "LIKE '%" & ex_ExternalTables.ex_EscapeSql( _
            VBA.UCase$(normalizedPart)) & "%' ORDER BY [ПІБ]"
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, shpoSessionConnection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY
    Do While Not recordset.EOF
        fioText = ex_ExternalTables.ex_ReadText(recordset, "ПІБ")
        If VBA.Len(fioText) > 0 Then outCandidates.Add fioText
        recordset.MoveNext
    Loop
    ex_TryFindFioCandidates = True

CleanExit:
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    Set recordset = Nothing
    On Error GoTo 0
    Exit Function
EH:
    ex_Helpers.LogError "FIO candidates lookup failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to get FIO candidates from SHPO: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

Public Function ex_TryResolveOrderReference( _
    ByVal orderInput As String, _
    ByRef outOrderNo As String, _
    ByRef outOrderDate As Date _
) As Boolean
    Dim ordersPath As String, ordersTableRef As String
    Dim connection As Object, recordset As Object
    Dim sqlText As String, inputIsDate As Boolean, inputDate As Date
    Dim orderYear As Long, candidateNo As String, candidateDate As Date
    Dim matchCount As Long

    On Error GoTo EH
    outOrderNo = VBA.vbNullString
    outOrderDate = 0
    inputIsDate = ex_Helpers.private_Date_LooksLikeFullDate(orderInput)
    If inputIsDate Then
        If Not ex_Helpers.private_Date_TryParse(orderInput, inputDate) Then
            ex_Helpers.ex_ShowErrorMessage "Order date must be valid and use dd.mm.yyyy format.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
        orderYear = VBA.Year(inputDate)
    Else
        orderYear = VBA.Year(VBA.Date)
        orderInput = private_Order_NormalizeNumber(orderInput)
        If VBA.Len(orderInput) = 0 Then
            ex_Helpers.LogError "Order number is empty after normalization"
            ex_Helpers.ex_ShowErrorMessage "Enter a valid order number.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
    End If

    If Not private_Order_TryGetTableRef(orderYear, ordersTableRef) Then Exit Function
    ordersPath = ex_Helpers.private_Path_ResolveFromWorkbook(ORDERS_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(ordersPath)) = 0 Then
        ex_Helpers.LogError "Orders workbook was not found: " & ordersPath
        ex_Helpers.ex_ShowErrorMessage "Orders workbook was not found: " & ordersPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_ExternalTables.ex_TryOpenConnection( _
        ordersPath, "Orders", connection) Then Exit Function

    sqlText = "SELECT [Номер наказу], [Дата наказу] FROM " & ordersTableRef
    ex_Helpers.LogDebug "Order lookup SQL: " & sqlText
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY
    Do While Not recordset.EOF
        candidateNo = private_Order_NormalizeNumber( _
            ex_ExternalTables.ex_ReadText(recordset, "Номер наказу"))
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
        ex_Helpers.ex_ShowErrorMessage "Order was not found for value '" & orderInput & _
            "' in the " & VBA.CStr(orderYear) & " orders table.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    If matchCount > 1 Then
        ex_Helpers.LogError "Order reference is ambiguous | Input=" & orderInput & _
            " | Matches=" & VBA.CStr(matchCount)
        ex_Helpers.ex_ShowErrorMessage "Multiple orders were found for value '" & orderInput & "'.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If

    ex_TryResolveOrderReference = True
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
    ex_Helpers.ex_ShowErrorMessage "Order lookup failed: [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

' Ищет единственный приказ по дате без ошибки, если приказ ещё не заведён.
Public Function ex_TryFindOrderNoByDate( _
    ByVal orderDate As Date, _
    ByRef outOrderNo As String _
) As Boolean
    Dim ordersPath As String, ordersTableRef As String
    Dim connection As Object, recordset As Object
    Dim sqlText As String
    Dim candidateNo As String, candidateDate As Date
    Dim matchCount As Long

    On Error GoTo EH
    outOrderNo = VBA.vbNullString
    Select Case VBA.Year(orderDate)
        Case 2025: ordersTableRef = ORDERS_2025_TABLE_REF
        Case 2026: ordersTableRef = ORDERS_2026_TABLE_REF
        Case Else
            ex_Helpers.LogDebug "Orders table is not configured for planned date " & _
                VBA.CStr(orderDate)
            ex_TryFindOrderNoByDate = True
            Exit Function
    End Select
    ordersPath = ex_Helpers.private_Path_ResolveFromWorkbook(ORDERS_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(ordersPath)) = 0 Then
        ex_Helpers.ex_ShowErrorMessage "Orders workbook was not found: " & ordersPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not ex_ExternalTables.ex_TryOpenConnection( _
        ordersPath, "Orders", connection) Then Exit Function
    sqlText = "SELECT [Номер наказу], [Дата наказу] FROM " & ordersTableRef
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY
    Do While Not recordset.EOF
        candidateDate = 0
        If ex_Helpers.private_Date_TryReadRecordsetDate( _
            recordset.Fields("Дата наказу").Value, candidateDate) Then
            If VBA.DateValue(candidateDate) = VBA.DateValue(orderDate) Then
                matchCount = matchCount + 1
                candidateNo = private_Order_NormalizeNumber( _
                    ex_ExternalTables.ex_ReadText(recordset, "Номер наказу"))
                If matchCount = 1 Then outOrderNo = candidateNo
            End If
        End If
        recordset.MoveNext
    Loop
    If matchCount <> 1 Then outOrderNo = VBA.vbNullString
    If matchCount > 1 Then
        ex_Helpers.LogDebug "Planned donation date has multiple orders | Date=" & _
            VBA.CStr(orderDate) & " | Matches=" & VBA.CStr(matchCount)
    End If
    ex_TryFindOrderNoByDate = True
CleanExit:
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    If Not connection Is Nothing Then connection.Close
    Set recordset = Nothing
    Set connection = Nothing
    On Error GoTo 0
    Exit Function
EH:
    ex_Helpers.LogError "Planned donation order lookup failed | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "Planned donation order lookup failed: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

Public Function ex_TryBuildTicketNo( _
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
        ex_Helpers.ex_ShowErrorMessage "Order number and ticket number must contain digits only.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outTicketNo = VBA.CStr(VBA.Year(orderDate)) & "/" & _
        orderNo & "/" & rawTicketNo
    ex_TryBuildTicketNo = True
End Function

Public Function ex_TryResolveFioNominative( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_TryResolveFioNominative = private_TryResolveFioCase( _
        ipnText, "ПІБ", outFioText)
End Function

Public Function ex_TryResolveFioGenitive( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_TryResolveFioGenitive = private_TryResolveFioCase( _
        ipnText, "Родовий", outFioText)
End Function

Public Function ex_TryResolveFioDative( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_TryResolveFioDative = private_TryResolveFioCase( _
        ipnText, "Давальний", outFioText)
End Function

Public Function ex_TryResolveFioAccusative( _
    ByVal ipnText As String, _
    ByRef outFioText As String _
) As Boolean
    ex_TryResolveFioAccusative = private_TryResolveFioCase( _
        ipnText, "Знахідний", outFioText)
End Function

Public Function ex_TryResolveRankNominative( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_TryResolveRankNominative = private_TryResolveRankCase( _
        ipnText, "Звання", outRankText)
End Function

Public Function ex_TryResolvePositionCode( _
    ByVal ipnText As String, _
    ByRef outPositionCode As String _
) As Boolean
    ex_TryResolvePositionCode = private_TryLookupShpoValue( _
        OS_TABLE_REF, "ІПН", ipnText, "Код посади", outPositionCode)
End Function

Public Function ex_TryResolveRankGenitive( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_TryResolveRankGenitive = private_TryResolveRankCase( _
        ipnText, "Родовий", outRankText)
End Function

Public Function ex_TryResolveRankDative( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_TryResolveRankDative = private_TryResolveRankCase( _
        ipnText, "Давальний", outRankText)
End Function

Public Function ex_TryResolveRankAccusative( _
    ByVal ipnText As String, _
    ByRef outRankText As String _
) As Boolean
    ex_TryResolveRankAccusative = private_TryResolveRankCase( _
        ipnText, "Знахідний", outRankText)
End Function

Public Function ex_TryResolvePositionNominative( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_TryResolvePositionNominative = private_TryResolvePositionCase( _
        positionCode, rankText, "Назва", outPositionText)
End Function

Public Function ex_TryResolvePositionGenitive( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_TryResolvePositionGenitive = private_TryResolvePositionCase( _
        positionCode, rankText, "Родовий", outPositionText)
End Function

Public Function ex_TryResolvePositionDative( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_TryResolvePositionDative = private_TryResolvePositionCase( _
        positionCode, rankText, "Давальний", outPositionText)
End Function

Public Function ex_TryResolvePositionAccusative( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByRef outPositionText As String _
) As Boolean
    ex_TryResolvePositionAccusative = private_TryResolvePositionCase( _
        positionCode, rankText, "Знахідний", outPositionText)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryLookupShpoValue( _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal resultHeader As String, _
    ByRef outValue As String _
) As Boolean
    Dim shpoPath As String
    Dim connection As Object
    Dim usesSessionConnection As Boolean

    On Error GoTo EH
    outValue = VBA.vbNullString
    shpoPath = ex_Helpers.private_Path_ResolveFromWorkbook( _
        SHPO_RELATIVE_PATH)
    If VBA.Len(VBA.Dir$(shpoPath)) = 0 Then
        ex_Helpers.LogError "SHPO file was not found: " & shpoPath
        ex_Helpers.ex_ShowErrorMessage "SHPO file was not found: " & shpoPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    If Not shpoSessionConnection Is Nothing Then
        Set connection = shpoSessionConnection
        usesSessionConnection = True
    Else
        If Not ex_ExternalTables.ex_TryOpenConnection( _
            shpoPath, "SHPO", connection) Then GoTo CleanExit
    End If
    private_TryLookupShpoValue = ex_ExternalTables.ex_TryLookup( _
        connection, "SHPO", tableRef, keyHeader, keyValue, resultHeader, outValue)

CleanExit:
    On Error Resume Next
    If Not usesSessionConnection And Not connection Is Nothing Then connection.Close
    Set connection = Nothing
    On Error GoTo 0
    Exit Function

EH:
    ex_Helpers.LogError "SHPO lookup failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    ex_Helpers.ex_ShowErrorMessage "SHPO lookup failed: [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

Private Function private_TryResolveFioCase( _
    ByVal ipnText As String, _
    ByVal resultHeader As String, _
    ByRef outFioText As String _
) As Boolean
    private_TryResolveFioCase = private_TryLookupShpoValue( _
        ALF_TABLE_REF, "ІПН", ipnText, resultHeader, outFioText)
End Function

Private Function private_TryResolveRankCase( _
    ByVal ipnText As String, _
    ByVal resultHeader As String, _
    ByRef outRankText As String _
) As Boolean
    Dim rankNominative As String

    If Not private_TryLookupShpoValue( _
        OS_TABLE_REF, "ІПН", ipnText, "Військове звання фактично", _
        rankNominative) Then Exit Function

    private_TryResolveRankCase = private_TryLookupShpoValue( _
        RANKS_TABLE_REF, "Звання", rankNominative, resultHeader, outRankText)
End Function

Private Function private_TryResolvePositionCase( _
    ByVal positionCode As String, _
    ByVal rankText As String, _
    ByVal resultHeader As String, _
    ByRef outPositionText As String _
) As Boolean
    Dim normalizedCode As String

    normalizedCode = VBA.UCase$(ex_Helpers.private_Text_Normalize(positionCode))
    If private_IsRozpPositionCode(normalizedCode) Then
        If private_IsOfficerRank(rankText) Then
            outPositionText = POSITION_ROZP_TEXT_PREFIX & _
                POSITION_ROZP_OFFICER_UNIT
        Else
            outPositionText = POSITION_ROZP_TEXT_PREFIX & _
                POSITION_ROZP_OTHER_UNIT
        End If
        private_TryResolvePositionCase = True
        Exit Function
    End If

    private_TryResolvePositionCase = private_TryLookupShpoValue( _
        POSITIONS_TABLE_REF, "Код", normalizedCode, resultHeader, outPositionText)
End Function

Private Function private_IsRozpPositionCode( _
    ByVal positionCode As String _
) As Boolean
    positionCode = VBA.Replace$(positionCode, " ", VBA.vbNullString)
    positionCode = VBA.Replace$(positionCode, "А", "A")
    private_IsRozpPositionCode = ( _
        VBA.Left$(positionCode, VBA.Len(POSITION_ROZP_PREFIX)) = _
        POSITION_ROZP_PREFIX)
End Function

Private Function private_IsOfficerRank(ByVal rankText As String) As Boolean
    Select Case VBA.LCase$(ex_Helpers.private_Text_Normalize(rankText))
        Case "молодший лейтенант", "лейтенант", "старший лейтенант", _
             "капітан", "майор", "підполковник", "полковник"
            private_IsOfficerRank = True
    End Select
End Function

' --------------------------------------
' namespace Order {
' --------------------------------------
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
            ex_Helpers.ex_ShowErrorMessage "Orders table is not configured for year " & _
                VBA.CStr(orderYear) & ".", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
    End Select
    private_Order_TryGetTableRef = True
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
' -------------------------------------