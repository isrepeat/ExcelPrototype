Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

' Явные ссылки текущего шаблона генерации документов.
Private Const INPUT_SHEET_NAME As String = "Відрядження"
Private Const PERSON_LOOKUP_CELL_ADDRESS As String = "C4"
Private Const POSITION_CODE_CELL_ADDRESS As String = "C5"
Private Const SHPO_RELATIVE_PATH As String = "ШПО.xlsx"

Private Const ALF_TABLE_REF As String = "[АЛФ$A1:J12000]"
Private Const OS_TABLE_REF As String = "[ОС$A1:AB12000]"
Private Const RANKS_TABLE_REF As String = "[Звання$A1:E12000]"
Private Const POSITIONS_TABLE_REF As String = "[Посади$A1:E12000]"
Private Const POSITION_ROZP_PREFIX As String = "A1A"
Private Const POSITION_ROZP_TEXT_PREFIX As String = _
    "у розпорядженні командира військової частини "
Private Const POSITION_ROZP_OFFICER_UNIT As String = "А3369"
Private Const POSITION_ROZP_OTHER_UNIT As String = "А7383"

Private Const AD_OPEN_STATIC As Long = 3
Private Const AD_LOCK_READ_ONLY As Long = 1
Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

' --------------------------------------
' namespace API {
' --------------------------------------
' Назначить этот макрос кнопке "Створити".
Public Sub fn_DocumentsGeneration_Create()
    Dim personLookup As String
    Dim shpoPath As String
    Dim connection As Object
    Dim ipnText As String
    Dim fioDative As String
    Dim rankText As String
    Dim rankDative As String
    Dim positionCode As String
    Dim positionText As String

    On Error GoTo EH

    ClearLog
    LogDebug "Generation started"
    private_Log_WorkbookContext
    personLookup = private_Input_ReadPersonLookup()
    If VBA.Len(personLookup) = 0 Then Exit Sub
    positionCode = private_Input_ReadPositionCode()
    If VBA.Len(positionCode) = 0 Then Exit Sub
    LogDebug "Person lookup: " & personLookup
    LogDebug "Person lookup Unicode: " & private_Text_ToUnicodeDebug(personLookup)
    LogDebug "Position code: " & positionCode

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
    LogDebug "Resolved rank: " & rankText
    LogDebug "Resolved position: " & positionText
    WriteLog "RESULT: " & rankDative & " " & fioDative & _
        " | " & positionText

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

    private_Input_ReadPersonLookup = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(PERSON_LOOKUP_CELL_ADDRESS).Text))
    If VBA.Len(private_Input_ReadPersonLookup) = 0 Then
        LogError "Person lookup cell is empty | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
            PERSON_LOOKUP_CELL_ADDRESS
        VBA.MsgBox "Enter FIO or IPN in cell " & _
            PERSON_LOOKUP_CELL_ADDRESS & ".", _
            VBA.vbExclamation, "Document Generation"
    End If
End Function

Private Function private_Input_ReadPositionCode() As String
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    private_Input_ReadPositionCode = private_Text_Normalize( _
        VBA.CStr(sourceSheet.Range(POSITION_CODE_CELL_ADDRESS).Text))
    If VBA.Len(private_Input_ReadPositionCode) = 0 Then
        LogError "Position code cell is empty | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & " | Cell=" & _
            POSITION_CODE_CELL_ADDRESS
        VBA.MsgBox "Enter the position code in cell " & _
            POSITION_CODE_CELL_ADDRESS & ".", _
            VBA.vbExclamation, "Document Generation"
    End If
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
    On Error GoTo EH
    Set outConnection = VBA.CreateObject("ADODB.Connection")
    outConnection.Open _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & shpoPath & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    private_Shpo_TryOpenConnection = True
    Exit Function
EH:
    LogError "Failed to open SHPO | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    VBA.MsgBox "Failed to open SHPO: [" & VBA.CStr(Err.Number) & _
        "] " & Err.Description, VBA.vbExclamation, "Document Generation"
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
' namespace Helpers {
' --------------------------------------
Private Function private_Text_IsIpn(ByVal valueText As String) As Boolean
    private_Text_IsIpn = (VBA.Len(valueText) > 0 And _
        Not valueText Like "*[!0-9]*")
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
        LogDebug "Input binding | SheetNameUnicode=" & _
            private_Text_ToUnicodeDebug(INPUT_SHEET_NAME) & _
            " | PersonCell=" & PERSON_LOOKUP_CELL_ADDRESS & _
            " | PositionCell=" & POSITION_CODE_CELL_ADDRESS
        LogDebug "Input raw values Unicode | Person=" & _
            private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(PERSON_LOOKUP_CELL_ADDRESS).Text)) & _
            " | Position=" & private_Text_ToUnicodeDebug(VBA.CStr( _
                worksheetObj.Range(POSITION_CODE_CELL_ADDRESS).Text))
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
