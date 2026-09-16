Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True

Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

' --------------------------------------
' namespace Word {
' --------------------------------------
' Общие операции форматирования, работы с путями, Word и журналом.
Public Function private_Word_TryGenerateDocument( _
    ByVal templatePathInput As String, _
    ByVal documentName As String, _
    ByVal placeholderNames As Variant, _
    ByVal placeholderValues As Variant, _
    ByRef outDocumentPath As String, _
    Optional ByVal outputFolderPathInput As String = "" _
) As Boolean
    Dim templatePath As String
    Dim outputFolderPath As String
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

    outputFolderPath = private_Text_Normalize(outputFolderPathInput)
    If VBA.Len(outputFolderPath) > 0 Then
        outputFolderPath = private_Path_ResolveFromWorkbook(outputFolderPath)
        If Not private_Path_TryEnsureFolder(outputFolderPath) Then Exit Function
    End If
    outDocumentPath = private_Path_BuildGeneratedDocumentPath( _
        templatePath, documentName, outputFolderPath)
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
    wordDoc.Close True
    Set wordDoc = Nothing
    If ownsWordApp Then wordApp.Quit
    Set wordApp = Nothing
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

Public Function private_Word_TryReplacePlaceholder( _
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
Public Function private_Date_TryFormat( _
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

Public Function private_Date_TryParse( _
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

Public Function private_Date_LooksLikeFullDate( _
    ByVal valueText As String _
) As Boolean
    private_Date_LooksLikeFullDate = ( _
        VBA.InStr(1, valueText, ".", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, valueText, "/", VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, valueText, "-", VBA.vbBinaryCompare) > 0)
End Function

Public Function private_Date_TryReadRecordsetDate( _
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

Public Function private_Date_GetUaMonthGenitive( _
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
' namespace Text {
' --------------------------------------
Public Function private_Text_IsIpn(ByVal valueText As String) As Boolean
    private_Text_IsIpn = (VBA.Len(valueText) > 0 And _
        Not valueText Like "*[!0-9]*")
End Function

Public Function private_Text_IsDigits(ByVal valueText As String) As Boolean
    valueText = private_Text_Normalize(valueText)
    private_Text_IsDigits = (VBA.Len(valueText) > 0 And _
        Not valueText Like "*[!0-9]*")
End Function

' Форматирует строку по именованным токенам: "{FIO} - {OrderNo}".
' Литеральные фигурные скобки задаются как "{{" и "}}".
Public Function private_Text_TryFormat( _
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

Public Function private_Path_ResolveFromWorkbook( _
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

Public Function private_Path_BuildGeneratedDocumentPath( _
    ByVal templatePath As String, _
    ByVal documentName As String, _
    Optional ByVal outputFolderPath As String = "" _
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

    If VBA.Len(outputFolderPath) > 0 Then
        folderPath = outputFolderPath
        If VBA.Right$(folderPath, 1) <> Application.PathSeparator Then _
            folderPath = folderPath & Application.PathSeparator
    Else
        slashPosition = VBA.InStrRev(templatePath, Application.PathSeparator)
        If slashPosition > 0 Then _
            folderPath = VBA.Left$(templatePath, slashPosition)
    End If
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

' Создаёт папку результата и отсутствующие родительские папки.
Private Function private_Path_TryEnsureFolder( _
    ByVal folderPath As String _
) As Boolean
    Dim parentPath As String
    Dim separatorPosition As Long

    On Error GoTo EH
    If VBA.Len(VBA.Dir$(folderPath, VBA.vbDirectory)) > 0 Then
        private_Path_TryEnsureFolder = True
        Exit Function
    End If
    separatorPosition = VBA.InStrRev(folderPath, Application.PathSeparator)
    If separatorPosition = 0 Then GoTo EH
    parentPath = VBA.Left$(folderPath, separatorPosition - 1)
    If VBA.Len(parentPath) = 0 Then GoTo EH
    If Not private_Path_TryEnsureFolder(parentPath) Then Exit Function
    VBA.MkDir folderPath
    LogDebug "Output folder created: " & folderPath
    private_Path_TryEnsureFolder = True
    Exit Function
EH:
    LogError "Failed to create output folder: " & folderPath & _
        " | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    VBA.MsgBox "Failed to create output folder: " & folderPath & _
        " | " & Err.Description, VBA.vbExclamation, "Document Generation"
End Function

Public Function private_Path_SanitizeFileName( _
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

Public Function private_Text_Normalize(ByVal valueText As String) As String
    valueText = VBA.Replace$(valueText, VBA.ChrW$(160), " ")
    valueText = VBA.Trim$(valueText)
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace$(valueText, "  ", " ")
    Loop
    private_Text_Normalize = valueText
End Function

Public Function private_Text_ToUnicodeDebug(ByVal valueText As String) As String
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

' --------------------------------------
' } // namespace Text
' --------------------------------------

' --------------------------------------
' namespace Logging {
' --------------------------------------
Public Sub ClearLog()
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Output As #fileNumber
    Close #fileNumber
#End If
End Sub

Public Sub LogError(ByVal messageText As String)
#If ENABLE_LOGGING Then
    WriteLog "ERROR: " & messageText
#End If
End Sub

Public Sub LogDebug(ByVal messageText As String)
#If ENABLE_LOGGING Then
#If ENABLE_DEBUG_LOGGING Then
    WriteLog "DEBUG: " & messageText
#End If
#End If
End Sub

Public Sub WriteLog(ByVal messageText As String)
#If ENABLE_LOGGING Then
    Dim fileNumber As Integer

    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Append As #fileNumber
    Print #fileNumber, VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & messageText
    Close #fileNumber
#End If
End Sub

Public Function GetLogFilePath() As String
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