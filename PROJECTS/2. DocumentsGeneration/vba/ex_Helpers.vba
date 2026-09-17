Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True
#Const CLEAR_LOG_ON_GENERATION = False

Private Const LOG_FILE_SUFFIX As String = "_logs.txt"

Private managedWordApp As Object
Private messageTargetRange As Range

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
    Optional ByVal outputFolderPathInput As String = "", _
    Optional ByVal overwriteDocumentPathInput As String = "", _
    Optional ByVal failIfDocumentExists As Boolean = False, _
    Optional ByVal renameUpdatedDocument As Boolean = False _
) As Boolean
    Dim templatePath As String
    Dim outputFolderPath As String
    Dim overwriteDocumentPath As String
    Dim temporaryDocumentPath As String
    Dim finalDocumentPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim placeholderIndex As Long

    On Error GoTo EH
    outDocumentPath = VBA.vbNullString
    templatePath = private_Path_ResolveFromWorkbook(templatePathInput)
    If VBA.Len(VBA.Dir$(templatePath)) = 0 Then
        LogError "Word template was not found: " & templatePath
        ex_ShowErrorMessage "Word template was not found: " & templatePath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    outputFolderPath = private_Text_Normalize(outputFolderPathInput)
    If VBA.Len(outputFolderPath) > 0 Then
        outputFolderPath = private_Path_ResolveFromWorkbook(outputFolderPath)
        If Not private_Path_TryEnsureFolder(outputFolderPath) Then Exit Function
    End If
    overwriteDocumentPath = private_Text_Normalize(overwriteDocumentPathInput)
    If VBA.Len(overwriteDocumentPath) > 0 Then
        overwriteDocumentPath = private_Path_ResolveFromWorkbook(overwriteDocumentPath)
        If VBA.Len(VBA.Dir$(overwriteDocumentPath)) = 0 Then
            LogError "Saved vacation ticket file was not found: " & overwriteDocumentPath
            ex_ShowErrorMessage "Saved vacation ticket file was not found: " & _
                overwriteDocumentPath, VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
        outDocumentPath = overwriteDocumentPath
        temporaryDocumentPath = private_Path_BuildTemporaryDocumentPath( _
            overwriteDocumentPath)
        If renameUpdatedDocument Then
            If Not private_Path_TryBuildExactDocumentPath( _
                templatePath, documentName, outputFolderPath, finalDocumentPath) Then Exit Function
            If VBA.StrComp(finalDocumentPath, outDocumentPath, _
                    VBA.vbTextCompare) <> 0 And _
               VBA.Len(VBA.Dir$(finalDocumentPath)) > 0 Then
                ex_ShowErrorMessage "The target vacation ticket file name is already in use: " & _
                    finalDocumentPath, VBA.vbExclamation, "Document Generation"
                Exit Function
            End If
        End If
    Else
        outDocumentPath = private_Path_BuildGeneratedDocumentPath( _
            templatePath, documentName, outputFolderPath, _
            Not failIfDocumentExists)
        temporaryDocumentPath = outDocumentPath
    End If
    If VBA.Len(outDocumentPath) = 0 Then Exit Function
    If VBA.Len(temporaryDocumentPath) = 0 Then Exit Function
    VBA.FileCopy templatePath, temporaryDocumentPath
    LogDebug "Word template copied | Source=" & templatePath & _
        " | Target=" & temporaryDocumentPath

    If Not managedWordApp Is Nothing Then
        Set wordApp = managedWordApp
        LogDebug "Managed Word application reused"
    Else
        On Error Resume Next
        Set wordApp = VBA.GetObject(, "Word.Application")
        On Error GoTo EH
        If wordApp Is Nothing Then
            Set wordApp = VBA.CreateObject("Word.Application")
            Set managedWordApp = wordApp
            LogDebug "Managed Word application started"
        Else
            LogDebug "Existing Word application used"
        End If
    End If

    Set wordDoc = wordApp.Documents.Open(temporaryDocumentPath)
    For placeholderIndex = LBound(placeholderNames) To UBound(placeholderNames)
        If Not private_Word_TryReplacePlaceholder( _
            wordDoc, VBA.CStr(placeholderNames(placeholderIndex)), _
            VBA.CStr(placeholderValues(placeholderIndex))) Then GoTo CleanFail
    Next placeholderIndex

    wordDoc.Save
    wordDoc.Close True
    Set wordDoc = Nothing
    If VBA.StrComp(temporaryDocumentPath, outDocumentPath, VBA.vbTextCompare) <> 0 Then
        VBA.Kill outDocumentPath
        Name temporaryDocumentPath As outDocumentPath
    End If
    If VBA.Len(finalDocumentPath) > 0 And _
       VBA.StrComp(finalDocumentPath, outDocumentPath, VBA.vbTextCompare) <> 0 Then
        Name outDocumentPath As finalDocumentPath
        outDocumentPath = finalDocumentPath
    End If
    Set wordApp = Nothing
    private_Word_TryGenerateDocument = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If VBA.Len(temporaryDocumentPath) > 0 Then
        If VBA.Len(VBA.Dir$(temporaryDocumentPath)) > 0 Then VBA.Kill temporaryDocumentPath
    End If
    On Error GoTo 0
    Exit Function
EH:
    LogError "Word generation failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    ex_ShowErrorMessage "Word generation failed: [" & VBA.CStr(Err.Number) & _
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
        ex_ShowErrorMessage "Required Word placeholder was not found: " & markerText, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    markerRange.Text = replacementText
    LogDebug "Word placeholder replaced | Name=" & placeholderName & _
        " | Value=" & replacementText
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
            formatPattern
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
    LogError "Invalid date value: " & VBA.CStr(dateValue)
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
            formatPattern
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
    Optional ByVal outputFolderPath As String = "", _
    Optional ByVal allowCopySuffix As Boolean = True _
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
        ex_ShowErrorMessage "Generated document file name is empty.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If

    candidatePath = folderPath & fileNameBase & extensionText
    If Not allowCopySuffix And VBA.Len(VBA.Dir$(candidatePath)) > 0 Then
        LogError "A generated document already exists: " & candidatePath
        ex_ShowErrorMessage "A document already exists for this vacation ticket: " & _
            candidatePath, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    copyIndex = 2
    Do While VBA.Len(VBA.Dir$(candidatePath)) > 0
        candidatePath = folderPath & fileNameBase & " (" & _
            VBA.CStr(copyIndex) & ")" & extensionText
        copyIndex = copyIndex + 1
    Loop
    private_Path_BuildGeneratedDocumentPath = candidatePath
End Function

' Формирует точный путь нового документа без автоматического добавления суффикса.
Private Function private_Path_TryBuildExactDocumentPath( _
    ByVal templatePath As String, _
    ByVal documentName As String, _
    ByVal outputFolderPath As String, _
    ByRef outDocumentPath As String _
) As Boolean
    Dim folderPath As String
    Dim fileNameBase As String
    Dim dotPosition As Long
    Dim slashPosition As Long
    Dim extensionText As String

    outDocumentPath = VBA.vbNullString
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
        If slashPosition = 0 Then
            ex_ShowErrorMessage "Cannot determine the document output folder.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
        folderPath = VBA.Left$(templatePath, slashPosition)
    End If
    fileNameBase = private_Path_SanitizeFileName(documentName)
    If VBA.Len(fileNameBase) = 0 Then
        ex_ShowErrorMessage "Generated document file name is empty.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outDocumentPath = folderPath & fileNameBase & extensionText
    private_Path_TryBuildExactDocumentPath = True
End Function

' Перемещает прежний документ в свободное имя с суффиксом «(old N)».
Public Function private_Path_TryArchiveDocument( _
    ByVal documentPath As String _
) As Boolean
    Dim dotPosition As Long
    Dim basePath As String
    Dim extensionText As String
    Dim archivedPath As String
    Dim archiveIndex As Long

    On Error GoTo EH
    If VBA.Len(VBA.Dir$(documentPath)) = 0 Then
        ex_ShowErrorMessage "Document to archive was not found: " & documentPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    dotPosition = VBA.InStrRev(documentPath, ".")
    If dotPosition = 0 Then
        ex_ShowErrorMessage "Document to archive has no extension: " & documentPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    basePath = VBA.Left$(documentPath, dotPosition - 1)
    extensionText = VBA.Mid$(documentPath, dotPosition)
    archiveIndex = 1
    Do
        archivedPath = basePath & " (old " & VBA.CStr(archiveIndex) & ")" & _
            extensionText
        archiveIndex = archiveIndex + 1
    Loop While VBA.Len(VBA.Dir$(archivedPath)) > 0
    Name documentPath As archivedPath
    private_Path_TryArchiveDocument = True
    Exit Function
EH:
    LogError "Failed to archive document | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    ex_ShowErrorMessage "Failed to archive document: " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

' Проверяет, что существующий Word-файл не открыт и доступен для замены.
Public Function private_Path_TryEnsureDocumentWritable( _
    ByVal documentPath As String _
) As Boolean
    Dim fileNumber As Integer
    Dim isOpen As Boolean

    On Error GoTo EH
    If VBA.Len(VBA.Dir$(documentPath)) = 0 Then
        ex_ShowErrorMessage "Document to overwrite was not found: " & documentPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    fileNumber = VBA.FreeFile
    Open documentPath For Binary Access Read Write Lock Read Write As #fileNumber
    isOpen = True
    Close #fileNumber
    private_Path_TryEnsureDocumentWritable = True
    Exit Function
EH:
    If isOpen Then Close #fileNumber
    LogError "Document is not writable | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description & " | Path=" & documentPath
    ex_ShowErrorMessage "The Word file is open or cannot be replaced. Close it " & _
        "and try again: " & documentPath, VBA.vbExclamation, _
        "Document Generation"
End Function

' Проверяет, что Word-шаблон доступен для чтения; открытый в Word файл допустим.
Public Function private_Path_TryEnsureDocumentReadable( _
    ByVal documentPathInput As String _
) As Boolean
    Dim documentPath As String
    Dim fileNumber As Integer
    Dim isOpen As Boolean

    On Error GoTo EH
    documentPath = private_Path_ResolveFromWorkbook(documentPathInput)
    If VBA.Len(VBA.Dir$(documentPath)) = 0 Then
        ex_ShowErrorMessage "Word template was not found: " & documentPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    fileNumber = VBA.FreeFile
    Open documentPath For Binary Access Read Shared As #fileNumber
    isOpen = True
    Close #fileNumber
    private_Path_TryEnsureDocumentReadable = True
    Exit Function
EH:
    If isOpen Then Close #fileNumber
    LogError "Document is not readable | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description & " | Path=" & documentPath
    ex_ShowErrorMessage "The Word template cannot be read. Close it if it is " & _
        "locked and try again: " & documentPath, VBA.vbExclamation, _
        "Document Generation"
End Function

' Пробной копией проверяет именно операцию FileCopy, используемую генератором Word.
Public Function private_Path_TryProbeTemplateCopy( _
    ByVal templatePathInput As String, _
    ByVal outputFolderPathInput As String _
) As Boolean
    Dim templatePath As String
    Dim outputFolderPath As String
    Dim probePath As String
    Dim fileSystem As Object

    On Error GoTo EH
    templatePath = private_Path_ResolveFromWorkbook(templatePathInput)
    If VBA.Len(VBA.Dir$(templatePath)) = 0 Then
        ex_ShowErrorMessage "Word template was not found: " & templatePath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outputFolderPath = private_Text_Normalize(outputFolderPathInput)
    If VBA.Len(outputFolderPath) = 0 Then
        outputFolderPath = VBA.Left$(templatePath, _
            VBA.InStrRev(templatePath, Application.PathSeparator) - 1)
    Else
        outputFolderPath = private_Path_ResolveFromWorkbook(outputFolderPath)
    End If
    If Not private_Path_TryEnsureFolder(outputFolderPath) Then Exit Function
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    probePath = fileSystem.BuildPath(outputFolderPath, fileSystem.GetTempName)
    VBA.FileCopy templatePath, probePath
    VBA.Kill probePath
    private_Path_TryProbeTemplateCopy = True
    Exit Function
EH:
    On Error Resume Next
    If VBA.Len(probePath) > 0 Then
        If VBA.Len(VBA.Dir$(probePath)) > 0 Then VBA.Kill probePath
    End If
    On Error GoTo 0
    LogError "Template copy preflight failed | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description & " | Template=" & templatePath
    ex_ShowErrorMessage "Cannot copy the Word template. Close it if it is " & _
        "locked and try again: " & templatePath, VBA.vbExclamation, _
        "Document Generation"
End Function

' Находит единственный Word-файл билета в указанной папке по номеру и ИПН.
Public Function private_Path_TryFindVacationTicketDocument( _
    ByVal outputFolderPathInput As String, _
    ByVal ticketNo As String, _
    ByVal ipnText As String, _
    ByRef outDocumentPath As String, _
    Optional ByVal isDocumentRequired As Boolean = True _
) As Boolean
    Dim outputFolderPath As String
    Dim ticketNoFileToken As String
    Dim fileName As String
    Dim matchCount As Long

    On Error GoTo EH
    outDocumentPath = VBA.vbNullString
    outputFolderPath = private_Path_ResolveFromWorkbook(outputFolderPathInput)
    If VBA.Len(VBA.Dir$(outputFolderPath, VBA.vbDirectory)) = 0 Then
        ex_ShowErrorMessage "Results folder was not found: " & outputFolderPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    ticketNoFileToken = private_Path_SanitizeFileName(ticketNo)
    fileName = VBA.Dir$(outputFolderPath & Application.PathSeparator & "*.docx")
    Do While VBA.Len(fileName) > 0
        If VBA.InStr(1, fileName, ticketNoFileToken & " ", _
                VBA.vbTextCompare) > 0 And _
           VBA.InStr(1, fileName, "(" & ipnText & ")", _
                VBA.vbTextCompare) > 0 And _
           Not private_Path_IsArchivedDocumentFileName(fileName) Then
            matchCount = matchCount + 1
            outDocumentPath = outputFolderPath & Application.PathSeparator & fileName
        End If
        fileName = VBA.Dir$()
    Loop
    If matchCount = 1 Then
        private_Path_TryFindVacationTicketDocument = True
        Exit Function
    End If
    If matchCount = 0 Then
        If Not isDocumentRequired Then
            private_Path_TryFindVacationTicketDocument = True
            Exit Function
        End If
        ex_ShowErrorMessage "Vacation ticket file was not found for ticket '" & _
            ticketNo & "' and IPN '" & ipnText & "' in folder: " & _
            outputFolderPath, VBA.vbExclamation, "Document Generation"
    Else
        ex_ShowErrorMessage "Multiple vacation ticket files were found for ticket '" & _
            ticketNo & "' and IPN '" & ipnText & "' in folder: " & _
            outputFolderPath, VBA.vbExclamation, "Document Generation"
    End If
    outDocumentPath = VBA.vbNullString
    Exit Function
EH:
    LogError "Failed to find vacation ticket file | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    ex_ShowErrorMessage "Failed to find vacation ticket file: " & _
        Err.Description, VBA.vbExclamation, "Document Generation"
End Function

' Исключает архивные копии, которым присваивается суффикс «(old N)».
Private Function private_Path_IsArchivedDocumentFileName( _
    ByVal fileName As String _
) As Boolean
    private_Path_IsArchivedDocumentFileName = ( _
        VBA.InStr(1, fileName, " (old ", VBA.vbTextCompare) > 0)
End Function

' Возвращает свободный временный путь в папке целевого Word-файла.
Private Function private_Path_BuildTemporaryDocumentPath( _
    ByVal documentPath As String _
) As String
    Dim dotPosition As Long
    Dim basePath As String
    Dim extensionText As String
    Dim candidatePath As String
    Dim copyIndex As Long

    dotPosition = VBA.InStrRev(documentPath, ".")
    If dotPosition = 0 Then
        LogError "Target vacation ticket file has no extension: " & documentPath
        ex_ShowErrorMessage "Target vacation ticket file has no extension: " & _
            documentPath, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    basePath = VBA.Left$(documentPath, dotPosition - 1)
    extensionText = VBA.Mid$(documentPath, dotPosition)
    candidatePath = basePath & ".updating" & extensionText
    copyIndex = 2
    Do While VBA.Len(VBA.Dir$(candidatePath)) > 0
        candidatePath = basePath & ".updating (" & VBA.CStr(copyIndex) & ")" & _
            extensionText
        copyIndex = copyIndex + 1
    Loop
    private_Path_BuildTemporaryDocumentPath = candidatePath
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
    ex_ShowErrorMessage "Failed to create output folder: " & folderPath & _
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

' Возвращает украинскую форму счётного слова для 1, 2-4 либо остальных чисел.
Public Function ex_GetUkrainianCountForm( _
    ByVal countValue As Long, _
    ByVal oneForm As String, _
    ByVal fewForm As String, _
    ByVal manyForm As String _
) As String
    Dim normalizedCount As Long

    If VBA.Len(oneForm) = 0 Or VBA.Len(fewForm) = 0 Or _
       VBA.Len(manyForm) = 0 Then
        Err.Raise VBA.vbObjectError + 4101, "ex_GetUkrainianCountForm", _
            "All Ukrainian count forms are required."
    End If
    normalizedCount = VBA.Abs(countValue)
    Select Case normalizedCount Mod 10
        Case 1
            If normalizedCount Mod 100 <> 11 Then
                ex_GetUkrainianCountForm = oneForm
            Else
                ex_GetUkrainianCountForm = manyForm
            End If
        Case 2 To 4
            If normalizedCount Mod 100 < 12 Or normalizedCount Mod 100 > 14 Then
                ex_GetUkrainianCountForm = fewForm
            Else
                ex_GetUkrainianCountForm = manyForm
            End If
        Case Else
            ex_GetUkrainianCountForm = manyForm
    End Select
End Function

' --------------------------------------
' } // namespace Text
' --------------------------------------

' --------------------------------------
' namespace Messaging {
' --------------------------------------
Public Function ex_TryConfigureMessageTarget( _
    ByVal worksheetName As String, _
    ByVal cellAddress As String _
) As Boolean
    Dim targetSheet As Worksheet

    On Error GoTo EH
    Set messageTargetRange = Nothing
    Set targetSheet = ThisWorkbook.Worksheets(worksheetName)
    Set messageTargetRange = targetSheet.Range(cellAddress).MergeArea
    ex_TryConfigureMessageTarget = True
    Exit Function
EH:
    Set messageTargetRange = Nothing
    LogError "Message target is unavailable | Sheet=" & worksheetName & _
        " | Cell=" & cellAddress & " | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    VBA.MsgBox "Message area '" & cellAddress & "' was not found on sheet '" & _
        worksheetName & "'.", VBA.vbExclamation, "Document Generation"
End Function

Public Sub ex_ClearMessageTarget()
    Set messageTargetRange = Nothing
End Sub

Public Sub ex_ShowStatusMessage( _
    ByVal messageText As String, _
    Optional ByVal isErrorMessage As Boolean = False _
)
    On Error GoTo EH
    If messageTargetRange Is Nothing Then
        LogError "Message target is not configured | Message=" & messageText
        Exit Sub
    End If
    messageTargetRange.Cells(1, 1).Value = private_Text_Normalize(messageText)
    If isErrorMessage Then
        messageTargetRange.Font.Color = VBA.RGB(255, 255, 0)
    Else
        messageTargetRange.Font.Color = VBA.RGB(116, 116, 116)
    End If
    Exit Sub
EH:
    LogError "Failed to write message target | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
End Sub

' Сохраняет совместимость с существующими модулями без использования StatusBar.
Public Sub ex_ShowStatusBarMessage(ByVal messageText As String)
    ex_ShowStatusMessage messageText
End Sub

' Выводит ошибку в область сообщений и диалоговое окно.
Public Sub ex_ShowErrorMessage( _
    ByVal messageText As String, _
    Optional ByVal buttons As VbMsgBoxStyle = VBA.vbExclamation, _
    Optional ByVal titleText As String = "Document Generation" _
)
    ex_ShowStatusMessage messageText, True
    VBA.MsgBox messageText, buttons, titleText
End Sub

Public Sub ClearLog()
#If ENABLE_LOGGING Then
#If CLEAR_LOG_ON_GENERATION Then
    Dim fileNumber As Integer

    fileNumber = VBA.FreeFile
    Open GetLogFilePath() For Output As #fileNumber
    Close #fileNumber
#End If
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