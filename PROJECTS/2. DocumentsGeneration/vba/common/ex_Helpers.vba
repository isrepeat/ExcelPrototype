Option Explicit

#Const ENABLE_LOGGING = True
#Const ENABLE_DEBUG_LOGGING = True
#Const CLEAR_LOG_ON_GENERATION = False

Private managedWordApp As Object
Private messageTargetRange As Range
Private configuredLogFileSuffix As String
Private logWriteFailureNotified As Boolean
Private logSessionStarted As Boolean

Private Const LOG_FOLDER_NAME As String = "2. DocumentsGeneration"
Private Const HELPERS_BUILD_ID As String = "2026-09-21.unicode-diagnostics.1"
' The log name is ASCII on purpose. VBA I/O on a PC with another ANSI locale
' may not read a workbook name with Cyrillic text correctly.
Private Const LOG_FILE_BASE_NAME As String = "documents_generation"

' --------------------------------------
' namespace Unicode {
' --------------------------------------
' Builds Unicode text from ASCII code points after VBA module import.
Public Function fn_FromCodePoints(ByVal codePointList As String) As String
    Dim codePointParts As Variant
    Dim codePointPart As Variant
    codePointParts = VBA.Split(codePointList, ",")
    For Each codePointPart In codePointParts
        fn_FromCodePoints = fn_FromCodePoints & VBA.ChrW$(VBA.CLng(codePointPart))
    Next codePointPart
End Function
' --------------------------------------
' } // namespace Unicode
' --------------------------------------

' --------------------------------------
' namespace Word {
' --------------------------------------
' Common formatting, path, Word, and log operations.
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
    Dim fileSystem As Object
    Dim placeholderIndex As Long

    On Error GoTo EH
    outDocumentPath = VBA.vbNullString
    templatePath = private_Path_ResolveFromWorkbook(templatePathInput)
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(templatePath) Then
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
        If Not fileSystem.FileExists(overwriteDocumentPath) Then
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
               fileSystem.FileExists(finalDocumentPath) Then
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
    fileSystem.CopyFile templatePath, temporaryDocumentPath, False
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
        fileSystem.DeleteFile outDocumentPath, True
        fileSystem.MoveFile temporaryDocumentPath, outDocumentPath
    End If
    If VBA.Len(finalDocumentPath) > 0 And _
       VBA.StrComp(finalDocumentPath, outDocumentPath, VBA.vbTextCompare) <> 0 Then
        fileSystem.MoveFile outDocumentPath, finalDocumentPath
        outDocumentPath = finalDocumentPath
    End If
    Set wordApp = Nothing
    private_Word_TryGenerateDocument = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If VBA.Len(temporaryDocumentPath) > 0 Then
        If Not fileSystem Is Nothing Then
            If fileSystem.FileExists(temporaryDocumentPath) Then _
                fileSystem.DeleteFile temporaryDocumentPath, True
        End If
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
    ' Replace long tokens first so short tokens do not change their parts.
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
    ' In a format string from a cell or external config, \" means a quote.
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
    ' DateSerial normalizes 31.02, so check parts after parsing.
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
        Case 1: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1089,1110,1095,1085,1103")
        Case 2: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1083,1102,1090,1086,1075,1086")
        Case 3: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1073,1077,1088,1077,1079,1085,1103")
        Case 4: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1082,1074,1110,1090,1085,1103")
        Case 5: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1090,1088,1072,1074,1085,1103")
        Case 6: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1095,1077,1088,1074,1085,1103")
        Case 7: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1083,1080,1087,1085,1103")
        Case 8: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1089,1077,1088,1087,1085,1103")
        Case 9: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1074,1077,1088,1077,1089,1085,1103")
        Case 10: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1078,1086,1074,1090,1085,1103")
        Case 11: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1083,1080,1089,1090,1086,1087,1072,1076,1072")
        Case 12: private_Date_GetUaMonthGenitive = fn_FromCodePoints("1075,1088,1091,1076,1085,1103")
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

' Formats text with named tokens: "{FIO} - {OrderNo}".
' Use "{{" and "}}" for literal braces.
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
    If VBA.Left$(pathText, 2) = ".\" Then pathText = VBA.Mid$(pathText, 3)
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
    Dim fileSystem As Object

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

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    candidatePath = folderPath & fileNameBase & extensionText
    If Not allowCopySuffix And fileSystem.FileExists(candidatePath) Then
        LogError "A generated document already exists: " & candidatePath
        ex_ShowErrorMessage "A document already exists for this vacation ticket: " & _
            candidatePath, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    copyIndex = 2
    Do While fileSystem.FileExists(candidatePath)
        candidatePath = folderPath & fileNameBase & " (" & _
            VBA.CStr(copyIndex) & ")" & extensionText
        copyIndex = copyIndex + 1
    Loop
    private_Path_BuildGeneratedDocumentPath = candidatePath
End Function

' Builds the exact path for a new document without an automatic suffix.
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

' Moves the old document to a free name with an "(old N)" suffix.
Public Function private_Path_TryArchiveDocument( _
    ByVal documentPath As String _
) As Boolean
    Dim dotPosition As Long
    Dim basePath As String
    Dim extensionText As String
    Dim archivedPath As String
    Dim archiveIndex As Long
    Dim fileSystem As Object

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(documentPath) Then
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
    Loop While fileSystem.FileExists(archivedPath)
    fileSystem.MoveFile documentPath, archivedPath
    private_Path_TryArchiveDocument = True
    Exit Function
EH:
    LogError "Failed to archive document | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description
    ex_ShowErrorMessage "Failed to archive document: " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

' Checks that an existing Word file is available for replacement.
Public Function private_Path_TryEnsureDocumentWritable( _
    ByVal documentPath As String _
) As Boolean
    Dim fileSystem As Object
    Dim targetFile As Object

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(documentPath) Then
        ex_ShowErrorMessage "Document to overwrite was not found: " & documentPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set targetFile = fileSystem.GetFile(documentPath)
    If targetFile.Size < 0 Then Err.Raise VBA.vbObjectError + 4107, _
        "private_Path_TryEnsureDocumentWritable", "Invalid document file size."
    private_Path_TryEnsureDocumentWritable = True
    Exit Function
EH:
    LogError "Document is not writable | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description & " | Path=" & documentPath
    ex_ShowErrorMessage "The Word file is open or cannot be replaced. Close it " & _
        "and try again: " & documentPath, VBA.vbExclamation, _
        "Document Generation"
End Function

' Checks that the Word template can be read. It may be open in Word.
Public Function private_Path_TryEnsureDocumentReadable( _
    ByVal documentPathInput As String _
) As Boolean
    Dim documentPath As String
    Dim fileSystem As Object
    Dim sourceFile As Object

    On Error GoTo EH
    documentPath = private_Path_ResolveFromWorkbook(documentPathInput)
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(documentPath) Then
        ex_ShowErrorMessage "Word template was not found: " & documentPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set sourceFile = fileSystem.GetFile(documentPath)
    If sourceFile.Size < 0 Then Err.Raise VBA.vbObjectError + 4106, _
        "private_Path_TryEnsureDocumentReadable", "Invalid template file size."
    private_Path_TryEnsureDocumentReadable = True
    Exit Function
EH:
    LogError "Document is not readable | Number=" & VBA.CStr(Err.Number) & _
        " | Description=" & Err.Description & " | Path=" & documentPath
    ex_ShowErrorMessage "The Word template cannot be read. Close it if it is " & _
        "locked and try again: " & documentPath, VBA.vbExclamation, _
        "Document Generation"
End Function

' Tests the FileCopy operation used by the Word generator.
Public Function private_Path_TryProbeTemplateCopy( _
    ByVal templatePathInput As String, _
    ByVal outputFolderPathInput As String _
) As Boolean
    Dim templatePath As String
    Dim outputFolderPath As String
    Dim probePath As String
    Dim fileSystem As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    templatePath = private_Path_ResolveFromWorkbook(templatePathInput)
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(templatePath) Then
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
    probePath = fileSystem.BuildPath(outputFolderPath, fileSystem.GetTempName)
    fileSystem.CopyFile templatePath, probePath, False
    fileSystem.DeleteFile probePath, True
    private_Path_TryProbeTemplateCopy = True
    Exit Function
EH:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If VBA.Len(probePath) > 0 And Not fileSystem Is Nothing Then
        If fileSystem.FileExists(probePath) Then fileSystem.DeleteFile probePath, True
    End If
    On Error GoTo 0
    LogError "Template copy preflight failed | Number=" & VBA.CStr(errorNumber) & _
        " | Description=" & errorDescription & " | Template=" & templatePath
    ex_ShowErrorMessage "Cannot copy the Word template. Close it if it is " & _
        "locked and try again: " & templatePath, VBA.vbExclamation, _
        "Document Generation"
End Function

' Finds one ticket Word file in the folder by ticket number and IPN.
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
    Dim fileSystem As Object
    Dim outputFolder As Object
    Dim folderFile As Object

    On Error GoTo EH
    outDocumentPath = VBA.vbNullString
    outputFolderPath = private_Path_ResolveFromWorkbook(outputFolderPathInput)
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(outputFolderPath) Then
        ex_ShowErrorMessage "Results folder was not found: " & outputFolderPath, _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    ticketNoFileToken = private_Path_SanitizeFileName(ticketNo)
    Set outputFolder = fileSystem.GetFolder(outputFolderPath)
    For Each folderFile In outputFolder.Files
        fileName = folderFile.Name
        If VBA.StrComp(fileSystem.GetExtensionName(fileName), "docx", _
                VBA.vbTextCompare) = 0 And _
           VBA.InStr(1, fileName, ticketNoFileToken & " ", _
                VBA.vbTextCompare) > 0 And _
           VBA.InStr(1, fileName, "(" & ipnText & ")", _
                VBA.vbTextCompare) > 0 And _
           Not private_Path_IsArchivedDocumentFileName(fileName) Then
            matchCount = matchCount + 1
            outDocumentPath = folderFile.Path
        End If
    Next folderFile
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

' Ignores archived copies with an "(old N)" suffix.
Private Function private_Path_IsArchivedDocumentFileName( _
    ByVal fileName As String _
) As Boolean
    private_Path_IsArchivedDocumentFileName = ( _
        VBA.InStr(1, fileName, " (old ", VBA.vbTextCompare) > 0)
End Function

' Returns a free temporary path in the target Word-file folder.
Private Function private_Path_BuildTemporaryDocumentPath( _
    ByVal documentPath As String _
) As String
    Dim dotPosition As Long
    Dim basePath As String
    Dim extensionText As String
    Dim candidatePath As String
    Dim copyIndex As Long
    Dim fileSystem As Object

    dotPosition = VBA.InStrRev(documentPath, ".")
    If dotPosition = 0 Then
        LogError "Target vacation ticket file has no extension: " & documentPath
        ex_ShowErrorMessage "Target vacation ticket file has no extension: " & _
            documentPath, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    basePath = VBA.Left$(documentPath, dotPosition - 1)
    extensionText = VBA.Mid$(documentPath, dotPosition)
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    candidatePath = basePath & ".updating" & extensionText
    copyIndex = 2
    Do While fileSystem.FileExists(candidatePath)
        candidatePath = basePath & ".updating (" & VBA.CStr(copyIndex) & ")" & _
            extensionText
        copyIndex = copyIndex + 1
    Loop
    private_Path_BuildTemporaryDocumentPath = candidatePath
End Function

' Creates the result folder and missing parent folders.
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

' Returns the Ukrainian count-word form for 1, 2-4, or other numbers.
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
    ex_ShowMessage "Message area '" & cellAddress & "' was not found on sheet '" & _
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

' Keeps compatibility with existing modules without using StatusBar.
Public Sub ex_ShowStatusBarMessage(ByVal messageText As String)
    ex_ShowStatusMessage messageText
End Sub

' --------------------------------------
' namespace Dialogs {
' --------------------------------------
' Shows a Windows dialog that supports Unicode text.
Public Function ex_ShowMessageBox( _
    ByVal messageText As String, _
    Optional ByVal buttons As VbMsgBoxStyle = VBA.vbExclamation, _
    Optional ByVal titleText As String = "Document Generation" _
) As VbMsgBoxResult
    Dim shell As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Fallback
    Set shell = VBA.CreateObject("WScript.Shell")
    LogDebug "DIALOG: Backend=WScript.Shell.Popup"
    ex_ShowMessageBox = shell.Popup(messageText, 0, titleText, VBA.CLng(buttons))
    Exit Function
Fallback:
    errorNumber = Err.Number
    errorDescription = Err.Description
    LogError "DIALOG: Backend=VBA.MsgBox fallback | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    ex_ShowMessageBox = VBA.MsgBox(messageText, buttons, titleText)
End Function

' Shows a message when the caller does not need the selected button.
Public Sub ex_ShowMessage( _
    ByVal messageText As String, _
    Optional ByVal buttons As VbMsgBoxStyle = VBA.vbExclamation, _
    Optional ByVal titleText As String = "Document Generation" _
)
    Call ex_ShowMessageBox(messageText, buttons, titleText)
End Sub
' --------------------------------------
' } // namespace Dialogs
' --------------------------------------

' Shows an error in the message area and a dialog box.
Public Sub ex_ShowErrorMessage( _
    ByVal messageText As String, _
    Optional ByVal buttons As VbMsgBoxStyle = VBA.vbExclamation, _
    Optional ByVal titleText As String = "Document Generation" _
)
    ex_ShowStatusMessage messageText, True
    Call ex_ShowMessageBox(messageText, buttons, titleText)
End Sub

Public Function ex_TryConfigureLogFileSuffix( _
    ByVal logFileSuffix As String _
) As Boolean
    logFileSuffix = private_Text_Normalize(logFileSuffix)
    If VBA.Len(logFileSuffix) = 0 Then
        ex_ShowMessage "Log file suffix is not configured.", VBA.vbExclamation, _
            "Document Generation"
        Exit Function
    End If
    configuredLogFileSuffix = logFileSuffix
    logWriteFailureNotified = False
    private_Log_WriteSessionHeader
    private_Log_WriteSystemInfo
    ex_TryConfigureLogFileSuffix = True
End Function

Public Sub ClearLog()
#If ENABLE_LOGGING Then
#If CLEAR_LOG_ON_GENERATION Then
    Dim fileSystem As Object
    Dim logStream As Object

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    ' TristateTrue creates a Unicode log without depending on the system ANSI code page.
    Set logStream = fileSystem.OpenTextFile(GetLogFilePath(), 2, True, -1)
    logStream.Close
    logSessionStarted = False
    private_Log_WriteSessionHeader
    private_Log_WriteSystemInfo
    Exit Sub
EH:
    On Error Resume Next
    If Not logStream Is Nothing Then logStream.Close
    private_Log_NotifyWriteFailure Err.Number, Err.Description
    On Error GoTo 0
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
    private_Log_WriteLine VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & messageText
#End If
End Sub

' Writes a log line without adding a timestamp or message type.
Private Sub private_Log_WriteRaw(ByVal messageText As String)
#If ENABLE_LOGGING Then
    private_Log_WriteLine messageText
#End If
End Sub

' Writes one prepared line to the configured Unicode log file.
Private Sub private_Log_WriteLine(ByVal lineText As String)
#If ENABLE_LOGGING Then
    Dim fileSystem As Object
    Dim logStream As Object

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    ' TristateTrue keeps Cyrillic text without depending on ACP.
    Set logStream = fileSystem.OpenTextFile(GetLogFilePath(), 8, True, -1)
    logStream.WriteLine lineText
    logStream.Close
    Exit Sub
EH:
    On Error Resume Next
    If Not logStream Is Nothing Then logStream.Close
    private_Log_NotifyWriteFailure Err.Number, Err.Description
    On Error GoTo 0
#End If
End Sub

Public Function GetLogFilePath() As String
    Dim logFolderPath As String

    If VBA.Len(configuredLogFileSuffix) = 0 Then
        Err.Raise VBA.vbObjectError + 4102, "GetLogFilePath", _
            "Log file suffix was not configured by the calling module."
    End If
    If Not private_Log_TryEnsureFolder(logFolderPath) Then
        Err.Raise VBA.vbObjectError + 4103, "GetLogFilePath", _
            "The local log folder is unavailable."
    End If
    GetLogFilePath = logFolderPath & Application.PathSeparator & _
        LOG_FILE_BASE_NAME & configuredLogFileSuffix
End Function

' Creates a local log folder that does not depend on the workbook path.
Private Function private_Log_TryEnsureFolder( _
    ByRef outFolderPath As String _
) As Boolean
    Dim tempFolderPath As String

    On Error GoTo EH
    outFolderPath = VBA.vbNullString
    tempFolderPath = VBA.Trim$(VBA.Environ$("TEMP"))
    If VBA.Len(tempFolderPath) = 0 Then
        Err.Raise VBA.vbObjectError + 4104, "private_Log_TryEnsureFolder", _
            "The TEMP environment variable is empty."
    End If
    If VBA.Len(VBA.Dir$(tempFolderPath, VBA.vbDirectory)) = 0 Then
        Err.Raise VBA.vbObjectError + 4105, "private_Log_TryEnsureFolder", _
            "The TEMP folder was not found: " & tempFolderPath
    End If
    outFolderPath = tempFolderPath & Application.PathSeparator & LOG_FOLDER_NAME
    If VBA.Len(VBA.Dir$(outFolderPath, VBA.vbDirectory)) = 0 Then
        VBA.MkDir outFolderPath
    End If
    private_Log_TryEnsureFolder = True
    Exit Function
EH:
    outFolderPath = VBA.vbNullString
End Function

' Writes an easy-to-find boundary between VBA sessions.
Private Sub private_Log_WriteSessionHeader()
    If logSessionStarted Then Exit Sub
    private_Log_WriteRaw "=============================================================================================================================="
    private_Log_WriteRaw "                                         New session started | " & _
        VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    private_Log_WriteRaw "=============================================================================================================================="
    logSessionStarted = True
End Sub

' Writes environment data that can affect Unicode behavior in VBA.
Private Sub private_Log_WriteSystemInfo()
    Dim officeUiLanguageId As Long
    Dim scriptShell As Object
    Dim scriptShellAvailable As Boolean

    On Error Resume Next
    officeUiLanguageId = Application.LanguageSettings.LanguageID(2)
    On Error GoTo 0
    scriptShellAvailable = private_Log_TryCreateWScriptShell(scriptShell)
    WriteLog "SYSTEM: Environment"
    WriteLog "SYSTEM: HelpersBuild=" & HELPERS_BUILD_ID
    WriteLog "SYSTEM: Windows=" & Application.OperatingSystem
    WriteLog "SYSTEM: OfficeVersion=" & Application.Version
    WriteLog "SYSTEM: OfficeBuild=" & VBA.CStr(Application.Build)
    WriteLog "SYSTEM: OfficeUILanguageId=" & VBA.CStr(officeUiLanguageId)
    WriteLog "SYSTEM: ACP=" & private_Log_ReadRegistryValue( _
        "HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage\ACP")
    WriteLog "SYSTEM: OEMCP=" & private_Log_ReadRegistryValue( _
        "HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage\OEMCP")
    WriteLog "SYSTEM: UserLocale=" & private_Log_ReadRegistryValue( _
        "HKCU\Control Panel\International\LocaleName")
    WriteLog "SYSTEM: SystemLocale=" & private_Log_ReadRegistryValue( _
        "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\Default")
    WriteLog "SYSTEM: ProcessArchitecture=" & _
        VBA.Environ$("PROCESSOR_ARCHITECTURE")
    WriteLog "SYSTEM: ProcessArchitectureWow64=" & _
        VBA.Environ$("PROCESSOR_ARCHITEW6432")
    WriteLog "SYSTEM: VBA7=" & private_Log_GetVba7State()
    WriteLog "SYSTEM: OfficeBitness=" & private_Log_GetOfficeBitness()
    WriteLog "SYSTEM: WScriptShellAvailable=" & _
        VBA.CStr(scriptShellAvailable)
    WriteLog "SYSTEM: WScriptEnabledHKCU=" & private_Log_ReadRegistryValue( _
        "HKCU\Software\Microsoft\Windows Script Host\Settings\Enabled")
    WriteLog "SYSTEM: WScriptEnabledHKLM=" & private_Log_ReadRegistryValue( _
        "HKLM\Software\Microsoft\Windows Script Host\Settings\Enabled")
    WriteLog "SYSTEM: WshomSystem32Version=" & private_Log_ReadFileVersion( _
        VBA.Environ$("SystemRoot") & "\System32\wshom.ocx")
    WriteLog "SYSTEM: WshomSysWow64Version=" & private_Log_ReadFileVersion( _
        VBA.Environ$("SystemRoot") & "\SysWOW64\wshom.ocx")
End Sub

Private Function private_Log_TryCreateWScriptShell( _
    ByRef outShell As Object _
) As Boolean
    On Error GoTo EH
    Set outShell = VBA.CreateObject("WScript.Shell")
    private_Log_TryCreateWScriptShell = Not outShell Is Nothing
    Exit Function
EH:
    Set outShell = Nothing
End Function

Private Function private_Log_GetVba7State() As String
#If VBA7 Then
    private_Log_GetVba7State = "True"
#Else
    private_Log_GetVba7State = "False"
#End If
End Function

Private Function private_Log_GetOfficeBitness() As String
#If Win64 Then
    private_Log_GetOfficeBitness = "64-bit"
#Else
    private_Log_GetOfficeBitness = "32-bit"
#End If
End Function

' Reads a registry value without affecting the main logging flow.
Private Function private_Log_ReadRegistryValue( _
    ByVal registryPath As String _
) As String
    Dim shell As Object

    On Error GoTo EH
    Set shell = VBA.CreateObject("WScript.Shell")
    private_Log_ReadRegistryValue = VBA.CStr(shell.RegRead(registryPath))
    Exit Function
EH:
    private_Log_ReadRegistryValue = "<unavailable>"
End Function

Private Function private_Log_ReadFileVersion( _
    ByVal filePath As String _
) As String
    Dim fileSystem As Object

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(filePath) Then
        private_Log_ReadFileVersion = "<not-found>"
        Exit Function
    End If
    private_Log_ReadFileVersion = fileSystem.GetFileVersion(filePath)
    Exit Function
EH:
    private_Log_ReadFileVersion = "<unavailable>"
End Function

' Logging must not hide the main operation error.
Private Sub private_Log_NotifyWriteFailure( _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String _
)
    If logWriteFailureNotified Then Exit Sub

    logWriteFailureNotified = True
    ex_ShowMessage "Unable to write the log to '%TEMP%\" & LOG_FOLDER_NAME & _
        "'. Generation will continue without logging." & VBA.vbCrLf & _
        VBA.vbCrLf & "Error [" & VBA.CStr(errorNumber) & "]: " & _
        errorDescription, VBA.vbExclamation, "Document Generation"
End Sub
' --------------------------------------
' } // namespace Logging
' --------------------------------------