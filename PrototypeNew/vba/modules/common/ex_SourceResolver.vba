Attribute VB_Name = "ex_SourceResolver"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private Const ERR_BASE As Long = vbObjectError + 3700

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_SourceResolver.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_ResolveLatestByDmyPattern( _
    ByVal filePathPattern As String, _
    Optional ByVal resolverArgs As String = vbNullString _
) As String
    Dim normalizedPattern As String
    Dim absolutePattern As String
    Dim folderPath As String
    Dim filePattern As String
    Dim searchMask As String
    Dim candidateName As String
    Dim candidatePath As String
    Dim candidateDate As Date
    Dim candidateWriteTime As Date
    Dim bestPath As String
    Dim bestDate As Date
    Dim bestWriteTime As Date
    Dim hasBest As Boolean
    Dim validationError As String

    resolverArgs = VBA.Trim$(resolverArgs)
    If VBA.Len(resolverArgs) > 0 Then
        ' Аргументы зарезервированы на будущее.
    End If

    normalizedPattern = private_NormalizeFilePath(filePathPattern)
    If VBA.Len(normalizedPattern) = 0 Then
        Err.Raise ERR_BASE + 1, "ex_SourceResolver", "Resolver pattern is empty."
    End If

    absolutePattern = private_ToAbsolutePath(normalizedPattern)
    If VBA.Len(absolutePattern) = 0 Then
        Err.Raise ERR_BASE + 2, "ex_SourceResolver", "Resolver pattern could not be converted to an absolute path: " & normalizedPattern
    End If

    folderPath = private_GetParentDirectory(absolutePattern)
    filePattern = private_GetFileName(absolutePattern)
    If VBA.Len(folderPath) = 0 Or VBA.Len(filePattern) = 0 Then
        Err.Raise ERR_BASE + 3, "ex_SourceResolver", "Resolver pattern must include both folder and file name: " & absolutePattern
    End If

    If Not private_ValidateDmyPattern(filePattern, validationError) Then
        Err.Raise ERR_BASE + 4, "ex_SourceResolver", validationError & " Pattern: " & absolutePattern
    End If

    If VBA.Len(VBA.Dir$(folderPath, vbDirectory)) = 0 Then
        Err.Raise ERR_BASE + 5, "ex_SourceResolver", "Resolver folder was not found: " & folderPath
    End If

    If (VBA.GetAttr(folderPath) And vbDirectory) = 0 Then
        Err.Raise ERR_BASE + 6, "ex_SourceResolver", "Resolver path is not a folder: " & folderPath
    End If

    searchMask = private_BuildSearchMask(filePattern)
    candidateName = VBA.Dir$(folderPath & "\" & searchMask, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)

    Do While VBA.Len(candidateName) > 0
        If private_TryExtractDateByPattern(filePattern, candidateName, candidateDate) Then
            candidatePath = folderPath & "\" & candidateName
            candidateWriteTime = VBA.FileDateTime(candidatePath)

            If (Not hasBest) _
               Or (candidateDate > bestDate) _
               Or (candidateDate = bestDate And candidateWriteTime > bestWriteTime) Then
                bestPath = candidatePath
                bestDate = candidateDate
                bestWriteTime = candidateWriteTime
                hasBest = True
            End If
        End If

        candidateName = VBA.Dir$
    Loop

    If Not hasBest Then
        Err.Raise ERR_BASE + 7, "ex_SourceResolver", _
            "No files matched the date pattern. Pattern: " & absolutePattern & ", search mask: " & searchMask
    End If

    fn_ResolveLatestByDmyPattern = bestPath
End Function

Public Function fn_ResolveAllByDmyPattern( _
    ByVal filePathPattern As String, _
    Optional ByVal resolverArgs As String = vbNullString _
) As Collection
    Dim normalizedPattern As String
    Dim absolutePattern As String
    Dim folderPath As String
    Dim filePattern As String
    Dim searchMask As String
    Dim candidateName As String
    Dim candidatePath As String
    Dim candidateDate As Date
    Dim candidateWriteTime As Date
    Dim validationError As String
    Dim sortDescending As Boolean
    Dim paths() As String
    Dim dates() As Date
    Dim writeTimes() As Date
    Dim itemCount As Long
    Dim i As Long
    Dim result As Collection
    Dim dateFrom As Date
    Dim dateTo As Date
    Dim hasDateFrom As Boolean
    Dim hasDateTo As Boolean

    normalizedPattern = private_NormalizeFilePath(filePathPattern)
    If VBA.Len(normalizedPattern) = 0 Then
        Err.Raise ERR_BASE + 1, "ex_SourceResolver", "Resolver pattern is empty."
    End If

    absolutePattern = private_ToAbsolutePath(normalizedPattern)
    folderPath = private_GetParentDirectory(absolutePattern)
    filePattern = private_GetFileName(absolutePattern)
    If VBA.Len(folderPath) = 0 Or VBA.Len(filePattern) = 0 Then
        Err.Raise ERR_BASE + 3, "ex_SourceResolver", _
            "Resolver pattern must include both folder and file name: " & absolutePattern
    End If
    If Not private_ValidateDmyPattern(filePattern, validationError) Then
        Err.Raise ERR_BASE + 4, "ex_SourceResolver", validationError & " Pattern: " & absolutePattern
    End If
    If Not private_TryReadResolverDateArg( _
        resolverArgs, "dateFrom", hasDateFrom, dateFrom, _
        validationError) Then
        Err.Raise ERR_BASE + 10, "ex_SourceResolver", validationError
    End If
    If Not private_TryReadResolverDateArg( _
        resolverArgs, "dateTo", hasDateTo, dateTo, _
        validationError) Then
        Err.Raise ERR_BASE + 11, "ex_SourceResolver", validationError
    End If
    If hasDateFrom And hasDateTo Then
        If dateFrom > dateTo Then
            Err.Raise ERR_BASE + 12, "ex_SourceResolver", _
                "dateFrom must not be later than dateTo."
        End If
    End If
    If VBA.Len(VBA.Dir$(folderPath, vbDirectory)) = 0 Then
        Err.Raise ERR_BASE + 5, "ex_SourceResolver", "Resolver folder was not found: " & folderPath
    End If

    sortDescending = (VBA.InStr(1, resolverArgs, "order=desc", VBA.vbTextCompare) > 0)
    searchMask = private_BuildSearchMask(filePattern)
    candidateName = VBA.Dir$(folderPath & "\" & searchMask, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)
    Do While VBA.Len(candidateName) > 0
        If private_TryExtractDateByPattern(filePattern, candidateName, candidateDate) Then
            If (Not hasDateFrom Or candidateDate >= dateFrom) And _
                (Not hasDateTo Or candidateDate <= dateTo) Then
                itemCount = itemCount + 1
                ReDim Preserve paths(1 To itemCount)
                ReDim Preserve dates(1 To itemCount)
                ReDim Preserve writeTimes(1 To itemCount)
                candidatePath = folderPath & "\" & candidateName
                paths(itemCount) = candidatePath
                dates(itemCount) = candidateDate
                writeTimes(itemCount) = VBA.FileDateTime(candidatePath)
            End If
        End If
        candidateName = VBA.Dir$
    Loop
    If itemCount = 0 And _
        VBA.InStr(1, resolverArgs, "allowEmpty=true", _
            VBA.vbTextCompare) = 0 Then
        Err.Raise ERR_BASE + 7, "ex_SourceResolver", _
            "No files matched the date pattern. Pattern: " & absolutePattern & ", search mask: " & searchMask
    End If

    private_SortResolvedPaths paths, dates, writeTimes, itemCount, sortDescending
    Set result = New Collection
    If itemCount > 0 Then
        For i = 1 To itemCount
            result.Add paths(i)
        Next i
    End If
    Set fn_ResolveAllByDmyPattern = result
End Function

' Возвращает дату, зашитую в имени уже разрешённого файла. Метод нужен
' mode-specific контроллерам, которые объединяют результаты нескольких
' шаблонов и затем выполняют общую сортировку.
Public Function fn_TryGetDmyDateByResolvedPath( _
    ByVal filePathPattern As String, _
    ByVal resolvedPath As String, _
    ByRef outDate As Date _
) As Boolean
    Dim patternFileName As String
    Dim resolvedFileName As String

    outDate = 0
    patternFileName = private_GetFileName( _
        private_NormalizeFilePath(filePathPattern))
    resolvedFileName = private_GetFileName( _
        private_NormalizeFilePath(resolvedPath))
    If VBA.Len(patternFileName) = 0 Or _
        VBA.Len(resolvedFileName) = 0 Then Exit Function

    fn_TryGetDmyDateByResolvedPath = private_TryExtractDateByPattern( _
        patternFileName, resolvedFileName, outDate)
End Function

Private Function private_TryReadResolverDateArg( _
    ByVal resolverArgs As String, _
    ByVal argName As String, _
    ByRef outHasValue As Boolean, _
    ByRef outDate As Date, _
    ByRef outErrorText As String _
) As Boolean
    Dim tokens As Variant
    Dim token As Variant
    Dim separatorPosition As Long
    Dim keyText As String
    Dim valueText As String

    outHasValue = False
    outDate = 0
    outErrorText = VBA.vbNullString
    tokens = VBA.Split(resolverArgs, ";")
    For Each token In tokens
        separatorPosition = VBA.InStr(1, VBA.CStr(token), "=", _
            VBA.vbBinaryCompare)
        If separatorPosition > 0 Then
            keyText = VBA.Trim$(VBA.Left$(VBA.CStr(token), _
                separatorPosition - 1))
            If VBA.StrComp(keyText, argName, VBA.vbTextCompare) = 0 Then
                valueText = VBA.Trim$(VBA.Mid$(VBA.CStr(token), _
                    separatorPosition + 1))
                If Not private_TryParseResolverDate( _
                    valueText, outDate) Then
                    outErrorText = "Invalid " & argName & _
                        " value. Expected dd.mm.yyyy or yyyy-mm-dd: " & _
                        valueText
                    Exit Function
                End If
                outHasValue = True
                Exit For
            End If
        End If
    Next token
    private_TryReadResolverDateArg = True
End Function

Private Function private_TryParseResolverDate( _
    ByVal valueText As String, _
    ByRef outDate As Date _
) As Boolean
    Dim normalizedText As String
    Dim parts As Variant
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim parsedDate As Date

    outDate = 0
    normalizedText = VBA.Replace(VBA.Replace( _
        VBA.Trim$(valueText), "-", "."), "/", ".")
    parts = VBA.Split(normalizedText, ".")
    If UBound(parts) <> 2 Then Exit Function
    If Not VBA.IsNumeric(parts(0)) Or _
        Not VBA.IsNumeric(parts(1)) Or _
        Not VBA.IsNumeric(parts(2)) Then Exit Function
    If VBA.Len(VBA.CStr(parts(0))) = 4 Then
        yearValue = VBA.CLng(parts(0))
        monthValue = VBA.CLng(parts(1))
        dayValue = VBA.CLng(parts(2))
    Else
        dayValue = VBA.CLng(parts(0))
        monthValue = VBA.CLng(parts(1))
        yearValue = VBA.CLng(parts(2))
    End If
    On Error GoTo EH
    parsedDate = VBA.DateSerial(yearValue, monthValue, dayValue)
    If VBA.Year(parsedDate) <> yearValue Or _
        VBA.Month(parsedDate) <> monthValue Or _
        VBA.Day(parsedDate) <> dayValue Then Exit Function
    outDate = parsedDate
    private_TryParseResolverDate = True
    Exit Function
EH:
    outDate = 0
End Function

Public Function fn_ExpandDmyRuntimeAliasByResolvedPath( _
    ByVal runtimeAliasPattern As String, _
    ByVal filePathPattern As String, _
    ByVal resolvedPath As String _
) As String
    Dim expandedAlias As String
    Dim patternFileName As String
    Dim resolvedFileName As String
    Dim resolvedDate As Date

    expandedAlias = VBA.Trim$(runtimeAliasPattern)
    If VBA.Len(expandedAlias) = 0 Then
        Err.Raise ERR_BASE + 8, "ex_SourceResolver", "Runtime alias pattern is empty."
    End If

    patternFileName = private_GetFileName(private_NormalizeFilePath(filePathPattern))
    resolvedFileName = private_GetFileName(private_NormalizeFilePath(resolvedPath))
    If Not private_TryExtractDateByPattern(patternFileName, resolvedFileName, resolvedDate) Then
        Err.Raise ERR_BASE + 9, "ex_SourceResolver", _
            "Resolved file does not match runtime alias date pattern: " & resolvedPath
    End If

    expandedAlias = VBA.Replace(expandedAlias, "{dd}", VBA.Format$(resolvedDate, "dd"), 1, -1, VBA.vbTextCompare)
    expandedAlias = VBA.Replace(expandedAlias, "{mm}", VBA.Format$(resolvedDate, "mm"), 1, -1, VBA.vbTextCompare)
    expandedAlias = VBA.Replace(expandedAlias, "{yyyy}", VBA.Format$(resolvedDate, "yyyy"), 1, -1, VBA.vbTextCompare)
    fn_ExpandDmyRuntimeAliasByResolvedPath = expandedAlias
End Function

' //
' // Internal
' //
Private Function private_ValidateDmyPattern(ByVal filePattern As String, ByRef outErrorText As String) As Boolean
    Dim pos As Long
    Dim closePos As Long
    Dim token As String
    Dim hasDd As Boolean
    Dim hasMm As Boolean
    Dim hasYyyy As Boolean

    If VBA.Len(filePattern) = 0 Then
        outErrorText = "File pattern is empty."
        Exit Function
    End If

    pos = 1
    Do While pos <= VBA.Len(filePattern)
        If VBA.Mid$(filePattern, pos, 1) = "{" Then
            closePos = VBA.InStr(pos + 1, filePattern, "}", vbBinaryCompare)
            If closePos <= pos + 1 Then
                outErrorText = "File pattern contains an unclosed placeholder."
                Exit Function
            End If

            token = VBA.LCase$(VBA.Trim$(VBA.Mid$(filePattern, pos + 1, closePos - pos - 1)))
            Select Case token
                Case "dd"
                    hasDd = True
                Case "mm", "monthua"
                    hasMm = True
                Case "yyyy", "yy"
                    hasYyyy = True
                Case Else
                    outErrorText = "Unsupported placeholder '{" & token & _
                        "}'. Supported date placeholders: {dd}, {mm}, " & _
                        "{monthUa}, {yy}, {yyyy}."
                    Exit Function
            End Select

            pos = closePos + 1
        Else
            pos = pos + 1
        End If
    Loop

    If Not hasDd Or Not hasMm Or Not hasYyyy Then
        outErrorText = "Date placeholders are required: {dd}, " & _
            "{mm} or {monthUa}, and {yy} or {yyyy}."
        Exit Function
    End If

    private_ValidateDmyPattern = True
End Function


Private Function private_TryExtractDateByPattern(ByVal filePattern As String, ByVal fileName As String, ByRef outDate As Date) As Boolean
    Dim dd As Long
    Dim mm As Long
    Dim yyyy As Long

    If Not private_TryMatchDmyPattern( _
        filePattern, fileName, 1, 1, dd, mm, yyyy) Then Exit Function
    If Not private_TryBuildExactDate(yyyy, mm, dd, outDate) Then Exit Function

    private_TryExtractDateByPattern = True
End Function

Private Function private_TryMatchDmyPattern( _
    ByVal filePattern As String, _
    ByVal fileName As String, _
    ByVal patternPos As Long, _
    ByVal filePos As Long, _
    ByRef outDd As Long, _
    ByRef outMm As Long, _
    ByRef outYyyy As Long _
) As Boolean
    Dim closePos As Long
    Dim token As String
    Dim nextFilePos As Long
    Dim parsedValue As Long
    Dim branchDd As Long
    Dim branchMm As Long
    Dim branchYyyy As Long
    Dim patternChar As String

    If patternPos > VBA.Len(filePattern) Then
        private_TryMatchDmyPattern = (filePos > VBA.Len(fileName))
        Exit Function
    End If

    patternChar = VBA.Mid$(filePattern, patternPos, 1)
    If patternChar = "*" Then
        For nextFilePos = filePos To VBA.Len(fileName) + 1
            branchDd = outDd
            branchMm = outMm
            branchYyyy = outYyyy
            If private_TryMatchDmyPattern(filePattern, fileName, _
                patternPos + 1, nextFilePos, branchDd, branchMm, _
                branchYyyy) Then
                outDd = branchDd
                outMm = branchMm
                outYyyy = branchYyyy
                private_TryMatchDmyPattern = True
                Exit Function
            End If
        Next nextFilePos
        Exit Function
    End If

    If patternChar = "?" Then
        If filePos > VBA.Len(fileName) Then Exit Function
        private_TryMatchDmyPattern = private_TryMatchDmyPattern( _
            filePattern, fileName, patternPos + 1, filePos + 1, _
            outDd, outMm, outYyyy)
        Exit Function
    End If

    If patternChar = "{" Then
        closePos = VBA.InStr(patternPos + 1, filePattern, "}", _
            VBA.vbBinaryCompare)
        If closePos <= patternPos + 1 Then Exit Function
        token = VBA.LCase$(VBA.Trim$(VBA.Mid$(filePattern, _
            patternPos + 1, closePos - patternPos - 1)))
        nextFilePos = filePos
        Select Case token
            Case "dd"
                If Not private_TryReadFixedDigits( _
                    fileName, nextFilePos, 2, parsedValue) Then Exit Function
                outDd = parsedValue
            Case "mm"
                If Not private_TryReadFixedDigits( _
                    fileName, nextFilePos, 2, parsedValue) Then Exit Function
                outMm = parsedValue
            Case "monthua"
                If Not private_TryReadUaMonth( _
                    fileName, nextFilePos, parsedValue) Then Exit Function
                outMm = parsedValue
            Case "yy"
                If Not private_TryReadFixedDigits( _
                    fileName, nextFilePos, 2, parsedValue) Then Exit Function
                outYyyy = 2000 + parsedValue
            Case "yyyy"
                If Not private_TryReadFixedDigits( _
                    fileName, nextFilePos, 4, parsedValue) Then Exit Function
                outYyyy = parsedValue
            Case Else
                Exit Function
        End Select
        private_TryMatchDmyPattern = private_TryMatchDmyPattern( _
            filePattern, fileName, closePos + 1, nextFilePos, _
            outDd, outMm, outYyyy)
        Exit Function
    End If

    If filePos > VBA.Len(fileName) Then Exit Function
    If VBA.StrComp(patternChar, VBA.Mid$(fileName, filePos, 1), _
        VBA.vbTextCompare) <> 0 Then Exit Function
    private_TryMatchDmyPattern = private_TryMatchDmyPattern( _
        filePattern, fileName, patternPos + 1, filePos + 1, _
        outDd, outMm, outYyyy)
End Function

Private Function private_TryReadUaMonth( _
    ByVal textValue As String, _
    ByRef ioPos As Long, _
    ByRef outMonth As Long _
) As Boolean
    Dim monthNames As Variant
    Dim monthIndex As Long
    Dim monthName As String

    monthNames = VBA.Array( _
        "січня", "лютого", "березня", "квітня", _
        "травня", "червня", "липня", "серпня", _
        "вересня", "жовтня", "листопада", "грудня")
    For monthIndex = LBound(monthNames) To UBound(monthNames)
        monthName = VBA.CStr(monthNames(monthIndex))
        If VBA.StrComp(VBA.Mid$(textValue, ioPos, VBA.Len(monthName)), _
            monthName, VBA.vbTextCompare) = 0 Then
            outMonth = monthIndex + 1
            ioPos = ioPos + VBA.Len(monthName)
            private_TryReadUaMonth = True
            Exit Function
        End If
    Next monthIndex
End Function


Private Function private_TryReadFixedDigits(ByVal textValue As String, ByRef ioPos As Long, ByVal digitsCount As Long, ByRef outValue As Long) As Boolean
    Dim chunk As String
    Dim i As Long
    Dim ch As String

    If digitsCount <= 0 Then Exit Function
    If ioPos < 1 Then Exit Function
    If ioPos + digitsCount - 1 > VBA.Len(textValue) Then Exit Function

    chunk = VBA.Mid$(textValue, ioPos, digitsCount)
    For i = 1 To VBA.Len(chunk)
        ch = VBA.Mid$(chunk, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    On Error GoTo ParseEH
    outValue = VBA.CLng(chunk)
    ioPos = ioPos + digitsCount
    private_TryReadFixedDigits = True
    Exit Function

ParseEH:
    private_TryReadFixedDigits = False
End Function


Private Function private_TryBuildExactDate(ByVal yyyy As Long, ByVal mm As Long, ByVal dd As Long, ByRef outDate As Date) As Boolean
    Dim parsedDate As Date

    If yyyy < 1900 Or yyyy > 9999 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    On Error GoTo ParseEH
    parsedDate = VBA.DateSerial(yyyy, mm, dd)
    On Error GoTo 0

    If VBA.Year(parsedDate) <> yyyy Then Exit Function
    If VBA.Month(parsedDate) <> mm Then Exit Function
    If VBA.Day(parsedDate) <> dd Then Exit Function

    outDate = parsedDate
    private_TryBuildExactDate = True
    Exit Function

ParseEH:
    private_TryBuildExactDate = False
End Function


Private Function private_BuildSearchMask(ByVal filePattern As String) As String
    Dim pos As Long
    Dim closePos As Long

    pos = 1
    Do While pos <= VBA.Len(filePattern)
        If VBA.Mid$(filePattern, pos, 1) = "{" Then
            closePos = VBA.InStr(pos + 1, filePattern, "}", vbBinaryCompare)
            If closePos > 0 Then
                private_BuildSearchMask = private_BuildSearchMask & "*"
                pos = closePos + 1
            Else
                private_BuildSearchMask = private_BuildSearchMask & VBA.Mid$(filePattern, pos, 1)
                pos = pos + 1
            End If
        Else
            private_BuildSearchMask = private_BuildSearchMask & VBA.Mid$(filePattern, pos, 1)
            pos = pos + 1
        End If
    Loop

    If VBA.Len(private_BuildSearchMask) = 0 Then
        private_BuildSearchMask = "*"
    End If
End Function


Private Function private_ToAbsolutePath(ByVal pathValue As String) As String
    Dim normalized As String
    Dim basePath As String
    Dim fso As Object

    normalized = private_NormalizeFilePath(pathValue)
    If VBA.Len(normalized) = 0 Then Exit Function

    If private_IsAbsolutePath(normalized) Then
        private_ToAbsolutePath = normalized
    Else
        basePath = VBA.Trim$(ThisWorkbook.Path)
        If VBA.Len(basePath) = 0 Then basePath = VBA.CurDir$
        If VBA.Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
        private_ToAbsolutePath = basePath & normalized
    End If

    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        private_ToAbsolutePath = fso.GetAbsolutePathName(private_ToAbsolutePath)
    End If
    On Error GoTo 0
End Function


Private Function private_NormalizeFilePath(ByVal pathValue As String) As String
    pathValue = VBA.Trim$(pathValue)
    If VBA.Len(pathValue) = 0 Then Exit Function

    private_NormalizeFilePath = VBA.Replace$(pathValue, "/", "\")
End Function


Private Function private_IsAbsolutePath(ByVal pathValue As String) As Boolean
    If VBA.Len(pathValue) < 2 Then Exit Function

    If VBA.Left$(pathValue, 2) = "\\" Then
        private_IsAbsolutePath = True
        Exit Function
    End If

    If VBA.Len(pathValue) >= 3 Then
        If VBA.Mid$(pathValue, 2, 1) = ":" Then
            If VBA.Mid$(pathValue, 3, 1) = "\" Or VBA.Mid$(pathValue, 3, 1) = "/" Then
                private_IsAbsolutePath = True
            End If
        End If
    End If
End Function


Private Function private_GetParentDirectory(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = VBA.InStrRev(filePath, "\", -1, vbBinaryCompare)
    If slashPos <= 1 Then Exit Function

    If slashPos = 3 And VBA.Mid$(filePath, 2, 1) = ":" Then
        private_GetParentDirectory = VBA.Left$(filePath, 3)
        Exit Function
    End If

    private_GetParentDirectory = VBA.Left$(filePath, slashPos - 1)
End Function


Private Function private_GetFileName(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = VBA.InStrRev(filePath, "\", -1, vbBinaryCompare)
    If slashPos <= 0 Then
        private_GetFileName = filePath
    Else
        private_GetFileName = VBA.Mid$(filePath, slashPos + 1)
    End If
End Function

Private Sub private_SortResolvedPaths( _
    ByRef paths() As String, _
    ByRef dates() As Date, _
    ByRef writeTimes() As Date, _
    ByVal itemCount As Long, _
    ByVal sortDescending As Boolean _
)
    Dim i As Long
    Dim j As Long
    Dim shouldSwap As Boolean
    Dim swapPath As String
    Dim swapDate As Date
    Dim swapWriteTime As Date

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If sortDescending Then
                shouldSwap = (dates(j) > dates(i)) Or _
                    (dates(j) = dates(i) And writeTimes(j) > writeTimes(i))
            Else
                shouldSwap = (dates(j) < dates(i)) Or _
                    (dates(j) = dates(i) And writeTimes(j) < writeTimes(i))
            End If
            If shouldSwap Then
                swapPath = paths(i)
                paths(i) = paths(j)
                paths(j) = swapPath
                swapDate = dates(i)
                dates(i) = dates(j)
                dates(j) = swapDate
                swapWriteTime = writeTimes(i)
                writeTimes(i) = writeTimes(j)
                writeTimes(j) = swapWriteTime
            End If
        Next j
    Next i
End Sub
