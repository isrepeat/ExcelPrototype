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
                Case "mm"
                    hasMm = True
                Case "yyyy"
                    hasYyyy = True
                Case Else
                    outErrorText = "Unsupported placeholder '{" & token & "}'. Only {dd}, {mm}, {yyyy} are allowed in this resolver."
                    Exit Function
            End Select

            pos = closePos + 1
        Else
            pos = pos + 1
        End If
    Loop

    If Not hasDd Or Not hasMm Or Not hasYyyy Then
        outErrorText = "Date placeholders are required: {dd}, {mm}, {yyyy}."
        Exit Function
    End If

    private_ValidateDmyPattern = True
End Function


Private Function private_TryExtractDateByPattern(ByVal filePattern As String, ByVal fileName As String, ByRef outDate As Date) As Boolean
    Dim patternPos As Long
    Dim filePos As Long
    Dim closePos As Long
    Dim token As String
    Dim dd As Long
    Dim mm As Long
    Dim yyyy As Long

    patternPos = 1
    filePos = 1

    Do While patternPos <= VBA.Len(filePattern)
        If VBA.Mid$(filePattern, patternPos, 1) = "{" Then
            closePos = VBA.InStr(patternPos + 1, filePattern, "}", vbBinaryCompare)
            If closePos <= patternPos + 1 Then Exit Function

            token = VBA.LCase$(VBA.Trim$(VBA.Mid$(filePattern, patternPos + 1, closePos - patternPos - 1)))
            Select Case token
                Case "dd"
                    If Not private_TryReadFixedDigits(fileName, filePos, 2, dd) Then Exit Function
                Case "mm"
                    If Not private_TryReadFixedDigits(fileName, filePos, 2, mm) Then Exit Function
                Case "yyyy"
                    If Not private_TryReadFixedDigits(fileName, filePos, 4, yyyy) Then Exit Function
                Case Else
                    Exit Function
            End Select

            patternPos = closePos + 1
        Else
            If filePos > VBA.Len(fileName) Then Exit Function
            If VBA.StrComp(VBA.Mid$(filePattern, patternPos, 1), VBA.Mid$(fileName, filePos, 1), vbTextCompare) <> 0 Then Exit Function
            patternPos = patternPos + 1
            filePos = filePos + 1
        End If
    Loop

    If filePos <= VBA.Len(fileName) Then Exit Function
    If Not private_TryBuildExactDate(yyyy, mm, dd, outDate) Then Exit Function

    private_TryExtractDateByPattern = True
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
    Set fso = CreateObject("Scripting.FileSystemObject")
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