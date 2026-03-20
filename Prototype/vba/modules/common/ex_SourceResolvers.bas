Attribute VB_Name = "ex_SourceResolvers"
Option Explicit

Private Const ERR_BASE As Long = vbObjectError + 3700

Public Function m_ResolveLatestByDmyPattern( _
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

    resolverArgs = Trim$(resolverArgs)
    If Len(resolverArgs) > 0 Then
        ' Reserved for future resolver options.
    End If

    normalizedPattern = mp_NormalizeFilePath(filePathPattern)
    If Len(normalizedPattern) = 0 Then
        Err.Raise ERR_BASE + 1, "ex_SourceResolvers", "Resolver pattern is empty."
    End If

    absolutePattern = mp_ToAbsolutePath(normalizedPattern)
    If Len(absolutePattern) = 0 Then
        Err.Raise ERR_BASE + 2, "ex_SourceResolvers", "Resolver pattern could not be converted to an absolute path: " & normalizedPattern
    End If

    folderPath = mp_GetParentDirectory(absolutePattern)
    filePattern = mp_GetFileName(absolutePattern)
    If Len(folderPath) = 0 Or Len(filePattern) = 0 Then
        Err.Raise ERR_BASE + 3, "ex_SourceResolvers", "Resolver pattern must include both folder and file name: " & absolutePattern
    End If

    If Not mp_ValidateDmyPattern(filePattern, validationError) Then
        Err.Raise ERR_BASE + 4, "ex_SourceResolvers", validationError & " Pattern: " & absolutePattern
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        Err.Raise ERR_BASE + 5, "ex_SourceResolvers", "Resolver folder was not found: " & folderPath
    End If

    If (GetAttr(folderPath) And vbDirectory) = 0 Then
        Err.Raise ERR_BASE + 6, "ex_SourceResolvers", "Resolver path is not a folder: " & folderPath
    End If

    searchMask = mp_BuildSearchMask(filePattern)
    candidateName = Dir$(folderPath & "\" & searchMask, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)

    Do While Len(candidateName) > 0
        If mp_TryExtractDateByPattern(filePattern, candidateName, candidateDate) Then
            candidatePath = folderPath & "\" & candidateName
            candidateWriteTime = FileDateTime(candidatePath)

            If (Not hasBest) _
               Or (candidateDate > bestDate) _
               Or (candidateDate = bestDate And candidateWriteTime > bestWriteTime) Then
                bestPath = candidatePath
                bestDate = candidateDate
                bestWriteTime = candidateWriteTime
                hasBest = True
            End If
        End If

        candidateName = Dir$
    Loop

    If Not hasBest Then
        Err.Raise ERR_BASE + 7, "ex_SourceResolvers", _
            "No files matched the date pattern. Pattern: " & absolutePattern & ", search mask: " & searchMask
    End If

    m_ResolveLatestByDmyPattern = bestPath
End Function

Private Function mp_ValidateDmyPattern(ByVal filePattern As String, ByRef outErrorText As String) As Boolean
    Dim pos As Long
    Dim closePos As Long
    Dim token As String
    Dim hasDd As Boolean
    Dim hasMm As Boolean
    Dim hasYyyy As Boolean

    If Len(filePattern) = 0 Then
        outErrorText = "File pattern is empty."
        Exit Function
    End If

    pos = 1
    Do While pos <= Len(filePattern)
        If Mid$(filePattern, pos, 1) = "{" Then
            closePos = InStr(pos + 1, filePattern, "}", vbBinaryCompare)
            If closePos <= pos + 1 Then
                outErrorText = "File pattern contains an unclosed placeholder."
                Exit Function
            End If

            token = LCase$(Trim$(Mid$(filePattern, pos + 1, closePos - pos - 1)))
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

    mp_ValidateDmyPattern = True
End Function

Private Function mp_TryExtractDateByPattern(ByVal filePattern As String, ByVal fileName As String, ByRef outDate As Date) As Boolean
    Dim patternPos As Long
    Dim filePos As Long
    Dim closePos As Long
    Dim token As String
    Dim dd As Long
    Dim mm As Long
    Dim yyyy As Long

    patternPos = 1
    filePos = 1

    Do While patternPos <= Len(filePattern)
        If Mid$(filePattern, patternPos, 1) = "{" Then
            closePos = InStr(patternPos + 1, filePattern, "}", vbBinaryCompare)
            If closePos <= patternPos + 1 Then Exit Function

            token = LCase$(Trim$(Mid$(filePattern, patternPos + 1, closePos - patternPos - 1)))
            Select Case token
                Case "dd"
                    If Not mp_TryReadFixedDigits(fileName, filePos, 2, dd) Then Exit Function
                Case "mm"
                    If Not mp_TryReadFixedDigits(fileName, filePos, 2, mm) Then Exit Function
                Case "yyyy"
                    If Not mp_TryReadFixedDigits(fileName, filePos, 4, yyyy) Then Exit Function
                Case Else
                    Exit Function
            End Select

            patternPos = closePos + 1
        Else
            If filePos > Len(fileName) Then Exit Function
            If StrComp(Mid$(filePattern, patternPos, 1), Mid$(fileName, filePos, 1), vbTextCompare) <> 0 Then Exit Function
            patternPos = patternPos + 1
            filePos = filePos + 1
        End If
    Loop

    If filePos <= Len(fileName) Then Exit Function
    If Not mp_TryBuildExactDate(yyyy, mm, dd, outDate) Then Exit Function

    mp_TryExtractDateByPattern = True
End Function

Private Function mp_TryReadFixedDigits(ByVal textValue As String, ByRef ioPos As Long, ByVal digitsCount As Long, ByRef outValue As Long) As Boolean
    Dim chunk As String
    Dim i As Long
    Dim ch As String

    If digitsCount <= 0 Then Exit Function
    If ioPos < 1 Then Exit Function
    If ioPos + digitsCount - 1 > Len(textValue) Then Exit Function

    chunk = Mid$(textValue, ioPos, digitsCount)
    For i = 1 To Len(chunk)
        ch = Mid$(chunk, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    On Error GoTo ParseEH
    outValue = CLng(chunk)
    ioPos = ioPos + digitsCount
    mp_TryReadFixedDigits = True
    Exit Function

ParseEH:
    mp_TryReadFixedDigits = False
End Function

Private Function mp_TryBuildExactDate(ByVal yyyy As Long, ByVal mm As Long, ByVal dd As Long, ByRef outDate As Date) As Boolean
    Dim parsedDate As Date

    If yyyy < 1900 Or yyyy > 9999 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function

    On Error GoTo ParseEH
    parsedDate = DateSerial(yyyy, mm, dd)
    On Error GoTo 0

    If Year(parsedDate) <> yyyy Then Exit Function
    If Month(parsedDate) <> mm Then Exit Function
    If Day(parsedDate) <> dd Then Exit Function

    outDate = parsedDate
    mp_TryBuildExactDate = True
    Exit Function

ParseEH:
    mp_TryBuildExactDate = False
End Function

Private Function mp_BuildSearchMask(ByVal filePattern As String) As String
    Dim pos As Long
    Dim closePos As Long

    pos = 1
    Do While pos <= Len(filePattern)
        If Mid$(filePattern, pos, 1) = "{" Then
            closePos = InStr(pos + 1, filePattern, "}", vbBinaryCompare)
            If closePos > 0 Then
                mp_BuildSearchMask = mp_BuildSearchMask & "*"
                pos = closePos + 1
            Else
                mp_BuildSearchMask = mp_BuildSearchMask & Mid$(filePattern, pos, 1)
                pos = pos + 1
            End If
        Else
            mp_BuildSearchMask = mp_BuildSearchMask & Mid$(filePattern, pos, 1)
            pos = pos + 1
        End If
    Loop

    If Len(mp_BuildSearchMask) = 0 Then
        mp_BuildSearchMask = "*"
    End If
End Function

Private Function mp_ToAbsolutePath(ByVal pathValue As String) As String
    Dim normalized As String
    Dim basePath As String
    Dim fso As Object

    normalized = mp_NormalizeFilePath(pathValue)
    If Len(normalized) = 0 Then Exit Function

    If mp_IsAbsolutePath(normalized) Then
        mp_ToAbsolutePath = normalized
    Else
        basePath = Trim$(ThisWorkbook.Path)
        If Len(basePath) = 0 Then basePath = CurDir$
        If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
        mp_ToAbsolutePath = basePath & normalized
    End If

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        mp_ToAbsolutePath = fso.GetAbsolutePathName(mp_ToAbsolutePath)
    End If
    On Error GoTo 0
End Function

Private Function mp_NormalizeFilePath(ByVal pathValue As String) As String
    pathValue = Trim$(pathValue)
    If Len(pathValue) = 0 Then Exit Function

    mp_NormalizeFilePath = Replace$(pathValue, "/", "\")
End Function

Private Function mp_IsAbsolutePath(ByVal pathValue As String) As Boolean
    If Len(pathValue) < 2 Then Exit Function

    If Left$(pathValue, 2) = "\\" Then
        mp_IsAbsolutePath = True
        Exit Function
    End If

    If Len(pathValue) >= 3 Then
        If Mid$(pathValue, 2, 1) = ":" Then
            If Mid$(pathValue, 3, 1) = "\" Or Mid$(pathValue, 3, 1) = "/" Then
                mp_IsAbsolutePath = True
            End If
        End If
    End If
End Function

Private Function mp_GetParentDirectory(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, "\", -1, vbBinaryCompare)
    If slashPos <= 1 Then Exit Function

    If slashPos = 3 And Mid$(filePath, 2, 1) = ":" Then
        mp_GetParentDirectory = Left$(filePath, 3)
        Exit Function
    End If

    mp_GetParentDirectory = Left$(filePath, slashPos - 1)
End Function

Private Function mp_GetFileName(ByVal filePath As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, "\", -1, vbBinaryCompare)
    If slashPos <= 0 Then
        mp_GetFileName = filePath
    Else
        mp_GetFileName = Mid$(filePath, slashPos + 1)
    End If
End Function
