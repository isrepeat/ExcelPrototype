Attribute VB_Name = "ex_SourceSnapshot"
Option Explicit

Private Const TEMP_FOLDER_NAME As String = "!TEMP"

Public Function m_GetSnapshotPath(ByVal sourcePath As String, Optional ByVal sourceTag As String = vbNullString) As String
    Dim normalizedSourcePath As String
    Dim tempFolderPath As String
    Dim snapshotPath As String

    normalizedSourcePath = Trim$(sourcePath)
    If Len(normalizedSourcePath) = 0 Then
        Err.Raise vbObjectError + 3600, "ex_SourceSnapshot", "Source path is empty."
    End If

    normalizedSourcePath = mp_NormalizePath(normalizedSourcePath)
    If Dir(normalizedSourcePath) = vbNullString Then
        Err.Raise vbObjectError + 3601, "ex_SourceSnapshot", "Source file not found: " & normalizedSourcePath
    End If

    tempFolderPath = mp_EnsureTempFolderPath()
    snapshotPath = tempFolderPath & "\" & mp_BuildSnapshotFileName(normalizedSourcePath)

    ' Snapshot name includes source signature (mtime+size), so if file exists we can reuse it.
    If Dir(snapshotPath) = vbNullString Then
        mp_CopySnapshot normalizedSourcePath, snapshotPath, sourceTag
    End If

    If Dir(snapshotPath) = vbNullString Then
        Err.Raise vbObjectError + 3602, "ex_SourceSnapshot", _
            "Snapshot file was not created in !TEMP. Source: " & normalizedSourcePath & ", Snapshot: " & snapshotPath
    End If

    m_GetSnapshotPath = snapshotPath
End Function

Private Function mp_NormalizePath(ByVal pathValue As String) As String
    Dim fso As Object

    On Error GoTo Fallback
    Set fso = CreateObject("Scripting.FileSystemObject")
    mp_NormalizePath = fso.GetAbsolutePathName(pathValue)
    Exit Function

Fallback:
    mp_NormalizePath = pathValue
End Function

Private Sub mp_CopySnapshot(ByVal sourcePath As String, ByVal snapshotPath As String, ByVal sourceTag As String)
    Dim contextText As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim wbSource As Workbook

    On Error GoTo FileCopyEH
    FileCopy sourcePath, snapshotPath
    Exit Sub

FileCopyEH:
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear

    ' If source is open in current Excel instance, SaveCopyAs can bypass file copy lock.
    If errNumber = 70 Then
        Set wbSource = mp_FindOpenWorkbookByPath(sourcePath)
        If Not wbSource Is Nothing Then
            On Error GoTo SaveCopyEH
            wbSource.SaveCopyAs snapshotPath
            Exit Sub
        End If
    End If

    GoTo RaiseCopyError

SaveCopyEH:
    errNumber = Err.Number
    errDescription = "SaveCopyAs failed: " & Err.Description
    Err.Clear

RaiseCopyError:
    contextText = Trim$(sourceTag)
    If Len(contextText) > 0 Then
        contextText = " [" & contextText & "]"
    End If

    Err.Raise vbObjectError + 3603, "ex_SourceSnapshot", _
        "Failed to copy source file to !TEMP" & contextText & ". Source: " & sourcePath & ", Snapshot: " & snapshotPath & ". " & _
        "InnerError #" & CStr(errNumber) & ": " & errDescription
End Sub

Private Function mp_FindOpenWorkbookByPath(ByVal sourcePath As String) As Workbook
    Dim wb As Workbook
    Dim sourceAbs As String
    Dim wbAbs As String

    sourceAbs = LCase$(mp_NormalizePath(sourcePath))

    For Each wb In Application.Workbooks
        On Error Resume Next
        wbAbs = LCase$(mp_NormalizePath(CStr(wb.FullName)))
        If Err.Number <> 0 Then
            Err.Clear
            wbAbs = vbNullString
        End If
        On Error GoTo 0

        If Len(wbAbs) > 0 Then
            If StrComp(wbAbs, sourceAbs, vbBinaryCompare) = 0 Then
                Set mp_FindOpenWorkbookByPath = wb
                Exit Function
            End If
        End If
    Next wb
End Function

Private Function mp_EnsureTempFolderPath() As String
    Dim basePath As String
    Dim tempPath As String

    basePath = Trim$(ThisWorkbook.Path)
    If Len(basePath) = 0 Then
        Err.Raise vbObjectError + 3604, "ex_SourceSnapshot", "Workbook path is empty. Save workbook before generating result."
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    tempPath = basePath & TEMP_FOLDER_NAME

    If Dir(tempPath, vbDirectory) = vbNullString Then
        MkDir tempPath
    End If

    If (GetAttr(tempPath) And vbDirectory) = 0 Then
        Err.Raise vbObjectError + 3605, "ex_SourceSnapshot", "Path exists but is not a folder: " & tempPath
    End If

    mp_EnsureTempFolderPath = tempPath
End Function

Private Function mp_BuildSnapshotFileName(ByVal sourcePath As String) As String
    Dim fileName As String
    Dim baseName As String
    Dim ext As String
    Dim dotPos As Long
    Dim pathFingerprint As String
    Dim sourceSignature As String

    fileName = mp_GetFileNameFromPath(sourcePath)
    dotPos = InStrRev(fileName, ".")

    If dotPos > 1 Then
        baseName = Left$(fileName, dotPos - 1)
        ext = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        ext = vbNullString
    End If

    baseName = mp_SanitizeFileNameToken(baseName)
    If Len(baseName) = 0 Then baseName = "source"
    If Len(baseName) > 80 Then baseName = Left$(baseName, 80)

    pathFingerprint = mp_BuildPathFingerprint(LCase$(sourcePath))
    sourceSignature = mp_BuildSourceSignatureToken(sourcePath)

    mp_BuildSnapshotFileName = baseName & "_" & pathFingerprint & "_" & sourceSignature & LCase$(ext)
End Function

Private Function mp_BuildSourceSignatureToken(ByVal sourcePath As String) As String
    Dim signatureRaw As String

    signatureRaw = CStr(CDbl(FileDateTime(sourcePath))) & "|" & CStr(FileLen(sourcePath))
    mp_BuildSourceSignatureToken = mp_BuildPathFingerprint(signatureRaw)
End Function

Private Function mp_GetFileNameFromPath(ByVal pathValue As String) As String
    Dim slashPos As Long
    Dim backslashPos As Long
    Dim cutPos As Long

    slashPos = InStrRev(pathValue, "/")
    backslashPos = InStrRev(pathValue, "\")
    cutPos = slashPos
    If backslashPos > cutPos Then cutPos = backslashPos

    If cutPos > 0 Then
        mp_GetFileNameFromPath = Mid$(pathValue, cutPos + 1)
    Else
        mp_GetFileNameFromPath = pathValue
    End If
End Function

Private Function mp_SanitizeFileNameToken(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", "(", ")"
                outText = outText & ch
            Case Else
                outText = outText & "_"
        End Select
    Next i

    Do While InStr(1, outText, "__", vbBinaryCompare) > 0
        outText = Replace$(outText, "__", "_")
    Loop

    mp_SanitizeFileNameToken = Trim$(outText)
End Function

Private Function mp_BuildPathFingerprint(ByVal valueText As String) As String
    Const MODULO As Double = 2147483629#
    Dim acc As Double
    Dim i As Long
    Dim codePoint As Long

    acc = 7#

    For i = 1 To Len(valueText)
        codePoint = AscW(Mid$(valueText, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536

        acc = (acc * 131#) + CDbl(codePoint)
        acc = acc - (Fix(acc / MODULO) * MODULO)
    Next i

    mp_BuildPathFingerprint = Right$("00000000" & Hex$(CLng(acc)), 8)
End Function
