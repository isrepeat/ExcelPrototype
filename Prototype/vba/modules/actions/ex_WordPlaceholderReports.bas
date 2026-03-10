Attribute VB_Name = "ex_WordPlaceholderReports"
Option Explicit

Private Const WD_FIND_STOP As Long = 0
Private Const WD_COLLAPSE_END As Long = 0
Private Const WD_ALERTS_NONE As Long = 0

Private g_WordApp As Object
Private g_WordAppOwnedByModule As Boolean

Public Sub m_API_ExportActiveSheetFooterPlaceholderReport()
    Dim ws As Worksheet
    Dim templatePath As String
    Dim outputPath As String
    Dim footerText As String
    Dim placeholderMap As Object
    Dim reportPath As String

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1751, "ex_WordPlaceholderReports", "Active sheet is not available for Word export."
    End If

    templatePath = mp_ResolvePath(Trim$(ex_ConfigProvider.m_GetConfigValue("Export.WordTemplatePath", vbNullString)))
    If Len(templatePath) = 0 Then
        Err.Raise vbObjectError + 1763, "ex_WordPlaceholderReports", "Missing required config key 'Export.WordTemplatePath'."
    End If
    If Not mp_FileExists(templatePath) Then
        Err.Raise vbObjectError + 1764, "ex_WordPlaceholderReports", "Word template not found by config key 'Export.WordTemplatePath': " & templatePath
    End If

    outputPath = mp_BuildDefaultOutputPath(templatePath)
    If Len(outputPath) = 0 Then
        Err.Raise vbObjectError + 1766, "ex_WordPlaceholderReports", "Unable to build output path from template path: " & templatePath
    End If

    footerText = ex_PostProcessActions.m_GetSinglePostProcessFooterText(ws)
    Set placeholderMap = m_BuildPlaceholderMapFromPairs( _
        "FromHospital", footerText _
    )

    reportPath = m_CreateWordReportFromTemplate(templatePath, outputPath, placeholderMap, True)
    ex_Messaging.m_ShowNotice "Word report created: " & reportPath, 5
    Exit Sub

EH:
    MsgBox "Word footer export failed: " & Err.Description, vbExclamation
End Sub

Public Function m_CreateWordReportFromTemplate( _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal placeholderMap As Object, _
    Optional ByVal failIfPlaceholderMissing As Boolean = True _
) As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim normalizedTemplatePath As String
    Dim normalizedOutputPath As String
    Dim totalReplacements As Long
    Dim missingTokens As String
    Dim normalizedMap As Object

    On Error GoTo EH

    normalizedTemplatePath = Trim$(templatePath)
    If Len(normalizedTemplatePath) = 0 Then
        Err.Raise vbObjectError + 1752, "ex_WordPlaceholderReports", "Template path is empty."
    End If
    If Not mp_FileExists(normalizedTemplatePath) Then
        Err.Raise vbObjectError + 1753, "ex_WordPlaceholderReports", "Template file not found: " & normalizedTemplatePath
    End If

    normalizedOutputPath = Trim$(outputPath)
    If Len(normalizedOutputPath) = 0 Then
        normalizedOutputPath = mp_BuildDefaultOutputPath(normalizedTemplatePath)
    End If
    If Len(normalizedOutputPath) = 0 Then
        Err.Raise vbObjectError + 1754, "ex_WordPlaceholderReports", "Output path is empty."
    End If

    If placeholderMap Is Nothing Then
        Err.Raise vbObjectError + 1755, "ex_WordPlaceholderReports", "Placeholder map is not provided."
    End If

    Set normalizedMap = mp_BuildNormalizedPlaceholderMap(placeholderMap)

    Set wdApp = mp_GetOrCreateWordApp()
    Set wdDoc = wdApp.Documents.Add(normalizedTemplatePath)

    totalReplacements = mp_ApplyPlaceholderMapInDocument(wdDoc, normalizedMap, missingTokens)
    If failIfPlaceholderMissing And Len(missingTokens) > 0 Then
        Err.Raise vbObjectError + 1756, "ex_WordPlaceholderReports", "Placeholders not found in template: " & missingTokens
    End If

    wdDoc.SaveAs2 normalizedOutputPath

    m_CreateWordReportFromTemplate = normalizedOutputPath

    On Error Resume Next
    wdDoc.Close False
    On Error GoTo 0
    Exit Function

EH:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 1757, "ex_WordPlaceholderReports", "Failed to build Word report: " & Err.Description
End Function

Public Sub m_ResetWordSession(Optional ByVal quitIfOwned As Boolean = True)
    On Error Resume Next
    If Not g_WordApp Is Nothing Then
        If quitIfOwned And g_WordAppOwnedByModule Then
            g_WordApp.Quit
        End If
    End If
    Set g_WordApp = Nothing
    g_WordAppOwnedByModule = False
    On Error GoTo 0
End Sub

Public Function m_BuildPlaceholderMapFromPairs(ParamArray keyValuePairs() As Variant) As Object
    Dim result As Object
    Dim i As Long
    Dim keyText As String
    Dim valueText As String
    Dim lowerBound As Long
    Dim upperBound As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1 ' vbTextCompare

    On Error GoTo NoPairs
    lowerBound = LBound(keyValuePairs)
    upperBound = UBound(keyValuePairs)
    On Error GoTo 0

    If upperBound < lowerBound Then
        Err.Raise vbObjectError + 1762, "ex_WordPlaceholderReports", "Placeholder map is empty."
    End If

    If (upperBound - lowerBound + 1) Mod 2 <> 0 Then
        Err.Raise vbObjectError + 1758, "ex_WordPlaceholderReports", "Placeholder pairs must be [key, value, key, value]."
    End If

    For i = lowerBound To upperBound Step 2
        keyText = Trim$(CStr(keyValuePairs(i)))
        If Len(keyText) = 0 Then
            Err.Raise vbObjectError + 1759, "ex_WordPlaceholderReports", "Placeholder key is empty in key/value pairs."
        End If
        valueText = CStr(keyValuePairs(i + 1))
        result(keyText) = valueText
    Next i

    Set m_BuildPlaceholderMapFromPairs = result
    Exit Function

NoPairs:
    Err.Raise vbObjectError + 1762, "ex_WordPlaceholderReports", "Placeholder map is empty."
End Function

Private Function mp_BuildNormalizedPlaceholderMap(ByVal sourceMap As Object) As Object
    Dim result As Object
    Dim token As Variant
    Dim normalizedToken As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1 ' vbTextCompare

    For Each token In sourceMap.Keys
        normalizedToken = mp_NormalizePlaceholderToken(CStr(token))
        If Len(normalizedToken) = 0 Then
            Err.Raise vbObjectError + 1760, "ex_WordPlaceholderReports", "Placeholder key cannot be empty."
        End If
        result(normalizedToken) = CStr(sourceMap(token))
    Next token

    If result.Count = 0 Then
        Err.Raise vbObjectError + 1761, "ex_WordPlaceholderReports", "Placeholder map is empty."
    End If

    Set mp_BuildNormalizedPlaceholderMap = result
End Function

Private Function mp_ApplyPlaceholderMapInDocument( _
    ByVal doc As Object, _
    ByVal placeholderMap As Object, _
    ByRef outMissingTokens As String _
) As Long
    Dim story As Object
    Dim currentRange As Object
    Dim token As Variant
    Dim perTokenHits As Object
    Dim replaceHits As Long

    Set perTokenHits = CreateObject("Scripting.Dictionary")
    perTokenHits.CompareMode = 1 ' vbTextCompare

    For Each token In placeholderMap.Keys
        perTokenHits(CStr(token)) = 0
    Next token

    For Each story In doc.StoryRanges
        Set currentRange = story

        Do While Not currentRange Is Nothing
            For Each token In placeholderMap.Keys
                replaceHits = mp_ReplaceTokenInRange(currentRange, CStr(token), CStr(placeholderMap(token)))
                If replaceHits > 0 Then
                    perTokenHits(CStr(token)) = CLng(perTokenHits(CStr(token))) + replaceHits
                    mp_ApplyPlaceholderMapInDocument = mp_ApplyPlaceholderMapInDocument + replaceHits
                End If
            Next token

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story

    For Each token In perTokenHits.Keys
        If CLng(perTokenHits(CStr(token))) <= 0 Then
            If Len(outMissingTokens) > 0 Then outMissingTokens = outMissingTokens & ", "
            outMissingTokens = outMissingTokens & CStr(token)
        End If
    Next token
End Function

Private Function mp_ReplaceTokenInRange( _
    ByVal sourceRange As Object, _
    ByVal token As String, _
    ByVal replacementText As String _
) As Long
    Dim findRange As Object

    Set findRange = sourceRange.Duplicate

    With findRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = token
        .Replacement.Text = vbNullString
        .Forward = True
        .Wrap = WD_FIND_STOP
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    Do While findRange.Find.Execute
        findRange.Text = replacementText
        mp_ReplaceTokenInRange = mp_ReplaceTokenInRange + 1
        findRange.Collapse WD_COLLAPSE_END
    Loop
End Function

Private Function mp_NormalizePlaceholderToken(ByVal tokenText As String) As String
    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then Exit Function

    If Left$(tokenText, 1) = "{" And Right$(tokenText, 1) = "}" Then
        mp_NormalizePlaceholderToken = tokenText
    Else
        mp_NormalizePlaceholderToken = "{" & tokenText & "}"
    End If
End Function

Private Function mp_BuildDefaultOutputPath(ByVal templatePath As String) As String
    Dim folderPath As String
    Dim fileNameOnly As String
    Dim dotPos As Long

    folderPath = mp_ExtractFolderPath(templatePath)
    fileNameOnly = mp_ExtractFileName(templatePath)
    If Len(fileNameOnly) = 0 Then Exit Function

    dotPos = InStrRev(fileNameOnly, ".")
    If dotPos > 1 Then
        fileNameOnly = Left$(fileNameOnly, dotPos - 1)
    End If

    mp_BuildDefaultOutputPath = folderPath & fileNameOnly & "_result.docx"
End Function

Private Function mp_ExtractFolderPath(ByVal filePath As String) As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos > 0 Then
        mp_ExtractFolderPath = Left$(filePath, sepPos)
    Else
        mp_ExtractFolderPath = ThisWorkbook.Path & "\"
    End If
End Function

Private Function mp_ExtractFileName(ByVal filePath As String) As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos > 0 Then
        mp_ExtractFileName = Mid$(filePath, sepPos + 1)
    Else
        mp_ExtractFileName = filePath
    End If
End Function

Private Function mp_FileExists(ByVal filePath As String) As Boolean
    mp_FileExists = Len(Dir$(filePath, vbNormal)) > 0
End Function

Private Function mp_GetOrCreateWordApp() As Object
    Dim wdApp As Object

    If mp_IsWordAppAlive(g_WordApp) Then
        Set mp_GetOrCreateWordApp = g_WordApp
        Exit Function
    End If

    Set g_WordApp = Nothing
    g_WordAppOwnedByModule = False

    Set wdApp = CreateObject("Word.Application")
    If wdApp Is Nothing Then
        Err.Raise vbObjectError + 1765, "ex_WordPlaceholderReports", "Unable to start Word.Application."
    End If
    wdApp.Visible = False
    On Error Resume Next
    wdApp.DisplayAlerts = WD_ALERTS_NONE
    On Error GoTo 0
    g_WordAppOwnedByModule = True

    Set g_WordApp = wdApp
    Set mp_GetOrCreateWordApp = g_WordApp
End Function

Private Function mp_IsWordAppAlive(ByVal wdApp As Object) As Boolean
    Dim docCount As Long

    If wdApp Is Nothing Then Exit Function

    On Error Resume Next
    docCount = CLng(wdApp.Documents.Count)
    If Err.Number = 0 Then mp_IsWordAppAlive = True
    On Error GoTo 0
End Function

Private Function mp_ResolvePath(ByVal inputPath As String) As String
    Dim basePath As String
    Dim normalized As String

    normalized = Trim$(inputPath)
    If Len(normalized) = 0 Then Exit Function
    normalized = Replace$(normalized, "/", "\")

    If Left$(normalized, 2) = "\\" Or InStr(1, normalized, ":\", vbTextCompare) > 0 Then
        mp_ResolvePath = normalized
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_ResolvePath = normalized
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    mp_ResolvePath = basePath & normalized
End Function

Private Function mp_TryPromptRequiredText( _
    ByVal prompt As String, _
    ByVal title As String, _
    ByVal defaultText As String, _
    ByRef outText As String _
) As Boolean
    If Not mp_TryPromptText(prompt, title, defaultText, outText) Then Exit Function
    If Len(outText) = 0 Then
        MsgBox "Input cannot be empty.", vbExclamation
        Exit Function
    End If
    mp_TryPromptRequiredText = True
End Function

Private Function mp_TryPromptText( _
    ByVal prompt As String, _
    ByVal title As String, _
    ByVal defaultText As String, _
    ByRef outText As String _
) As Boolean
    Dim response As Variant

    response = Application.InputBox(prompt, title, defaultText, Type:=2)
    If VarType(response) = vbBoolean Then
        If CBool(response) = False Then Exit Function
    End If

    outText = Trim$(CStr(response))
    mp_TryPromptText = True
End Function
