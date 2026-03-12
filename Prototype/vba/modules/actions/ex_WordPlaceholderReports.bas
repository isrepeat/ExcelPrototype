Attribute VB_Name = "ex_WordPlaceholderReports"
Option Explicit

Private Const WD_FIND_STOP As Long = 0
Private Const WD_COLLAPSE_END As Long = 0
Private Const WD_ALERTS_NONE As Long = 0
Private Const EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX As String = "CREATEWITHPOSTFIX"
Private Const EXPORT_OUTPUT_MODE_OVERWRITE_TEMPLATE As String = "OVERWRITETEMPLATE"
Private Const EXPORT_INSERT_MODE_REPLACE_ALL As String = "REPLACEALL"
Private Const EXPORT_INSERT_MODE_APPEND_TOP As String = "APPENDTOTOP"
Private Const EXPORT_INSERT_MODE_APPEND_BOTTOM As String = "APPENDTOBOTTOM"
Private Const EXPORT_RUNTIME_WORD_RESULTS_PLACE As String = "Export.RuntimeDataBase.WordResultsPlace"
Private Const EXPORT_RUNTIME_WORD_PASTE_ANCHOR As String = "Export.RuntimeDataBase.WordPasteAnchor"
Private Const EXPORT_APPEND_SEPARATOR As String = vbCrLf & vbCrLf
Private Const EXPORT_BOOKMARK_PREFIX As String = "EP_Anchor_"
Private Const EXPORT_BOOKMARK_MAX_LEN As Long = 40
Private Const EXPORT_DUPLICATE_CONFIRM_TITLE As String = "Duplicate Export Confirmation"
Private Const EXPORT_HASH_MODULO As Double = 2147483629#

Private g_WordApp As Object
Private g_WordAppOwnedByModule As Boolean
Private g_LastExportHashBySheet As Object

Public Sub m_API_ExportActiveSheetFooterPlaceholderReport()
    Dim ws As Worksheet
    Dim templatePath As String
    Dim outputPath As String
    Dim placeholderMap As Object
    Dim reportPath As String
    Dim insertMode As String
    Dim outputMode As String
    Dim outputPostfix As String
    Dim wordResultsPlace As String
    Dim wordPasteAnchor As String
    Dim sourceText As String
    Dim currentExportHash As String

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

    insertMode = mp_NormalizeInsertMode(ex_ConfigProvider.m_GetConfigValue("Export.InsertMode", "ReplaceAll"))
    outputMode = mp_NormalizeOutputMode(ex_ConfigProvider.m_GetConfigValue("Export.OutputMode", "CreateWithPostfix"))
    outputPostfix = CStr(ex_ConfigProvider.m_GetConfigValue("Export.OutputPostfix", "_result"))

    outputPath = mp_BuildOutputPathByMode(templatePath, outputMode, outputPostfix)
    If Len(outputPath) = 0 Then
        Err.Raise vbObjectError + 1766, "ex_WordPlaceholderReports", "Unable to build output path by mode '" & outputMode & "' from template path: " & templatePath
    End If

    wordResultsPlace = Trim$(ex_PostProcessActions.m_GetRuntimeData(EXPORT_RUNTIME_WORD_RESULTS_PLACE, vbNullString, ws))
    If Len(wordResultsPlace) = 0 Then
        Err.Raise vbObjectError + 1774, "ex_WordPlaceholderReports", _
            "Missing runtime export value '" & EXPORT_RUNTIME_WORD_RESULTS_PLACE & "'. " & _
            "Run Search -> Post Process and ensure at least one export results block was generated."
    End If

    wordPasteAnchor = Trim$(ex_PostProcessActions.m_GetRuntimeData(EXPORT_RUNTIME_WORD_PASTE_ANCHOR, vbNullString, ws))
    If Len(wordPasteAnchor) = 0 Then
        Err.Raise vbObjectError + 1775, "ex_WordPlaceholderReports", _
            "Missing runtime export value '" & EXPORT_RUNTIME_WORD_PASTE_ANCHOR & "'. " & _
            "Run Post Process to prepare export anchor."
    End If

    sourceText = mp_ReadSheetTextByRuntimePointer(ws, wordResultsPlace)
    currentExportHash = mp_ComputeTextHashHex(wordPasteAnchor & vbLf & sourceText)
    If Not mp_ConfirmProceedForExportHash(ws, currentExportHash) Then
        ex_Messaging.m_ShowNotice "Export canceled.", 3
        Exit Sub
    End If

    Set placeholderMap = m_BuildPlaceholderMapFromPairs(wordPasteAnchor, sourceText)

    reportPath = m_CreateWordReportFromTemplate(templatePath, outputPath, placeholderMap, True, insertMode)
    mp_SaveLastExportHash ws, currentExportHash
    ex_Messaging.m_ShowNotice "Word report created: " & reportPath, 5
    Exit Sub

EH:
    MsgBox "Word footer export failed: " & Err.Description, vbExclamation
End Sub

Public Function m_CreateWordReportFromTemplate( _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal placeholderMap As Object, _
    Optional ByVal failIfPlaceholderMissing As Boolean = True, _
    Optional ByVal insertMode As String = "ReplaceAll" _
) As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim normalizedTemplatePath As String
    Dim normalizedOutputPath As String
    Dim normalizedInsertMode As String
    Dim totalReplacements As Long
    Dim missingTokens As String
    Dim normalizedMap As Object
    Dim saveInPlace As Boolean
    Dim useExistingOutput As Boolean
    Dim retriedFromTemplate As Boolean
    Dim failureSource As String
    Dim failureDescription As String
    Dim failureNumber As Long

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

    normalizedInsertMode = mp_NormalizeInsertMode(insertMode)

    If placeholderMap Is Nothing Then
        Err.Raise vbObjectError + 1755, "ex_WordPlaceholderReports", "Placeholder map is not provided."
    End If

    Set normalizedMap = mp_BuildNormalizedPlaceholderMap(placeholderMap)

    Set wdApp = mp_GetOrCreateWordApp()
    saveInPlace = (StrComp(normalizedOutputPath, normalizedTemplatePath, vbTextCompare) = 0)
    useExistingOutput = False
    If Not saveInPlace Then
        If mp_FileExists(normalizedOutputPath) Then
            If StrComp(normalizedInsertMode, EXPORT_INSERT_MODE_APPEND_TOP, vbTextCompare) = 0 Or _
               StrComp(normalizedInsertMode, EXPORT_INSERT_MODE_APPEND_BOTTOM, vbTextCompare) = 0 Then
                useExistingOutput = True
            End If
        End If
    End If

    If saveInPlace Then
        Set wdDoc = wdApp.Documents.Open(normalizedTemplatePath, False, False)
    ElseIf useExistingOutput Then
        Set wdDoc = wdApp.Documents.Open(normalizedOutputPath, False, False)
    Else
        Set wdDoc = wdApp.Documents.Add(normalizedTemplatePath)
    End If

    totalReplacements = mp_ApplyPlaceholderMapInDocument(wdDoc, normalizedMap, missingTokens, normalizedInsertMode)
    If Len(missingTokens) > 0 And useExistingOutput Then
        On Error Resume Next
        wdDoc.Close False
        On Error GoTo EH

        Set wdDoc = wdApp.Documents.Add(normalizedTemplatePath)
        useExistingOutput = False
        retriedFromTemplate = True
        missingTokens = vbNullString
        totalReplacements = mp_ApplyPlaceholderMapInDocument(wdDoc, normalizedMap, missingTokens, normalizedInsertMode)
    End If

    If failIfPlaceholderMissing And Len(missingTokens) > 0 Then
        If retriedFromTemplate Then
            Err.Raise vbObjectError + 1756, "ex_WordPlaceholderReports", "Placeholders not found in template: " & missingTokens
        Else
            Err.Raise vbObjectError + 1756, "ex_WordPlaceholderReports", "Placeholders not found in template/output: " & missingTokens
        End If
    End If

    If saveInPlace Or useExistingOutput Then
        wdDoc.Save
    Else
        wdDoc.SaveAs2 normalizedOutputPath
    End If

    m_CreateWordReportFromTemplate = normalizedOutputPath

    On Error Resume Next
    wdDoc.Close False
    On Error GoTo 0
    Exit Function

EH:
    failureSource = Err.Source
    failureDescription = Err.Description
    failureNumber = Err.Number
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 1757, "ex_WordPlaceholderReports", _
        "Failed to build Word report. Cause: [" & failureSource & " #" & CStr(failureNumber) & "] " & failureDescription
End Function

Private Function mp_ConfirmProceedForExportHash(ByVal ws As Worksheet, ByVal currentExportHash As String) As Boolean
    Dim previousHash As String
    Dim messageText As String
    Dim response As VbMsgBoxResult

    previousHash = mp_GetLastExportHash(ws)
    If Len(previousHash) = 0 Then
        mp_ConfirmProceedForExportHash = True
        Exit Function
    End If

    If StrComp(previousHash, currentExportHash, vbBinaryCompare) <> 0 Then
        mp_ConfirmProceedForExportHash = True
        Exit Function
    End If

    messageText = "The current Post Process result matches the previously exported content (hash: " & currentExportHash & ")." & vbCrLf & vbCrLf & _
                  "Do you want to export the same content again?"
    response = MsgBox(messageText, vbQuestion + vbYesNo + vbDefaultButton2, EXPORT_DUPLICATE_CONFIRM_TITLE)
    mp_ConfirmProceedForExportHash = (response = vbYes)
End Function

Private Function mp_ComputeTextHashHex(ByVal valueText As String) As String
    Dim acc As Double
    Dim i As Long
    Dim codePoint As Long

    acc = 7#

    For i = 1 To Len(valueText)
        codePoint = AscW(Mid$(valueText, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536

        acc = (acc * 131#) + CDbl(codePoint)
        acc = acc - (Fix(acc / EXPORT_HASH_MODULO) * EXPORT_HASH_MODULO)
    Next i

    mp_ComputeTextHashHex = Right$("00000000" & Hex$(CLng(acc)), 8)
End Function

Private Function mp_GetLastExportHash(ByVal ws As Worksheet) As String
    Dim cache As Object
    Dim cacheKey As String

    If ws Is Nothing Then Exit Function
    Set cache = mp_EnsureLastExportHashCache()
    cacheKey = mp_BuildSheetExportCacheKey(ws)
    If Len(cacheKey) = 0 Then Exit Function

    If cache.Exists(cacheKey) Then
        mp_GetLastExportHash = CStr(cache(cacheKey))
    End If
End Function

Private Sub mp_SaveLastExportHash(ByVal ws As Worksheet, ByVal exportHash As String)
    Dim cache As Object
    Dim cacheKey As String

    If ws Is Nothing Then Exit Sub
    Set cache = mp_EnsureLastExportHashCache()
    cacheKey = mp_BuildSheetExportCacheKey(ws)
    If Len(cacheKey) = 0 Then Exit Sub

    cache(cacheKey) = CStr(exportHash)
End Sub

Private Function mp_EnsureLastExportHashCache() As Object
    If g_LastExportHashBySheet Is Nothing Then
        Set g_LastExportHashBySheet = CreateObject("Scripting.Dictionary")
        g_LastExportHashBySheet.CompareMode = 1 ' vbTextCompare
    End If
    Set mp_EnsureLastExportHashCache = g_LastExportHashBySheet
End Function

Private Function mp_BuildSheetExportCacheKey(ByVal ws As Worksheet) As String
    Dim wbName As String
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    wbName = CStr(ws.Parent.Name)
    On Error GoTo 0

    mp_BuildSheetExportCacheKey = LCase$(Trim$(wbName) & "|" & Trim$(ws.Name))
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
    Set g_LastExportHashBySheet = Nothing
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
    ByRef outMissingTokens As String, _
    Optional ByVal insertMode As String = EXPORT_INSERT_MODE_REPLACE_ALL _
) As Long
    Dim token As Variant
    Dim perTokenHits As Object
    Dim tokenHits As Long
    Dim normalizedMode As String

    Set perTokenHits = CreateObject("Scripting.Dictionary")
    perTokenHits.CompareMode = 1 ' vbTextCompare

    normalizedMode = mp_NormalizeInsertMode(insertMode)

    For Each token In placeholderMap.Keys
        perTokenHits(CStr(token)) = 0
    Next token

    For Each token In placeholderMap.Keys
        tokenHits = mp_ReplaceTokenInDocumentByMode(doc, CStr(token), CStr(placeholderMap(token)), normalizedMode)
        If tokenHits > 0 Then
            perTokenHits(CStr(token)) = CLng(perTokenHits(CStr(token))) + tokenHits
            mp_ApplyPlaceholderMapInDocument = mp_ApplyPlaceholderMapInDocument + tokenHits
        End If
    Next token

    For Each token In perTokenHits.Keys
        If CLng(perTokenHits(CStr(token))) <= 0 Then
            If Len(outMissingTokens) > 0 Then outMissingTokens = outMissingTokens & ", "
            outMissingTokens = outMissingTokens & CStr(token)
        End If
    Next token
End Function

Private Function mp_ReplaceTokenInDocumentByMode( _
    ByVal doc As Object, _
    ByVal token As String, _
    ByVal replacementText As String, _
    ByVal insertMode As String _
) As Long
    Dim story As Object
    Dim currentRange As Object
    Dim firstHit As Object
    Dim lastHit As Object
    Dim bookmarkRange As Object
    Dim appliedRange As Object
    Dim existingText As String
    Dim mergedText As String
    Dim hitCount As Long

    If StrComp(insertMode, EXPORT_INSERT_MODE_REPLACE_ALL, vbTextCompare) = 0 Then
        For Each story In doc.StoryRanges
            Set currentRange = story
            Do While Not currentRange Is Nothing
                mp_ReplaceTokenInDocumentByMode = mp_ReplaceTokenInDocumentByMode + mp_ReplaceTokenInRange(currentRange, token, replacementText)
                Set currentRange = currentRange.NextStoryRange
            Loop
        Next story
        Exit Function
    End If

    For Each story In doc.StoryRanges
        Set currentRange = story
        Do While Not currentRange Is Nothing
            hitCount = hitCount + mp_CaptureTokenHitsInRange(currentRange, token, firstHit, lastHit)
            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story

    If hitCount > 0 Then
        If StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_TOP, vbTextCompare) = 0 Then
            firstHit.Text = replacementText
            Set appliedRange = firstHit.Duplicate
            mp_UpsertAnchorBookmark doc, token, appliedRange
        ElseIf StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_BOTTOM, vbTextCompare) = 0 Then
            lastHit.Text = replacementText
            Set appliedRange = lastHit.Duplicate
            mp_UpsertAnchorBookmark doc, token, appliedRange
        Else
            Err.Raise vbObjectError + 1769, "ex_WordPlaceholderReports", "Unsupported insert mode: " & insertMode
        End If
        mp_ReplaceTokenInDocumentByMode = 1
        Exit Function
    End If

    If Not mp_TryGetAnchorBookmarkRange(doc, token, bookmarkRange) Then Exit Function

    existingText = CStr(bookmarkRange.Text)
    If Len(existingText) > 0 Then
        If StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_TOP, vbTextCompare) = 0 Then
            mergedText = replacementText & EXPORT_APPEND_SEPARATOR & existingText
        ElseIf StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_BOTTOM, vbTextCompare) = 0 Then
            mergedText = existingText & EXPORT_APPEND_SEPARATOR & replacementText
        Else
            Err.Raise vbObjectError + 1769, "ex_WordPlaceholderReports", "Unsupported insert mode: " & insertMode
        End If
    Else
        mergedText = replacementText
    End If

    bookmarkRange.Text = mergedText
    Set appliedRange = bookmarkRange.Duplicate
    mp_UpsertAnchorBookmark doc, token, appliedRange
    mp_ReplaceTokenInDocumentByMode = 1
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

Private Function mp_CaptureTokenHitsInRange( _
    ByVal sourceRange As Object, _
    ByVal token As String, _
    ByRef firstHit As Object, _
    ByRef lastHit As Object _
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
        If firstHit Is Nothing Then
            Set firstHit = findRange.Duplicate
        End If
        Set lastHit = findRange.Duplicate
        mp_CaptureTokenHitsInRange = mp_CaptureTokenHitsInRange + 1
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

Private Sub mp_UpsertAnchorBookmark(ByVal doc As Object, ByVal token As String, ByVal targetRange As Object)
    Dim bookmarkName As String

    If doc Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub

    bookmarkName = mp_BuildAnchorBookmarkName(token)
    If Len(bookmarkName) = 0 Then Exit Sub

    On Error Resume Next
    If doc.Bookmarks.Exists(bookmarkName) Then
        doc.Bookmarks(bookmarkName).Delete
    End If
    doc.Bookmarks.Add bookmarkName, targetRange
    On Error GoTo 0
End Sub

Private Function mp_TryGetAnchorBookmarkRange(ByVal doc As Object, ByVal token As String, ByRef outRange As Object) As Boolean
    Dim bookmarkName As String

    If doc Is Nothing Then Exit Function
    bookmarkName = mp_BuildAnchorBookmarkName(token)
    If Len(bookmarkName) = 0 Then Exit Function

    On Error GoTo CleanFail
    If doc.Bookmarks.Exists(bookmarkName) Then
        Set outRange = doc.Bookmarks(bookmarkName).Range
        mp_TryGetAnchorBookmarkRange = Not (outRange Is Nothing)
    End If
    Exit Function

CleanFail:
    Set outRange = Nothing
    mp_TryGetAnchorBookmarkRange = False
End Function

Private Function mp_BuildAnchorBookmarkName(ByVal token As String) As String
    Dim rawToken As String
    Dim cleanToken As String
    Dim i As Long
    Dim ch As String
    Dim resultName As String

    rawToken = Trim$(token)
    If Len(rawToken) = 0 Then Exit Function
    If Left$(rawToken, 1) = "{" And Right$(rawToken, 1) = "}" Then
        rawToken = Mid$(rawToken, 2, Len(rawToken) - 2)
    End If
    rawToken = Trim$(rawToken)
    If Len(rawToken) = 0 Then Exit Function

    For i = 1 To Len(rawToken)
        ch = Mid$(rawToken, i, 1)
        If mp_IsAsciiAlphaNum(ch) Then
            cleanToken = cleanToken & ch
        Else
            cleanToken = cleanToken & "_"
        End If
    Next i

    resultName = EXPORT_BOOKMARK_PREFIX & cleanToken
    If Len(resultName) > EXPORT_BOOKMARK_MAX_LEN Then
        resultName = Left$(resultName, EXPORT_BOOKMARK_MAX_LEN)
    End If
    mp_BuildAnchorBookmarkName = resultName
End Function

Private Function mp_IsAsciiAlphaNum(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
        mp_IsAsciiAlphaNum = True
    End If
End Function

Private Function mp_BuildDefaultOutputPath(ByVal templatePath As String) As String
    mp_BuildDefaultOutputPath = mp_BuildPostfixedOutputPath(templatePath, "_result")
End Function

Private Function mp_ReadSheetTextByRuntimePointer( _
    ByVal ws As Worksheet, _
    ByVal sourcePointer As String _
) As String
    Dim normalizedPointer As String
    Dim addressText As String
    Dim pointerRange As Range
    Dim colonPos As Long
    Dim prefixText As String
    Dim prefixCandidate As String

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1776, "ex_WordPlaceholderReports", "Target worksheet is missing for runtime source pointer."
    End If

    normalizedPointer = Trim$(sourcePointer)
    If Len(normalizedPointer) = 0 Then
        Err.Raise vbObjectError + 1777, "ex_WordPlaceholderReports", "Runtime source pointer is empty."
    End If

    colonPos = InStr(1, normalizedPointer, ":", vbBinaryCompare)
    If colonPos > 0 Then
        prefixCandidate = Trim$(Left$(normalizedPointer, colonPos - 1))
        prefixText = UCase$(prefixCandidate)
        If StrComp(prefixText, "CELL", vbBinaryCompare) = 0 Or StrComp(prefixText, "RANGE", vbBinaryCompare) = 0 Then
            addressText = Trim$(Mid$(normalizedPointer, colonPos + 1))
        Else
            addressText = normalizedPointer
        End If
    Else
        addressText = normalizedPointer
    End If

    If Len(addressText) = 0 Then
        Err.Raise vbObjectError + 1779, "ex_WordPlaceholderReports", "Runtime source pointer address is empty in '" & normalizedPointer & "'."
    End If

    On Error GoTo ResolveErr
    Set pointerRange = ws.Range(addressText)
    On Error GoTo 0
    If pointerRange Is Nothing Then
        Err.Raise vbObjectError + 1780, "ex_WordPlaceholderReports", "Unable to resolve runtime source pointer '" & normalizedPointer & "' on sheet '" & ws.Name & "'."
    End If

    mp_ReadSheetTextByRuntimePointer = CStr(pointerRange.Cells(1, 1).Value)
    Exit Function

ResolveErr:
    Err.Raise vbObjectError + 1781, "ex_WordPlaceholderReports", "Invalid runtime source pointer '" & normalizedPointer & "' on sheet '" & ws.Name & "': " & Err.Description
End Function

Private Function mp_BuildOutputPathByMode( _
    ByVal templatePath As String, _
    ByVal outputMode As String, _
    ByVal outputPostfix As String _
) As String
    Select Case mp_NormalizeOutputMode(outputMode)
        Case EXPORT_OUTPUT_MODE_OVERWRITE_TEMPLATE
            mp_BuildOutputPathByMode = templatePath
        Case EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX
            mp_BuildOutputPathByMode = mp_BuildPostfixedOutputPath(templatePath, outputPostfix)
        Case Else
            Err.Raise vbObjectError + 1770, "ex_WordPlaceholderReports", "Unsupported output mode: " & outputMode
    End Select
End Function

Private Function mp_BuildPostfixedOutputPath(ByVal templatePath As String, ByVal outputPostfix As String) As String
    Dim folderPath As String
    Dim fileNameOnly As String
    Dim dotPos As Long
    Dim baseName As String
    Dim extensionText As String

    folderPath = mp_ExtractFolderPath(templatePath)
    fileNameOnly = mp_ExtractFileName(templatePath)
    If Len(fileNameOnly) = 0 Then Exit Function

    outputPostfix = Trim$(outputPostfix)
    If Len(outputPostfix) = 0 Then outputPostfix = "_result"
    If InStr(1, outputPostfix, "\", vbBinaryCompare) > 0 Or InStr(1, outputPostfix, "/", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1771, "ex_WordPlaceholderReports", "Export.OutputPostfix must not contain path separators: " & outputPostfix
    End If

    dotPos = InStrRev(fileNameOnly, ".")
    If dotPos > 1 Then
        baseName = Left$(fileNameOnly, dotPos - 1)
        extensionText = Mid$(fileNameOnly, dotPos)
    Else
        baseName = fileNameOnly
        extensionText = vbNullString
    End If

    mp_BuildPostfixedOutputPath = folderPath & baseName & outputPostfix & extensionText
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

Private Function mp_NormalizeInsertMode(ByVal modeText As String) As String
    modeText = UCase$(Trim$(modeText))
    If Len(modeText) = 0 Then modeText = EXPORT_INSERT_MODE_REPLACE_ALL

    Select Case modeText
        Case EXPORT_INSERT_MODE_REPLACE_ALL, EXPORT_INSERT_MODE_APPEND_TOP, EXPORT_INSERT_MODE_APPEND_BOTTOM
            mp_NormalizeInsertMode = modeText
        Case Else
            Err.Raise vbObjectError + 1772, "ex_WordPlaceholderReports", "Invalid Export.InsertMode '" & modeText & "'. Expected: ReplaceAll, AppendToTop, AppendToBottom."
    End Select
End Function

Private Function mp_NormalizeOutputMode(ByVal modeText As String) As String
    modeText = UCase$(Trim$(modeText))
    If Len(modeText) = 0 Then modeText = EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX

    Select Case modeText
        Case EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX, EXPORT_OUTPUT_MODE_OVERWRITE_TEMPLATE
            mp_NormalizeOutputMode = modeText
        Case Else
            Err.Raise vbObjectError + 1773, "ex_WordPlaceholderReports", "Invalid Export.OutputMode '" & modeText & "'. Expected: CreateWithPostfix, OverwriteTemplate."
    End Select
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
