Attribute VB_Name = "ex_WordPlaceholderReports"
Option Explicit

Private Const WD_FIND_STOP As Long = 0
Private Const WD_COLLAPSE_END As Long = 0
Private Const WD_COLLAPSE_START As Long = 1
Private Const WD_ALERTS_NONE As Long = 0
Private Const EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX As String = "CREATEWITHPOSTFIX"
Private Const EXPORT_OUTPUT_MODE_OVERWRITE_TEMPLATE As String = "OVERWRITETEMPLATE"
Private Const EXPORT_INSERT_MODE_REPLACE_ALL As String = "REPLACEALL"
Private Const EXPORT_INSERT_MODE_APPEND_TOP As String = "APPENDTOTOP"
Private Const EXPORT_INSERT_MODE_APPEND_BOTTOM As String = "APPENDTOBOTTOM"
Private Const EXPORT_RUNTIME_WORD_RESULTS_PLACE As String = "Export.RuntimeDataBase.WordResultsPlace"
Private Const EXPORT_RUNTIME_WORD_PASTE_ANCHOR As String = "Export.RuntimeDataBase.WordPasteAnchor"
Private Const EXPORT_RUNTIME_WORD_ANCHOR_PREFIX As String = "Export.RuntimeDataBase.WordAnchor."
Private Const EXPORT_RUNTIME_WORD_RECORD_KEY_PREFIX As String = "Export.RuntimeDataBase.WordRecordKey."
Private Const EXPORT_APPEND_SEPARATOR As String = vbCrLf & vbCrLf
Private Const EXPORT_ANCHOR_MARKER_PREFIX As String = "{\export:"
Private Const EXPORT_ANCHOR_MARKER_BEGIN_SUFFIX As String = "_Begin}"
Private Const EXPORT_ANCHOR_MARKER_END_SUFFIX As String = "_End}"
Private Const EXPORT_BOOKMARK_PREFIX As String = "EP_Anchor_"
Private Const EXPORT_BOOKMARK_RECORD_PREFIX As String = "EP_Record_"
Private Const EXPORT_BOOKMARK_TOP_SUFFIX As String = "_Top"
Private Const EXPORT_BOOKMARK_BOTTOM_SUFFIX As String = "_Bottom"
Private Const EXPORT_BOOKMARK_MAX_LEN As Long = 40
Private Const EXPORT_DUPLICATE_CONFIRM_TITLE As String = "Duplicate Export Confirmation"
Private Const EXPORT_HASH_MODULO As Double = 2147483629#

Private g_WordApp As Object
Private g_WordAppOwnedByModule As Boolean

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
    Dim placeholderMapFromAnchors As Object
    Dim currentRecordKeySet As Object
    Dim hasMultiRuntimeAnchors As Boolean

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
    If mp_FileExists(outputPath) Then
        If mp_IsFileLocked(outputPath) Then
            MsgBox "Word export is blocked: target file is currently open." & vbCrLf & _
                   "Close the file and retry: " & outputPath, vbExclamation
            Exit Sub
        End If
    End If

    Set placeholderMapFromAnchors = mp_BuildPlaceholderMapFromRuntimeAnchors(ws)
    If Not placeholderMapFromAnchors Is Nothing Then
        If placeholderMapFromAnchors.Count > 0 Then
            hasMultiRuntimeAnchors = True
            Set placeholderMap = placeholderMapFromAnchors
        End If
    End If

    If Not hasMultiRuntimeAnchors Then
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
        Set placeholderMap = m_BuildPlaceholderMapFromPairs(wordPasteAnchor, sourceText)
    End If

    Set currentRecordKeySet = mp_BuildRecordKeySetFromRuntime(ws)
    If Not mp_ConfirmProceedForExport(templatePath, outputPath, outputMode, insertMode, currentRecordKeySet) Then
        ex_Messaging.m_ShowNotice "Export canceled.", 3
        Exit Sub
    End If

    reportPath = m_CreateWordReportFromTemplate(templatePath, outputPath, placeholderMap, True, insertMode, currentRecordKeySet)
    ex_Messaging.m_ShowNotice "Word report created: " & reportPath, 5
    Exit Sub

EH:
    MsgBox "Word footer export failed: " & Err.Description, vbExclamation
End Sub

Public Sub m_API_CleanupExportAnchorMarkers()
    Dim templatePath As String
    Dim outputPath As String
    Dim outputMode As String
    Dim outputPostfix As String
    Dim targetPath As String
    Dim removedCount As Long

    On Error GoTo EH

    templatePath = mp_ResolvePath(Trim$(ex_ConfigProvider.m_GetConfigValue("Export.WordTemplatePath", vbNullString)))
    If Len(templatePath) = 0 Then
        Err.Raise vbObjectError + 1787, "ex_WordPlaceholderReports", "Missing required config key 'Export.WordTemplatePath'."
    End If
    If Not mp_FileExists(templatePath) Then
        Err.Raise vbObjectError + 1788, "ex_WordPlaceholderReports", "Word template not found by config key 'Export.WordTemplatePath': " & templatePath
    End If

    outputMode = mp_NormalizeOutputMode(ex_ConfigProvider.m_GetConfigValue("Export.OutputMode", "CreateWithPostfix"))
    outputPostfix = CStr(ex_ConfigProvider.m_GetConfigValue("Export.OutputPostfix", "_result"))
    outputPath = mp_BuildOutputPathByMode(templatePath, outputMode, outputPostfix)
    targetPath = mp_BuildDuplicateCheckDocumentPath(templatePath, outputPath, outputMode)
    If Len(targetPath) = 0 Then
        Err.Raise vbObjectError + 1789, "ex_WordPlaceholderReports", "Unable to resolve target document path for export anchor cleanup."
    End If
    If Not mp_FileExists(targetPath) Then
        If StrComp(outputMode, EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX, vbTextCompare) = 0 Then
            targetPath = templatePath
        End If
    End If
    If Not mp_FileExists(targetPath) Then
        Err.Raise vbObjectError + 1790, "ex_WordPlaceholderReports", "Target document for export anchor cleanup was not found: " & targetPath
    End If

    removedCount = mp_RemoveExportAnchorMarkersFromDocumentPath(targetPath)
    ex_Messaging.m_ShowNotice "Export anchors cleanup finished. Removed markers: " & CStr(removedCount), 5
    Exit Sub

EH:
    MsgBox "Word export anchors cleanup failed: " & Err.Description, vbExclamation
End Sub

Private Function mp_BuildPlaceholderMapFromRuntimeAnchors(ByVal ws As Worksheet) As Object
    Dim runtimeEntries As Object
    Dim result As Object
    Dim dataKey As Variant
    Dim normalizedPrefix As String
    Dim anchorName As String
    Dim valueText As String
    Dim keyText As String

    If ws Is Nothing Then Exit Function

    Set runtimeEntries = ex_PostProcessActions.m_GetRuntimeDataEntriesByPrefix(EXPORT_RUNTIME_WORD_ANCHOR_PREFIX, ws)
    If runtimeEntries Is Nothing Then Exit Function
    If runtimeEntries.Count = 0 Then Exit Function

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1 ' vbTextCompare
    normalizedPrefix = LCase$(EXPORT_RUNTIME_WORD_ANCHOR_PREFIX)

    For Each dataKey In runtimeEntries.Keys
        keyText = LCase$(CStr(dataKey))
        If StrComp(Left$(keyText, Len(normalizedPrefix)), normalizedPrefix, vbBinaryCompare) <> 0 Then
            GoTo ContinueKey
        End If
        anchorName = Mid$(CStr(dataKey), Len(normalizedPrefix) + 1)
        If Len(Trim$(anchorName)) = 0 Then GoTo ContinueKey

        valueText = CStr(runtimeEntries(CStr(dataKey)))
        result(anchorName) = valueText
ContinueKey:
    Next dataKey

    If result.Count > 0 Then Set mp_BuildPlaceholderMapFromRuntimeAnchors = result
End Function

Private Function mp_BuildRecordKeySetFromRuntime(ByVal ws As Worksheet) As Object
    Dim runtimeEntries As Object
    Dim result As Object
    Dim dataKey As Variant
    Dim valueText As String
    Dim recordKey As String

    If ws Is Nothing Then Exit Function

    Set runtimeEntries = ex_PostProcessActions.m_GetRuntimeDataEntriesByPrefix(EXPORT_RUNTIME_WORD_RECORD_KEY_PREFIX, ws)
    If runtimeEntries Is Nothing Then Exit Function
    If runtimeEntries.Count = 0 Then Exit Function

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1 ' vbTextCompare

    For Each dataKey In runtimeEntries.Keys
        valueText = CStr(runtimeEntries(CStr(dataKey)))
        recordKey = valueText
        recordKey = mp_NormalizeRecordKey(recordKey)
        If Len(recordKey) = 0 Then GoTo ContinueKey
        result(recordKey) = "1"
ContinueKey:
    Next dataKey

    If result.Count > 0 Then Set mp_BuildRecordKeySetFromRuntime = result
End Function

Private Function mp_NormalizeRecordKey(ByVal recordKey As String) As String
    Dim delimiterPos As Long
    Dim sectionName As String
    Dim fioText As String

    recordKey = Replace$(recordKey, vbCr, " ")
    recordKey = Replace$(recordKey, vbLf, " ")
    recordKey = Replace$(recordKey, vbTab, " ")

    delimiterPos = InStr(1, recordKey, "|", vbBinaryCompare)
    If delimiterPos <= 0 Then Exit Function

    sectionName = mp_NormalizeRecordKeyPart(Left$(recordKey, delimiterPos - 1))
    fioText = mp_NormalizeRecordKeyPart(Mid$(recordKey, delimiterPos + 1))
    If Len(sectionName) = 0 Or Len(fioText) = 0 Then Exit Function

    mp_NormalizeRecordKey = sectionName & "|" & fioText
End Function

Private Function mp_NormalizeRecordKeyPart(ByVal valueText As String) As String
    valueText = Replace$(valueText, Chr$(160), " ")
    valueText = Trim$(LCase$(valueText))
    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
        valueText = Replace$(valueText, "  ", " ")
    Loop
    mp_NormalizeRecordKeyPart = valueText
End Function

Public Function m_CreateWordReportFromTemplate( _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal placeholderMap As Object, _
    Optional ByVal failIfPlaceholderMissing As Boolean = True, _
    Optional ByVal insertMode As String = "ReplaceAll", _
    Optional ByVal recordKeySet As Object = Nothing _
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

    If failIfPlaceholderMissing And Len(missingTokens) > 0 Then
        Err.Raise vbObjectError + 1756, "ex_WordPlaceholderReports", _
            "Export anchor(s) were not found in target document. Missing placeholders: " & missingTokens
    End If

    mp_UpsertRecordKeyBookmarks wdDoc, recordKeySet

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

Private Function mp_ConfirmProceedForExport( _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal outputMode As String, _
    ByVal insertMode As String, _
    ByVal currentRecordKeySet As Object _
) As Boolean
    Dim docPathForDuplicateCheck As String
    Dim duplicateRecordKeys As Object
    Dim messageText As String
    Dim response As VbMsgBoxResult
    Dim duplicatesPreview As String
    Dim duplicateCount As Long
    Dim totalCount As Long

    mp_ConfirmProceedForExport = True

    If currentRecordKeySet Is Nothing Then Exit Function
    If currentRecordKeySet.Count = 0 Then Exit Function
    If Not mp_ShouldCheckDuplicatesByBookmarks(insertMode) Then Exit Function

    docPathForDuplicateCheck = mp_BuildDuplicateCheckDocumentPath(templatePath, outputPath, outputMode)
    If Len(docPathForDuplicateCheck) = 0 Then Exit Function
    If Not mp_FileExists(docPathForDuplicateCheck) Then Exit Function

    Set duplicateRecordKeys = mp_FindDuplicateRecordKeysInDocument(docPathForDuplicateCheck, currentRecordKeySet)
    If duplicateRecordKeys Is Nothing Then Exit Function

    duplicateCount = duplicateRecordKeys.Count
    If duplicateCount <= 0 Then Exit Function

    totalCount = currentRecordKeySet.Count
    duplicatesPreview = mp_BuildDuplicateRecordPreview(duplicateRecordKeys, 5)

    messageText = "The current export has " & CStr(duplicateCount) & " duplicate record(s) out of " & CStr(totalCount) & _
                  " (matched by Section + FIO bookmarks)." & vbCrLf & vbCrLf & _
                  "Document: " & docPathForDuplicateCheck & vbCrLf
    If Len(duplicatesPreview) > 0 Then
        messageText = messageText & vbCrLf & "Examples:" & vbCrLf & duplicatesPreview
    End If
    messageText = messageText & vbCrLf & vbCrLf & "Do you want to export anyway?"

    response = MsgBox(messageText, vbQuestion + vbYesNo + vbDefaultButton2, EXPORT_DUPLICATE_CONFIRM_TITLE)
    mp_ConfirmProceedForExport = (response = vbYes)
End Function

Private Function mp_ShouldCheckDuplicatesByBookmarks(ByVal insertMode As String) As Boolean
    If StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_TOP, vbTextCompare) = 0 Then
        mp_ShouldCheckDuplicatesByBookmarks = True
    ElseIf StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_BOTTOM, vbTextCompare) = 0 Then
        mp_ShouldCheckDuplicatesByBookmarks = True
    End If
End Function

Private Function mp_BuildDuplicateCheckDocumentPath( _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal outputMode As String _
) As String
    Select Case mp_NormalizeOutputMode(outputMode)
        Case EXPORT_OUTPUT_MODE_OVERWRITE_TEMPLATE
            mp_BuildDuplicateCheckDocumentPath = Trim$(templatePath)
        Case EXPORT_OUTPUT_MODE_CREATE_WITH_POSTFIX
            mp_BuildDuplicateCheckDocumentPath = Trim$(outputPath)
    End Select
End Function

Private Function mp_FindDuplicateRecordKeysInDocument(ByVal documentPath As String, ByVal currentRecordKeySet As Object) As Object
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim result As Object
    Dim recordKey As Variant
    Dim bookmarkName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1 ' vbTextCompare

    If currentRecordKeySet Is Nothing Then
        Set mp_FindDuplicateRecordKeysInDocument = result
        Exit Function
    End If
    If currentRecordKeySet.Count = 0 Then
        Set mp_FindDuplicateRecordKeysInDocument = result
        Exit Function
    End If

    On Error GoTo EH
    Set wdApp = mp_GetOrCreateWordApp()
    Set wdDoc = wdApp.Documents.Open(documentPath, False, True)

    For Each recordKey In currentRecordKeySet.Keys
        bookmarkName = mp_BuildRecordBookmarkName(CStr(recordKey))
        If Len(bookmarkName) > 0 Then
            If wdDoc.Bookmarks.Exists(bookmarkName) Then
                result(CStr(recordKey)) = bookmarkName
            End If
        End If
    Next recordKey

    On Error Resume Next
    wdDoc.Close False
    On Error GoTo 0

    Set mp_FindDuplicateRecordKeysInDocument = result
    Exit Function

EH:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 1782, "ex_WordPlaceholderReports", _
        "Failed to check duplicate record bookmarks in document '" & documentPath & "': " & Err.Description
End Function

Private Function mp_BuildDuplicateRecordPreview(ByVal duplicateRecordKeys As Object, ByVal maxItems As Long) As String
    Dim key As Variant
    Dim i As Long

    If duplicateRecordKeys Is Nothing Then Exit Function
    If duplicateRecordKeys.Count = 0 Then Exit Function
    If maxItems <= 0 Then maxItems = 1

    For Each key In duplicateRecordKeys.Keys
        mp_BuildDuplicateRecordPreview = mp_BuildDuplicateRecordPreview & " - " & CStr(key) & vbCrLf
        i = i + 1
        If i >= maxItems Then Exit For
    Next key

    If duplicateRecordKeys.Count > i Then
        mp_BuildDuplicateRecordPreview = mp_BuildDuplicateRecordPreview & " - ... (" & CStr(duplicateRecordKeys.Count - i) & " more)"
    ElseIf Len(mp_BuildDuplicateRecordPreview) > 0 Then
        mp_BuildDuplicateRecordPreview = Left$(mp_BuildDuplicateRecordPreview, Len(mp_BuildDuplicateRecordPreview) - Len(vbCrLf))
    End If
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
    Dim anchorName As String
    Dim beginMarker As String
    Dim endMarker As String
    Dim beginMarkerRange As Object
    Dim endMarkerRange As Object
    Dim contentBetweenMarkers As Object
    Dim insertionRange As Object
    Dim appendSeparatorText As String
    Dim insertionText As String
    Dim hasMeaningfulContentBetweenMarkers As Boolean
    Dim insertionPos As Long
    Dim charAfterBegin As String
    Dim charBeforeEnd As String
    Dim needPrefixLineBreak As Boolean
    Dim needSuffixLineBreak As Boolean

    replacementText = mp_NormalizeTextForWord(replacementText)
    appendSeparatorText = mp_NormalizeTextForWord(EXPORT_APPEND_SEPARATOR)

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

    anchorName = mp_ExtractAnchorNameFromToken(token)
    If Len(anchorName) = 0 Then Exit Function

    beginMarker = mp_BuildExplicitAnchorMarker(anchorName, True)
    endMarker = mp_BuildExplicitAnchorMarker(anchorName, False)
    If Not mp_TryFindExplicitAnchorMarkerPair(doc, beginMarker, endMarker, beginMarkerRange, endMarkerRange) Then Exit Function
    mp_EnsureExplicitAnchorGapForSingleLineBreak doc, beginMarker, endMarker, beginMarkerRange, endMarkerRange

    Set contentBetweenMarkers = doc.Range(beginMarkerRange.End, endMarkerRange.Start)
    hasMeaningfulContentBetweenMarkers = mp_RangeHasMeaningfulText(contentBetweenMarkers)

    If StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_TOP, vbTextCompare) = 0 Then
        insertionPos = beginMarkerRange.End
        charAfterBegin = mp_GetDocumentCharAt(doc, insertionPos)
        If StrComp(charAfterBegin, vbCr, vbBinaryCompare) = 0 Then
            insertionPos = insertionPos + 1
        Else
            needPrefixLineBreak = True
        End If

        insertionText = replacementText
        If needPrefixLineBreak Then
            insertionText = vbCr & insertionText
        End If
        If hasMeaningfulContentBetweenMarkers Then
            insertionText = insertionText & appendSeparatorText
        End If
        Set insertionRange = doc.Range(insertionPos, insertionPos)
        insertionRange.Text = insertionText
    ElseIf StrComp(insertMode, EXPORT_INSERT_MODE_APPEND_BOTTOM, vbTextCompare) = 0 Then
        insertionPos = endMarkerRange.Start
        If insertionPos > 0 Then
            charBeforeEnd = mp_GetDocumentCharAt(doc, insertionPos - 1)
            If StrComp(charBeforeEnd, vbCr, vbBinaryCompare) = 0 Then
                insertionPos = insertionPos - 1
            Else
                needSuffixLineBreak = True
            End If
        Else
            needSuffixLineBreak = True
        End If

        insertionText = replacementText
        If hasMeaningfulContentBetweenMarkers Then
            insertionText = appendSeparatorText & insertionText
        End If
        If needSuffixLineBreak Then
            insertionText = insertionText & vbCr
        End If
        Set insertionRange = doc.Range(insertionPos, insertionPos)
        insertionRange.Text = insertionText
    Else
        Err.Raise vbObjectError + 1769, "ex_WordPlaceholderReports", "Unsupported insert mode: " & insertMode
    End If

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

Private Function mp_ExtractAnchorNameFromToken(ByVal tokenText As String) As String
    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then Exit Function

    If Left$(tokenText, 1) = "{" And Right$(tokenText, 1) = "}" Then
        tokenText = Mid$(tokenText, 2, Len(tokenText) - 2)
    End If

    mp_ExtractAnchorNameFromToken = Trim$(tokenText)
End Function

Private Function mp_BuildExplicitAnchorMarker(ByVal anchorName As String, ByVal isBeginMarker As Boolean) As String
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Function

    If isBeginMarker Then
        mp_BuildExplicitAnchorMarker = EXPORT_ANCHOR_MARKER_PREFIX & anchorName & EXPORT_ANCHOR_MARKER_BEGIN_SUFFIX
    Else
        mp_BuildExplicitAnchorMarker = EXPORT_ANCHOR_MARKER_PREFIX & anchorName & EXPORT_ANCHOR_MARKER_END_SUFFIX
    End If
End Function

Private Function mp_TryFindExplicitAnchorMarkerPair( _
    ByVal doc As Object, _
    ByVal beginMarker As String, _
    ByVal endMarker As String, _
    ByRef outBeginMarkerRange As Object, _
    ByRef outEndMarkerRange As Object _
) As Boolean
    Dim story As Object
    Dim currentRange As Object
    Dim rangeAfterBeginMarker As Object

    If doc Is Nothing Then Exit Function
    If Len(beginMarker) = 0 Or Len(endMarker) = 0 Then Exit Function

    For Each story In doc.StoryRanges
        Set currentRange = story
        Do While Not currentRange Is Nothing
            If mp_TryFindTextInRange(currentRange, beginMarker, outBeginMarkerRange) Then
                Set rangeAfterBeginMarker = currentRange.Duplicate
                rangeAfterBeginMarker.Start = outBeginMarkerRange.End
                If rangeAfterBeginMarker.Start <= rangeAfterBeginMarker.End Then
                    If mp_TryFindTextInRange(rangeAfterBeginMarker, endMarker, outEndMarkerRange) Then
                        mp_TryFindExplicitAnchorMarkerPair = True
                        Exit Function
                    End If
                End If
            End If
            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story
End Function

Private Function mp_TryFindTextInRange(ByVal sourceRange As Object, ByVal targetText As String, ByRef outFoundRange As Object) As Boolean
    Dim findRange As Object

    If sourceRange Is Nothing Then Exit Function
    If Len(targetText) = 0 Then Exit Function

    Set findRange = sourceRange.Duplicate
    With findRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = targetText
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

    If findRange.Find.Execute Then
        Set outFoundRange = findRange.Duplicate
        mp_TryFindTextInRange = True
    End If
End Function

Private Sub mp_EnsureExplicitAnchorGapForSingleLineBreak( _
    ByVal doc As Object, _
    ByVal beginMarker As String, _
    ByVal endMarker As String, _
    ByRef beginMarkerRange As Object, _
    ByRef endMarkerRange As Object _
)
    Dim betweenRange As Object
    Dim betweenText As String

    If doc Is Nothing Then Exit Sub
    If beginMarkerRange Is Nothing Then Exit Sub
    If endMarkerRange Is Nothing Then Exit Sub

    Set betweenRange = doc.Range(beginMarkerRange.End, endMarkerRange.Start)
    betweenText = CStr(betweenRange.Text)
    If mp_IsOnlySingleLineBreakBetweenAnchors(betweenText) Then
        betweenRange.Text = vbCr & vbCr
        mp_TryFindExplicitAnchorMarkerPair doc, beginMarker, endMarker, beginMarkerRange, endMarkerRange
    End If
End Sub

Private Function mp_IsOnlySingleLineBreakBetweenAnchors(ByVal betweenText As String) As Boolean
    betweenText = Replace$(betweenText, vbTab, vbNullString)
    betweenText = Replace$(betweenText, " ", vbNullString)
    betweenText = Replace$(betweenText, Chr$(160), vbNullString)
    betweenText = Replace$(betweenText, vbLf, vbNullString)

    mp_IsOnlySingleLineBreakBetweenAnchors = (StrComp(betweenText, vbCr, vbBinaryCompare) = 0)
End Function

Private Function mp_RangeHasMeaningfulText(ByVal sourceRange As Object) As Boolean
    Dim textValue As String

    If sourceRange Is Nothing Then Exit Function

    textValue = CStr(sourceRange.Text)
    textValue = Replace$(textValue, vbCr, vbNullString)
    textValue = Replace$(textValue, vbLf, vbNullString)
    textValue = Replace$(textValue, vbTab, vbNullString)
    textValue = Replace$(textValue, Chr$(160), vbNullString)
    textValue = Replace$(textValue, " ", vbNullString)

    mp_RangeHasMeaningfulText = (Len(textValue) > 0)
End Function

Private Function mp_GetDocumentCharAt(ByVal doc As Object, ByVal charPos As Long) As String
    Dim docEnd As Long
    Dim charRange As Object

    If doc Is Nothing Then Exit Function
    If charPos < 0 Then Exit Function

    On Error GoTo CleanFail
    docEnd = CLng(doc.Content.End)
    If charPos >= docEnd Then Exit Function

    Set charRange = doc.Range(charPos, charPos + 1)
    mp_GetDocumentCharAt = CStr(charRange.Text)
    Exit Function

CleanFail:
    mp_GetDocumentCharAt = vbNullString
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

Private Sub mp_TrimInsertedRangeForAppendTop(ByVal targetRange As Object, ByVal separatorLength As Long)
    If targetRange Is Nothing Then Exit Sub
    If separatorLength <= 0 Then Exit Sub

    On Error Resume Next
    targetRange.End = targetRange.End - separatorLength
    If targetRange.End < targetRange.Start Then targetRange.End = targetRange.Start
    On Error GoTo 0
End Sub

Private Sub mp_TrimInsertedRangeForAppendBottom(ByVal targetRange As Object, ByVal separatorLength As Long)
    If targetRange Is Nothing Then Exit Sub
    If separatorLength <= 0 Then Exit Sub

    On Error Resume Next
    targetRange.Start = targetRange.Start + separatorLength
    If targetRange.End < targetRange.Start Then targetRange.Start = targetRange.End
    On Error GoTo 0
End Sub

Private Function mp_NormalizeTextForWord(ByVal sourceText As String) As String
    sourceText = Replace$(sourceText, vbCrLf, vbCr)
    sourceText = Replace$(sourceText, vbLf, vbCr)
    mp_NormalizeTextForWord = sourceText
End Function

Private Function mp_CreateRangeSafe(ByVal doc As Object, ByVal startPos As Long, ByVal endPos As Long) As Object
    Dim docEnd As Long

    If doc Is Nothing Then Exit Function

    On Error GoTo CleanFail
    docEnd = CLng(doc.Content.End)
    If startPos < 0 Then startPos = 0
    If endPos < startPos Then endPos = startPos
    If startPos > docEnd Then startPos = docEnd
    If endPos > docEnd Then endPos = docEnd

    Set mp_CreateRangeSafe = doc.Range(startPos, endPos)
    Exit Function

CleanFail:
    Set mp_CreateRangeSafe = Nothing
End Function

Private Sub mp_UpsertAnchorBookmark(ByVal doc As Object, ByVal token As String, ByVal targetRange As Object)
    If doc Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub

    mp_UpsertAnchorBookmarkPair doc, token, targetRange, targetRange
End Sub

Private Sub mp_UpsertAnchorBookmarkPair(ByVal doc As Object, ByVal token As String, ByVal topRange As Object, ByVal bottomRange As Object)
    If doc Is Nothing Then Exit Sub
    If topRange Is Nothing And bottomRange Is Nothing Then Exit Sub

    If topRange Is Nothing Then Set topRange = bottomRange.Duplicate
    If bottomRange Is Nothing Then Set bottomRange = topRange.Duplicate

    mp_UpsertAnchorBookmarkByRole doc, token, True, topRange
    mp_UpsertAnchorBookmarkByRole doc, token, False, bottomRange
End Sub

Private Sub mp_UpsertAnchorBookmarkByRole(ByVal doc As Object, ByVal token As String, ByVal isTopRole As Boolean, ByVal targetRange As Object)
    Dim bookmarkName As String

    If doc Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub

    If isTopRole Then
        bookmarkName = mp_BuildAnchorBookmarkTopName(token)
    Else
        bookmarkName = mp_BuildAnchorBookmarkBottomName(token)
    End If
    If Len(bookmarkName) = 0 Then Exit Sub

    mp_UpsertBookmarkByName doc, bookmarkName, targetRange
End Sub

Private Sub mp_UpsertBookmarkByName(ByVal doc As Object, ByVal bookmarkName As String, ByVal targetRange As Object)
    If doc Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    If Len(Trim$(bookmarkName)) = 0 Then Exit Sub

    On Error Resume Next
    If doc.Bookmarks.Exists(bookmarkName) Then
        doc.Bookmarks(bookmarkName).Delete
    End If
    doc.Bookmarks.Add bookmarkName, targetRange
    On Error GoTo 0
End Sub

Private Sub mp_UpsertRecordKeyBookmarks(ByVal doc As Object, ByVal recordKeySet As Object)
    Dim markerRange As Object
    Dim recordKey As Variant
    Dim bookmarkName As String

    If doc Is Nothing Then Exit Sub
    If recordKeySet Is Nothing Then Exit Sub
    If recordKeySet.Count = 0 Then Exit Sub

    Set markerRange = doc.Content.Duplicate
    markerRange.Collapse WD_COLLAPSE_END

    For Each recordKey In recordKeySet.Keys
        bookmarkName = mp_BuildRecordBookmarkName(CStr(recordKey))
        If Len(bookmarkName) > 0 Then
            mp_UpsertBookmarkByName doc, bookmarkName, markerRange
        End If
    Next recordKey
End Sub

Private Function mp_BuildRecordBookmarkName(ByVal recordKey As String) As String
    Dim normalizedRecordKey As String
    Dim hashA As String
    Dim hashB As String

    normalizedRecordKey = mp_NormalizeRecordKey(recordKey)
    If Len(normalizedRecordKey) = 0 Then Exit Function

    hashA = mp_ComputeTextHashHex(normalizedRecordKey)
    hashB = mp_ComputeTextHashHex("record|" & normalizedRecordKey)
    mp_BuildRecordBookmarkName = EXPORT_BOOKMARK_RECORD_PREFIX & hashA & hashB
End Function

Private Function mp_RemoveExportAnchorMarkersFromDocumentPath(ByVal documentPath As String) As Long
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim previousScreenUpdating As Boolean

    On Error GoTo EH

    Set wdApp = mp_GetOrCreateWordApp()
    On Error Resume Next
    previousScreenUpdating = CBool(wdApp.ScreenUpdating)
    wdApp.ScreenUpdating = False
    On Error GoTo EH

    Set wdDoc = wdApp.Documents.Open(documentPath, False, False)
    mp_RemoveExportAnchorMarkersFromDocumentPath = mp_RemoveExportAnchorMarkersFromDocument(wdDoc)
    wdDoc.Save

    On Error Resume Next
    wdDoc.Close False
    wdApp.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Function

EH:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Err.Raise vbObjectError + 1791, "ex_WordPlaceholderReports", _
        "Failed to cleanup export anchor markers in document '" & documentPath & "': " & Err.Description
End Function

Private Function mp_RemoveExportAnchorMarkersFromDocument(ByVal doc As Object) As Long
    Dim mainRange As Object

    If doc Is Nothing Then Exit Function

    ' Export anchors are generated in the main document body.
    Set mainRange = doc.Content
    mp_RemoveExportAnchorMarkersFromDocument = mp_RemoveExportAnchorMarkersInRange(doc, mainRange)
End Function

Private Function mp_RemoveExportAnchorMarkersInRange(ByVal doc As Object, ByVal sourceRange As Object) As Long
    Dim sourceText As String
    Dim prefixText As String
    Dim scanPos As Long
    Dim markerEndPos As Long
    Dim tokenText As String
    Dim tokenRemoveStartPos As Long
    Dim tokenRemoveEndPos As Long
    Dim markerCount As Long
    Dim mergedCount As Long
    Dim deleteStart() As Long
    Dim deleteEnd() As Long
    Dim i As Long
    Dim baseStartPos As Long
    Dim absDeleteStart As Long
    Dim absDeleteEnd As Long

    If doc Is Nothing Then Exit Function
    If sourceRange Is Nothing Then Exit Function

    sourceText = CStr(sourceRange.Text)
    If Len(sourceText) = 0 Then Exit Function

    prefixText = EXPORT_ANCHOR_MARKER_PREFIX
    scanPos = InStr(1, sourceText, prefixText, vbTextCompare)
    Do While scanPos > 0
        markerEndPos = InStr(scanPos + Len(prefixText), sourceText, "}", vbBinaryCompare)
        If markerEndPos <= 0 Then Exit Do

        tokenText = Mid$(sourceText, scanPos, markerEndPos - scanPos + 1)
        If mp_IsExplicitExportMarkerToken(tokenText) Then
            tokenRemoveStartPos = scanPos
            tokenRemoveEndPos = markerEndPos

            Do While tokenRemoveEndPos < Len(sourceText)
                If StrComp(Mid$(sourceText, tokenRemoveEndPos + 1, 1), "}", vbBinaryCompare) <> 0 Then Exit Do
                tokenRemoveEndPos = tokenRemoveEndPos + 1
            Loop

            ' Remove the marker line itself: marker token + trailing line break (if present).
            If tokenRemoveEndPos < Len(sourceText) Then
                If StrComp(Mid$(sourceText, tokenRemoveEndPos + 1, 1), vbCr, vbBinaryCompare) = 0 Then
                    tokenRemoveEndPos = tokenRemoveEndPos + 1
                End If
            End If

            markerCount = markerCount + 1
            mp_AddDeleteInterval deleteStart, deleteEnd, mergedCount, tokenRemoveStartPos, tokenRemoveEndPos
        End If

        scanPos = InStr(markerEndPos + 1, sourceText, prefixText, vbTextCompare)
    Loop

    If markerCount <= 0 Then Exit Function
    If mergedCount <= 0 Then Exit Function

    mp_MergeDeleteIntervals deleteStart, deleteEnd, mergedCount
    baseStartPos = CLng(sourceRange.Start)
    For i = mergedCount To 1 Step -1
        absDeleteStart = baseStartPos + deleteStart(i) - 1
        absDeleteEnd = baseStartPos + deleteEnd(i)

        Do While StrComp(mp_GetDocumentCharAt(doc, absDeleteEnd), "}", vbBinaryCompare) = 0
            absDeleteEnd = absDeleteEnd + 1
        Loop

        doc.Range(absDeleteStart, absDeleteEnd).Text = vbNullString
    Next i

    mp_RemoveExportAnchorMarkersInRange = markerCount
End Function

Private Sub mp_AddDeleteInterval( _
    ByRef deleteStart() As Long, _
    ByRef deleteEnd() As Long, _
    ByRef intervalCount As Long, _
    ByVal startPos As Long, _
    ByVal endPos As Long _
)
    If endPos < startPos Then Exit Sub

    intervalCount = intervalCount + 1
    If intervalCount = 1 Then
        ReDim deleteStart(1 To 1)
        ReDim deleteEnd(1 To 1)
    Else
        ReDim Preserve deleteStart(1 To intervalCount)
        ReDim Preserve deleteEnd(1 To intervalCount)
    End If

    deleteStart(intervalCount) = startPos
    deleteEnd(intervalCount) = endPos
End Sub

Private Sub mp_MergeDeleteIntervals( _
    ByRef deleteStart() As Long, _
    ByRef deleteEnd() As Long, _
    ByRef intervalCount As Long _
)
    Dim i As Long
    Dim writePos As Long

    If intervalCount <= 1 Then Exit Sub

    writePos = 1
    For i = 2 To intervalCount
        If deleteStart(i) <= deleteEnd(writePos) + 1 Then
            If deleteEnd(i) > deleteEnd(writePos) Then
                deleteEnd(writePos) = deleteEnd(i)
            End If
        Else
            writePos = writePos + 1
            deleteStart(writePos) = deleteStart(i)
            deleteEnd(writePos) = deleteEnd(i)
        End If
    Next i

    intervalCount = writePos
    ReDim Preserve deleteStart(1 To intervalCount)
    ReDim Preserve deleteEnd(1 To intervalCount)
End Sub

Private Function mp_IsExplicitExportMarkerToken(ByVal candidateToken As String) As Boolean
    Dim tokenLower As String

    candidateToken = Trim$(candidateToken)
    If Len(candidateToken) = 0 Then Exit Function

    tokenLower = LCase$(candidateToken)
    If Left$(tokenLower, Len(LCase$(EXPORT_ANCHOR_MARKER_PREFIX))) <> LCase$(EXPORT_ANCHOR_MARKER_PREFIX) Then Exit Function
    If Right$(tokenLower, Len(LCase$(EXPORT_ANCHOR_MARKER_BEGIN_SUFFIX))) = LCase$(EXPORT_ANCHOR_MARKER_BEGIN_SUFFIX) Then
        mp_IsExplicitExportMarkerToken = True
        Exit Function
    End If
    If Right$(tokenLower, Len(LCase$(EXPORT_ANCHOR_MARKER_END_SUFFIX))) = LCase$(EXPORT_ANCHOR_MARKER_END_SUFFIX) Then
        mp_IsExplicitExportMarkerToken = True
    End If
End Function

Private Function mp_TryGetAnchorBookmarkRange(ByVal doc As Object, ByVal token As String, ByRef outRange As Object) As Boolean
    Dim topRange As Object
    Dim bottomRange As Object

    If mp_TryGetAnchorBookmarkPair(doc, token, topRange, bottomRange) Then
        Set outRange = topRange
        mp_TryGetAnchorBookmarkRange = Not (outRange Is Nothing)
    End If
End Function

Private Function mp_TryGetAnchorBookmarkPair(ByVal doc As Object, ByVal token As String, ByRef outTopRange As Object, ByRef outBottomRange As Object) As Boolean
    Dim topBookmarkName As String
    Dim bottomBookmarkName As String
    Dim legacyBookmarkName As String
    Dim hasTop As Boolean
    Dim hasBottom As Boolean
    Dim legacyRange As Object

    If doc Is Nothing Then Exit Function

    topBookmarkName = mp_BuildAnchorBookmarkTopName(token)
    bottomBookmarkName = mp_BuildAnchorBookmarkBottomName(token)
    legacyBookmarkName = mp_BuildAnchorBookmarkName(token)

    hasTop = mp_TryGetBookmarkRangeByName(doc, topBookmarkName, outTopRange)
    hasBottom = mp_TryGetBookmarkRangeByName(doc, bottomBookmarkName, outBottomRange)

    If hasTop And hasBottom Then
        mp_TryGetAnchorBookmarkPair = True
        Exit Function
    End If

    If hasTop And Not hasBottom Then
        Set outBottomRange = outTopRange.Duplicate
        mp_UpsertBookmarkByName doc, bottomBookmarkName, outBottomRange
        mp_TryGetAnchorBookmarkPair = True
        Exit Function
    End If

    If hasBottom And Not hasTop Then
        Set outTopRange = outBottomRange.Duplicate
        mp_UpsertBookmarkByName doc, topBookmarkName, outTopRange
        mp_TryGetAnchorBookmarkPair = True
        Exit Function
    End If

    If mp_TryGetBookmarkRangeByName(doc, legacyBookmarkName, legacyRange) Then
        Set outTopRange = legacyRange.Duplicate
        Set outBottomRange = legacyRange.Duplicate
        mp_UpsertBookmarkByName doc, topBookmarkName, outTopRange
        mp_UpsertBookmarkByName doc, bottomBookmarkName, outBottomRange
        mp_TryGetAnchorBookmarkPair = True
    End If
End Function

Private Function mp_TryGetBookmarkRangeByName(ByVal doc As Object, ByVal bookmarkName As String, ByRef outRange As Object) As Boolean
    If doc Is Nothing Then Exit Function
    bookmarkName = Trim$(bookmarkName)
    If Len(bookmarkName) = 0 Then Exit Function

    On Error GoTo CleanFail
    If doc.Bookmarks.Exists(bookmarkName) Then
        Set outRange = doc.Bookmarks(bookmarkName).Range
        mp_TryGetBookmarkRangeByName = Not (outRange Is Nothing)
    End If
    Exit Function

CleanFail:
    Set outRange = Nothing
    mp_TryGetBookmarkRangeByName = False
End Function

Private Function mp_BuildAnchorBookmarkTopName(ByVal token As String) As String
    mp_BuildAnchorBookmarkTopName = mp_BuildAnchorBookmarkName(token, EXPORT_BOOKMARK_TOP_SUFFIX)
End Function

Private Function mp_BuildAnchorBookmarkBottomName(ByVal token As String) As String
    mp_BuildAnchorBookmarkBottomName = mp_BuildAnchorBookmarkName(token, EXPORT_BOOKMARK_BOTTOM_SUFFIX)
End Function

Private Function mp_BuildAnchorBookmarkName(ByVal token As String, Optional ByVal roleSuffix As String = vbNullString) As String
    Dim rawToken As String
    Dim cleanToken As String
    Dim i As Long
    Dim ch As String
    Dim resultName As String
    Dim maxBaseLen As Long

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

    roleSuffix = CStr(roleSuffix)
    maxBaseLen = EXPORT_BOOKMARK_MAX_LEN - Len(roleSuffix)
    If maxBaseLen < Len(EXPORT_BOOKMARK_PREFIX) Then
        maxBaseLen = Len(EXPORT_BOOKMARK_PREFIX)
    End If

    resultName = EXPORT_BOOKMARK_PREFIX & cleanToken
    If Len(resultName) > maxBaseLen Then
        resultName = Left$(resultName, maxBaseLen)
    End If
    mp_BuildAnchorBookmarkName = resultName & roleSuffix
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
    Dim nameText As String
    Dim pointerRange As Range
    Dim namedEntry As Name
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
        ElseIf StrComp(prefixText, "NAME", vbBinaryCompare) = 0 Then
            nameText = Trim$(Mid$(normalizedPointer, colonPos + 1))
        Else
            addressText = normalizedPointer
        End If
    Else
        addressText = normalizedPointer
    End If

    If Len(nameText) > 0 Then
        On Error Resume Next
        Set namedEntry = ws.Names(nameText)
        On Error GoTo ResolveErr
        If namedEntry Is Nothing Then
            Err.Raise vbObjectError + 1780, "ex_WordPlaceholderReports", "Unable to resolve runtime source pointer '" & normalizedPointer & "' on sheet '" & ws.Name & "'."
        End If
        On Error Resume Next
        Set pointerRange = namedEntry.RefersToRange
        On Error GoTo ResolveErr
        If pointerRange Is Nothing Then
            Err.Raise vbObjectError + 1780, "ex_WordPlaceholderReports", "Unable to resolve runtime source pointer '" & normalizedPointer & "' on sheet '" & ws.Name & "'."
        End If
        mp_ReadSheetTextByRuntimePointer = CStr(pointerRange.Cells(1, 1).Value)
        Exit Function
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

Private Function mp_IsFileLocked(ByVal filePath As String) As Boolean
    Dim fileHandle As Integer

    If Len(Trim$(filePath)) = 0 Then Exit Function
    If Not mp_FileExists(filePath) Then Exit Function

    On Error GoTo Locked
    fileHandle = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #fileHandle
    Close #fileHandle
    Exit Function

Locked:
    mp_IsFileLocked = True
    On Error Resume Next
    If fileHandle > 0 Then Close #fileHandle
    On Error GoTo 0
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
