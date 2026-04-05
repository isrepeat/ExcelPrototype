Attribute VB_Name = "ex_Messaging"
Option Explicit

Private g_StatusClearTime As Date
Private g_StatusClearScheduled As Boolean
Private g_StatusClearProcedureName As String
Private g_CloseUntil As Date

Private Const STATUS_STORE_APP As String = "ExcelPrototype"
Private Const STATUS_STORE_SECTION_PREFIX As String = "ex_Messaging_"
Private Const STATUS_STORE_KEY_CLEAR_AT As String = "status_clear_at"
Private Const STATUS_STORE_KEY_CLEAR_PROC As String = "status_clear_proc"
Private Const STATUS_STORE_KEY_CLOSE_UNTIL As String = "close_until"
Private Const STATUS_CLOSE_HOLD_SECONDS As Double = 15#
Private Const DEFAULT_LOG_FILE_RELATIVE_PATH As String = "Logs\personalcard_pipeline.log"
Private Const SETTINGS_KEY_FILE_LOG_ENABLED As String = "st_FileLogEnabled"
Private Const BANNER_STYLE_STAGE_NAME As String = "banners"
Private Const BANNER_KIND_ERROR As String = "errorbanner"
Private Const BANNER_KIND_WARNING As String = "warningbanner"
Private Const STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP As String = "customautoheight-margin-top"
Private Const STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM As String = "customautoheight-margin-bottom"
Private Const POST_PROCESS_HEADER_ANCHOR_NAME As String = "__pcPostProcessSingleHeader"
Private Const BANNER_ANCHOR_PREFIX As String = "__pcBanner_"
Private Const BANNER_MESSAGE_ANCHOR_PREFIX As String = "__pcBannerMsg_"
Private Const TABLE_ANCHOR_PREFIX As String = "__pcTable_"
Private Const ROW_ANCHOR_PREFIX As String = "__pcRow_"
Private Const BANNER_MAX_ANCHOR_INDEX As Long = 9999
Private Const DEFAULT_BANNER_COLUMNS As Long = 8
Private Const DEFAULT_BANNER_ROWS As Long = 3

' =============================================================================
' Status bar notification
' =============================================================================

Public Sub m_ShowNotice(ByVal msg As String, Optional ByVal seconds As Double = 3)
    Dim procedureName As String

    If seconds <= 0 Then seconds = 3
    If mp_IsClosingActive() Then Exit Sub

    Application.StatusBar = msg
    mp_CancelPendingStatusClearCore True

    procedureName = mp_GetStatusClearProcedureName()
    On Error Resume Next
    g_StatusClearTime = Now + (seconds / 86400#)
    Application.OnTime EarliestTime:=g_StatusClearTime, Procedure:=procedureName, Schedule:=True
    g_StatusClearScheduled = (Err.Number = 0)
    If g_StatusClearScheduled Then
        mp_SavePersistedSchedule g_StatusClearTime, procedureName
    Else
        g_StatusClearTime = 0
        mp_ClearPersistedSchedule
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub m_ClearStatusBar()
    ' Очищает статус бар
    g_StatusClearScheduled = False
    g_StatusClearTime = 0
    mp_ClearPersistedSchedule
    Application.StatusBar = False
End Sub

Public Sub m_CancelPendingStatusClear()
    mp_CancelPendingStatusClearCore True
End Sub

Public Sub m_BeginWorkbookClose(Optional ByVal holdSeconds As Double = STATUS_CLOSE_HOLD_SECONDS)
    If holdSeconds <= 0 Then holdSeconds = STATUS_CLOSE_HOLD_SECONDS
    g_CloseUntil = Now + (holdSeconds / 86400#)
    mp_SaveStoreValue STATUS_STORE_KEY_CLOSE_UNTIL, CStr(CDbl(g_CloseUntil))
End Sub

Public Sub m_EndWorkbookClose()
    g_CloseUntil = 0
    mp_DeleteStoreValue STATUS_STORE_KEY_CLOSE_UNTIL
End Sub

' =============================================================================
' File logging (simple runtime diagnostics)
' =============================================================================

Public Sub m_ClearLogFile(Optional ByVal relativeOrAbsolutePath As String = DEFAULT_LOG_FILE_RELATIVE_PATH)
    Dim logPath As String
    Dim fileNo As Integer

    logPath = mp_ResolveLogFilePath(relativeOrAbsolutePath)
    mp_EnsureLogParentFolder logPath

    fileNo = FreeFile
    Open logPath For Output As #fileNo
    Close #fileNo
End Sub

Public Function m_LogToFile( _
    ByVal messageText As String, _
    Optional ByVal relativeOrAbsolutePath As String = DEFAULT_LOG_FILE_RELATIVE_PATH _
) As String
    Dim logPath As String
    Dim fileNo As Integer
    Dim lineText As String

    logPath = mp_ResolveLogFilePath(relativeOrAbsolutePath)
    If Not mp_IsFileLogEnabled() Then
        m_LogToFile = logPath
        Exit Function
    End If

    mp_EnsureLogParentFolder logPath

    lineText = Format$(Now, "yyyy-mm-dd HH:nn:ss") & " | " & CStr(messageText)
    fileNo = FreeFile
    Open logPath For Append As #fileNo
    Print #fileNo, lineText
    Close #fileNo

    m_LogToFile = logPath
End Function

Private Function mp_IsFileLogEnabled() As Boolean
    Dim rawValue As String

    rawValue = LCase$(Trim$(ex_XmlCore.m_GetSettingsValue(SETTINGS_KEY_FILE_LOG_ENABLED, "false")))
    Select Case rawValue
        Case "1", "true", "yes", "on"
            mp_IsFileLogEnabled = True
        Case Else
            mp_IsFileLogEnabled = False
    End Select
End Function

Private Function mp_GetStatusClearProcedureName() As String
    If Len(g_StatusClearProcedureName) = 0 Then
        g_StatusClearProcedureName = "'" & ThisWorkbook.Name & "'!ex_Messaging.m_ClearStatusBar"
    End If
    mp_GetStatusClearProcedureName = g_StatusClearProcedureName
End Function

Private Sub mp_CancelPendingStatusClearCore(Optional ByVal clearPersistedState As Boolean = True)
    Dim persistedTime As Date
    Dim persistedProcedure As String

    On Error Resume Next
    If g_StatusClearScheduled And g_StatusClearTime > 0 Then
        Application.OnTime EarliestTime:=g_StatusClearTime, Procedure:=mp_GetStatusClearProcedureName(), Schedule:=False
    End If
    If mp_TryReadPersistedSchedule(persistedTime, persistedProcedure) Then
        Application.OnTime EarliestTime:=persistedTime, Procedure:=persistedProcedure, Schedule:=False
    End If
    Err.Clear
    On Error GoTo 0

    g_StatusClearScheduled = False
    g_StatusClearTime = 0
    If clearPersistedState Then
        mp_ClearPersistedSchedule
    End If
End Sub

Private Function mp_IsClosingActive() As Boolean
    Dim persistedCloseUntilText As String
    Dim persistedCloseUntil As Date

    If g_CloseUntil > 0 Then
        If g_CloseUntil > Now Then
            mp_IsClosingActive = True
            Exit Function
        End If
        g_CloseUntil = 0
    End If

    persistedCloseUntilText = Trim$(mp_GetStoreValue(STATUS_STORE_KEY_CLOSE_UNTIL))
    If Len(persistedCloseUntilText) = 0 Then Exit Function
    If Not IsNumeric(persistedCloseUntilText) Then
        mp_DeleteStoreValue STATUS_STORE_KEY_CLOSE_UNTIL
        Exit Function
    End If

    persistedCloseUntil = CDate(CDbl(persistedCloseUntilText))
    If persistedCloseUntil > Now Then
        g_CloseUntil = persistedCloseUntil
        mp_IsClosingActive = True
    Else
        mp_DeleteStoreValue STATUS_STORE_KEY_CLOSE_UNTIL
    End If
End Function

Private Sub mp_SavePersistedSchedule(ByVal scheduledTime As Date, ByVal procedureName As String)
    mp_SaveStoreValue STATUS_STORE_KEY_CLEAR_AT, CStr(CDbl(scheduledTime))
    mp_SaveStoreValue STATUS_STORE_KEY_CLEAR_PROC, procedureName
End Sub

Private Function mp_TryReadPersistedSchedule(ByRef outScheduledTime As Date, ByRef outProcedureName As String) As Boolean
    Dim scheduledTimeText As String

    scheduledTimeText = Trim$(mp_GetStoreValue(STATUS_STORE_KEY_CLEAR_AT))
    outProcedureName = Trim$(mp_GetStoreValue(STATUS_STORE_KEY_CLEAR_PROC))
    If Len(scheduledTimeText) = 0 Or Len(outProcedureName) = 0 Then Exit Function
    If Not IsNumeric(scheduledTimeText) Then Exit Function

    outScheduledTime = CDate(CDbl(scheduledTimeText))
    mp_TryReadPersistedSchedule = True
End Function

Private Sub mp_ClearPersistedSchedule()
    mp_DeleteStoreValue STATUS_STORE_KEY_CLEAR_AT
    mp_DeleteStoreValue STATUS_STORE_KEY_CLEAR_PROC
End Sub

Private Function mp_GetStoreValue(ByVal keyName As String) As String
    On Error Resume Next
    mp_GetStoreValue = GetSetting(STATUS_STORE_APP, mp_GetStoreSection(), keyName, vbNullString)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub mp_SaveStoreValue(ByVal keyName As String, ByVal valueText As String)
    On Error Resume Next
    SaveSetting STATUS_STORE_APP, mp_GetStoreSection(), keyName, valueText
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub mp_DeleteStoreValue(ByVal keyName As String)
    On Error Resume Next
    DeleteSetting STATUS_STORE_APP, mp_GetStoreSection(), keyName
    Err.Clear
    On Error GoTo 0
End Sub

Private Function mp_GetStoreSection() As String
    mp_GetStoreSection = STATUS_STORE_SECTION_PREFIX & ThisWorkbook.Name
End Function

Private Function mp_ResolveLogFilePath(ByVal relativeOrAbsolutePath As String) As String
    Dim normalized As String
    Dim basePath As String

    normalized = Trim$(relativeOrAbsolutePath)
    If Len(normalized) = 0 Then normalized = DEFAULT_LOG_FILE_RELATIVE_PATH
    normalized = Replace$(normalized, "/", "\")

    If Left$(normalized, 2) = "\\" Or InStr(1, normalized, ":\", vbTextCompare) > 0 Then
        mp_ResolveLogFilePath = normalized
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_ResolveLogFilePath = normalized
        Exit Function
    End If
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    mp_ResolveLogFilePath = basePath & normalized
End Function

Private Sub mp_EnsureLogParentFolder(ByVal filePath As String)
    Dim fso As Object
    Dim parentPath As String

    If Len(Trim$(filePath)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    parentPath = fso.GetParentFolderName(filePath)
    If Len(parentPath) = 0 Then Exit Sub
    mp_EnsureFolderExists fso, parentPath
End Sub

Private Sub mp_EnsureFolderExists(ByVal fso As Object, ByVal folderPath As String)
    Dim parentPath As String

    If fso Is Nothing Then Exit Sub
    folderPath = Trim$(folderPath)
    If Len(folderPath) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        mp_EnsureFolderExists fso, parentPath
    End If
    fso.CreateFolder folderPath
End Sub

Public Sub m_RenderErrorBanner( _
    ByVal ws As Worksheet, _
    ByVal errDescription As String, _
    Optional ByVal errSource As String = vbNullString, _
    Optional ByVal errNumber As Long = 0, _
    Optional ByVal titleText As String = "ERROR: Operation failed", _
    Optional ByVal bannerRangeAddress As String = vbNullString, _
    Optional ByVal prepareOutputPanel As Boolean = False, _
    Optional ByVal wb As Workbook _
)
    Dim messageText As String
    Dim bodyLines As Collection

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(errDescription)
    If Len(messageText) = 0 Then
        messageText = "Unknown error."
    End If

    Set bodyLines = New Collection
    bodyLines.Add messageText
    bodyLines.Add "Source: " & IIf(Len(Trim$(errSource)) > 0, errSource, "n/a")
    bodyLines.Add "Code: " & CStr(errNumber)

    If prepareOutputPanel Then
        On Error Resume Next
        ex_OutputFormattingPipeline.m_ApplySheetPipeline ws
        On Error GoTo 0
    End If

    m_RenderBanner ws, titleText, bodyLines, bannerRangeAddress, BANNER_KIND_ERROR, messageText
End Sub

Public Sub m_RenderWarningBanner( _
    ByVal ws As Worksheet, _
    ByVal warningText As String, _
    Optional ByVal titleText As String = "WARNING", _
    Optional ByVal bannerRangeAddress As String = vbNullString _
)
    Dim messageText As String
    Dim bodyLines As Collection

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(warningText)
    If Len(messageText) = 0 Then
        messageText = "Action required."
    End If

    Set bodyLines = New Collection
    bodyLines.Add messageText

    m_RenderBanner ws, titleText, bodyLines, bannerRangeAddress, BANNER_KIND_WARNING, messageText
End Sub

Public Sub m_ClearBannerAnchors(ByVal ws As Worksheet)
    mp_ClearAnchorsByPrefix ws, BANNER_ANCHOR_PREFIX
    mp_ClearAnchorsByPrefix ws, BANNER_MESSAGE_ANCHOR_PREFIX
End Sub

Public Function m_TryGetBannerRangeAddressByText( _
    ByVal ws As Worksheet, _
    ByVal bannerText As String, _
    ByRef outRangeAddress As String _
) As Boolean
    Dim anchorRange As Range

    If ws Is Nothing Then Exit Function
    If Not mp_TryGetBannerRangeByMessage(ws, bannerText, anchorRange) Then Exit Function
    If anchorRange Is Nothing Then Exit Function

    outRangeAddress = anchorRange.Address(False, False, xlA1)
    m_TryGetBannerRangeAddressByText = True
End Function

Public Sub m_ClearResultTableAnchors(ByVal ws As Worksheet)
    mp_ClearAnchorsByPrefix ws, TABLE_ANCHOR_PREFIX
End Sub

Public Sub m_ClearResultRowAnchors(ByVal ws As Worksheet)
    mp_ClearAnchorsByPrefix ws, ROW_ANCHOR_PREFIX
End Sub

Public Function m_BuildResultRowAnchorName( _
    ByVal tableRef As String, _
    ByVal rowOrdinal As Long _
) As String
    Dim normalized As String

    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then Exit Function
    If rowOrdinal < 1 Then Exit Function

    normalized = mp_SanitizeNameToken(tableRef)
    If Len(normalized) = 0 Then Exit Function
    If Len(normalized) > 150 Then normalized = Left$(normalized, 150)

    m_BuildResultRowAnchorName = ROW_ANCHOR_PREFIX & normalized & "_" & mp_ChecksumHex4(tableRef) & "_" & CStr(rowOrdinal)
End Function

Public Sub m_RegisterResultRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByVal rowIndex As Long _
)
    Dim anchorRange As Range

    If ws Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    If rowIndex > ws.Rows.Count Then Exit Sub

    Set anchorRange = ws.Cells(rowIndex, 1)
    mp_SetNamedRangeAnchor ws, anchorName, anchorRange
End Sub

Public Function m_TryResolveResultRowAnchorRow( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim anchorRange As Range

    If ws Is Nothing Then Exit Function
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Function
    If Not mp_TryGetNamedRangeAnchor(ws, anchorName, anchorRange) Then Exit Function
    If anchorRange Is Nothing Then Exit Function

    outRowIndex = anchorRange.Row
    If outRowIndex < 1 Then Exit Function
    If outRowIndex > ws.Rows.Count Then Exit Function
    m_TryResolveResultRowAnchorRow = True
End Function

Public Sub m_RegisterResultTableAnchor( _
    ByVal ws As Worksheet, _
    ByVal tableRef As String, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long _
)
    Dim anchorName As String
    Dim anchorRange As Range

    If ws Is Nothing Then Exit Sub
    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then Exit Sub
    If rowStart < 1 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    If rowEnd > ws.Rows.Count Then rowEnd = ws.Rows.Count

    anchorName = mp_BuildResultTableAnchorName(tableRef)
    If Len(anchorName) = 0 Then Exit Sub

    Set anchorRange = ws.Range(ws.Cells(rowStart, 1), ws.Cells(rowEnd, 1))
    mp_SetNamedRangeAnchor ws, anchorName, anchorRange
End Sub

Public Function m_HasResultTableAnchor( _
    ByVal ws As Worksheet, _
    ByVal tableRef As String _
) As Boolean
    Dim rowStart As Long
    Dim rowEnd As Long

    m_HasResultTableAnchor = mp_TryGetResultTableBounds(ws, tableRef, rowStart, rowEnd)
End Function

Public Sub m_RenderTextBanner( _
    ByVal ws As Worksheet, _
    ByVal bannerText As String, _
    Optional ByVal titleText As String = "NOTICE", _
    Optional ByVal bannerRangeAddress As String = vbNullString, _
    Optional ByVal bannerKind As String = BANNER_KIND_WARNING _
)
    Dim bodyLines As Collection

    If ws Is Nothing Then Exit Sub

    Set bodyLines = New Collection
    bannerText = Trim$(bannerText)
    If Len(bannerText) = 0 Then
        bodyLines.Add "Action required."
    Else
        bodyLines.Add bannerText
    End If

    m_RenderBanner ws, titleText, bodyLines, bannerRangeAddress, bannerKind, bannerText
End Sub

Public Sub m_RenderTextBannerAtCell( _
    ByVal ws As Worksheet, _
    ByVal bannerText As String, _
    ByVal topLeftCellRef As String, _
    Optional ByVal titleText As String = "NOTICE", _
    Optional ByVal bannerKind As String = BANNER_KIND_WARNING, _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 0, _
    Optional ByVal insertRows As Boolean = True _
)
    Dim targetCell As Range
    Dim existingBannerRange As Range
    Dim bannerColumns As Long
    Dim bannerRows As Long
    Dim requiredRows As Long
    Dim rowsToInsert As Long
    Dim insertAtRow As Long
    Dim bannerStartRow As Long
    Dim rangeAddress As String

    If ws Is Nothing Then Exit Sub
    If gapRowsBefore < 0 Then gapRowsBefore = 0
    If gapRowsAfter < 0 Then gapRowsAfter = 0

    If mp_TryGetBannerRangeByMessage(ws, bannerText, existingBannerRange) Then
        m_RenderTextBanner ws, bannerText, titleText, existingBannerRange.Address(False, False, xlA1), bannerKind
        Exit Sub
    End If

    topLeftCellRef = Trim$(topLeftCellRef)
    If Len(topLeftCellRef) = 0 Then topLeftCellRef = "A1"

    Set targetCell = ws.Range(topLeftCellRef).Cells(1, 1)
    mp_GetBannerDimensions bannerKind, bannerColumns, bannerRows
    requiredRows = mp_GetRequiredBannerRowsFromText(bannerText, bannerRows)

    If insertRows Then
        insertAtRow = targetCell.Row
        rowsToInsert = gapRowsBefore + requiredRows + gapRowsAfter
        If rowsToInsert > 0 Then
            mp_InsertRowsSafe ws, insertAtRow, rowsToInsert
            bannerStartRow = insertAtRow + gapRowsBefore
            mp_UnmergeSpacerRows ws, insertAtRow, gapRowsBefore
            mp_UnmergeSpacerRows ws, bannerStartRow + requiredRows, gapRowsAfter
        Else
            bannerStartRow = insertAtRow
        End If
    Else
        bannerStartRow = targetCell.Row + gapRowsBefore
    End If

    rangeAddress = mp_BuildAddress( _
        bannerStartRow, _
        targetCell.Column, _
        bannerStartRow + requiredRows - 1, _
        targetCell.Column + bannerColumns - 1 _
    )

    m_RenderTextBanner ws, bannerText, titleText, rangeAddress, bannerKind
End Sub

Public Sub m_RenderTextBannerAfterBanner( _
    ByVal ws As Worksheet, _
    ByVal bannerText As String, _
    ByVal afterBannerIndex As Long, _
    Optional ByVal titleText As String = "NOTICE", _
    Optional ByVal bannerKind As String = BANNER_KIND_WARNING, _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 1, _
    Optional ByVal insertRows As Boolean = True _
)
    Dim existingBannerRange As Range
    Dim afterStartRow As Long
    Dim afterEndRow As Long
    Dim bannerColumns As Long
    Dim bannerRows As Long
    Dim requiredRows As Long
    Dim rowsToInsert As Long
    Dim bannerStartRow As Long
    Dim topStartRow As Long
    Dim insertAtRow As Long
    Dim rangeAddress As String

    If ws Is Nothing Then Exit Sub
    If gapRowsBefore < 0 Then gapRowsBefore = 0
    If gapRowsAfter < 0 Then gapRowsAfter = 0

    If mp_TryGetBannerRangeByMessage(ws, bannerText, existingBannerRange) Then
        m_RenderTextBanner ws, bannerText, titleText, existingBannerRange.Address(False, False, xlA1), bannerKind
        Exit Sub
    End If

    mp_GetBannerDimensions bannerKind, bannerColumns, bannerRows
    requiredRows = mp_GetRequiredBannerRowsFromText(bannerText, bannerRows)

    If afterBannerIndex <= 0 Then
        topStartRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
        If topStartRow < 1 Then topStartRow = 1
        topStartRow = mp_AdjustBannerStartForHeader(ws, topStartRow, requiredRows, True)

        If insertRows Then
            insertAtRow = topStartRow
            rowsToInsert = gapRowsBefore + requiredRows + gapRowsAfter
            mp_InsertRowsSafe ws, insertAtRow, rowsToInsert
            bannerStartRow = insertAtRow + gapRowsBefore
            mp_UnmergeSpacerRows ws, insertAtRow, gapRowsBefore
            mp_UnmergeSpacerRows ws, bannerStartRow + requiredRows, gapRowsAfter
        Else
            bannerStartRow = topStartRow + gapRowsBefore
            bannerStartRow = mp_AdjustBannerStartForExistingAnchors(ws, bannerStartRow, requiredRows)
        End If
    Else
        If Not mp_TryGetBannerBoundsByIndex(ws, afterBannerIndex, afterStartRow, afterEndRow) Then
            Err.Raise vbObjectError + 1761, "ex_Messaging", "Banner #" & CStr(afterBannerIndex) & " was not found on sheet '" & ws.Name & "'."
        End If

        If insertRows Then
            insertAtRow = afterEndRow + 1
            rowsToInsert = gapRowsBefore + requiredRows + gapRowsAfter
            mp_InsertRowsSafe ws, insertAtRow, rowsToInsert
            bannerStartRow = insertAtRow + gapRowsBefore
            mp_UnmergeSpacerRows ws, insertAtRow, gapRowsBefore
            mp_UnmergeSpacerRows ws, bannerStartRow + requiredRows, gapRowsAfter
        Else
            bannerStartRow = afterEndRow + 1 + gapRowsBefore
        End If
    End If

    rangeAddress = mp_BuildAddress(bannerStartRow, 1, bannerStartRow + requiredRows - 1, bannerColumns)
    m_RenderTextBanner ws, bannerText, titleText, rangeAddress, bannerKind
End Sub

Public Sub m_RenderTextBannerAtTable( _
    ByVal ws As Worksheet, _
    ByVal bannerText As String, _
    ByVal tableRef As String, _
    Optional ByVal positionText As String = "before", _
    Optional ByVal titleText As String = "NOTICE", _
    Optional ByVal bannerKind As String = BANNER_KIND_WARNING, _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 1, _
    Optional ByVal insertRows As Boolean = True _
)
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    Dim bannerColumns As Long
    Dim bannerRows As Long
    Dim requiredRows As Long
    Dim normalizedPos As String
    Dim rowsToInsert As Long
    Dim insertAtRow As Long
    Dim bannerStartRow As Long
    Dim rangeAddress As String

    If ws Is Nothing Then Exit Sub
    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then
        Err.Raise vbObjectError + 1762, "ex_Messaging", "Table reference is required for banner placement."
    End If
    If gapRowsBefore < 0 Then gapRowsBefore = 0
    If gapRowsAfter < 0 Then gapRowsAfter = 0

    If Not mp_TryGetResultTableBounds(ws, tableRef, tableStartRow, tableEndRow) Then
        Err.Raise vbObjectError + 1763, "ex_Messaging", "Table anchor is not found for tableRef '" & tableRef & "' on sheet '" & ws.Name & "'."
    End If

    normalizedPos = LCase$(Trim$(positionText))
    If Len(normalizedPos) = 0 Then normalizedPos = "before"
    If normalizedPos <> "before" And normalizedPos <> "after" Then
        Err.Raise vbObjectError + 1764, "ex_Messaging", "Unsupported table banner position '" & positionText & "'. Use 'before' or 'after'."
    End If

    mp_GetBannerDimensions bannerKind, bannerColumns, bannerRows
    requiredRows = mp_GetRequiredBannerRowsFromText(bannerText, bannerRows)

    If normalizedPos = "before" Then
        If insertRows Then
            insertAtRow = tableStartRow
            rowsToInsert = gapRowsBefore + requiredRows + gapRowsAfter
            mp_InsertRowsSafe ws, insertAtRow, rowsToInsert
            bannerStartRow = insertAtRow + gapRowsBefore
            mp_UnmergeSpacerRows ws, insertAtRow, gapRowsBefore
            mp_UnmergeSpacerRows ws, bannerStartRow + requiredRows, gapRowsAfter
        Else
            bannerStartRow = tableStartRow - gapRowsAfter - requiredRows - gapRowsBefore
            If bannerStartRow < 1 Then bannerStartRow = 1
        End If
    Else
        If insertRows Then
            insertAtRow = tableEndRow + 1
            rowsToInsert = gapRowsBefore + requiredRows + gapRowsAfter
            mp_InsertRowsSafe ws, insertAtRow, rowsToInsert
            bannerStartRow = insertAtRow + gapRowsBefore
            mp_UnmergeSpacerRows ws, insertAtRow, gapRowsBefore
            mp_UnmergeSpacerRows ws, bannerStartRow + requiredRows, gapRowsAfter
        Else
            bannerStartRow = tableEndRow + 1 + gapRowsBefore
        End If
    End If

    rangeAddress = mp_BuildAddress(bannerStartRow, 1, bannerStartRow + requiredRows - 1, bannerColumns)
    m_RenderTextBanner ws, bannerText, titleText, rangeAddress, bannerKind
End Sub

Public Sub m_RenderBanner( _
    ByVal ws As Worksheet, _
    ByVal titleText As String, _
    ByVal bodyLines As Collection, _
    Optional ByVal bannerRangeAddress As String = vbNullString, _
    Optional ByVal bannerKind As String = BANNER_KIND_WARNING, _
    Optional ByVal bannerIdentityText As String = vbNullString _
)
    Dim requestedRange As Range
    Dim bannerRange As Range
    Dim bodyText As String
    Dim combinedText As String
    Dim rowCount As Long
    Dim rowOffset As Long
    Dim startRow As Long
    Dim startCol As Long
    Dim colCount As Long
    Dim isAutoAddress As Boolean

    If ws Is Nothing Then Exit Sub

    isAutoAddress = (Len(Trim$(bannerRangeAddress)) = 0)
    bodyText = mp_GetBannerBodyText(bodyLines)
    combinedText = mp_ComposeBannerText(titleText, bodyText)
    If isAutoAddress Then
        bannerRangeAddress = ex_SheetStylesXmlProvider.m_GetOutputBannerRangeAddress(ThisWorkbook, 1)
    End If

    Set requestedRange = ws.Range(bannerRangeAddress)
    startRow = requestedRange.Row
    startCol = requestedRange.Column
    colCount = requestedRange.Columns.Count
    rowCount = 1

    startRow = mp_AdjustBannerStartForHeader(ws, startRow, rowCount, isAutoAddress)
    If isAutoAddress Then
        startRow = mp_AdjustBannerStartForExistingAnchors(ws, startRow, rowCount)
    End If

    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1
    If colCount < 1 Then colCount = 1
    If startRow > ws.Rows.Count Then startRow = ws.Rows.Count
    If startCol > ws.Columns.Count Then startCol = ws.Columns.Count
    If startCol + colCount - 1 > ws.Columns.Count Then
        colCount = ws.Columns.Count - startCol + 1
    End If
    If rowCount > ws.Rows.Count - startRow + 1 Then
        rowCount = ws.Rows.Count - startRow + 1
    End If
    If rowCount < 1 Then Exit Sub

    Set bannerRange = ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow + rowCount - 1, startCol + colCount - 1) _
    )

    On Error Resume Next
    ws.Rows(CStr(startRow) & ":" & CStr(startRow + rowCount - 1)).RowHeight = ws.StandardHeight
    On Error GoTo 0

    bannerRange.ClearContents
    bannerRange.UnMerge
    For rowOffset = 0 To rowCount - 1
        ws.Range(ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column), ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    Next rowOffset

    ws.Cells(bannerRange.Row, bannerRange.Column).Value = combinedText

    mp_ApplyBannerKindPipeline ws, bannerRange.Row, rowCount, bannerKind
    mp_ApplyBannerAutoHeight ws, bannerRange, combinedText, bannerKind
    mp_RegisterBannerAnchor ws, bannerRange
    mp_RegisterBannerMessageAnchor ws, bannerRange, bannerIdentityText
End Sub

Private Function mp_GetBannerBodyText(ByVal bodyLines As Collection) As String
    Dim i As Long
    Dim lineText As String

    If bodyLines Is Nothing Then Exit Function
    If bodyLines.Count = 0 Then Exit Function

    For i = 1 To bodyLines.Count
        lineText = CStr(bodyLines(i))
        If i = 1 Then
            mp_GetBannerBodyText = lineText
        Else
            mp_GetBannerBodyText = mp_GetBannerBodyText & vbLf & lineText
        End If
    Next i
End Function

Private Function mp_ComposeBannerText( _
    ByVal titleText As String, _
    ByVal bodyText As String _
) As String
    titleText = Trim$(titleText)
    bodyText = Trim$(bodyText)

    If Len(titleText) = 0 Then
        mp_ComposeBannerText = bodyText
    ElseIf Len(bodyText) = 0 Then
        mp_ComposeBannerText = titleText
    Else
        mp_ComposeBannerText = titleText & vbLf & vbLf & bodyText
    End If
End Function

Public Sub m_ApplyBannerAutoHeightForRange( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal bannerText As String, _
    Optional ByVal bannerKind As String = BANNER_KIND_WARNING _
)
    Dim effectiveRange As Range

    If ws Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub

    Set effectiveRange = targetRange
    On Error Resume Next
    If CBool(effectiveRange.MergeCells) Then
        Set effectiveRange = effectiveRange.MergeArea
    End If
    On Error GoTo 0
    If effectiveRange Is Nothing Then Exit Sub

    mp_ApplyBannerAutoHeight ws, effectiveRange, CStr(bannerText), CStr(bannerKind)
End Sub

Private Sub mp_ApplyBannerAutoHeight( _
    ByVal ws As Worksheet, _
    ByVal bannerRange As Range, _
    ByVal bannerText As String, _
    ByVal bannerKind As String _
)
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim hasBannerStyle As Boolean
    Dim baseRowHeight As Double
    Dim normalizedKind As String
    Dim autoHeightMarginTop As Double
    Dim autoHeightMarginBottom As Double

    If ws Is Nothing Then Exit Sub
    If bannerRange Is Nothing Then Exit Sub

    normalizedKind = LCase$(Trim$(bannerKind))
    If normalizedKind = BANNER_KIND_ERROR Then
        If ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
            hasBannerStyle = True
        ElseIf ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
            hasBannerStyle = True
        End If
    Else
        If ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
            hasBannerStyle = True
        ElseIf ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
            hasBannerStyle = True
        End If
    End If

    baseRowHeight = ws.StandardHeight
    If hasBannerStyle Then
        If bannerStyle.RowHeight > 0 Then baseRowHeight = bannerStyle.RowHeight
    End If

    mp_LoadBannerAutoHeightMargins ws, normalizedKind, autoHeightMarginTop, autoHeightMarginBottom
    ex_SheetHelpers.m_ApplySingleRowTextAutoHeight _
        ws, _
        bannerRange, _
        bannerText, _
        baseRowHeight, _
        0, _
        True, _
        True, _
        autoHeightMarginTop, _
        autoHeightMarginBottom, _
        0, _
        0, _
        2, _
        0
End Sub

Private Sub mp_LoadBannerAutoHeightMargins( _
    ByVal ws As Worksheet, _
    ByVal bannerKind As String, _
    ByRef outMarginTop As Double, _
    ByRef outMarginBottom As Double _
)
    Dim stageLayers As Collection
    Dim layerObj As obj_StyleLayer
    Dim ruleObj As obj_StyleRule
    Dim declarations As Object
    Dim parsedValue As Double

    outMarginTop = 0
    outMarginBottom = 0
    If ws Is Nothing Then Exit Sub

    Set stageLayers = ex_StylePipelineEngine.m_LoadSheetPipelineLayers(ws.Name, ThisWorkbook, BANNER_STYLE_STAGE_NAME)
    If stageLayers Is Nothing Then Exit Sub
    If stageLayers.Count = 0 Then Exit Sub

    For Each layerObj In stageLayers
        If layerObj Is Nothing Then GoTo ContinueLayer
        For Each ruleObj In layerObj.Rules
            If ruleObj Is Nothing Then GoTo ContinueRule
            If StrComp(LCase$(Trim$(ruleObj.Target)), "row", vbBinaryCompare) <> 0 Then GoTo ContinueRule
            If Not mp_BannerRuleMatchesKind(ruleObj.Selector, bannerKind) Then GoTo ContinueRule

            Set declarations = ruleObj.Declarations
            If declarations Is Nothing Then GoTo ContinueRule

            If declarations.Exists(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP) Then
                If Not mp_TryParseNonNegativeDouble(CStr(declarations(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP)), parsedValue) Then
                    Err.Raise vbObjectError + 1769, "ex_Messaging", _
                        "Invalid '" & STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_TOP & "' for banner rule '" & ruleObj.RuleId & "'."
                End If
                outMarginTop = parsedValue
            End If
            If declarations.Exists(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM) Then
                If Not mp_TryParseNonNegativeDouble(CStr(declarations(STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM)), parsedValue) Then
                    Err.Raise vbObjectError + 1770, "ex_Messaging", _
                        "Invalid '" & STYLE_PROP_CUSTOM_AUTO_HEIGHT_MARGIN_BOTTOM & "' for banner rule '" & ruleObj.RuleId & "'."
                End If
                outMarginBottom = parsedValue
            End If
ContinueRule:
        Next ruleObj
ContinueLayer:
    Next layerObj
End Sub

Private Function mp_BannerRuleMatchesKind( _
    ByVal selectorText As String, _
    ByVal bannerKind As String _
) As Boolean
    Dim selectorParts() As String
    Dim selectorPart As Variant
    Dim keyText As String
    Dim valueText As String
    Dim eqPos As Long
    Dim kindClause As String
    Dim kindTokens() As String
    Dim token As Variant
    Dim normalizedKind As String

    normalizedKind = LCase$(Trim$(bannerKind))
    If Len(normalizedKind) = 0 Then normalizedKind = BANNER_KIND_WARNING

    selectorText = Trim$(selectorText)
    If Len(selectorText) = 0 Then
        mp_BannerRuleMatchesKind = True
        Exit Function
    End If

    selectorParts = Split(selectorText, ";")
    For Each selectorPart In selectorParts
        eqPos = InStr(1, CStr(selectorPart), "=", vbBinaryCompare)
        If eqPos <= 1 Then GoTo ContinuePart
        keyText = LCase$(Trim$(Left$(CStr(selectorPart), eqPos - 1)))
        If StrComp(keyText, "kind", vbBinaryCompare) <> 0 Then GoTo ContinuePart
        valueText = Trim$(Mid$(CStr(selectorPart), eqPos + 1))
        If Len(valueText) = 0 Then
            mp_BannerRuleMatchesKind = False
            Exit Function
        End If
        kindClause = Replace(valueText, ",", "|")
        kindTokens = Split(kindClause, "|")
        For Each token In kindTokens
            token = LCase$(Trim$(CStr(token)))
            If Len(token) = 0 Then GoTo ContinueToken
            If CStr(token) = "*" Then
                mp_BannerRuleMatchesKind = True
                Exit Function
            End If
            If StrComp(CStr(token), normalizedKind, vbBinaryCompare) = 0 Then
                mp_BannerRuleMatchesKind = True
                Exit Function
            End If
ContinueToken:
        Next token
        mp_BannerRuleMatchesKind = False
        Exit Function
ContinuePart:
    Next selectorPart

    ' Rules without explicit kind selector are treated as global for all banner kinds.
    mp_BannerRuleMatchesKind = True
End Function

Private Function mp_TryParseNonNegativeDouble( _
    ByVal textValue As String, _
    ByRef outValue As Double _
) As Boolean
    textValue = Trim$(textValue)
    If Len(textValue) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseDouble(textValue, outValue, True) Then Exit Function
    If outValue < 0 Then Exit Function
    mp_TryParseNonNegativeDouble = True
End Function

Private Sub mp_ApplyBannerKindPipeline( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal rowCount As Long, _
    ByVal bannerKind As String _
)
    Dim stageLayers As Collection
    Dim bannerPipeline As Collection
    Dim layerObj As obj_StyleLayer
    Dim kindRanges As Object
    Dim emptyTargets As Collection
    Dim normalizedKind As String

    If ws Is Nothing Then Exit Sub
    If startRow < 1 Then Exit Sub
    If rowCount < 1 Then Exit Sub

    normalizedKind = LCase$(Trim$(bannerKind))
    If Len(normalizedKind) = 0 Then normalizedKind = BANNER_KIND_WARNING

    Set stageLayers = ex_StylePipelineEngine.m_LoadSheetPipelineLayers(ws.Name, ThisWorkbook, BANNER_STYLE_STAGE_NAME)
    If stageLayers Is Nothing Or stageLayers.Count = 0 Then
        MsgBox "StylePipeline has no stage '" & BANNER_STYLE_STAGE_NAME & "' for page '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    Set bannerPipeline = ex_StylePipelineEngine.m_CreatePipeline()
    For Each layerObj In stageLayers
        ex_StylePipelineEngine.m_AddLayer bannerPipeline, layerObj
    Next layerObj

    Set kindRanges = ex_StylePipelineEngine.m_CreateKindRanges()
    ex_StylePipelineEngine.m_AddKindRange kindRanges, normalizedKind, startRow, 1, startRow + rowCount - 1, 0

    Set emptyTargets = New Collection
    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, emptyTargets, bannerPipeline, vbNullString, kindRanges
End Sub

Private Function mp_GetRequiredBannerRowsFromText( _
    ByVal bannerText As String, _
    ByVal configuredRows As Long _
) As Long
    ' Banner layout is fixed to one row with plain text:
    ' Title + blank line + content.
    mp_GetRequiredBannerRowsFromText = 1
End Function

Private Sub mp_GetBannerDimensions( _
    ByVal bannerKind As String, _
    ByRef outColumns As Long, _
    ByRef outRows As Long _
)
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim normalizedKind As String

    normalizedKind = LCase$(Trim$(bannerKind))
    If Len(normalizedKind) = 0 Then normalizedKind = BANNER_KIND_WARNING

    If normalizedKind = BANNER_KIND_ERROR Then
        If ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        ElseIf ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        End If
    Else
        If ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        ElseIf ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        End If
    End If

    If outColumns < 1 Then outColumns = DEFAULT_BANNER_COLUMNS
    If outRows < 1 Then outRows = DEFAULT_BANNER_ROWS
End Sub

Private Function mp_AdjustBannerStartForHeader( _
    ByVal ws As Worksheet, _
    ByVal desiredStartRow As Long, _
    ByVal rowCount As Long, _
    Optional ByVal enforceViewZone As Boolean = True _
) As Long
    Dim headerStartRow As Long
    Dim headerEndRow As Long
    Dim bannerEndRow As Long
    Dim viewStartRow As Long

    mp_AdjustBannerStartForHeader = desiredStartRow
    If ws Is Nothing Then Exit Function
    If rowCount < 1 Then rowCount = 1

    viewStartRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    If viewStartRow < 1 Then viewStartRow = 1

    If enforceViewZone And mp_AdjustBannerStartForHeader < viewStartRow Then
        mp_AdjustBannerStartForHeader = viewStartRow
    End If

    If Not mp_TryGetScriptHeaderBounds(ws, headerStartRow, headerEndRow) Then Exit Function
    If enforceViewZone Then
        If mp_AdjustBannerStartForHeader <= (headerEndRow + 1) Then
            mp_AdjustBannerStartForHeader = headerEndRow + 2
            Exit Function
        End If
    End If
    bannerEndRow = mp_AdjustBannerStartForHeader + rowCount - 1
    If mp_RowsOverlap(mp_AdjustBannerStartForHeader, bannerEndRow, headerStartRow, headerEndRow) Then
        ' Keep one spacer row below single-header block.
        mp_AdjustBannerStartForHeader = headerEndRow + 2
    End If
End Function

Private Function mp_AdjustBannerStartForExistingAnchors( _
    ByVal ws As Worksheet, _
    ByVal desiredStartRow As Long, _
    ByVal rowCount As Long _
) As Long
    Dim anchors As Collection
    Dim entry As Object
    Dim startRow As Long
    Dim endRow As Long
    Dim hasOverlap As Boolean

    startRow = desiredStartRow
    If ws Is Nothing Then
        mp_AdjustBannerStartForExistingAnchors = startRow
        Exit Function
    End If
    If rowCount < 1 Then
        mp_AdjustBannerStartForExistingAnchors = startRow
        Exit Function
    End If

    Set anchors = mp_GetSortedBannerAnchors(ws)

    Do
        hasOverlap = False
        endRow = startRow + rowCount - 1
        For Each entry In anchors
            If entry Is Nothing Then GoTo ContinueEntry
            If mp_RowsOverlap(startRow, endRow, CLng(entry("RowStart")), CLng(entry("RowEnd"))) Then
                startRow = CLng(entry("RowEnd")) + 1
                hasOverlap = True
                Exit For
            End If
ContinueEntry:
        Next entry
    Loop While hasOverlap

    mp_AdjustBannerStartForExistingAnchors = startRow
End Function

Private Function mp_RowsOverlap( _
    ByVal rowStartA As Long, _
    ByVal rowEndA As Long, _
    ByVal rowStartB As Long, _
    ByVal rowEndB As Long _
) As Boolean
    If rowEndA < rowStartA Then rowEndA = rowStartA
    If rowEndB < rowStartB Then rowEndB = rowStartB
    mp_RowsOverlap = Not (rowEndA < rowStartB Or rowEndB < rowStartA)
End Function

Private Function mp_TryGetScriptHeaderBounds( _
    ByVal ws As Worksheet, _
    ByRef outStartRow As Long, _
    ByRef outEndRow As Long _
) As Boolean
    Dim headerRange As Range

    If ws Is Nothing Then Exit Function
    If Not mp_TryGetNamedRangeAnchor(ws, POST_PROCESS_HEADER_ANCHOR_NAME, headerRange) Then Exit Function
    If headerRange Is Nothing Then Exit Function

    outStartRow = headerRange.Row
    outEndRow = headerRange.Row + headerRange.Rows.Count - 1
    mp_TryGetScriptHeaderBounds = True
End Function

Private Function mp_GetSortedBannerAnchors(ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Dim namedEntry As Name
    Dim anchorRange As Range
    Dim nameText As String
    Dim entry As Object

    Set result = New Collection
    If ws Is Nothing Then
        Set mp_GetSortedBannerAnchors = result
        Exit Function
    End If

    For Each namedEntry In ws.Names
        If namedEntry Is Nothing Then GoTo ContinueName
        nameText = CStr(namedEntry.Name)
        If InStr(1, nameText, "!", vbBinaryCompare) > 0 Then
            nameText = Mid$(nameText, InStrRev(nameText, "!", -1, vbBinaryCompare) + 1)
        End If
        If StrComp(Left$(LCase$(nameText), Len(BANNER_ANCHOR_PREFIX)), LCase$(BANNER_ANCHOR_PREFIX), vbBinaryCompare) <> 0 Then GoTo ContinueName

        On Error Resume Next
        Set anchorRange = namedEntry.RefersToRange
        On Error GoTo 0
        If anchorRange Is Nothing Then GoTo ContinueName

        Set entry = CreateObject("Scripting.Dictionary")
        entry.CompareMode = 1
        entry("Name") = nameText
        entry("RowStart") = anchorRange.Row
        entry("RowEnd") = anchorRange.Row + anchorRange.Rows.Count - 1
        entry("ColStart") = anchorRange.Column
        entry("ColEnd") = anchorRange.Column + anchorRange.Columns.Count - 1
        mp_AddBannerAnchorEntrySorted result, entry
ContinueName:
        Set anchorRange = Nothing
    Next namedEntry

    Set mp_GetSortedBannerAnchors = result
End Function

Private Sub mp_AddBannerAnchorEntrySorted(ByVal target As Collection, ByVal entry As Object)
    Dim i As Long

    If target Is Nothing Then Exit Sub
    If entry Is Nothing Then Exit Sub

    For i = 1 To target.Count
        If CLng(entry("RowStart")) < CLng(target(i)("RowStart")) Then
            target.Add entry, Before:=i
            Exit Sub
        End If
    Next i
    target.Add entry
End Sub

Private Function mp_TryGetBannerBoundsByIndex( _
    ByVal ws As Worksheet, _
    ByVal bannerIndex As Long, _
    ByRef outRowStart As Long, _
    ByRef outRowEnd As Long _
) As Boolean
    Dim anchors As Collection
    Dim entry As Object

    If ws Is Nothing Then Exit Function
    If bannerIndex < 1 Then Exit Function

    Set anchors = mp_GetSortedBannerAnchors(ws)
    If anchors Is Nothing Then Exit Function
    If bannerIndex > anchors.Count Then Exit Function

    Set entry = anchors(bannerIndex)
    If entry Is Nothing Then Exit Function

    outRowStart = CLng(entry("RowStart"))
    outRowEnd = CLng(entry("RowEnd"))
    mp_TryGetBannerBoundsByIndex = True
End Function

Private Sub mp_RegisterBannerAnchor(ByVal ws As Worksheet, ByVal bannerRange As Range)
    Dim anchors As Collection
    Dim entry As Object
    Dim anchorName As String
    Dim startRow As Long
    Dim endRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim nextIndex As Long

    If ws Is Nothing Then Exit Sub
    If bannerRange Is Nothing Then Exit Sub

    startRow = bannerRange.Row
    endRow = bannerRange.Row + bannerRange.Rows.Count - 1
    startCol = bannerRange.Column
    endCol = bannerRange.Column + bannerRange.Columns.Count - 1

    Set anchors = mp_GetSortedBannerAnchors(ws)
    For Each entry In anchors
        If entry Is Nothing Then GoTo ContinueEntry
        If CLng(entry("RowStart")) = startRow _
            And CLng(entry("RowEnd")) = endRow _
            And CLng(entry("ColStart")) = startCol _
            And CLng(entry("ColEnd")) = endCol Then
            anchorName = CStr(entry("Name"))
            Exit For
        End If
ContinueEntry:
    Next entry

    If Len(anchorName) = 0 Then
        anchorName = vbNullString
        For nextIndex = 1 To BANNER_MAX_ANCHOR_INDEX
            If Not mp_AnchorNameExists(ws, BANNER_ANCHOR_PREFIX & Format$(nextIndex, "0000")) Then
                anchorName = BANNER_ANCHOR_PREFIX & Format$(nextIndex, "0000")
                Exit For
            End If
        Next nextIndex
    End If
    If Len(anchorName) = 0 Then Exit Sub

    mp_SetNamedRangeAnchor ws, anchorName, bannerRange
End Sub

Private Sub mp_RegisterBannerMessageAnchor( _
    ByVal ws As Worksheet, _
    ByVal bannerRange As Range, _
    ByVal bannerText As String _
)
    Dim anchorName As String

    If ws Is Nothing Then Exit Sub
    If bannerRange Is Nothing Then Exit Sub

    anchorName = mp_BuildBannerMessageAnchorName(bannerText)
    If Len(anchorName) = 0 Then Exit Sub
    mp_SetNamedRangeAnchor ws, anchorName, bannerRange
End Sub

Private Function mp_TryGetBannerRangeByMessage( _
    ByVal ws As Worksheet, _
    ByVal bannerText As String, _
    ByRef outRange As Range _
) As Boolean
    Dim anchorName As String

    If ws Is Nothing Then Exit Function
    anchorName = mp_BuildBannerMessageAnchorName(bannerText)
    If Len(anchorName) = 0 Then Exit Function

    If Not mp_TryGetNamedRangeAnchor(ws, anchorName, outRange) Then Exit Function
    If outRange Is Nothing Then Exit Function
    mp_TryGetBannerRangeByMessage = True
End Function

Private Function mp_BuildBannerMessageAnchorName(ByVal bannerText As String) As String
    Dim normalized As String

    bannerText = Trim$(bannerText)
    If Len(bannerText) = 0 Then Exit Function

    normalized = mp_SanitizeNameToken(LCase$(bannerText))
    If Len(normalized) = 0 Then normalized = "banner"
    If Len(normalized) > 150 Then normalized = Left$(normalized, 150)

    mp_BuildBannerMessageAnchorName = BANNER_MESSAGE_ANCHOR_PREFIX & normalized & "_" & mp_ChecksumHex4(bannerText)
End Function

Private Function mp_BuildResultTableAnchorName(ByVal tableRef As String) As String
    Dim normalized As String

    normalized = mp_SanitizeNameToken(tableRef)
    If Len(normalized) = 0 Then Exit Function

    If Len(normalized) > 180 Then normalized = Left$(normalized, 180)
    mp_BuildResultTableAnchorName = TABLE_ANCHOR_PREFIX & normalized & "_" & mp_ChecksumHex4(tableRef)
End Function

Private Function mp_TryGetResultTableBounds( _
    ByVal ws As Worksheet, _
    ByVal tableRef As String, _
    ByRef outRowStart As Long, _
    ByRef outRowEnd As Long _
) As Boolean
    Dim anchorName As String
    Dim anchorRange As Range

    If ws Is Nothing Then Exit Function
    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then Exit Function

    anchorName = mp_BuildResultTableAnchorName(tableRef)
    If Len(anchorName) = 0 Then Exit Function
    If Not mp_TryGetNamedRangeAnchor(ws, anchorName, anchorRange) Then Exit Function
    If anchorRange Is Nothing Then Exit Function

    outRowStart = anchorRange.Row
    outRowEnd = anchorRange.Row + anchorRange.Rows.Count - 1
    mp_TryGetResultTableBounds = True
End Function

Private Sub mp_ClearAnchorsByPrefix(ByVal ws As Worksheet, ByVal namePrefix As String)
    Dim namesToDelete As Collection
    Dim namedEntry As Name
    Dim nameText As String
    Dim i As Long

    If ws Is Nothing Then Exit Sub
    namePrefix = LCase$(Trim$(namePrefix))
    If Len(namePrefix) = 0 Then Exit Sub

    Set namesToDelete = New Collection
    For Each namedEntry In ws.Names
        If namedEntry Is Nothing Then GoTo ContinueName
        nameText = CStr(namedEntry.Name)
        If InStr(1, nameText, "!", vbBinaryCompare) > 0 Then
            nameText = Mid$(nameText, InStrRev(nameText, "!", -1, vbBinaryCompare) + 1)
        End If
        If StrComp(Left$(LCase$(nameText), Len(namePrefix)), namePrefix, vbBinaryCompare) = 0 Then
            namesToDelete.Add nameText
        End If
ContinueName:
    Next namedEntry

    For i = 1 To namesToDelete.Count
        On Error Resume Next
        ws.Names(CStr(namesToDelete(i))).Delete
        On Error GoTo 0
    Next i
End Sub

Private Function mp_TryGetNamedRangeAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByRef outRange As Range _
) As Boolean
    Dim namedEntry As Name

    If ws Is Nothing Then Exit Function
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Function

    On Error Resume Next
    Set namedEntry = ws.Names(anchorName)
    On Error GoTo 0
    If namedEntry Is Nothing Then Exit Function

    On Error Resume Next
    Set outRange = namedEntry.RefersToRange
    On Error GoTo 0
    If outRange Is Nothing Then Exit Function

    mp_TryGetNamedRangeAnchor = True
End Function

Private Sub mp_SetNamedRangeAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByVal anchorRange As Range _
)
    If ws Is Nothing Then Exit Sub
    If anchorRange Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub

    On Error Resume Next
    ws.Names(anchorName).Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Names.Add Name:=anchorName, RefersTo:="=" & anchorRange.Address(True, True, xlA1, True)
    On Error GoTo 0
End Sub

Private Function mp_AnchorNameExists(ByVal ws As Worksheet, ByVal anchorName As String) As Boolean
    Dim namedEntry As Name

    If ws Is Nothing Then Exit Function
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Function

    On Error Resume Next
    Set namedEntry = ws.Names(anchorName)
    On Error GoTo 0
    mp_AnchorNameExists = Not (namedEntry Is Nothing)
End Function

Private Sub mp_InsertRowsSafe(ByVal ws As Worksheet, ByVal atRow As Long, ByVal rowCount As Long)
    Dim insertStart As Long
    Dim insertEnd As Long

    If ws Is Nothing Then Exit Sub
    If rowCount <= 0 Then Exit Sub

    If atRow < 1 Then atRow = 1
    If atRow > ws.Rows.Count Then atRow = ws.Rows.Count

    insertStart = atRow
    insertEnd = atRow + rowCount - 1
    If insertEnd > ws.Rows.Count Then insertEnd = ws.Rows.Count
    If insertEnd < insertStart Then Exit Sub

    ws.Rows(CStr(insertStart) & ":" & CStr(insertEnd)).Insert Shift:=xlDown
End Sub

Private Sub mp_UnmergeSpacerRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal rowCount As Long)
    Dim endRow As Long

    If ws Is Nothing Then Exit Sub
    If rowCount <= 0 Then Exit Sub
    If startRow < 1 Then startRow = 1
    If startRow > ws.Rows.Count Then Exit Sub

    endRow = startRow + rowCount - 1
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count
    If endRow < startRow Then Exit Sub

    On Error Resume Next
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).UnMerge
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).RowHeight = ws.StandardHeight
    On Error GoTo 0
End Sub

Private Function mp_SanitizeNameToken(ByVal sourceText As String) As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long
    Dim resultText As String

    sourceText = Trim$(sourceText)
    If Len(sourceText) = 0 Then Exit Function

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        codePoint = AscW(ch)
        If (codePoint >= 48 And codePoint <= 57) _
            Or (codePoint >= 65 And codePoint <= 90) _
            Or (codePoint >= 97 And codePoint <= 122) _
            Or codePoint = 95 Then
            resultText = resultText & ch
        Else
            resultText = resultText & "_"
        End If
    Next i

    If Len(resultText) = 0 Then
        resultText = "anchor"
    ElseIf Mid$(resultText, 1, 1) Like "#" Then
        resultText = "_" & resultText
    End If

    mp_SanitizeNameToken = resultText
End Function

Private Function mp_ChecksumHex4(ByVal sourceText As String) As String
    Dim i As Long
    Dim checksum As Long

    checksum = 0
    For i = 1 To Len(sourceText)
        checksum = ((checksum * 31) + AscW(Mid$(sourceText, i, 1))) And &HFFFF&
    Next i

    mp_ChecksumHex4 = Right$("0000" & Hex$(checksum), 4)
End Function

Public Function m_CreateWarningBannersRuntimeLayer( _
    ByVal ws As Worksheet, _
    ByVal pendingWarningBanners As Collection, _
    Optional ByVal layerId As String = "runtime-warning-banners", _
    Optional ByVal priority As Long = 850 _
) As obj_StyleLayer
    Dim runtimeLayer As obj_StyleLayer
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim hasBannerStyle As Boolean
    Dim entry As Object
    Dim bannerRange As Range
    Dim ruleIndex As Long
    Dim rowCount As Long
    Dim startRow As Long
    Dim startCol As Long
    Dim endRow As Long
    Dim endCol As Long
    Dim declarations As Object
    Dim rowSelector As String
    Dim showGrid As Boolean
    Dim wrapTextEnabled As Boolean
    Dim rowHeightValue As Double
    Dim titleBold As Boolean
    Dim horizontalToken As String
    Dim verticalToken As String
    Dim backColorHex As String
    Dim fontColorHex As String
    Dim gridColorHex As String

    If ws Is Nothing Then Exit Function
    If pendingWarningBanners Is Nothing Then Exit Function
    If pendingWarningBanners.Count = 0 Then Exit Function

    hasBannerStyle = ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook)
    wrapTextEnabled = IIf(hasBannerStyle, bannerStyle.WrapText, True)
    rowHeightValue = IIf(hasBannerStyle, bannerStyle.RowHeight, 24#)
    titleBold = IIf(hasBannerStyle, bannerStyle.TitleBold, True)
    showGrid = IIf(hasBannerStyle, bannerStyle.ShowGrid, False)
    horizontalToken = mp_HAlignToToken(IIf(hasBannerStyle, bannerStyle.HorizontalAlignment, xlLeft))
    verticalToken = mp_VAlignToToken(IIf(hasBannerStyle, bannerStyle.VerticalAlignment, xlCenter))
    backColorHex = mp_ColorToHex(IIf(hasBannerStyle, bannerStyle.BackColor, RGB(76, 63, 16)))
    fontColorHex = mp_ColorToHex(IIf(hasBannerStyle, bannerStyle.FontColor, RGB(255, 229, 153)))
    gridColorHex = mp_ColorToHex(IIf(hasBannerStyle, bannerStyle.GridColor, RGB(80, 80, 80)))

    Set runtimeLayer = New obj_StyleLayer
    runtimeLayer.Initialize layerId, priority, "runtime", True

    For Each entry In pendingWarningBanners
        If entry Is Nothing Then GoTo ContinueEntry
        If Not entry.Exists("RangeAddress") Then GoTo ContinueEntry
        If Len(Trim$(CStr(entry("RangeAddress")))) = 0 Then GoTo ContinueEntry

        Set bannerRange = ws.Range(CStr(entry("RangeAddress")))
        rowCount = bannerRange.Rows.Count
        If hasBannerStyle Then
            If bannerStyle.Rows > rowCount Then rowCount = bannerStyle.Rows
        End If
        If rowCount < 2 Then rowCount = 2

        startRow = bannerRange.Row
        startCol = bannerRange.Column
        endRow = startRow + rowCount - 1
        endCol = startCol + bannerRange.Columns.Count - 1

        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("overflow") = IIf(wrapTextEnabled, "wrap", "clip")
        declarations("horizontal") = horizontalToken
        declarations("vertical") = verticalToken
        declarations("backColor") = backColorHex
        declarations("fontColor") = fontColorHex
        declarations("rowHeight") = mp_ToInvariantDoubleText(rowHeightValue)
        If showGrid Then
            declarations("borderColor") = gridColorHex
            declarations("borderWeight") = "thin"
        End If
        mp_AddRuntimeRangeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), mp_BuildAddress(startRow, startCol, endRow, endCol), declarations

        ruleIndex = ruleIndex + 1
        Set declarations = mp_CreateDeclarations()
        declarations("fontBold") = IIf(titleBold, "true", "false")
        mp_AddRuntimeRangeRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), mp_BuildAddress(startRow, startCol, startRow, endCol), declarations

        If wrapTextEnabled Then
            ruleIndex = ruleIndex + 1
            Set declarations = mp_CreateDeclarations()
            declarations("autoHeight") = "true"
            rowSelector = "row=" & CStr(startRow + 1) & ":" & CStr(startRow + 1) & ";col=" & CStr(startCol) & ":" & CStr(endCol)
            mp_AddRuntimeRowRule runtimeLayer, layerId & ".rule" & CStr(ruleIndex), rowSelector, declarations
        End If
ContinueEntry:
    Next entry

    If runtimeLayer.RuleCount = 0 Then Exit Function
    Set m_CreateWarningBannersRuntimeLayer = runtimeLayer
End Function

Private Sub mp_ApplyBannerRowHeights( _
    ByVal ws As Worksheet, _
    ByVal bannerRange As Range, _
    ByVal rowCount As Long, _
    ByVal baseRowHeight As Double, _
    ByVal messageRowIndex As Long, _
    ByVal messageText As String, _
    ByVal wrapTextEnabled As Boolean _
)
    Dim measuredHeight As Double
    Dim messageRange As Range
    Dim rowStart As Long
    Dim rowEnd As Long

    If ws Is Nothing Then Exit Sub
    If bannerRange Is Nothing Then Exit Sub
    If rowCount < 1 Then Exit Sub
    If baseRowHeight <= 0 Then baseRowHeight = 24

    rowStart = bannerRange.Row
    rowEnd = bannerRange.Row + rowCount - 1
    ws.Rows(CStr(rowStart) & ":" & CStr(rowEnd)).RowHeight = baseRowHeight

    If Not wrapTextEnabled Then Exit Sub
    If Len(Trim$(messageText)) = 0 Then Exit Sub
    If messageRowIndex < rowStart Or messageRowIndex > rowEnd Then Exit Sub

    Set messageRange = ws.Range( _
        ws.Cells(messageRowIndex, bannerRange.Column), _
        ws.Cells(messageRowIndex, bannerRange.Column + bannerRange.Columns.Count - 1) _
    )
    measuredHeight = mp_MeasureBannerTextHeight(ws, messageRange, messageText)
    If measuredHeight > baseRowHeight Then
        ws.Rows(messageRowIndex).RowHeight = measuredHeight
    End If
End Sub

Private Function mp_MeasureBannerTextHeight( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal messageText As String _
) As Double
    mp_MeasureBannerTextHeight = ex_SheetHelpers.m_MeasureTextHeight(ws, targetRange, messageText, 0, 0) + 2
End Function

Private Function mp_CreateDeclarations() As Object
    Set mp_CreateDeclarations = CreateObject("Scripting.Dictionary")
    mp_CreateDeclarations.CompareMode = 1
End Function

Private Sub mp_AddRuntimeRangeRule( _
    ByVal layer As obj_StyleLayer, _
    ByVal ruleId As String, _
    ByVal addressText As String, _
    ByVal declarations As Object _
)
    Dim ruleObj As obj_StyleRule

    If layer Is Nothing Then Exit Sub
    If declarations Is Nothing Then Exit Sub
    If Len(Trim$(ruleId)) = 0 Then Exit Sub
    If Len(Trim$(addressText)) = 0 Then Exit Sub

    Set ruleObj = New obj_StyleRule
    ruleObj.Initialize ruleId, "range", "address=" & addressText, declarations
    layer.AddRule ruleObj
End Sub

Private Sub mp_AddRuntimeRowRule( _
    ByVal layer As obj_StyleLayer, _
    ByVal ruleId As String, _
    ByVal selectorText As String, _
    ByVal declarations As Object _
)
    Dim ruleObj As obj_StyleRule

    If layer Is Nothing Then Exit Sub
    If declarations Is Nothing Then Exit Sub
    If Len(Trim$(ruleId)) = 0 Then Exit Sub
    If Len(Trim$(selectorText)) = 0 Then Exit Sub

    Set ruleObj = New obj_StyleRule
    ruleObj.Initialize ruleId, "row", selectorText, declarations
    layer.AddRule ruleObj
End Sub

Private Function mp_BuildAddress( _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
) As String
    If rowStart < 1 Then rowStart = 1
    If colStart < 1 Then colStart = 1
    If rowEnd < rowStart Then rowEnd = rowStart
    If colEnd < colStart Then colEnd = colStart

    mp_BuildAddress = mp_ToColumnLetter(colStart) & CStr(rowStart) & ":" & mp_ToColumnLetter(colEnd) & CStr(rowEnd)
End Function

Private Function mp_ToColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long

    If columnIndex < 1 Then columnIndex = 1
    n = columnIndex
    Do While n > 0
        remainder = (n - 1) Mod 26
        mp_ToColumnLetter = Chr$(65 + remainder) & mp_ToColumnLetter
        n = (n - remainder - 1) \ 26
    Loop
End Function

Private Function mp_ColorToHex(ByVal colorValue As Long) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = colorValue Mod 256
    g = (colorValue \ 256) Mod 256
    b = (colorValue \ 65536) Mod 256
    mp_ColorToHex = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

Private Function mp_ToInvariantDoubleText(ByVal value As Double) As String
    mp_ToInvariantDoubleText = Replace$(Trim$(CStr(value)), ",", ".")
End Function

Private Function mp_HAlignToToken(ByVal value As Long) As String
    Select Case value
        Case xlCenter, xlCenterAcrossSelection
            mp_HAlignToToken = "center"
        Case xlRight
            mp_HAlignToToken = "right"
        Case xlFill
            mp_HAlignToToken = "fill"
        Case xlJustify
            mp_HAlignToToken = "justify"
        Case xlDistributed
            mp_HAlignToToken = "distributed"
        Case Else
            mp_HAlignToToken = "left"
    End Select
End Function

Private Function mp_VAlignToToken(ByVal value As Long) As String
    Select Case value
        Case xlTop
            mp_VAlignToToken = "top"
        Case xlBottom
            mp_VAlignToToken = "bottom"
        Case xlJustify
            mp_VAlignToToken = "justify"
        Case xlDistributed
            mp_VAlignToToken = "distributed"
        Case Else
            mp_VAlignToToken = "center"
    End Select
End Function
