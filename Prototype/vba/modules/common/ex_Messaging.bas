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
Private Const DEFAULT_LOG_FILE_RELATIVE_PATH As String = "Logs\postprocess_debug.log"

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
    mp_EnsureLogParentFolder logPath

    lineText = Format$(Now, "yyyy-mm-dd HH:nn:ss") & " | " & CStr(messageText)
    fileNo = FreeFile
    Open logPath For Append As #fileNo
    Print #fileNo, lineText
    Close #fileNo

    m_LogToFile = logPath
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
    Optional ByVal bannerRangeAddress As String = "A1:H4" _
)
    Dim bannerRange As Range
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim hasBannerStyle As Boolean
    Dim rowCount As Long
    Dim messageText As String
    Dim rowOffset As Long

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(errDescription)
    If Len(messageText) = 0 Then
        messageText = "Unknown error."
    End If

    hasBannerStyle = ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook)
    If Len(Trim$(bannerRangeAddress)) = 0 Then
        bannerRangeAddress = ex_SheetStylesXmlProvider.m_GetOutputErrorBannerRangeAddress(ThisWorkbook)
    End If

    Set bannerRange = ws.Range(bannerRangeAddress)
    rowCount = bannerRange.Rows.Count
    If hasBannerStyle Then
        If bannerStyle.Rows > rowCount Then rowCount = bannerStyle.Rows
    End If
    If rowCount < 4 Then rowCount = 4

    bannerRange.ClearContents
    bannerRange.UnMerge
    For rowOffset = 0 To rowCount - 1
        ws.Range(ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column), ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    Next rowOffset

    ws.Cells(bannerRange.Row, bannerRange.Column).Value = titleText
    ws.Cells(bannerRange.Row + 1, bannerRange.Column).Value = messageText
    ws.Cells(bannerRange.Row + 2, bannerRange.Column).Value = "Source: " & IIf(Len(Trim$(errSource)) > 0, errSource, "n/a")
    ws.Cells(bannerRange.Row + 3, bannerRange.Column).Value = "Code: " & CStr(errNumber)

    With ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row + rowCount - 1, bannerRange.Column + bannerRange.Columns.Count - 1))
        .WrapText = IIf(hasBannerStyle, bannerStyle.WrapText, True)
        .VerticalAlignment = IIf(hasBannerStyle, bannerStyle.VerticalAlignment, xlCenter)
        .HorizontalAlignment = IIf(hasBannerStyle, bannerStyle.HorizontalAlignment, xlLeft)
        .Interior.Pattern = xlSolid
        .Interior.Color = IIf(hasBannerStyle, bannerStyle.BackColor, RGB(192, 0, 0))
        .Font.Color = IIf(hasBannerStyle, bannerStyle.FontColor, RGB(255, 255, 255))
        .Font.Bold = False

        If hasBannerStyle And bannerStyle.ShowGrid Then
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders.Color = bannerStyle.GridColor
            .Borders.Weight = xlThin
        ElseIf hasBannerStyle Then
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End If
    End With

    ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row, bannerRange.Column + bannerRange.Columns.Count - 1)).Font.Bold = IIf(hasBannerStyle, bannerStyle.TitleBold, True)
    mp_ApplyBannerRowHeights ws, bannerRange, rowCount, IIf(hasBannerStyle, bannerStyle.RowHeight, 24), bannerRange.Row + 1, messageText, IIf(hasBannerStyle, bannerStyle.WrapText, True)
End Sub

Public Sub m_RenderWarningBanner( _
    ByVal ws As Worksheet, _
    ByVal warningText As String, _
    Optional ByVal titleText As String = "WARNING", _
    Optional ByVal bannerRangeAddress As String = "A1:H3" _
)
    Dim bannerRange As Range
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim hasBannerStyle As Boolean
    Dim rowCount As Long
    Dim messageText As String
    Dim rowOffset As Long

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(warningText)
    If Len(messageText) = 0 Then
        messageText = "Action required."
    End If

    hasBannerStyle = ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook)
    If Len(Trim$(bannerRangeAddress)) = 0 Then
        bannerRangeAddress = ex_SheetStylesXmlProvider.m_GetOutputWarningBannerRangeAddress(ThisWorkbook)
    End If

    Set bannerRange = ws.Range(bannerRangeAddress)
    rowCount = bannerRange.Rows.Count
    If hasBannerStyle Then
        If bannerStyle.Rows > rowCount Then rowCount = bannerStyle.Rows
    End If
    If rowCount < 2 Then rowCount = 2

    bannerRange.ClearContents
    bannerRange.UnMerge
    For rowOffset = 0 To rowCount - 1
        ws.Range(ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column), ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    Next rowOffset

    ws.Cells(bannerRange.Row, bannerRange.Column).Value = titleText
    ws.Cells(bannerRange.Row + 1, bannerRange.Column).Value = messageText

    With ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row + rowCount - 1, bannerRange.Column + bannerRange.Columns.Count - 1))
        .WrapText = IIf(hasBannerStyle, bannerStyle.WrapText, True)
        .VerticalAlignment = IIf(hasBannerStyle, bannerStyle.VerticalAlignment, xlCenter)
        .HorizontalAlignment = IIf(hasBannerStyle, bannerStyle.HorizontalAlignment, xlLeft)
        .Interior.Pattern = xlSolid
        .Interior.Color = IIf(hasBannerStyle, bannerStyle.BackColor, RGB(76, 63, 16))
        .Font.Color = IIf(hasBannerStyle, bannerStyle.FontColor, RGB(255, 229, 153))
        .Font.Bold = False

        If hasBannerStyle And bannerStyle.ShowGrid Then
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders.Color = bannerStyle.GridColor
            .Borders.Weight = xlThin
        ElseIf hasBannerStyle Then
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End If
    End With

    ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row, bannerRange.Column + bannerRange.Columns.Count - 1)).Font.Bold = IIf(hasBannerStyle, bannerStyle.TitleBold, True)
    mp_ApplyBannerRowHeights ws, bannerRange, rowCount, IIf(hasBannerStyle, bannerStyle.RowHeight, 24), bannerRange.Row + 1, messageText, IIf(hasBannerStyle, bannerStyle.WrapText, True)
End Sub

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
    Dim textBoxShape As Object

    On Error GoTo EH
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    If Len(messageText) = 0 Then Exit Function

    Set textBoxShape = ws.Shapes.AddTextbox(1, targetRange.Left, targetRange.Top, targetRange.Width, 8)
    textBoxShape.Line.Visible = 0
    textBoxShape.Fill.Visible = 0
    textBoxShape.TextFrame2.MarginLeft = 0
    textBoxShape.TextFrame2.MarginRight = 0
    textBoxShape.TextFrame2.MarginTop = 0
    textBoxShape.TextFrame2.MarginBottom = 0
    textBoxShape.TextFrame2.WordWrap = -1
    textBoxShape.TextFrame2.AutoSize = 1
    textBoxShape.TextFrame2.TextRange.Text = messageText
    textBoxShape.TextFrame2.TextRange.Font.Size = targetRange.Font.Size
    textBoxShape.TextFrame2.TextRange.Font.Name = CStr(targetRange.Font.Name)

    mp_MeasureBannerTextHeight = textBoxShape.Height + 2

Cleanup:
    On Error Resume Next
    If Not textBoxShape Is Nothing Then textBoxShape.Delete
    On Error GoTo 0
    Exit Function

EH:
    mp_MeasureBannerTextHeight = 0
    Resume Cleanup
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
