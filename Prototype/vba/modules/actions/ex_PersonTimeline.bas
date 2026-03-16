Attribute VB_Name = "ex_PersonTimeline"
Option Explicit

Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_COL_MARKER As Long = 1
Private Const DEV_COL_KEY As Long = 2
Private Const DEV_COL_VALUE As Long = 3
Private Const RESULT_SHEET_NAME As String = "g_PersonTimeline"
Private Const POST_PROCESS_SCRIPT_KEY_IMPLICIT As String = "PostProcess.Script.Implicit"
Private Const POST_PROCESS_SCRIPT_KEY_EXPLICIT As String = "PostProcess.Script.Explicit"

Private g_LastPostProcessCfg As Object
Private g_LastPostProcessTables As Collection
Private g_LastResultHasPartialMatchCandidates As Boolean
Private g_AdoLookupCacheSignature As String
Private g_AdoResolvedTableRefByConfigured As Object
Private g_AdoAutoTableRefBySignature As Object
Private g_AdoFieldMapByTableRef As Object
Private g_AdoFieldListByTableRef As Object
Private g_AdoFieldGenericByTableRef As Object
Private g_AdoMarkerRangeRefBySignature As Object

Private Sub mp_EnsureAdoLookupCaches(ByVal cfg As Object)
    Dim signature As String

    signature = mp_BuildAdoLookupCacheSignature(cfg)
    If StrComp(g_AdoLookupCacheSignature, signature, vbBinaryCompare) <> 0 Then
        mp_ResetAdoLookupCaches
        g_AdoLookupCacheSignature = signature
        Exit Sub
    End If

    mp_EnsureAdoLookupCacheContainers
End Sub

Private Sub mp_ResetAdoLookupCaches()
    g_AdoLookupCacheSignature = vbNullString

    Set g_AdoResolvedTableRefByConfigured = Nothing
    Set g_AdoAutoTableRefBySignature = Nothing
    Set g_AdoFieldMapByTableRef = Nothing
    Set g_AdoFieldListByTableRef = Nothing
    Set g_AdoFieldGenericByTableRef = Nothing
    Set g_AdoMarkerRangeRefBySignature = Nothing
    ex_FetchDslEngine.m_ResetPlanCache

    mp_EnsureAdoLookupCacheContainers
End Sub

Private Sub mp_EnsureAdoLookupCacheContainers()
    If g_AdoResolvedTableRefByConfigured Is Nothing Then
        Set g_AdoResolvedTableRefByConfigured = CreateObject("Scripting.Dictionary")
        g_AdoResolvedTableRefByConfigured.CompareMode = 1
    End If
    If g_AdoAutoTableRefBySignature Is Nothing Then
        Set g_AdoAutoTableRefBySignature = CreateObject("Scripting.Dictionary")
        g_AdoAutoTableRefBySignature.CompareMode = 1
    End If
    If g_AdoFieldMapByTableRef Is Nothing Then
        Set g_AdoFieldMapByTableRef = CreateObject("Scripting.Dictionary")
        g_AdoFieldMapByTableRef.CompareMode = 1
    End If
    If g_AdoFieldListByTableRef Is Nothing Then
        Set g_AdoFieldListByTableRef = CreateObject("Scripting.Dictionary")
        g_AdoFieldListByTableRef.CompareMode = 1
    End If
    If g_AdoFieldGenericByTableRef Is Nothing Then
        Set g_AdoFieldGenericByTableRef = CreateObject("Scripting.Dictionary")
        g_AdoFieldGenericByTableRef.CompareMode = 1
    End If
    If g_AdoMarkerRangeRefBySignature Is Nothing Then
        Set g_AdoMarkerRangeRefBySignature = CreateObject("Scripting.Dictionary")
        g_AdoMarkerRangeRefBySignature.CompareMode = 1
    End If
End Sub

Private Function mp_BuildAdoLookupCacheSignature(ByVal cfg As Object) As String
    Dim outputAliases As Variant
    Dim fieldAliases As Variant
    Dim i As Long
    Dim j As Long
    Dim tableAlias As String
    Dim sourceAlias As String
    Dim sourcePrefix As String
    Dim keyAlias As String
    Dim columnsRaw As String
    Dim fieldAlias As String
    Dim signature As String

    On Error GoTo EH

    If cfg Is Nothing Then
        mp_BuildAdoLookupCacheSignature = "cfg:none"
        Exit Function
    End If

    If cfg.Count = 0 Then
        mp_BuildAdoLookupCacheSignature = "cfg:empty"
        Exit Function
    End If

    signature = "cfg:" & CStr(cfg.Count) & "|out=" & mp_GetCfgOptional(cfg, "Output.Sheets", vbNullString)
    outputAliases = mp_SplitList(mp_GetCfgOptional(cfg, "Output.Sheets", vbNullString))
    If mp_IsEmptyVariantArray(outputAliases) Then
        mp_BuildAdoLookupCacheSignature = signature
        Exit Function
    End If

    For i = LBound(outputAliases) To UBound(outputAliases)
        tableAlias = Trim$(CStr(outputAliases(i)))
        If Len(tableAlias) = 0 Then GoTo ContinueTable

        sourceAlias = mp_GetCfgOptional(cfg, "Output.Sheet[" & tableAlias & "].SourceAlias", vbNullString)
        If Len(sourceAlias) = 0 Then sourceAlias = tableAlias
        sourcePrefix = "Source." & sourceAlias

        keyAlias = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Key", vbNullString)
        columnsRaw = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Columns", vbNullString)
        signature = signature & "|tbl=" & sourceAlias & ":" & tableAlias & _
                    "|fp=" & mp_GetCfgOptional(cfg, sourcePrefix & ".FilePath", vbNullString) & _
                    "|fr=" & mp_GetCfgOptional(cfg, sourcePrefix & ".FileResolver", vbNullString) & _
                    "|fra=" & mp_GetCfgOptional(cfg, sourcePrefix & ".FileResolverArgs", vbNullString) & _
                    "|fpr=" & mp_GetSourcePathSignatureValue(cfg, sourceAlias) & _
                    "|hh=" & mp_GetCfgOptional(cfg, sourcePrefix & ".HasHeaders", vbNullString) & _
                    "|sn=" & mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].SheetName", vbNullString) & _
                    "|rsm=" & mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeStartMarker", vbNullString) & _
                    "|rem=" & mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeEndMarker", vbNullString) & _
                    "|k=" & keyAlias & _
                    "|c=" & columnsRaw

        If Len(keyAlias) > 0 Then
            signature = signature & "|m[" & keyAlias & "]=" & mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & keyAlias & "]", vbNullString)
        End If

        fieldAliases = mp_SplitList(columnsRaw)
        If Not mp_IsEmptyVariantArray(fieldAliases) Then
            For j = LBound(fieldAliases) To UBound(fieldAliases)
                fieldAlias = Trim$(CStr(fieldAliases(j)))
                If Len(fieldAlias) = 0 Then GoTo ContinueField
                signature = signature & "|m[" & fieldAlias & "]=" & mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]", vbNullString)
ContinueField:
            Next j
        End If
ContinueTable:
    Next i

    mp_BuildAdoLookupCacheSignature = signature
    Exit Function

EH:
    mp_BuildAdoLookupCacheSignature = "cfg:error|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Sub mp_SortVariantTextArray(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    If mp_IsEmptyVariantArray(arr) Then Exit Sub
    If UBound(arr) <= LBound(arr) Then Exit Sub

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Public Sub m_ShowPersonTimeline_UI()

    Dim fio As String

    fio = Trim$(ex_ConfigProvider.m_GetConfigValue("CommonKey", vbNullString))
    If Len(fio) = 0 Then
        fio = Trim$(ex_ConfigProvider.m_GetConfigValue("PersonFIO", vbNullString))
    End If

    m_ShowPersonTimeline fio

End Sub

Private Function mp_ValidateTimelineEntryConfig(ByRef outErrorText As String) As Boolean
    Dim cfg As Object
    Dim mode As OutputMode

    On Error GoTo EH

    Set cfg = mp_LoadConfigDictionary()
    mode = ex_Settings.m_GetOutputMode()

    mp_ValidateTimelineEntryConfig = mp_ValidateTimelineConfig(cfg, mode, outErrorText)
    Exit Function

EH:
    outErrorText = "Config validation failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    mp_ValidateTimelineEntryConfig = False
End Function

Public Sub m_ShowPersonTimeline(ByVal fio As String)

    On Error GoTo EH

    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevCalculateBeforeSave As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevCalculateBeforeSave = Application.CalculateBeforeSave

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = False

    Dim wsOut As Worksheet
    Dim resultSheetExistedBeforeRender As Boolean

    If Len(Trim$(fio)) = 0 Then
        Err.Raise vbObjectError + 1300, "ex_PersonTimeline", _
            "Config key 'CommonKey' (or fallback 'PersonFIO') is empty."
    End If

    Dim cfg As Object
    Set cfg = mp_LoadConfigDictionary()
    mp_EnsureAdoLookupCaches cfg

    Dim mode As OutputMode
    mode = ex_Settings.m_GetOutputMode()

    Dim validationError As String
    If Not mp_ValidateTimelineConfig(cfg, mode, validationError) Then
        Application.EnableEvents = prevEnableEvents
        Application.DisplayAlerts = prevDisplayAlerts
        Application.ScreenUpdating = prevScreenUpdating
        Application.CalculateBeforeSave = prevCalculateBeforeSave
        Application.Calculation = prevCalculation
        MsgBox validationError, vbExclamation
        Exit Sub
    End If

    resultSheetExistedBeforeRender = mp_WorksheetExists(RESULT_SHEET_NAME)
    Set wsOut = mp_CreateOrClearSheet(RESULT_SHEET_NAME)
    ex_SheetViewZoom.m_ApplyProfileZoomForResultSheet wsOut, resultSheetExistedBeforeRender
    ex_Messaging.m_ClearBannerAnchors wsOut
    ex_Messaging.m_ClearResultTableAnchors wsOut
    ex_Messaging.m_ClearResultRowAnchors wsOut

    Dim resultFieldRanges As Collection
    Set resultFieldRanges = New Collection
    Dim resultTables As Collection
    Set resultTables = New Collection
    Dim resultTablesByRef As Object
    Set resultTablesByRef = CreateObject("Scripting.Dictionary")
    resultTablesByRef.CompareMode = 1

    Dim outputStyle As t_OutputSheetStyle
    Dim hasOutputStyle As Boolean

    If Not ex_SheetStylesXmlProvider.m_EnsureInitialized(ThisWorkbook) Then
        Err.Raise vbObjectError + 1304, "ex_PersonTimeline", "Failed to initialize style registry."
    End If
    hasOutputStyle = ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook)

    Dim outputAliases As Variant
    outputAliases = mp_GetListRequired(cfg, "Output.Sheets")

    Dim tableSourceMap As Object
    Set tableSourceMap = CreateObject("Scripting.Dictionary")
    tableSourceMap.CompareMode = 1

    Dim connCache As Object
    Set connCache = CreateObject("Scripting.Dictionary")
    connCache.CompareMode = 1

    Dim rowIndex As Long
    rowIndex = 1
    If hasOutputStyle Then
        rowIndex = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    End If

    Dim headerRows As Collection
    Set headerRows = New Collection

    Dim sectionRows As Collection
    Set sectionRows = New Collection

    Dim pendingWarningBanners As Collection
    Set pendingWarningBanners = New Collection

    Dim partialMatchRowRanges As Collection
    Set partialMatchRowRanges = New Collection

    Dim renderedCount As Long
    renderedCount = 0

    Dim i As Long
    For i = LBound(outputAliases) To UBound(outputAliases)
        Dim tableAlias As String
        tableAlias = Trim$(CStr(outputAliases(i)))
        If Len(tableAlias) = 0 Then
            GoTo ContinueAlias
        End If

        Dim sourceAlias As String
        sourceAlias = mp_GetSourceAliasCached(cfg, tableAlias, tableSourceMap)

        Dim tableType As String
        tableType = LCase$(mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Type"))

        If Not mp_IsSupportedOutputTableType(tableType) Then
            Err.Raise vbObjectError + 1301, "ex_PersonTimeline", _
                "Unsupported table type for alias '" & tableAlias & "': " & tableType
        End If
        If Not mp_ShouldRenderTableForMode(mode, tableType) Then
            GoTo ContinueAlias
        End If

        Dim adoObjectName As String
        adoObjectName = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].SheetName", vbNullString)

        Dim sourceConn As Object
        Set sourceConn = mp_GetConnectionForSource(connCache, cfg, sourceAlias)

        If tableType = "state" Then
            Dim stateRendered As Boolean
            rowIndex = mp_WriteStateCardGeneric(wsOut, sourceConn, adoObjectName, fio, rowIndex, cfg, resultFieldRanges, resultTables, resultTablesByRef, sourceAlias, tableAlias, headerRows, sectionRows, pendingWarningBanners, partialMatchRowRanges, stateRendered)
            If stateRendered Then
                rowIndex = mp_AdvanceRowIndexAfterRenderedTable(cfg, outputAliases, i, mode, tableSourceMap, tableAlias, tableType, rowIndex)
                renderedCount = renderedCount + 1
            End If
        Else
            Dim eventsSectionRow As Long
            Dim eventsSectionAdded As Boolean
            Dim eventsRendered As Boolean

            eventsSectionRow = rowIndex

            If mode <> StateTableOnly Then
                wsOut.Cells(rowIndex, 1).Value = "Events [" & tableAlias & "]"
                wsOut.Cells(rowIndex, 1).Font.Bold = True
                sectionRows.Add rowIndex
                eventsSectionAdded = True
                rowIndex = rowIndex + 1
            End If
            rowIndex = mp_WriteEventsGeneric(wsOut, sourceConn, adoObjectName, fio, rowIndex, cfg, resultFieldRanges, resultTables, resultTablesByRef, sourceAlias, tableAlias, headerRows, pendingWarningBanners, partialMatchRowRanges, eventsRendered)
            If eventsRendered Then
                rowIndex = mp_AdvanceRowIndexAfterRenderedTable(cfg, outputAliases, i, mode, tableSourceMap, tableAlias, tableType, rowIndex)
                renderedCount = renderedCount + 1
            Else
                If eventsSectionAdded Then
                    wsOut.Rows(eventsSectionRow).ClearContents
                    sectionRows.Remove sectionRows.Count
                End If
                rowIndex = eventsSectionRow
            End If
        End If

ContinueAlias:
    Next i

    If renderedCount = 0 Then
        Err.Raise vbObjectError + 1303, "ex_PersonTimeline", _
            "No sheets were rendered for mode '" & ex_Settings.m_GetOutputModeDisplay() & "'. Check Output.Sheets and sheet Type."
    End If

    If hasOutputStyle Then
        ex_OutputPanel.m_RenderForSheet wsOut, outputStyle
    End If
    mp_ApplyTimelineStyleLayers wsOut, headerRows, sectionRows, resultFieldRanges, resultTables, partialMatchRowRanges, hasOutputStyle, outputStyle, pendingWarningBanners
    mp_RenderPendingWarningBanners wsOut, pendingWarningBanners

    mp_StorePostProcessContext cfg, resultTables, (partialMatchRowRanges.Count > 0)
    If partialMatchRowRanges.Count = 0 Then
        ex_PostProcessDsl.m_ApplyScriptToSheet wsOut, cfg, resultTables, POST_PROCESS_SCRIPT_KEY_IMPLICIT
    End If

    mp_CloseConnections connCache

    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    Application.CalculateBeforeSave = prevCalculateBeforeSave
    Application.Calculation = prevCalculation

    Exit Sub

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errOutputStyle As t_OutputSheetStyle

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    mp_CloseConnections connCache
    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    Application.CalculateBeforeSave = prevCalculateBeforeSave
    Application.Calculation = prevCalculation
    On Error GoTo 0

    If mp_IsConfigValidationError(errNumber) Then
        MsgBox errDescription, vbExclamation
        Exit Sub
    End If

    If wsOut Is Nothing Then
        resultSheetExistedBeforeRender = mp_WorksheetExists(RESULT_SHEET_NAME)
        Set wsOut = mp_CreateOrClearSheet(RESULT_SHEET_NAME)
        ex_SheetViewZoom.m_ApplyProfileZoomForResultSheet wsOut, resultSheetExistedBeforeRender
    End If
    On Error Resume Next
    ex_OutputFormattingPipeline.m_ApplySheetPipeline wsOut
    If ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(errOutputStyle, ThisWorkbook) Then
        ex_OutputPanel.m_RenderForSheet wsOut, errOutputStyle
    End If
    On Error GoTo 0
    ex_Messaging.m_RenderErrorBanner wsOut, errDescription, errSource, errNumber, "ERROR: Timeline generation failed", ex_SheetStylesXmlProvider.m_GetOutputErrorBannerRangeAddress(ThisWorkbook)

End Sub

Public Sub m_RunPostProcessForActiveSheet()
    Dim ws As Worksheet
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim prevScreenUpdating As Boolean
    Dim outputStyle As t_OutputSheetStyle
    Dim hasOutputStyle As Boolean

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Active sheet is not available for post-process.", vbExclamation
        Exit Sub
    End If

    If g_LastPostProcessCfg Is Nothing Or g_LastPostProcessTables Is Nothing Then
        MsgBox "Post-process context is not prepared. Generate Personal Card result first.", vbExclamation
        Exit Sub
    End If

    If g_LastResultHasPartialMatchCandidates Then
        MsgBox "Post-process is unavailable for partial-match candidate output. Select the full key from candidates and run search again.", vbExclamation
        Exit Sub
    End If

    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    hasOutputStyle = ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook)
    If hasOutputStyle Then
        ex_OutputPanel.m_DeletePanelButtonsForSheet ws
    End If

    ex_PostProcessDsl.m_ApplyScriptToSheet ws, g_LastPostProcessCfg, g_LastPostProcessTables, POST_PROCESS_SCRIPT_KEY_EXPLICIT
    If hasOutputStyle Then
        ex_OutputPanel.m_RenderForSheet ws, outputStyle
    End If
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    If hasOutputStyle And Not ws Is Nothing Then ex_OutputPanel.m_RenderForSheet ws, outputStyle
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    MsgBox "Post-process failed: [" & errSource & " #" & CStr(errNumber) & "] " & errDescription, vbExclamation
End Sub

Public Sub m_ResetResultPageSessionState()
    mp_ResetAdoLookupCaches
    ex_PostProcessDsl.m_ResetScriptCache
    Set g_LastPostProcessCfg = Nothing
    Set g_LastPostProcessTables = Nothing
    g_LastResultHasPartialMatchCandidates = False
    ex_SheetViewZoom.m_ResetZoomCache RESULT_SHEET_NAME
End Sub

Private Sub mp_StorePostProcessContext(ByVal cfg As Object, ByVal resultTables As Collection, ByVal hasPartialMatchCandidates As Boolean)
    Set g_LastPostProcessCfg = cfg
    Set g_LastPostProcessTables = resultTables
    g_LastResultHasPartialMatchCandidates = hasPartialMatchCandidates
End Sub

Private Function mp_ShouldStrictPreflightValidation() As Boolean
    Dim valueText As String
    Dim parsed As Boolean

    valueText = ex_XmlCore.m_GetSettingsValue("st_StrictPreflightValidation", "false")
    If ex_XmlCore.m_TryParseBoolean(valueText, parsed) Then
        mp_ShouldStrictPreflightValidation = parsed
    Else
        mp_ShouldStrictPreflightValidation = False
    End If
End Function

Private Function mp_ValidateTimelineConfig( _
    ByVal cfg As Object, _
    ByVal mode As OutputMode, _
    ByRef outErrorText As String _
) As Boolean
    Dim outputAliases As Variant
    Dim tableSourceMap As Object
    Dim resultFieldRanges As Collection
    Dim activeModeKey As String
    Dim strictPreflight As Boolean
    Dim i As Long

    On Error GoTo EH

    outputAliases = mp_GetListRequired(cfg, "Output.Sheets")

    Set tableSourceMap = CreateObject("Scripting.Dictionary")
    tableSourceMap.CompareMode = 1

    Set resultFieldRanges = New Collection

    For i = LBound(outputAliases) To UBound(outputAliases)
        Dim tableAlias As String
        Dim sourceAlias As String
        Dim tableType As String
        Dim fields As Variant
    Dim fieldIndex As Long
    Dim fieldAlias As String
    Dim isVirtualField As Boolean

        tableAlias = Trim$(CStr(outputAliases(i)))
        If Len(tableAlias) = 0 Then GoTo ContinueAlias

        sourceAlias = mp_GetSourceAliasCached(cfg, tableAlias, tableSourceMap)
        tableType = LCase$(mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Type"))

        If Not mp_IsSupportedOutputTableType(tableType) Then
            Err.Raise vbObjectError + 1301, "ex_PersonTimeline", _
                "Unsupported table type for alias '" & tableAlias & "': " & tableType
        End If
        If Not mp_ShouldRenderTableForMode(mode, tableType) Then GoTo ContinueAlias

        Call mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].SheetName")
        Call mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Key")
        Dim rangeStartMarker As String
        Dim rangeEndMarker As String
        Dim configuredSheetName As String
        configuredSheetName = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].SheetName")
        rangeStartMarker = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeStartMarker", vbNullString)
        rangeEndMarker = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeEndMarker", vbNullString)
        If (Len(rangeStartMarker) > 0 Xor Len(rangeEndMarker) > 0) Then
            Err.Raise vbObjectError + 1747, "ex_PersonTimeline", _
                "Both markers must be provided for auto-range mode: '" & _
                sourceAlias & ".Sheet[" & tableAlias & "].RangeStartMarker' and '" & _
                sourceAlias & ".Sheet[" & tableAlias & "].RangeEndMarker'."
        End If
        If Len(rangeStartMarker) > 0 Then
            If mp_IsExplicitAdoRangeReference(configuredSheetName) Then
                Err.Raise vbObjectError + 1748, "ex_PersonTimeline", _
                    "Auto-range markers are not allowed with explicit range SheetName for " & _
                    sourceAlias & ".Sheet[" & tableAlias & "].SheetName."
            End If
        End If

        fields = mp_GetEffectiveFieldAliases(cfg, sourceAlias, tableAlias)

        For fieldIndex = LBound(fields) To UBound(fields)
            fieldAlias = Trim$(CStr(fields(fieldIndex)))
            If Len(fieldAlias) = 0 Then GoTo ContinueField

                isVirtualField = mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias)
                If Not isVirtualField Then
                    Call mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldAlias)
                End If
                If isVirtualField Then
                    mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, fieldAlias, 1, 1, 1, ex_FetchDslEngine.m_GetGeneratedKindValue()
                Else
                    mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, fieldAlias, 1, 1, 1
                End If
ContinueField:
        Next fieldIndex

ContinueAlias:
    Next i

    strictPreflight = mp_ShouldStrictPreflightValidation()
    If strictPreflight Then
        activeModeKey = ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev)
        If Not ex_StylePipelineEngine.m_ValidateColumnStylesPipeline( _
            resultFieldRanges, _
            Nothing, _
            activeModeKey, _
            outErrorText, _
            ThisWorkbook, _
            RESULT_SHEET_NAME _
        ) Then
            Exit Function
        End If

    End If

    mp_ValidateTimelineConfig = True
    Exit Function

EH:
    outErrorText = "Config validation failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    mp_ValidateTimelineConfig = False
End Function

Private Function mp_IsConfigValidationError(ByVal errNumber As Long) As Boolean
    Select Case errNumber
        Case vbObjectError + 1300, vbObjectError + 1301, vbObjectError + 1335, _
             vbObjectError + 1340, vbObjectError + 1341, vbObjectError + 1360, _
             vbObjectError + 1370, vbObjectError + 1371, vbObjectError + 1380, _
             vbObjectError + 1390, vbObjectError + 1491, vbObjectError + 1492, _
             vbObjectError + 1710, vbObjectError + 1711, vbObjectError + 1712, _
             vbObjectError + 1713, vbObjectError + 1714, vbObjectError + 1715, _
             vbObjectError + 1716, vbObjectError + 1717, vbObjectError + 1718, _
             vbObjectError + 1719, vbObjectError + 1734, vbObjectError + 1735, _
             vbObjectError + 1736, vbObjectError + 1737, vbObjectError + 1738, _
             vbObjectError + 1739, vbObjectError + 1740, vbObjectError + 1741, _
             vbObjectError + 1742, vbObjectError + 1743, vbObjectError + 1744, _
             vbObjectError + 1745, vbObjectError + 1746, vbObjectError + 1747, _
             vbObjectError + 1748, vbObjectError + 1749, vbObjectError + 1750, _
             vbObjectError + 1751, vbObjectError + 1752, vbObjectError + 1753, _
             vbObjectError + 1754, vbObjectError + 1755, vbObjectError + 1756, _
             vbObjectError + 1590, vbObjectError + 1591, vbObjectError + 1592, _
             vbObjectError + 1593, vbObjectError + 1594, vbObjectError + 1595, _
             vbObjectError + 1596, vbObjectError + 1597, _
             vbObjectError + 1493, vbObjectError + 1494
            mp_IsConfigValidationError = True
    End Select
End Function

Private Function mp_WriteStateCardGeneric( _
    ByVal wsOut As Worksheet, _
    ByVal adoConn As Object, _
    ByVal adoObjectName As String, _
    ByVal fio As String, _
    ByVal rowIndex As Long, _
    ByVal cfg As Object, _
    ByVal resultFieldRanges As Collection, _
    ByVal resultTables As Collection, _
    ByVal resultTablesByRef As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByVal pendingWarningBanners As Collection, _
    ByVal partialMatchRowRanges As Collection, _
    ByRef outTableRendered As Boolean _
) As Long
    On Error GoTo EH

    Dim rs As Object
    Dim sql As String
    Dim tableRef As String
    Dim expectedHeaders As Variant
    Dim stateRows As Variant
    Dim rowCount As Long
    Dim stepName As String
    Dim showNoStateRow As Boolean

    Dim sourceFields As Variant
    sourceFields = mp_GetOrderedFieldAliases(cfg, sourceAlias, tableAlias)
    Dim fields As Variant
    fields = mp_GetEffectiveFieldAliases(cfg, sourceAlias, tableAlias)
    Dim resultTable As obj_ResultTable
    Set resultTable = mp_EnsureResultTable(resultTables, resultTablesByRef, sourceAlias, tableAlias)
    mp_RegisterResultTableFieldAliases resultTable, sourceAlias, tableAlias, fields

    showNoStateRow = mp_IsLikelyFullPersonKey(fio)
    outTableRendered = False
    mp_WriteStateCardGeneric = rowIndex

    Dim keyAlias As String
    keyAlias = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Key")

    Dim keyHeader As String
    keyHeader = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, keyAlias)

    expectedHeaders = mp_BuildExpectedHeaders(cfg, sourceAlias, tableAlias, sourceFields, keyHeader)
    tableRef = mp_GetAdoTableReference(cfg, adoConn, adoObjectName, expectedHeaders, keyHeader, fio, sourceAlias, tableAlias)
    stepName = "resolve-key-header"
    keyHeader = mp_ResolveAdoMappedHeader(cfg, sourceAlias, tableAlias, keyAlias, adoConn, tableRef)

    stepName = "open-exact-recordset"
    sql = "SELECT * FROM " & tableRef & " WHERE " & mp_BuildAdoWhereEquals(keyHeader, fio)
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, adoConn, 0, 1

    Dim headerRow As Long
    Dim valueRow As Long
    Dim dataEndRow As Long
    Dim keyLabel As String
    keyLabel = mp_GetLabel(cfg, sourceAlias, tableAlias, keyAlias)
    Dim fieldCount As Long
    fieldCount = UBound(fields) - LBound(fields) + 1
    Dim i As Long
    Dim outCol As Long
    Dim fieldAlias As String
    Dim sourceHeader As String
    Dim fieldOrdinals() As Long
    Dim outValues() As Variant
    Dim outRow As Long
    Dim fetchKindsByOutRow As Object
    Dim fetchKindsBySheetRow As Object

    If Not rs.EOF Then
        stepName = "fetch-exact-rows"
        stateRows = rs.GetRows
        rowCount = UBound(stateRows, 2) - LBound(stateRows, 2) + 1

        wsOut.Cells(rowIndex, 1).Value = fio
        sectionRows.Add rowIndex
        rowIndex = rowIndex + 1

        headerRow = rowIndex
        headerRows.Add headerRow

        valueRow = headerRow + 1

        ReDim fieldOrdinals(LBound(fields) To UBound(fields))
        For i = LBound(fields) To UBound(fields)
            fieldAlias = Trim$(CStr(fields(i)))
            If Len(fieldAlias) = 0 Then GoTo ContinueExactField

            outCol = 1 + (i - LBound(fields))
            wsOut.Cells(headerRow, outCol).Value = mp_GetFieldLabel(cfg, sourceAlias, tableAlias, fieldAlias)
            If mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then
                fieldOrdinals(i) = -2
            Else
                sourceHeader = mp_ResolveAdoMappedHeader(cfg, sourceAlias, tableAlias, fieldAlias, adoConn, tableRef)
                fieldOrdinals(i) = mp_RecordsetGetFieldOrdinal(rs, sourceHeader)
            End If
ContinueExactField:
        Next i

        ReDim outValues(1 To rowCount, 1 To fieldCount)
        For outRow = 1 To rowCount
            For i = LBound(fields) To UBound(fields)
                fieldAlias = Trim$(CStr(fields(i)))
                If mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then
                    outValues(outRow, 1 + (i - LBound(fields))) = vbNullString
                ElseIf fieldOrdinals(i) >= 0 Then
                    outValues(outRow, 1 + (i - LBound(fields))) = mp_ToCellValue(stateRows(fieldOrdinals(i), outRow - 1))
                Else
                    outValues(outRow, 1 + (i - LBound(fields))) = "(missing column)"
                End If
            Next i
        Next outRow

        stepName = "collect-fetch-metadata"
        mp_AppendFetchRowsFromSource cfg, sourceAlias, tableAlias, adoConn, tableRef, keyHeader, fio, fields, outValues, rowCount, fieldCount, fetchKindsByOutRow

        dataEndRow = valueRow + rowCount - 1
        mp_AddResultFieldRangesForFields resultFieldRanges, cfg, sourceAlias, tableAlias, fields, headerRow, dataEndRow

        wsOut.Range(wsOut.Cells(valueRow, 1), wsOut.Cells(dataEndRow, fieldCount)).Value2 = outValues
        Set fetchKindsBySheetRow = mp_BuildSheetRowKindsMap(fetchKindsByOutRow, valueRow)
        mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, fields, valueRow, dataEndRow, fetchKindsBySheetRow, headerRow, dataEndRow

        rs.Close
        outTableRendered = True
        mp_WriteStateCardGeneric = dataEndRow + 1
        Exit Function
    End If

    rs.Close
    Set rs = Nothing

    stepName = "open-partial-recordset"
    sql = "SELECT DISTINCT " & mp_QuoteSqlIdentifier(keyHeader) & _
          " FROM " & tableRef & _
          " WHERE " & mp_BuildAdoWhereLike(keyHeader, fio) & _
          " ORDER BY " & mp_QuoteSqlIdentifier(keyHeader)
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, adoConn, 0, 1

    If rs.EOF Then
        rs.Close
        If showNoStateRow Then
            wsOut.Cells(rowIndex, 1).Value = fio
            sectionRows.Add rowIndex
            rowIndex = rowIndex + 1
            mp_WriteStateCardGeneric = mp_RenderStateNoData(wsOut, rowIndex, cfg, sourceAlias, tableAlias, fields, headerRows, resultFieldRanges, resultTable)
            outTableRendered = True
        End If
        Exit Function
    End If

    Dim partialRows As Variant
    Dim partialCount As Long
    Dim partialValues() As Variant
    Dim keyOnlyFields(0 To 0) As String
    Dim candidateIndex As Long

    partialRows = rs.GetRows
    partialCount = UBound(partialRows, 2) - LBound(partialRows, 2) + 1

    wsOut.Cells(rowIndex, 1).Value = "Candidates [State " & tableAlias & "] (" & CStr(partialCount) & ")"
    sectionRows.Add rowIndex
    rowIndex = rowIndex + 1
    rowIndex = mp_RenderStateCandidatesWarningBanner(wsOut, rowIndex, fio, partialCount, pendingWarningBanners)

    headerRow = rowIndex
    headerRows.Add headerRow
    wsOut.Cells(headerRow, 1).Value = keyLabel

    valueRow = headerRow + 1
    dataEndRow = valueRow + partialCount - 1

    ReDim partialValues(1 To partialCount, 1 To 1)
    keyOnlyFields(0) = keyAlias
    For candidateIndex = 1 To partialCount
        partialValues(candidateIndex, 1) = mp_ToCellValue(partialRows(0, candidateIndex - 1))
    Next candidateIndex

    wsOut.Range(wsOut.Cells(valueRow, 1), wsOut.Cells(dataEndRow, 1)).Value2 = partialValues
    mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, keyAlias, 1, headerRow, dataEndRow
    mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, keyOnlyFields, valueRow, dataEndRow, visualRowStart:=headerRow, visualRowEnd:=dataEndRow
    mp_AddPartialMatchRowRange partialMatchRowRanges, headerRow, dataEndRow

    rs.Close
    outTableRendered = True
    mp_WriteStateCardGeneric = dataEndRow + 1
    Exit Function

EH:
    Dim innerErrNumber As Long
    Dim innerErrDescription As String
    innerErrNumber = Err.Number
    innerErrDescription = Err.Description

    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 1319, "ex_PersonTimeline.mp_WriteStateCardGeneric", _
        "Failed for table alias '" & tableAlias & "' (source '" & sourceAlias & "') at step '" & stepName & "'. " & _
        "SQL=[" & sql & "]. InnerError #" & CStr(innerErrNumber) & ": " & innerErrDescription
End Function

Private Function mp_WriteEventsGeneric( _
    ByVal wsOut As Worksheet, _
    ByVal adoConn As Object, _
    ByVal adoObjectName As String, _
    ByVal fio As String, _
    ByVal rowIndex As Long, _
    ByVal cfg As Object, _
    ByVal resultFieldRanges As Collection, _
    ByVal resultTables As Collection, _
    ByVal resultTablesByRef As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal headerRows As Collection, _
    ByVal pendingWarningBanners As Collection, _
    ByVal partialMatchRowRanges As Collection, _
    ByRef outTableRendered As Boolean _
) As Long
    On Error GoTo EH

    Dim rs As Object
    Dim sql As String
    Dim tableRef As String
    Dim sourceHeader As String
    Dim fieldData As Variant
    Dim rowCount As Long
    Dim expectedHeaders As Variant
    Dim stepName As String
    Dim showNoEventsRow As Boolean

    Dim sourceFields As Variant
    sourceFields = mp_GetOrderedFieldAliases(cfg, sourceAlias, tableAlias)
    Dim fields As Variant
    fields = mp_GetEffectiveFieldAliases(cfg, sourceAlias, tableAlias)
    Dim resultTable As obj_ResultTable
    Set resultTable = mp_EnsureResultTable(resultTables, resultTablesByRef, sourceAlias, tableAlias)
    mp_RegisterResultTableFieldAliases resultTable, sourceAlias, tableAlias, fields

    Dim fieldCount As Long
    fieldCount = UBound(fields) - LBound(fields) + 1

    showNoEventsRow = mp_IsLikelyFullPersonKey(fio)
    outTableRendered = False
    mp_WriteEventsGeneric = rowIndex

    Dim keyAlias As String
    keyAlias = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Key")

    Dim keyHeader As String
    keyHeader = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, keyAlias)

    expectedHeaders = mp_BuildExpectedHeaders(cfg, sourceAlias, tableAlias, sourceFields, keyHeader)
    stepName = "resolve-table"
    tableRef = mp_GetAdoTableReference(cfg, adoConn, adoObjectName, expectedHeaders, keyHeader, fio, sourceAlias, tableAlias)
    stepName = "resolve-key-header"
    keyHeader = mp_ResolveAdoMappedHeader(cfg, sourceAlias, tableAlias, keyAlias, adoConn, tableRef)

    stepName = "open-exact-recordset"
    sql = "SELECT * FROM " & tableRef & " WHERE " & mp_BuildAdoWhereEquals(keyHeader, fio)
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, adoConn, 0, 1

    Dim outHeaderRow As Long
    Dim outDataRow As Long
    Dim i As Long
    Dim fieldAlias As String
    Dim headerValues() As Variant
    Dim fieldOrdinals() As Long
    Dim keyDataEndRow As Long
    Dim partialValues() As Variant
    Dim keyOnlyFields(0 To 0) As String
    Dim candidateIndex As Long
    Dim partialRows As Variant
    Dim partialCount As Long
    Dim fetchDslApplied As Boolean
    Dim fetchKindsByOutRow As Object
    Dim fetchKindsBySheetRow As Object

    If rs.EOF Then
        rs.Close
        Set rs = Nothing

        stepName = "open-partial-recordset"
        sql = "SELECT DISTINCT " & mp_QuoteSqlIdentifier(keyHeader) & _
              " FROM " & tableRef & _
              " WHERE " & mp_BuildAdoWhereLike(keyHeader, fio) & _
              " ORDER BY " & mp_QuoteSqlIdentifier(keyHeader)
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open sql, adoConn, 0, 1

        If rs.EOF Then
            rs.Close
            If showNoEventsRow Then
                mp_WriteEventsGeneric = mp_RenderEventsNoData(wsOut, rowIndex, cfg, sourceAlias, tableAlias, fields, headerRows, resultFieldRanges, resultTable)
                outTableRendered = True
            End If
            Exit Function
        End If

        partialRows = rs.GetRows
        partialCount = UBound(partialRows, 2) - LBound(partialRows, 2) + 1
        rs.Close

        rowIndex = mp_RenderStateCandidatesWarningBanner(wsOut, rowIndex, fio, partialCount, pendingWarningBanners)

        outHeaderRow = rowIndex
        headerRows.Add outHeaderRow
        wsOut.Cells(outHeaderRow, 1).Value = mp_GetLabel(cfg, sourceAlias, tableAlias, keyAlias)

        outDataRow = outHeaderRow + 1
        keyDataEndRow = outDataRow + partialCount - 1

        ReDim partialValues(1 To partialCount, 1 To 1)
        keyOnlyFields(0) = keyAlias
        For candidateIndex = 1 To partialCount
            partialValues(candidateIndex, 1) = mp_ToCellValue(partialRows(0, candidateIndex - 1))
        Next candidateIndex

        wsOut.Range(wsOut.Cells(outDataRow, 1), wsOut.Cells(keyDataEndRow, 1)).Value2 = partialValues
        mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, keyAlias, 1, outHeaderRow, keyDataEndRow
        mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, keyOnlyFields, outDataRow, keyDataEndRow, visualRowStart:=outHeaderRow, visualRowEnd:=keyDataEndRow
        mp_AddPartialMatchRowRange partialMatchRowRanges, outHeaderRow, keyDataEndRow

        outTableRendered = True
        mp_WriteEventsGeneric = keyDataEndRow + 1
        Exit Function
    End If

    stepName = "fetch-exact-rows"
    fieldData = rs.GetRows
    rowCount = UBound(fieldData, 2) - LBound(fieldData, 2) + 1

    outHeaderRow = rowIndex
    headerRows.Add outHeaderRow

    ReDim headerValues(1 To 1, 1 To fieldCount)
    ReDim fieldOrdinals(LBound(fields) To UBound(fields))

    For i = LBound(fields) To UBound(fields)
        fieldAlias = Trim$(CStr(fields(i)))
        headerValues(1, 1 + (i - LBound(fields))) = mp_GetFieldLabel(cfg, sourceAlias, tableAlias, fieldAlias)
        If mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then
            fieldOrdinals(i) = -2
        Else
            sourceHeader = mp_ResolveAdoMappedHeader(cfg, sourceAlias, tableAlias, fieldAlias, adoConn, tableRef)
            fieldOrdinals(i) = mp_RecordsetGetFieldOrdinal(rs, sourceHeader)
        End If
    Next i
    wsOut.Range(wsOut.Cells(outHeaderRow, 1), wsOut.Cells(outHeaderRow, fieldCount)).Value = headerValues

    outDataRow = outHeaderRow + 1

    Dim outValues() As Variant
    Dim outIndex As Long
    ReDim outValues(1 To rowCount, 1 To fieldCount)

    For outIndex = 1 To rowCount
        For i = LBound(fields) To UBound(fields)
            fieldAlias = Trim$(CStr(fields(i)))
            If mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then
                outValues(outIndex, 1 + (i - LBound(fields))) = vbNullString
            ElseIf fieldOrdinals(i) >= 0 Then
                outValues(outIndex, 1 + (i - LBound(fields))) = mp_ToCellValue(fieldData(fieldOrdinals(i), outIndex - 1))
            Else
                outValues(outIndex, 1 + (i - LBound(fields))) = "(missing column)"
            End If
        Next i
    Next outIndex

    stepName = "collect-fetch-metadata"
    fetchDslApplied = mp_AppendFetchRowsFromSource( _
        cfg, sourceAlias, tableAlias, adoConn, tableRef, keyHeader, fio, fields, outValues, rowCount, fieldCount, fetchKindsByOutRow)

    stepName = "write-output"
    wsOut.Range(wsOut.Cells(outDataRow, 1), wsOut.Cells(outDataRow + rowCount - 1, fieldCount)).Value2 = outValues
    outDataRow = outDataRow + rowCount
    rs.Close

    Dim sortAlias As String
    sortAlias = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Sort", vbNullString)

    If Len(sortAlias) > 0 And Not fetchDslApplied Then
        Dim sortOutCol As Long
        sortOutCol = -1

        For i = LBound(fields) To UBound(fields)
            If StrComp(Trim$(CStr(fields(i))), sortAlias, vbTextCompare) = 0 Then
                sortOutCol = 1 + (i - LBound(fields))
                Exit For
            End If
        Next i

        If sortOutCol > 0 Then
            On Error GoTo SortEH
            mp_NormalizeDateColumn wsOut, outHeaderRow + 1, outDataRow - 1, sortOutCol
            mp_SortRangeByColumnIndex wsOut, outHeaderRow, outDataRow - 1, 1, fieldCount, sortOutCol
            On Error GoTo EH
        End If
    End If

    mp_AddResultFieldRangesForFields resultFieldRanges, cfg, sourceAlias, tableAlias, fields, outHeaderRow, outDataRow - 1
    Set fetchKindsBySheetRow = mp_BuildSheetRowKindsMap(fetchKindsByOutRow, outHeaderRow + 1)
    mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, fields, outHeaderRow + 1, outDataRow - 1, fetchKindsBySheetRow, outHeaderRow, outDataRow - 1

    outTableRendered = True
    mp_WriteEventsGeneric = outDataRow
    Exit Function

SortEH:
    Err.Clear
    mp_AddResultFieldRangesForFields resultFieldRanges, cfg, sourceAlias, tableAlias, fields, outHeaderRow, outDataRow - 1
    Set fetchKindsBySheetRow = mp_BuildSheetRowKindsMap(fetchKindsByOutRow, outHeaderRow + 1)
    mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, fields, outHeaderRow + 1, outDataRow - 1, fetchKindsBySheetRow, outHeaderRow, outDataRow - 1
    On Error GoTo EH
    outTableRendered = True
    mp_WriteEventsGeneric = outDataRow
    Exit Function

EH:
    Dim innerErrNumber As Long
    Dim innerErrDescription As String
    innerErrNumber = Err.Number
    innerErrDescription = Err.Description

    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 1329, "ex_PersonTimeline.mp_WriteEventsGeneric", _
        "Failed for table alias '" & tableAlias & "' (source '" & sourceAlias & "') at step '" & stepName & "'. " & _
        "SQL=[" & sql & "]. InnerError #" & CStr(innerErrNumber) & ": " & innerErrDescription

End Function

Private Function mp_GetAdoTableReference( _
    ByVal cfg As Object, _
    ByVal adoConn As Object, _
    ByVal adoObjectName As String, _
    ByVal expectedHeaders As Variant, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As String
    Dim resolvedRef As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String

    adoObjectName = Trim$(adoObjectName)
    If Len(adoObjectName) > 0 Then
        rangeStartMarker = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeStartMarker", vbNullString)
        rangeEndMarker = mp_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeEndMarker", vbNullString)

        If (Len(rangeStartMarker) > 0 Xor Len(rangeEndMarker) > 0) Then
            Err.Raise vbObjectError + 1747, "ex_PersonTimeline", _
                "Both markers must be provided for auto-range mode: '" & _
                sourceAlias & ".Sheet[" & tableAlias & "].RangeStartMarker' and '" & _
                sourceAlias & ".Sheet[" & tableAlias & "].RangeEndMarker'."
        End If

        If Len(rangeStartMarker) > 0 Then
            resolvedRef = mp_BuildAdoRangeReferenceFromMarkers( _
                cfg, sourceAlias, tableAlias, adoObjectName, rangeStartMarker, rangeEndMarker)
            mp_GetAdoTableReference = mp_QuoteSqlIdentifier(resolvedRef)
            Exit Function
        End If

        resolvedRef = mp_ResolveExplicitAdoObjectReference(adoConn, adoObjectName, sourceAlias, tableAlias)
        mp_GetAdoTableReference = mp_TryAutoDetectHeaderRangeReference( _
            adoConn, resolvedRef, expectedHeaders, keyHeader, sourceAlias, tableAlias)
        Exit Function
    End If

    Err.Raise vbObjectError + 1335, "ex_PersonTimeline", _
        "Missing required config key '" & sourceAlias & ".Sheet[" & tableAlias & "].SheetName'."
End Function

Private Function mp_BuildAdoRangeReferenceFromMarkers( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal configuredSheetName As String, _
    ByVal startMarker As String, _
    ByVal endMarker As String _
) As String
    Dim sourcePath As String
    Dim snapshotPath As String
    Dim cacheKey As String
    Dim normalizedSheetName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim startCell As Range
    Dim endCell As Range
    Dim markerCol As Long
    Dim headerRow As Long
    Dim dataLastRow As Long
    Dim firstHeaderCell As Range
    Dim firstCol As Long
    Dim markerColLetter As String
    Dim firstColLetter As String
    Dim sheetToken As String

    On Error GoTo EH

    If mp_IsExplicitAdoRangeReference(configuredSheetName) Then
        Err.Raise vbObjectError + 1748, "ex_PersonTimeline", _
            "Auto-range markers are not allowed with explicit range SheetName for " & _
            sourceAlias & ".Sheet[" & tableAlias & "].SheetName."
    End If

    sourcePath = mp_GetResolvedSourcePath(cfg, sourceAlias)
    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 1360, "ex_PersonTimeline", "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)
    normalizedSheetName = mp_NormalizeAdoObjectNameExact(configuredSheetName)
    cacheKey = LCase$(snapshotPath & "|" & normalizedSheetName & "|" & Trim$(startMarker) & "|" & Trim$(endMarker))

    mp_EnsureAdoLookupCacheContainers
    If g_AdoMarkerRangeRefBySignature.Exists(cacheKey) Then
        mp_BuildAdoRangeReferenceFromMarkers = CStr(g_AdoMarkerRangeRefBySignature(cacheKey))
        Exit Function
    End If

    Set wb = Workbooks.Open(Filename:=snapshotPath, ReadOnly:=True, UpdateLinks:=0)

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo EH

    Set ws = mp_FindWorksheetByConfiguredAdoName(wb, configuredSheetName)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1749, "ex_PersonTimeline", _
            "Worksheet for configured SheetName '" & configuredSheetName & "' was not found in source '" & sourceAlias & "'."
    End If

    Set startCell = mp_FindFirstMarkerCell(ws, startMarker)
    If startCell Is Nothing Then
        Err.Raise vbObjectError + 1750, "ex_PersonTimeline", _
            "Start marker '" & startMarker & "' was not found on sheet '" & ws.Name & "'."
    End If

    markerCol = startCell.Column
    Set endCell = mp_FindMarkerCellInColumnAfterRow(ws, markerCol, endMarker, startCell.Row)
    If endCell Is Nothing Then
        Err.Raise vbObjectError + 1751, "ex_PersonTimeline", _
            "End marker '" & endMarker & "' was not found below start marker '" & startMarker & "' in column " & CStr(markerCol) & " on sheet '" & ws.Name & "'."
    End If

    headerRow = startCell.Row - 1
    If headerRow < 1 Then
        Err.Raise vbObjectError + 1752, "ex_PersonTimeline", _
            "Header row cannot be determined: start marker '" & startMarker & "' is located at row " & CStr(startCell.Row) & "."
    End If

    dataLastRow = endCell.Row - 1
    If dataLastRow < startCell.Row Then
        Err.Raise vbObjectError + 1753, "ex_PersonTimeline", _
            "Detected marker range is empty: end marker '" & endMarker & "' is at row " & CStr(endCell.Row) & "."
    End If

    Set firstHeaderCell = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, markerCol)).Find( _
        What:="*", _
        After:=ws.Cells(headerRow, markerCol), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    If firstHeaderCell Is Nothing Then
        Err.Raise vbObjectError + 1754, "ex_PersonTimeline", _
            "Header row " & CStr(headerRow) & " contains no header cells before marker column."
    End If

    firstCol = firstHeaderCell.Column
    If firstCol > markerCol Then
        Err.Raise vbObjectError + 1755, "ex_PersonTimeline", _
            "Invalid detected range bounds: first column " & CStr(firstCol) & " is greater than marker column " & CStr(markerCol) & "."
    End If

    firstColLetter = mp_ToColumnLetter(firstCol)
    markerColLetter = mp_ToColumnLetter(markerCol)
    sheetToken = mp_BuildAdoSheetToken(configuredSheetName)
    If Len(sheetToken) = 0 Then
        Err.Raise vbObjectError + 1756, "ex_PersonTimeline", _
            "Failed to build ADO sheet token from SheetName '" & configuredSheetName & "'."
    End If

    mp_BuildAdoRangeReferenceFromMarkers = sheetToken & firstColLetter & CStr(headerRow) & ":" & markerColLetter & CStr(dataLastRow)
    g_AdoMarkerRangeRefBySignature(cacheKey) = mp_BuildAdoRangeReferenceFromMarkers

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

EH:
    Dim errNo As Long
    Dim errSrc As String
    Dim innerErr As String
    errNo = Err.Number
    errSrc = Err.Source
    innerErr = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    If errNo <> 0 Then Err.Raise errNo, errSrc, innerErr
End Function

Private Function mp_FindWorksheetByConfiguredAdoName(ByVal wb As Workbook, ByVal configuredSheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim needle As String
    Dim needleAlt As String

    needle = mp_ExtractSheetNameToken(configuredSheetName)
    If Len(needle) = 0 Then Exit Function
    needleAlt = Replace$(needle, "#", ".")

    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), needle, vbTextCompare) = 0 Then
            Set mp_FindWorksheetByConfiguredAdoName = ws
            Exit Function
        End If
        If StrComp(Replace$(Trim$(ws.Name), ".", "#"), needle, vbTextCompare) = 0 Then
            Set mp_FindWorksheetByConfiguredAdoName = ws
            Exit Function
        End If
        If StrComp(Trim$(ws.Name), needleAlt, vbTextCompare) = 0 Then
            Set mp_FindWorksheetByConfiguredAdoName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function mp_FindFirstMarkerCell(ByVal ws As Worksheet, ByVal markerText As String) As Range
    Dim searchRange As Range

    If ws Is Nothing Then Exit Function
    markerText = Trim$(markerText)
    If Len(markerText) = 0 Then Exit Function

    Set searchRange = ws.UsedRange
    If searchRange Is Nothing Then Exit Function

    Set mp_FindFirstMarkerCell = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
End Function

Private Function mp_FindMarkerCellInColumnAfterRow( _
    ByVal ws As Worksheet, _
    ByVal markerColumn As Long, _
    ByVal markerText As String, _
    ByVal minExclusiveRow As Long _
) As Range
    Dim searchRange As Range
    Dim firstFound As Range
    Dim currentFound As Range
    Dim firstAddress As String
    Dim bestRow As Long

    If ws Is Nothing Then Exit Function
    If markerColumn <= 0 Then Exit Function
    markerText = Trim$(markerText)
    If Len(markerText) = 0 Then Exit Function

    On Error Resume Next
    Set searchRange = Intersect(ws.Columns(markerColumn), ws.UsedRange)
    On Error GoTo 0
    If searchRange Is Nothing Then
        Set searchRange = ws.Columns(markerColumn)
    End If

    Set firstFound = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    If firstFound Is Nothing Then Exit Function

    bestRow = 0
    firstAddress = firstFound.Address
    Set currentFound = firstFound

    Do
        If currentFound.Row > minExclusiveRow Then
            If bestRow = 0 Or currentFound.Row < bestRow Then
                bestRow = currentFound.Row
                Set mp_FindMarkerCellInColumnAfterRow = currentFound
            End If
        End If
        Set currentFound = searchRange.FindNext(currentFound)
        If currentFound Is Nothing Then Exit Do
    Loop While currentFound.Address <> firstAddress
End Function

Private Function mp_ExtractSheetNameToken(ByVal configuredSheetName As String) As String
    Dim token As String
    Dim dollarPos As Long

    token = Trim$(configuredSheetName)
    If Len(token) = 0 Then Exit Function

    If Left$(token, 1) = "[" And Right$(token, 1) = "]" Then
        token = Mid$(token, 2, Len(token) - 2)
    End If
    token = mp_CleanAdoSchemaObjectName(token)

    dollarPos = InStr(1, token, "$", vbBinaryCompare)
    If dollarPos > 0 Then
        token = Left$(token, dollarPos - 1)
    End If

    mp_ExtractSheetNameToken = Trim$(token)
End Function

Private Function mp_BuildAdoSheetToken(ByVal configuredSheetName As String) As String
    Dim token As String
    Dim dollarPos As Long

    token = Trim$(configuredSheetName)
    If Len(token) = 0 Then Exit Function

    If Left$(token, 1) = "[" And Right$(token, 1) = "]" Then
        token = Mid$(token, 2, Len(token) - 2)
    End If
    token = mp_CleanAdoSchemaObjectName(token)
    token = Trim$(token)
    If Len(token) = 0 Then Exit Function

    dollarPos = InStr(1, token, "$", vbBinaryCompare)
    If dollarPos > 0 Then
        mp_BuildAdoSheetToken = Left$(token, dollarPos)
    Else
        mp_BuildAdoSheetToken = token & "$"
    End If
End Function

Private Function mp_ResolveExplicitAdoObjectReference( _
    ByVal adoConn As Object, _
    ByVal configuredName As String, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As String
    Dim schemaRs As Object
    Dim schemaName As String
    Dim schemaNameClean As String
    Dim listedNames As String
    Dim listedCount As Long
    Dim cacheKey As String

    configuredName = mp_CleanAdoSchemaObjectName(Trim$(configuredName))
    cacheKey = LCase$(Trim$(sourceAlias) & "|" & Trim$(tableAlias) & "|" & mp_NormalizeAdoObjectNameExact(configuredName))

    If Not g_AdoResolvedTableRefByConfigured Is Nothing Then
        If g_AdoResolvedTableRefByConfigured.Exists(cacheKey) Then
            mp_ResolveExplicitAdoObjectReference = CStr(g_AdoResolvedTableRefByConfigured(cacheKey))
            Exit Function
        End If
    End If

    If mp_IsExplicitAdoRangeReference(configuredName) Then
        mp_ResolveExplicitAdoObjectReference = mp_QuoteSqlIdentifier(configuredName)
        If Not g_AdoResolvedTableRefByConfigured Is Nothing Then
            g_AdoResolvedTableRefByConfigured(cacheKey) = mp_ResolveExplicitAdoObjectReference
        End If
        Exit Function
    End If

    Set schemaRs = adoConn.OpenSchema(20)
    Do While Not schemaRs.EOF
        schemaName = CStr(schemaRs.Fields("TABLE_NAME").Value)
        schemaNameClean = mp_CleanAdoSchemaObjectName(schemaName)

        If listedCount < 20 Then
            If Len(listedNames) > 0 Then listedNames = listedNames & ", "
            listedNames = listedNames & schemaNameClean
            listedCount = listedCount + 1
        End If

        If mp_IsMatchingConfiguredAdoObject(schemaNameClean, configuredName) Then
            schemaRs.Close
            mp_ResolveExplicitAdoObjectReference = mp_QuoteSqlIdentifier(schemaNameClean)
            If Not g_AdoResolvedTableRefByConfigured Is Nothing Then
                g_AdoResolvedTableRefByConfigured(cacheKey) = mp_ResolveExplicitAdoObjectReference
            End If
            Exit Function
        End If

        schemaRs.MoveNext
    Loop
    schemaRs.Close

    Err.Raise vbObjectError + 1336, "ex_PersonTimeline", _
        "Configured SheetName '" & configuredName & "' for " & sourceAlias & ".Sheet[" & tableAlias & "] was not found. Available objects: " & listedNames
End Function

Private Function mp_TryAutoDetectHeaderRangeReference( _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal expectedHeaders As Variant, _
    ByVal keyHeader As String, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As String
    Dim rs As Object
    Dim keyOrdinal As Long
    Dim detectedRef As String
    Dim sourceRef As String
    Dim autoKey As String

    mp_TryAutoDetectHeaderRangeReference = tableRef
    If Len(Trim$(tableRef)) = 0 Then Exit Function
    If Len(Trim$(keyHeader)) = 0 Then Exit Function

    sourceRef = mp_UnquoteSqlIdentifier(tableRef)
    If InStr(1, sourceRef, "$", vbBinaryCompare) = 0 Then Exit Function
    ' Explicit range reference means user already constrained source bounds.
    ' Keep strict object as-is and skip any auto-detection expansion.
    If mp_IsExplicitAdoRangeReference(sourceRef) Then Exit Function

    autoKey = LCase$(Trim$(sourceAlias) & "|" & Trim$(tableAlias) & "|" & mp_NormalizeAdoObjectNameExact(sourceRef) & "|" & _
                     mp_NormalizeHeader(keyHeader) & "|" & mp_BuildExpectedHeadersSignature(expectedHeaders))
    If Not g_AdoAutoTableRefBySignature Is Nothing Then
        If g_AdoAutoTableRefBySignature.Exists(autoKey) Then
            mp_TryAutoDetectHeaderRangeReference = CStr(g_AdoAutoTableRefBySignature(autoKey))
            Exit Function
        End If
    End If

    On Error GoTo EH
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & tableRef & " WHERE 1=0", adoConn, 0, 1
    keyOrdinal = mp_RecordsetGetFieldOrdinal(rs, keyHeader)
    If keyOrdinal >= 0 Then
        rs.Close
        If Not g_AdoAutoTableRefBySignature Is Nothing Then
            g_AdoAutoTableRefBySignature(autoKey) = tableRef
        End If
        Exit Function
    End If
    rs.Close

    If mp_TryDetectHeaderRangeFromTopRows(adoConn, tableRef, expectedHeaders, keyHeader, detectedRef) Then
        mp_TryAutoDetectHeaderRangeReference = detectedRef
    End If
    If Not g_AdoAutoTableRefBySignature Is Nothing Then
        g_AdoAutoTableRefBySignature(autoKey) = mp_TryAutoDetectHeaderRangeReference
    End If
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_TryDetectHeaderRangeFromTopRows( _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal expectedHeaders As Variant, _
    ByVal keyHeader As String, _
    ByRef outDetectedRef As String _
) As Boolean
    Const MAX_HEADER_ALIGNMENT_SHIFT As Long = 20
    Dim sheetPrefix As String
    Dim probeRef As String
    Dim rs As Object
    Dim rowsData As Variant
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim fieldLower As Long
    Dim fieldUpper As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim bestRowIndex As Long
    Dim bestScore As Long
    Dim bestLastCol As Long
    Dim rowTokens As Object
    Dim expectedSet As Object
    Dim keyToken As String
    Dim cellText As String
    Dim normalized As String
    Dim lastNonEmptyCol As Long
    Dim currentScore As Long
    Dim token As Variant
    Dim headerRowAbs As Long
    Dim colLetter As String
    Dim fallbackRowAbs As Long
    Dim alignmentShift As Long

    sheetPrefix = mp_ExtractAdoSheetPrefix(tableRef)
    If Len(sheetPrefix) = 0 Then Exit Function

    probeRef = "[" & sheetPrefix & "A1:ZZ200]"

    Set expectedSet = mp_BuildNormalizedHeaderTokenSet(expectedHeaders, keyHeader)
    If expectedSet Is Nothing Then Exit Function
    If expectedSet.Count = 0 Then Exit Function

    keyToken = mp_NormalizeHeader(keyHeader)
    If Len(keyToken) = 0 Then Exit Function

    On Error GoTo EH
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & probeRef, adoConn, 0, 1
    If rs.EOF Then
        rs.Close
        Exit Function
    End If

    rowsData = rs.GetRows
    rs.Close

    rowLower = LBound(rowsData, 2)
    rowUpper = UBound(rowsData, 2)
    fieldLower = LBound(rowsData, 1)
    fieldUpper = UBound(rowsData, 1)
    bestRowIndex = -1
    bestScore = 0

    For rowIndex = rowLower To rowUpper
        Set rowTokens = CreateObject("Scripting.Dictionary")
        rowTokens.CompareMode = 1
        lastNonEmptyCol = 0

        For colIndex = fieldLower To fieldUpper
            cellText = mp_ToSafeText(rowsData(colIndex, rowIndex))
            normalized = mp_NormalizeHeader(cellText)
            If Len(normalized) > 0 Then
                rowTokens(normalized) = True
                lastNonEmptyCol = (colIndex - fieldLower + 1)
            End If
        Next colIndex

        If rowTokens.Exists(keyToken) Then
            currentScore = 0
            For Each token In expectedSet.Keys
                If rowTokens.Exists(CStr(token)) Then
                    currentScore = currentScore + 1
                End If
            Next token

            If currentScore > bestScore Then
                bestScore = currentScore
                bestRowIndex = rowIndex
                bestLastCol = lastNonEmptyCol
            End If
        End If
    Next rowIndex

    If bestRowIndex < 0 Then Exit Function
    If bestLastCol <= 0 Then bestLastCol = (fieldUpper - fieldLower + 1)
    If bestLastCol <= 0 Then Exit Function

    colLetter = mp_ToColumnLetter(bestLastCol)
    If Len(colLetter) = 0 Then Exit Function

    ' Provider alignment can drift from A1 (A1, A2, A3...).
    ' Probe a bounded set of shifts and keep the first candidate
    ' that exposes keyHeader in WHERE 1=0 metadata.
    For alignmentShift = 1 To MAX_HEADER_ALIGNMENT_SHIFT
        headerRowAbs = (bestRowIndex - rowLower) + alignmentShift
        If headerRowAbs > 0 Then
            If mp_TryBuildValidatedHeaderRangeRef(adoConn, sheetPrefix, headerRowAbs, colLetter, keyHeader, outDetectedRef) Then
                mp_TryDetectHeaderRangeFromTopRows = True
                Exit Function
            End If
        End If
    Next alignmentShift

    ' Last resort: keep deterministic fallback based on +1.
    fallbackRowAbs = (bestRowIndex - rowLower) + 1
    If fallbackRowAbs <= 0 Then Exit Function
    outDetectedRef = "[" & sheetPrefix & "A" & CStr(fallbackRowAbs) & ":" & colLetter & "1048576]"
    mp_TryDetectHeaderRangeFromTopRows = True
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_TryBuildValidatedHeaderRangeRef( _
    ByVal adoConn As Object, _
    ByVal sheetPrefix As String, _
    ByVal headerRowAbs As Long, _
    ByVal colLetter As String, _
    ByVal keyHeader As String, _
    ByRef outRangeRef As String _
) As Boolean
    Dim rs As Object
    Dim candidateRef As String

    If adoConn Is Nothing Then Exit Function
    If headerRowAbs <= 0 Then Exit Function
    If Len(Trim$(sheetPrefix)) = 0 Then Exit Function
    If Len(Trim$(colLetter)) = 0 Then Exit Function
    If Len(Trim$(keyHeader)) = 0 Then Exit Function

    candidateRef = "[" & sheetPrefix & "A" & CStr(headerRowAbs) & ":" & colLetter & "1048576]"

    On Error GoTo EH
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & candidateRef & " WHERE 1=0", adoConn, 0, 1
    If mp_RecordsetGetFieldOrdinal(rs, keyHeader) >= 0 Then
        outRangeRef = candidateRef
        mp_TryBuildValidatedHeaderRangeRef = True
    End If
    rs.Close
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_GetSourceAliasCached(ByVal cfg As Object, ByVal tableAlias As String, ByVal cache As Object) As String
    If Not cache Is Nothing Then
        If cache.Exists(tableAlias) Then
            mp_GetSourceAliasCached = CStr(cache(tableAlias))
            Exit Function
        End If
    End If

    mp_GetSourceAliasCached = mp_FindSourceAliasForTable(cfg, tableAlias)

    If Not cache Is Nothing Then
        cache(tableAlias) = mp_GetSourceAliasCached
    End If
End Function

Private Sub mp_ApplyTimelineStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByVal resultFieldRanges As Collection, _
    ByVal resultTables As Collection, _
    ByVal partialMatchRowRanges As Collection, _
    ByVal hasOutputStyle As Boolean, _
    ByRef outputStyle As t_OutputSheetStyle, _
    ByVal pendingWarningBanners As Collection _
)
    Dim rowKindRanges As Object
    Dim partialRowKindRanges As Object
    Dim fetchDslRowKindRanges As Object
    Dim runtimeLayers As Collection
    Dim runtimeLayer As obj_StyleLayer

    Set rowKindRanges = mp_BuildTimelineRowKindRanges(headerRows, sectionRows, resultFieldRanges)
    Set partialRowKindRanges = mp_BuildPartialMatchRowKindRanges(partialMatchRowRanges)
    Set fetchDslRowKindRanges = mp_BuildFetchDslRowKindRanges(resultTables)
    mp_MergeRowKindRanges rowKindRanges, partialRowKindRanges
    mp_MergeRowKindRanges rowKindRanges, fetchDslRowKindRanges

    Set runtimeLayers = New Collection

    If hasOutputStyle Then
        Set runtimeLayer = ex_OutputPanel.m_CreateRuntimeLayer(ws, outputStyle, "runtime-control-panel", 800)
        If Not runtimeLayer Is Nothing Then runtimeLayers.Add runtimeLayer
    End If

    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, resultFieldRanges, Nothing, rowKindRanges, vbNullString, False, runtimeLayers
End Sub

Private Function mp_BuildFetchDslRowKindRanges(ByVal resultTables As Collection) As Object
    Dim result As Object
    Dim tableObj As obj_ResultTable
    Dim rowObj As obj_ResultRow
    Dim kindRows As Collection
    Dim rowKindValue As String
    Dim rowKindTokens As Variant
    Dim tokenIndex As Long
    Dim tokenText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not resultTables Is Nothing Then
        For Each tableObj In resultTables
            If tableObj Is Nothing Then GoTo ContinueTable
            If tableObj.Rows Is Nothing Then GoTo ContinueTable

            For Each rowObj In tableObj.Rows
                If rowObj Is Nothing Then GoTo ContinueRow
                rowKindValue = Trim$(rowObj.Kind)
                If Len(rowKindValue) = 0 Then GoTo ContinueRow

                rowKindTokens = Split(rowKindValue, "|")
                For tokenIndex = LBound(rowKindTokens) To UBound(rowKindTokens)
                    tokenText = LCase$(Trim$(CStr(rowKindTokens(tokenIndex))))
                    If Len(tokenText) = 0 Then GoTo ContinueToken
                    If Not result.Exists(tokenText) Then
                        Set kindRows = New Collection
                        Set result(tokenText) = kindRows
                    End If
                    result(tokenText).Add CLng(rowObj.RowIndex)
ContinueToken:
                Next tokenIndex
ContinueRow:
            Next rowObj
ContinueTable:
        Next tableObj
    End If

    Set mp_BuildFetchDslRowKindRanges = result
End Function

Private Function mp_BuildTimelineRowKindRanges( _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByVal resultFieldRanges As Collection _
) As Object
    Dim result As Object
    Dim headerRowsMap As Object
    Dim sectionRowsMap As Object
    Dim contentRowsMap As Object
    Dim target As Object
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowIndex As Long
    Dim rowKey As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Set headerRowsMap = mp_BuildRowsMap(headerRows)
    Set sectionRowsMap = mp_BuildRowsMap(sectionRows)
    Set contentRowsMap = CreateObject("Scripting.Dictionary")
    contentRowsMap.CompareMode = 0

    If Not resultFieldRanges Is Nothing Then
        For Each target In resultFieldRanges
            If target Is Nothing Then GoTo ContinueTarget
            If Not target.Exists("RowStart") Then GoTo ContinueTarget
            If Not target.Exists("RowEnd") Then GoTo ContinueTarget

            rowStart = CLng(target("RowStart"))
            rowEnd = CLng(target("RowEnd"))
            If rowStart <= 0 Then GoTo ContinueTarget
            If rowEnd < rowStart Then rowEnd = rowStart

            For rowIndex = rowStart To rowEnd
                rowKey = CStr(rowIndex)
                If headerRowsMap.Exists(rowKey) Then GoTo ContinueRow
                If sectionRowsMap.Exists(rowKey) Then GoTo ContinueRow
                contentRowsMap(rowKey) = True
ContinueRow:
            Next rowIndex
ContinueTarget:
        Next target
    End If

    Set result("header") = mp_RowsMapToRangeCollection(headerRowsMap)
    Set result("section") = mp_RowsMapToRangeCollection(sectionRowsMap)
    Set result("content") = mp_RowsMapToRangeCollection(contentRowsMap)
    Set mp_BuildTimelineRowKindRanges = result
End Function

Private Sub mp_MergeRowKindRanges(ByVal targetRanges As Object, ByVal sourceRanges As Object)
    Dim kindName As Variant
    Dim targetCollection As Collection
    Dim sourceCollection As Collection
    Dim rowItem As Variant

    If targetRanges Is Nothing Then Exit Sub
    If sourceRanges Is Nothing Then Exit Sub

    For Each kindName In sourceRanges.Keys
        If targetRanges.Exists(CStr(kindName)) Then
            Set targetCollection = targetRanges(CStr(kindName))
        Else
            Set targetCollection = New Collection
            Set targetRanges(CStr(kindName)) = targetCollection
        End If

        Set sourceCollection = sourceRanges(CStr(kindName))
        If sourceCollection Is Nothing Then GoTo ContinueKind

        For Each rowItem In sourceCollection
            targetCollection.Add rowItem
        Next rowItem
ContinueKind:
    Next kindName
End Sub

Private Function mp_BuildPartialMatchRowKindRanges(ByVal partialMatchRowRanges As Collection) As Object
    Dim result As Object
    Dim partialRanges As Collection
    Dim rowItem As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set partialRanges = New Collection

    If Not partialMatchRowRanges Is Nothing Then
        For Each rowItem In partialMatchRowRanges
            partialRanges.Add rowItem
        Next rowItem
    End If

    Set result("partialmatch") = partialRanges
    Set mp_BuildPartialMatchRowKindRanges = result
End Function

Private Function mp_BuildRowsMap(ByVal rowsCollection As Collection) As Object
    Dim rowsMap As Object
    Dim itemValue As Variant

    Set rowsMap = CreateObject("Scripting.Dictionary")
    rowsMap.CompareMode = 0

    If rowsCollection Is Nothing Then
        Set mp_BuildRowsMap = rowsMap
        Exit Function
    End If

    For Each itemValue In rowsCollection
        rowsMap(CStr(CLng(itemValue))) = True
    Next itemValue

    Set mp_BuildRowsMap = rowsMap
End Function

Private Function mp_RowsMapToRangeCollection(ByVal rowsMap As Object) As Collection
    Dim result As Collection
    Dim keys() As Long
    Dim keyValue As Variant
    Dim i As Long
    Dim count As Long
    Dim rangeItem As Object
    Dim runStart As Long
    Dim runEnd As Long

    Set result = New Collection
    If rowsMap Is Nothing Then
        Set mp_RowsMapToRangeCollection = result
        Exit Function
    End If
    If rowsMap.Count = 0 Then
        Set mp_RowsMapToRangeCollection = result
        Exit Function
    End If

    ReDim keys(1 To rowsMap.Count)
    For Each keyValue In rowsMap.Keys
        count = count + 1
        keys(count) = CLng(keyValue)
    Next keyValue
    If count = 0 Then
        Set mp_RowsMapToRangeCollection = result
        Exit Function
    End If

    mp_SortLongArray keys

    runStart = keys(1)
    runEnd = runStart
    For i = 2 To UBound(keys)
        If keys(i) = runEnd + 1 Then
            runEnd = keys(i)
        Else
            Set rangeItem = CreateObject("Scripting.Dictionary")
            rangeItem.CompareMode = 1
            rangeItem("RowStart") = runStart
            rangeItem("RowEnd") = runEnd
            result.Add rangeItem

            runStart = keys(i)
            runEnd = runStart
        End If
    Next i

    Set rangeItem = CreateObject("Scripting.Dictionary")
    rangeItem.CompareMode = 1
    rangeItem("RowStart") = runStart
    rangeItem("RowEnd") = runEnd
    result.Add rangeItem

    Set mp_RowsMapToRangeCollection = result
End Function

Private Sub mp_SortLongArray(ByRef values() As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    If UBound(values) <= LBound(values) Then Exit Sub

    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(j) < values(i) Then
                tmp = values(i)
                values(i) = values(j)
                values(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function mp_LoadConfigDictionary() As Object

    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ws_Dev

    On Error Resume Next
    Set tbl = ws.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        Err.Raise vbObjectError + 1330, "ex_PersonTimeline", _
            "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'."
    End If

    If tbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1331, "ex_PersonTimeline", _
            "Config table '" & DEV_CONFIG_TABLE_NAME & "' has no data rows."
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim dataRange As Range
    Set dataRange = tbl.DataBodyRange

    Dim r As Long
    For r = 1 To dataRange.Rows.Count
        Dim markerText As String
        markerText = Trim$(CStr(dataRange.Cells(r, DEV_COL_MARKER).Value))
        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then
            GoTo ContinueRow
        End If

        Dim keyText As String
        keyText = Trim$(CStr(dataRange.Cells(r, DEV_COL_KEY).Value))
        If Len(keyText) = 0 Then
            GoTo ContinueRow
        End If

        dict(keyText) = CStr(dataRange.Cells(r, DEV_COL_VALUE).Value)

ContinueRow:
    Next r

    Set mp_LoadConfigDictionary = dict

End Function

Private Sub mp_AddResultFieldRange( _
    ByVal resultFieldRanges As Collection, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
        ByVal outCol As Long, _
        ByVal rowStart As Long, _
        ByVal rowEnd As Long, _
    Optional ByVal fieldKind As String = vbNullString _
)
    Dim target As Object

    If resultFieldRanges Is Nothing Then Exit Sub
    If Len(Trim$(fieldAlias)) = 0 Then Exit Sub
    If outCol <= 0 Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    Set target = CreateObject("Scripting.Dictionary")
    target.CompareMode = 1
    target("MapKey") = sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    target("ColumnIndex") = outCol
    target("RowStart") = rowStart
    target("RowEnd") = rowEnd
    fieldKind = mp_CombineKindTags(fieldKind)
    If Len(fieldKind) > 0 Then target("Kind") = fieldKind

    resultFieldRanges.Add target
End Sub

Private Sub mp_AddPartialMatchRowRange( _
    ByVal partialMatchRowRanges As Collection, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long _
)
    Dim target As Object

    If partialMatchRowRanges Is Nothing Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    Set target = CreateObject("Scripting.Dictionary")
    target.CompareMode = 1
    target("RowStart") = rowStart
    target("RowEnd") = rowEnd
    partialMatchRowRanges.Add target
End Sub

Private Function mp_EnsureResultTable( _
    ByVal resultTables As Collection, _
    ByVal resultTablesByRef As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As obj_ResultTable
    Dim tableRef As String
    tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"

    If resultTablesByRef Is Nothing Then
        Err.Raise vbObjectError + 1338, "ex_PersonTimeline", "Result tables index dictionary is not initialized."
    End If

    If Not resultTablesByRef.Exists(tableRef) Then
        Dim tableObj As obj_ResultTable
        Set tableObj = New obj_ResultTable
        tableObj.Initialize tableRef
        resultTablesByRef.Add tableRef, tableObj
        If Not resultTables Is Nothing Then resultTables.Add tableObj
    End If

    Set mp_EnsureResultTable = resultTablesByRef(tableRef)
End Function

Private Sub mp_RegisterResultTableFieldAliases( _
    ByVal resultTable As obj_ResultTable, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Variant _
)
    Dim i As Long
    Dim fieldAlias As String
    Dim mapKey As String

    If resultTable Is Nothing Then Exit Sub
    If mp_IsEmptyVariantArray(fields) Then Exit Sub

    For i = LBound(fields) To UBound(fields)
        fieldAlias = Trim$(CStr(fields(i)))
        If Len(fieldAlias) = 0 Then GoTo ContinueField
        mapKey = sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
        resultTable.AddFieldMap fieldAlias, mapKey
ContinueField:
    Next i
End Sub

Private Sub mp_CaptureResultTableRowsFromOutput( _
    ByVal wsOut As Worksheet, _
    ByVal resultTable As obj_ResultTable, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Variant, _
    ByVal dataRowStart As Long, _
    ByVal dataRowEnd As Long, _
    Optional ByVal rowKindsBySheetRow As Object = Nothing, _
    Optional ByVal visualRowStart As Long = 0, _
    Optional ByVal visualRowEnd As Long = 0 _
)
    Dim r As Long
    Dim i As Long
    Dim outCol As Long
    Dim fieldCount As Long
    Dim rowOffset As Long
    Dim fieldAlias As String
    Dim mapKey As String
    Dim valueText As String
    Dim rowObj As obj_ResultRow
    Dim dataRange As Range
    Dim capturedValues As Variant
    Dim isScalarRange As Boolean
    Dim rowOrdinal As Long
    Dim rowAnchorName As String

    If wsOut Is Nothing Then Exit Sub
    If resultTable Is Nothing Then Exit Sub
    If mp_IsEmptyVariantArray(fields) Then Exit Sub
    If dataRowStart <= 0 Then Exit Sub
    If dataRowEnd < dataRowStart Then Exit Sub

    fieldCount = UBound(fields) - LBound(fields) + 1
    If fieldCount <= 0 Then Exit Sub

    Set dataRange = wsOut.Range(wsOut.Cells(dataRowStart, 1), wsOut.Cells(dataRowEnd, fieldCount))
    capturedValues = dataRange.Value2
    isScalarRange = Not IsArray(capturedValues)

    For r = dataRowStart To dataRowEnd
        Set rowObj = resultTable.EnsureRow(r)
        rowOrdinal = r - dataRowStart + 1
        rowAnchorName = ex_Messaging.m_BuildResultRowAnchorName(resultTable.TableRef, rowOrdinal)
        If Len(rowAnchorName) = 0 Then
            Err.Raise vbObjectError + 1316, "ex_PersonTimeline", "Unable to build row anchor name for table '" & resultTable.TableRef & "' row ordinal " & CStr(rowOrdinal) & "."
        End If
        rowObj.RowAnchorName = rowAnchorName
        ex_Messaging.m_RegisterResultRowAnchor wsOut, rowAnchorName, r
        If Not rowKindsBySheetRow Is Nothing Then
            If rowKindsBySheetRow.Exists(CStr(r)) Then
                rowObj.Kind = CStr(rowKindsBySheetRow(CStr(r)))
            Else
                rowObj.Kind = vbNullString
            End If
        Else
            rowObj.Kind = vbNullString
        End If
        rowOffset = 1 + (r - dataRowStart)
        For i = LBound(fields) To UBound(fields)
            fieldAlias = Trim$(CStr(fields(i)))
            If Len(fieldAlias) = 0 Then GoTo ContinueField
            outCol = 1 + (i - LBound(fields))
            mapKey = sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
            If isScalarRange Then
                valueText = CStr(capturedValues)
            Else
                valueText = CStr(capturedValues(rowOffset, outCol))
            End If
            rowObj.SetValue fieldAlias, mapKey, valueText
ContinueField:
        Next i
    Next r

    If visualRowStart <= 0 Then visualRowStart = dataRowStart
    If visualRowEnd < visualRowStart Then visualRowEnd = dataRowEnd
    ex_Messaging.m_RegisterResultTableAnchor wsOut, resultTable.TableRef, visualRowStart, visualRowEnd
End Sub

Private Sub mp_AddResultFieldRangesForFields( _
    ByVal resultFieldRanges As Collection, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Variant, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long _
)
    Dim i As Long
    Dim fieldAlias As String
    Dim virtualKind As String
    Dim headerKind As String
    Dim contentKind As String

    If resultFieldRanges Is Nothing Then Exit Sub
    If mp_IsEmptyVariantArray(fields) Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    For i = LBound(fields) To UBound(fields)
        fieldAlias = Trim$(CStr(fields(i)))
        If Len(fieldAlias) = 0 Then GoTo ContinueField
        If mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then
            virtualKind = ex_FetchDslEngine.m_GetGeneratedKindValue()
        Else
            virtualKind = vbNullString
        End If

        If Len(virtualKind) > 0 Then
            headerKind = mp_CombineKindTags("header", virtualKind)
            mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, fieldAlias, 1 + (i - LBound(fields)), rowStart, rowStart, headerKind

            If rowEnd >= (rowStart + 1) Then
                contentKind = mp_CombineKindTags("content", virtualKind)
                mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, fieldAlias, 1 + (i - LBound(fields)), rowStart + 1, rowEnd, contentKind
            End If
        Else
            mp_AddResultFieldRange resultFieldRanges, sourceAlias, tableAlias, fieldAlias, 1 + (i - LBound(fields)), rowStart, rowEnd
        End If
ContinueField:
    Next i
End Sub

Private Function mp_CombineKindTags( _
    ByVal primaryTag As String, _
    Optional ByVal secondaryTag As String = vbNullString _
) As String
    Dim tags As Object
    Dim raw As Variant
    Dim parts As Variant
    Dim i As Long
    Dim tokenText As String

    Set tags = CreateObject("Scripting.Dictionary")
    tags.CompareMode = 1

    For Each raw In Array(primaryTag, secondaryTag)
        If Len(Trim$(CStr(raw))) = 0 Then GoTo ContinueRaw
        parts = Split(CStr(raw), "|")
        For i = LBound(parts) To UBound(parts)
            tokenText = LCase$(Trim$(CStr(parts(i))))
            If Len(tokenText) > 0 Then tags(tokenText) = True
        Next i
ContinueRaw:
    Next raw

    If tags.Count > 0 Then
        mp_CombineKindTags = Join(tags.Keys, "|")
    Else
        mp_CombineKindTags = vbNullString
    End If
End Function

Private Function mp_AppendFetchRowsFromSource( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal fields As Variant, _
    ByRef ioOutValues() As Variant, _
    ByRef ioRowCount As Long, _
    ByVal fieldCount As Long, _
    Optional ByRef outKindsByOutRow As Object = Nothing _
) As Boolean
    If ioRowCount <= 0 Then Exit Function
    If fieldCount <= 0 Then Exit Function
    If mp_IsEmptyVariantArray(fields) Then Exit Function

    mp_AppendFetchRowsFromSource = ex_FetchDslEngine.m_ApplyFetchRowsFromSource( _
        cfg, sourceAlias, tableAlias, adoConn, tableRef, keyHeader, keyValue, fields, ioOutValues, ioRowCount, fieldCount, outKindsByOutRow)
End Function

Private Function mp_BuildSheetRowKindsMap(ByVal outKindsByOutRow As Object, ByVal dataRowStart As Long) As Object
    Dim result As Object
    Dim outRowKey As Variant
    Dim outRowIndex As Long
    Dim sheetRowIndex As Long

    If outKindsByOutRow Is Nothing Then Exit Function
    If dataRowStart <= 0 Then Exit Function

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    For Each outRowKey In outKindsByOutRow.Keys
        outRowIndex = CLng(outRowKey)
        If outRowIndex <= 0 Then GoTo ContinueKey
        sheetRowIndex = dataRowStart + outRowIndex - 1
        result(CStr(sheetRowIndex)) = CStr(outKindsByOutRow(outRowKey))
ContinueKey:
    Next outRowKey

    Set mp_BuildSheetRowKindsMap = result
End Function

Private Function mp_AdvanceRowIndexAfterRenderedTable( _
    ByVal cfg As Object, _
    ByVal outputAliases As Variant, _
    ByVal currentIndex As Long, _
    ByVal mode As OutputMode, _
    ByVal tableSourceMap As Object, _
    ByVal currentTableAlias As String, _
    ByVal currentTableType As String, _
    ByVal rowIndexAfterCurrentTable As Long _
) As Long
    Dim nextSourceAlias As String
    Dim nextTableAlias As String
    Dim nextTableType As String
    Dim gapRows As Long

    mp_AdvanceRowIndexAfterRenderedTable = rowIndexAfterCurrentTable

    If Not mp_TryGetNextRenderableOutputTable(cfg, outputAliases, currentIndex, mode, tableSourceMap, nextSourceAlias, nextTableAlias, nextTableType) Then
        Exit Function
    End If

    gapRows = mp_GetOutputTablesGapRows(cfg, currentTableAlias, currentTableType, nextTableAlias, nextTableType)
    If gapRows > 0 Then
        mp_AdvanceRowIndexAfterRenderedTable = rowIndexAfterCurrentTable + gapRows
    End If
End Function

Private Function mp_TryGetNextRenderableOutputTable( _
    ByVal cfg As Object, _
    ByVal outputAliases As Variant, _
    ByVal fromIndex As Long, _
    ByVal mode As OutputMode, _
    ByVal tableSourceMap As Object, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outTableType As String _
) As Boolean
    Dim i As Long
    Dim candidateAlias As String
    Dim candidateSourceAlias As String
    Dim candidateTableType As String

    For i = fromIndex + 1 To UBound(outputAliases)
        candidateAlias = Trim$(CStr(outputAliases(i)))
        If Len(candidateAlias) = 0 Then GoTo ContinueCandidate

        candidateSourceAlias = mp_GetSourceAliasCached(cfg, candidateAlias, tableSourceMap)
        candidateTableType = LCase$(mp_GetCfgRequired(cfg, candidateSourceAlias & ".Sheet[" & candidateAlias & "].Type"))

        If Not mp_IsSupportedOutputTableType(candidateTableType) Then
            Err.Raise vbObjectError + 1301, "ex_PersonTimeline", _
                "Unsupported table type for alias '" & candidateAlias & "': " & candidateTableType
        End If

        If Not mp_ShouldRenderTableForMode(mode, candidateTableType) Then
            GoTo ContinueCandidate
        End If

        outSourceAlias = candidateSourceAlias
        outTableAlias = candidateAlias
        outTableType = candidateTableType
        mp_TryGetNextRenderableOutputTable = True
        Exit Function
ContinueCandidate:
    Next i
End Function

Private Function mp_GetOutputTablesGapRows( _
    ByVal cfg As Object, _
    ByVal currentTableAlias As String, _
    ByVal currentTableType As String, _
    ByVal nextTableAlias As String, _
    ByVal nextTableType As String _
) As Long
    Dim keyName As String
    Dim currentTypeToken As String
    Dim nextTypeToken As String

    currentTypeToken = mp_ToOutputTableTypeToken(currentTableType)
    nextTypeToken = mp_ToOutputTableTypeToken(nextTableType)

    keyName = "Output.Layout.Gap.Between[" & currentTableAlias & "->" & nextTableAlias & "]"
    If cfg.Exists(keyName) Then
        mp_GetOutputTablesGapRows = mp_GetCfgNonNegativeLongValue(cfg, keyName)
        Exit Function
    End If

    keyName = "Output.Layout.Gap.BetweenType[" & currentTypeToken & "->" & nextTypeToken & "]"
    If cfg.Exists(keyName) Then
        mp_GetOutputTablesGapRows = mp_GetCfgNonNegativeLongValue(cfg, keyName)
        Exit Function
    End If

    keyName = "Output.Layout.Gap.After[" & currentTableAlias & "]"
    If cfg.Exists(keyName) Then
        mp_GetOutputTablesGapRows = mp_GetCfgNonNegativeLongValue(cfg, keyName)
        Exit Function
    End If

    keyName = "Output.Layout.Gap.AfterType[" & currentTypeToken & "]"
    If cfg.Exists(keyName) Then
        mp_GetOutputTablesGapRows = mp_GetCfgNonNegativeLongValue(cfg, keyName)
        Exit Function
    End If

    keyName = "Output.Layout.Gap.Default"
    If cfg.Exists(keyName) Then
        mp_GetOutputTablesGapRows = mp_GetCfgNonNegativeLongValue(cfg, keyName)
        Exit Function
    End If

    mp_GetOutputTablesGapRows = 1
End Function

Private Function mp_IsSupportedOutputTableType(ByVal tableType As String) As Boolean
    mp_IsSupportedOutputTableType = (tableType = "state" Or tableType = "events")
End Function

Private Function mp_ShouldRenderTableForMode(ByVal mode As OutputMode, ByVal tableType As String) As Boolean
    If mode = StateTableOnly And tableType <> "state" Then Exit Function
    If mode = EventsTableOnly And tableType <> "events" Then Exit Function
    mp_ShouldRenderTableForMode = True
End Function

Private Function mp_ToOutputTableTypeToken(ByVal tableType As String) As String
    Select Case LCase$(Trim$(tableType))
        Case "state"
            mp_ToOutputTableTypeToken = "State"
        Case "events"
            mp_ToOutputTableTypeToken = "Events"
        Case Else
            mp_ToOutputTableTypeToken = Trim$(tableType)
    End Select
End Function

Private Function mp_GetCfgNonNegativeLongValue(ByVal cfg As Object, ByVal keyName As String) As Long
    Dim rawValue As String
    Dim parsedValue As Long

    rawValue = mp_GetCfgRequired(cfg, keyName)
    If Not mp_TryParseNonNegativeLong(rawValue, parsedValue) Then
        Err.Raise vbObjectError + 1760, "ex_PersonTimeline", _
            "Config key '" & keyName & "' must be a non-negative integer, got: '" & rawValue & "'."
    End If

    mp_GetCfgNonNegativeLongValue = parsedValue
End Function

Private Function mp_TryParseNonNegativeLong(ByVal rawValue As String, ByRef outValue As Long) As Boolean
    Dim textValue As String
    Dim i As Long
    Dim ch As String

    textValue = Trim$(rawValue)
    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    On Error GoTo ParseEH
    outValue = CLng(textValue)
    mp_TryParseNonNegativeLong = True
    Exit Function

ParseEH:
    mp_TryParseNonNegativeLong = False
End Function

Private Function mp_FindSourceAliasForTable(ByVal cfg As Object, ByVal tableAlias As String) As String
    Dim sourceAliases As Variant
    Dim aliases As Variant
    Dim i As Long
    Dim src As String
    Dim found As String

    tableAlias = Trim$(tableAlias)
    If Len(tableAlias) = 0 Then
        Err.Raise vbObjectError + 1335, "ex_PersonTimeline", "Output.Sheets contains an empty table alias."
    End If

    sourceAliases = mp_GetSourceAliases(cfg)
    For i = LBound(sourceAliases) To UBound(sourceAliases)
        src = CStr(sourceAliases(i))
        aliases = mp_GetListRequired(cfg, "Source." & src & ".SheetAliases")
        If mp_ArrayContainsText(aliases, tableAlias) Then
            If Len(found) > 0 Then
                Err.Raise vbObjectError + 1340, "ex_PersonTimeline", _
                    "Sheet alias '" & tableAlias & "' is declared in multiple sources: '" & found & "' and '" & src & "'."
            End If
            found = src
        End If
    Next i

    If Len(found) = 0 Then
        Err.Raise vbObjectError + 1341, "ex_PersonTimeline", _
            "Sheet alias '" & tableAlias & "' is not declared in any Source.*.SheetAliases."
    End If

    mp_FindSourceAliasForTable = found
End Function

Private Function mp_GetSourceAliases(ByVal cfg As Object) As Variant

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim key As Variant
    For Each key In cfg.Keys
        Dim k As String
        k = CStr(key)

        If LCase$(Left$(k, 7)) = "source." Then
            Dim p As Long
            p = InStr(8, k, ".", vbBinaryCompare)
            If p > 8 Then
                Dim srcAlias As String
                srcAlias = Mid$(k, 8, p - 8)
                If Len(srcAlias) > 0 Then
                    d(srcAlias) = srcAlias
                End If
            End If
        End If
    Next key

    If d.Count = 0 Then
        Err.Raise vbObjectError + 1350, "ex_PersonTimeline", "No Source.* keys found in config."
    End If

    Dim arr() As String
    ReDim arr(0 To d.Count - 1)

    Dim i As Long
    i = 0
    For Each key In d.Keys
        arr(i) = CStr(key)
        i = i + 1
    Next key

    mp_GetSourceAliases = arr

End Function

Private Function mp_GetConnectionForSource(ByVal connCache As Object, ByVal cfg As Object, ByVal sourceAlias As String) As Object
    Dim sourcePath As String
    Dim snapshotPath As String
    Dim conn As Object

    If connCache.Exists(sourceAlias) Then
        Set mp_GetConnectionForSource = connCache(sourceAlias)
        Exit Function
    End If

    sourcePath = mp_GetResolvedSourcePath(cfg, sourceAlias)

    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 1360, "ex_PersonTimeline", "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)

    Set conn = CreateObject("ADODB.Connection")
    On Error GoTo EH
    conn.Open mp_BuildAdoConnectionString(snapshotPath)
    On Error GoTo 0

    connCache.Add sourceAlias, conn
    Set mp_GetConnectionForSource = conn
    Exit Function

EH:
    Err.Raise vbObjectError + 1362, "ex_PersonTimeline", _
        "ADO connection failed for source '" & sourceAlias & "' (source: " & sourcePath & ", snapshot: " & snapshotPath & "): " & Err.Description
End Function

Private Sub mp_CloseConnections(ByVal connCache As Object)
    Dim key As Variant
    Dim conn As Object

    If connCache Is Nothing Then Exit Sub

    On Error Resume Next
    For Each key In connCache.Keys
        Set conn = connCache(key)
        If Not conn Is Nothing Then
            If conn.State <> 0 Then conn.Close
        End If
    Next key
    connCache.RemoveAll
    On Error GoTo 0
End Sub

Private Function mp_BuildAdoConnectionString(ByVal sourcePath As String) As String
    Dim ext As String
    Dim props As String

    ext = LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))

    Select Case ext
        Case "xls"
            props = "Excel 8.0;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsx"
            props = "Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsm"
            props = "Excel 12.0 Macro;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsb"
            props = "Excel 12.0;HDR=YES;IMEX=1;ReadOnly=True"
        Case Else
            Err.Raise vbObjectError + 1363, "ex_PersonTimeline", "Unsupported source file extension for ADO: ." & ext
    End Select

    mp_BuildAdoConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""" & props & """;"
End Function

Private Function mp_ResolveAdoTableReference(ByVal adoConn As Object, ByVal tableName As String, Optional ByVal expectedHeaders As Variant, Optional ByVal keyHeader As String = vbNullString, Optional ByVal keyValue As String = vbNullString) As String
    Dim schemaRs As Object
    Dim schemaName As String
    Dim schemaNameClean As String
    Dim resolvedName As String
    Dim listedNames As String
    Dim listedCount As Long
    Dim fallbackName As String
    Dim fallbackScore As Long
    Dim currentScore As Long
    Dim keyMatchCount As Long

    Set schemaRs = adoConn.OpenSchema(20)

    Do While Not schemaRs.EOF
        schemaName = CStr(schemaRs.Fields("TABLE_NAME").Value)
        schemaNameClean = mp_CleanAdoSchemaObjectName(schemaName)
        If listedCount < 15 Then
            If Len(listedNames) > 0 Then listedNames = listedNames & ", "
            listedNames = listedNames & schemaNameClean
            listedCount = listedCount + 1
        End If
        If mp_IsMatchingAdoTableName(schemaNameClean, tableName) Then
            resolvedName = schemaNameClean
            Exit Do
        End If

        If Not mp_IsEmptyVariantArray(expectedHeaders) Then
            currentScore = mp_AdoObjectMatchScore(adoConn, schemaNameClean, expectedHeaders)
        Else
            currentScore = 0
        End If

        If Len(keyHeader) > 0 And Len(keyValue) > 0 Then
            keyMatchCount = mp_AdoObjectKeyMatchCount(adoConn, schemaNameClean, keyHeader, keyValue)
            If keyMatchCount > 0 Then
                currentScore = currentScore + (100000 + keyMatchCount)
            End If
        End If

        If currentScore > fallbackScore Then
            fallbackScore = currentScore
            fallbackName = schemaNameClean
        End If
        schemaRs.MoveNext
    Loop

    schemaRs.Close

    If Len(resolvedName) = 0 And Len(fallbackName) > 0 And fallbackScore > 0 Then
        resolvedName = fallbackName
    End If

    If Len(resolvedName) = 0 Then
        Err.Raise vbObjectError + 1302, "ex_PersonTimeline", _
            "ADO table/range not found: '" & tableName & "'. Available objects: " & listedNames
    End If

    mp_ResolveAdoTableReference = mp_QuoteSqlIdentifier(resolvedName)
End Function

Private Function mp_AdoObjectKeyMatchCount(ByVal adoConn As Object, ByVal objectName As String, ByVal keyHeader As String, ByVal keyValue As String) As Long
    Dim rs As Object
    Dim sql As String

    On Error GoTo EH

    If Len(Trim$(keyHeader)) = 0 Then Exit Function

    Set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT COUNT(*) AS Cnt FROM " & mp_QuoteSqlIdentifier(objectName) & " WHERE " & mp_BuildAdoWhereEquals(keyHeader, keyValue)

    rs.Open sql, adoConn, 0, 1
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0).Value) Then
            mp_AdoObjectKeyMatchCount = CLng(rs.Fields(0).Value)
        End If
    End If

    rs.Close
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_BuildExpectedHeaders(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fields As Variant, ByVal keyHeader As String) As Variant
    Dim d As Object
    Dim i As Long
    Dim h As Variant
    Dim out() As String
    Dim idx As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    h = Trim$(keyHeader)
    If Len(h) > 0 Then d(h) = h

    For i = LBound(fields) To UBound(fields)
        h = Trim$(mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, Trim$(CStr(fields(i)))))
        If Len(h) > 0 Then d(h) = h
    Next i

    If d.Count = 0 Then
        mp_BuildExpectedHeaders = Array()
        Exit Function
    End If

    ReDim out(0 To d.Count - 1)
    idx = 0
    For Each h In d.Keys
        out(idx) = CStr(h)
        idx = idx + 1
    Next h

    mp_BuildExpectedHeaders = out
End Function

Private Function mp_AdoObjectMatchScore(ByVal adoConn As Object, ByVal objectName As String, ByVal expectedHeaders As Variant) As Long
    Dim rs As Object
    Dim i As Long
    Dim headerToken As String

    If mp_IsEmptyVariantArray(expectedHeaders) Then Exit Function

    On Error GoTo EH
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & mp_QuoteSqlIdentifier(objectName) & " WHERE 1=0", adoConn, 0, 1

    For i = LBound(expectedHeaders) To UBound(expectedHeaders)
        headerToken = Trim$(CStr(expectedHeaders(i)))
        If mp_AdoRecordsetHasField(rs, headerToken) Then
            mp_AdoObjectMatchScore = mp_AdoObjectMatchScore + 1
        End If
    Next i

    rs.Close
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_AdoRecordsetHasField(ByVal rs As Object, ByVal fieldName As String) As Boolean
    mp_AdoRecordsetHasField = (mp_RecordsetGetFieldOrdinal(rs, fieldName) >= 0)
End Function

Private Function mp_CleanAdoSchemaObjectName(ByVal value As String) As String
    value = Trim$(value)

    If Len(value) >= 2 Then
        If Left$(value, 1) = "'" And Right$(value, 1) = "'" Then
            value = Mid$(value, 2, Len(value) - 2)
        End If
    End If

    mp_CleanAdoSchemaObjectName = value
End Function

Private Function mp_AdoObjectHasField(ByVal adoConn As Object, ByVal objectName As String, ByVal fieldName As String) As Boolean
    Dim rs As Object
    Dim sql As String
    Dim i As Long

    On Error GoTo EH

    sql = "SELECT * FROM " & mp_QuoteSqlIdentifier(objectName) & " WHERE 1=0"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, adoConn, 0, 1

    For i = 0 To rs.Fields.Count - 1
        If StrComp(mp_NormalizeHeader(CStr(rs.Fields(i).Name)), mp_NormalizeHeader(fieldName), vbTextCompare) = 0 Then
            mp_AdoObjectHasField = True
            Exit For
        End If
    Next i

    rs.Close
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_IsMatchingAdoTableName(ByVal candidate As String, ByVal requested As String) As Boolean
    mp_IsMatchingAdoTableName = (StrComp(mp_NormalizeAdoObjectName(candidate), mp_NormalizeAdoObjectName(requested), vbTextCompare) = 0)
End Function

Private Function mp_IsMatchingConfiguredAdoObject(ByVal candidate As String, ByVal configured As String) As Boolean
    Dim candidateExact As String
    Dim configuredExact As String

    candidateExact = mp_NormalizeAdoObjectNameExact(candidate)
    configuredExact = mp_NormalizeAdoObjectNameExact(configured)

    If StrComp(candidateExact, configuredExact, vbTextCompare) = 0 Then
        mp_IsMatchingConfiguredAdoObject = True
        Exit Function
    End If

    ' Configured short sheet name (without "$") should still match sheet object names like "Sheet1$".
    If InStr(1, configuredExact, "$", vbBinaryCompare) = 0 Then
        If StrComp(mp_NormalizeAdoObjectName(candidateExact), mp_NormalizeAdoObjectName(configuredExact), vbTextCompare) = 0 Then
            mp_IsMatchingConfiguredAdoObject = True
            Exit Function
        End If
    End If
End Function

Private Function mp_IsExplicitAdoRangeReference(ByVal value As String) As Boolean
    value = Trim$(value)
    If InStr(1, value, "$", vbBinaryCompare) <= 0 Then Exit Function
    If InStr(1, value, ":", vbBinaryCompare) <= 0 Then Exit Function
    mp_IsExplicitAdoRangeReference = True
End Function

Private Function mp_NormalizeAdoObjectName(ByVal value As String) As String
    Dim dollarPos As Long

    value = mp_NormalizeAdoObjectNameExact(value)

    dollarPos = InStr(1, value, "$", vbBinaryCompare)
    If dollarPos > 0 Then
        value = Left$(value, dollarPos - 1)
    End If

    mp_NormalizeAdoObjectName = LCase$(Trim$(value))
End Function

Private Function mp_NormalizeAdoObjectNameExact(ByVal value As String) As String
    value = mp_CleanAdoSchemaObjectName(value)
    value = mp_StripAdoObjectOrdinalPrefix(value)

    If Len(value) >= 2 Then
        If Left$(value, 1) = "[" And Right$(value, 1) = "]" Then
            value = Mid$(value, 2, Len(value) - 2)
        End If
    End If

    If Len(value) >= 2 Then
        If Left$(value, 1) = "'" And Right$(value, 1) = "'" Then
            value = Mid$(value, 2, Len(value) - 2)
        End If
    End If

    mp_NormalizeAdoObjectNameExact = LCase$(Trim$(value))
End Function

Private Function mp_StripAdoObjectOrdinalPrefix(ByVal value As String) As String
    Dim i As Long
    Dim j As Long
    Dim prefix As String
    Dim marker As String

    value = Trim$(value)
    If Len(value) = 0 Then
        mp_StripAdoObjectOrdinalPrefix = value
        Exit Function
    End If

    j = 1
    Do While j <= Len(value)
        If Mid$(value, j, 1) < "0" Or Mid$(value, j, 1) > "9" Then Exit Do
        j = j + 1
    Loop

    If j > 1 And j <= Len(value) Then
        marker = Mid$(value, j, 1)
        If marker = "#" Or marker = "." Or marker = ")" Or marker = "-" Then
            prefix = Trim$(Left$(value, j - 1))
            If IsNumeric(prefix) Then
                i = j + 1
                Do While i <= Len(value) And Mid$(value, i, 1) = " "
                    i = i + 1
                Loop
                value = Mid$(value, i)
            End If
        End If
    End If

    mp_StripAdoObjectOrdinalPrefix = value
End Function

Private Function mp_UnquoteSqlIdentifier(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) >= 2 Then
        If Left$(value, 1) = "[" And Right$(value, 1) = "]" Then
            value = Mid$(value, 2, Len(value) - 2)
        End If
    End If
    mp_UnquoteSqlIdentifier = Replace$(value, "]]", "]")
End Function

Private Function mp_ExtractAdoSheetPrefix(ByVal tableRef As String) As String
    Dim objectName As String
    Dim dollarPos As Long

    objectName = mp_UnquoteSqlIdentifier(tableRef)
    If Len(objectName) = 0 Then Exit Function

    dollarPos = InStr(1, objectName, "$", vbBinaryCompare)
    If dollarPos <= 0 Then Exit Function

    mp_ExtractAdoSheetPrefix = Left$(objectName, dollarPos)
End Function

Private Function mp_BuildNormalizedHeaderTokenSet(ByVal expectedHeaders As Variant, ByVal keyHeader As String) As Object
    Dim d As Object
    Dim i As Long
    Dim token As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    token = mp_NormalizeHeader(keyHeader)
    If Len(token) > 0 Then d(token) = True

    If Not mp_IsEmptyVariantArray(expectedHeaders) Then
        For i = LBound(expectedHeaders) To UBound(expectedHeaders)
            token = mp_NormalizeHeader(CStr(expectedHeaders(i)))
            If Len(token) > 0 Then d(token) = True
        Next i
    End If

    Set mp_BuildNormalizedHeaderTokenSet = d
End Function

Private Function mp_BuildExpectedHeadersSignature(ByVal expectedHeaders As Variant) As String
    Dim d As Object
    Dim keys As Variant
    Dim i As Long
    Dim token As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    If Not mp_IsEmptyVariantArray(expectedHeaders) Then
        For i = LBound(expectedHeaders) To UBound(expectedHeaders)
            token = mp_NormalizeHeader(CStr(expectedHeaders(i)))
            If Len(token) > 0 Then d(token) = True
        Next i
    End If

    If d.Count = 0 Then
        mp_BuildExpectedHeadersSignature = "-"
        Exit Function
    End If

    keys = d.Keys
    mp_SortVariantTextArray keys
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then mp_BuildExpectedHeadersSignature = mp_BuildExpectedHeadersSignature & "|"
        mp_BuildExpectedHeadersSignature = mp_BuildExpectedHeadersSignature & CStr(keys(i))
    Next i
End Function

Private Function mp_QuoteSqlIdentifier(ByVal value As String) As String
    mp_QuoteSqlIdentifier = "[" & Replace$(value, "]", "]]" ) & "]"
End Function

Private Function mp_BuildAdoWhereEquals(ByVal columnName As String, ByVal valueText As String) As String
    mp_BuildAdoWhereEquals = mp_QuoteSqlIdentifier(columnName) & " = '" & Replace$(valueText, "'", "''") & "'"
End Function

Private Function mp_ResolveAdoMappedHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
    ByVal adoConn As Object, _
    ByVal tableRef As String _
) As String
    Dim desiredHeader As String
    Dim desiredToken As String
    Dim rs As Object
    Dim fieldMap As Object
    Dim tableCacheKey As String
    Dim availableFields As String
    Dim hasGenericFields As Boolean
    Dim hintText As String
    Dim i As Long
    Dim fieldName As String
    Dim fieldToken As String

    ' The mapping format is SourceHeader|Label, where Label is display-only.
    ' For SQL and field resolution we must always use the source header (left token).
    desiredHeader = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldAlias)
    desiredToken = mp_NormalizeHeader(desiredHeader)
    If Len(desiredToken) = 0 Then
        Err.Raise vbObjectError + 1391, "ex_PersonTimeline", _
            "Configured source header is empty for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]."
    End If

    mp_EnsureAdoLookupCacheContainers
    tableCacheKey = LCase$(Trim$(sourceAlias) & "|" & Trim$(tableAlias) & "|" & mp_NormalizeAdoObjectNameExact(tableRef))

    If g_AdoFieldMapByTableRef.Exists(tableCacheKey) Then
        Set fieldMap = g_AdoFieldMapByTableRef.Item(tableCacheKey)
        If g_AdoFieldListByTableRef.Exists(tableCacheKey) Then
            availableFields = CStr(g_AdoFieldListByTableRef(tableCacheKey))
        End If
        If g_AdoFieldGenericByTableRef.Exists(tableCacheKey) Then
            hasGenericFields = CBool(g_AdoFieldGenericByTableRef(tableCacheKey))
        End If
    End If

    On Error GoTo EH
    If fieldMap Is Nothing Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM " & tableRef & " WHERE 1=0", adoConn, 0, 1

        Set fieldMap = CreateObject("Scripting.Dictionary")
        fieldMap.CompareMode = 1

        For i = 0 To rs.Fields.Count - 1
            fieldName = CStr(rs.Fields(i).Name)
            fieldToken = mp_NormalizeHeader(fieldName)
            If Len(fieldToken) > 0 Then
                If Not fieldMap.Exists(fieldToken) Then
                    fieldMap.Add fieldToken, fieldName
                End If
            End If
        Next i

        availableFields = mp_ListAdoRecordsetFields(rs, 25)
        hasGenericFields = mp_RecordsetLooksLikeUnnamedFields(rs)
        rs.Close

        If g_AdoFieldMapByTableRef.Exists(tableCacheKey) Then
            Set g_AdoFieldMapByTableRef.Item(tableCacheKey) = fieldMap
        Else
            g_AdoFieldMapByTableRef.Add tableCacheKey, fieldMap
        End If
        g_AdoFieldListByTableRef(tableCacheKey) = availableFields
        g_AdoFieldGenericByTableRef(tableCacheKey) = hasGenericFields
    End If

    If fieldMap.Exists(desiredToken) Then
        mp_ResolveAdoMappedHeader = CStr(fieldMap(desiredToken))
        Exit Function
    End If

    If hasGenericFields Then
        hintText = " Hint: ADO returned generic fields (F1..Fn). Set '" & sourceAlias & ".Sheet[" & tableAlias & "].SheetName' " & _
                   "to an explicit range where the first row contains real headers (example: ШПС$A10:K5000)."
    End If
    If Len(Trim$(availableFields)) = 0 Then availableFields = "(none)"

    Err.Raise vbObjectError + 1391, "ex_PersonTimeline", _
        "Configured source header '" & desiredHeader & "' is not found for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]. " & _
        "Available fields: " & availableFields & "." & hintText
    Exit Function

EH:
    Dim innerErrNumber As Long
    Dim innerErrSource As String
    Dim innerErrDescription As String
    innerErrNumber = Err.Number
    innerErrSource = Err.Source
    innerErrDescription = Err.Description

    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0

    If innerErrNumber = vbObjectError + 1391 Then
        Err.Raise innerErrNumber, innerErrSource, innerErrDescription
    End If
    Err.Raise vbObjectError + 1391, "ex_PersonTimeline", _
        "Failed to resolve source header for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]: " & innerErrDescription
End Function

Private Function mp_ListAdoRecordsetFields(ByVal rs As Object, Optional ByVal maxCount As Long = 25) As String
    Dim i As Long
    Dim count As Long
    Dim nameText As String

    If rs Is Nothing Then Exit Function
    If maxCount <= 0 Then maxCount = 25

    For i = 0 To rs.Fields.Count - 1
        If count > 0 Then mp_ListAdoRecordsetFields = mp_ListAdoRecordsetFields & ", "
        nameText = Trim$(CStr(rs.Fields(i).Name))
        If Len(nameText) = 0 Then nameText = "(empty)"
        mp_ListAdoRecordsetFields = mp_ListAdoRecordsetFields & "[" & nameText & "]"
        count = count + 1
        If count >= maxCount Then Exit For
    Next i

    If rs.Fields.Count > maxCount Then
        mp_ListAdoRecordsetFields = mp_ListAdoRecordsetFields & ", ..."
    End If
End Function

Private Function mp_RecordsetLooksLikeUnnamedFields(ByVal rs As Object) As Boolean
    Dim i As Long
    Dim fieldName As String
    Dim genericCount As Long
    Dim nonEmptyCount As Long

    If rs Is Nothing Then Exit Function

    For i = 0 To rs.Fields.Count - 1
        fieldName = Trim$(CStr(rs.Fields(i).Name))
        If Len(fieldName) = 0 Then GoTo ContinueField
        nonEmptyCount = nonEmptyCount + 1
        If mp_IsAdoGenericFieldName(fieldName) Then genericCount = genericCount + 1
ContinueField:
    Next i

    If nonEmptyCount = 0 Then Exit Function
    mp_RecordsetLooksLikeUnnamedFields = ((genericCount * 2) >= nonEmptyCount)
End Function

Private Function mp_IsAdoGenericFieldName(ByVal value As String) As Boolean
    Dim i As Long
    value = Trim$(value)
    If Len(value) < 2 Then Exit Function
    If Left$(value, 1) <> "F" And Left$(value, 1) <> "f" Then Exit Function
    For i = 2 To Len(value)
        If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Exit Function
    Next i
    mp_IsAdoGenericFieldName = True
End Function

Private Function mp_BuildAdoWhereLike(ByVal columnName As String, ByVal valueText As String) As String
    Dim escaped As String
    Dim likeStar As String
    Dim likePercent As String
    Dim colExpr As String

    escaped = mp_EscapeAdoLikeValue(valueText)
    colExpr = mp_QuoteSqlIdentifier(columnName)
    likeStar = "'*" & Replace$(escaped, "'", "''") & "*'"
    likePercent = "'%" & Replace$(escaped, "'", "''") & "%'"

    ' Keep partial-match SQL simple to avoid ACE parser/type issues on numeric-like column names (e.g. [56]).
    mp_BuildAdoWhereLike = "(" & colExpr & " LIKE " & likeStar & " OR " & colExpr & " LIKE " & likePercent & ")"
End Function

Private Function mp_EscapeAdoLikeValue(ByVal valueText As String) As String
    valueText = Replace$(valueText, "[", "[[]")
    valueText = Replace$(valueText, "*", "[*]")
    valueText = Replace$(valueText, "?", "[?]")
    valueText = Replace$(valueText, "#", "[#]")
    valueText = Replace$(valueText, "%", "[%]")
    valueText = Replace$(valueText, "_", "[_]")
    mp_EscapeAdoLikeValue = valueText
End Function

Private Function mp_RecordsetGetFieldOrdinal(ByVal rs As Object, ByVal fieldName As String) As Long
    Dim i As Long

    mp_RecordsetGetFieldOrdinal = -1
    If rs Is Nothing Then Exit Function

    For i = 0 To rs.Fields.Count - 1
        If StrComp(mp_NormalizeHeader(CStr(rs.Fields(i).Name)), mp_NormalizeHeader(fieldName), vbTextCompare) = 0 Then
            mp_RecordsetGetFieldOrdinal = i
            Exit Function
        End If
    Next i
End Function

Private Function mp_ToSafeText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function

    mp_ToSafeText = Trim$(CStr(valueIn))
End Function

Private Function mp_NormalizeSearchToken(ByVal valueText As String) As String
    valueText = Replace$(valueText, vbCr, " ")
    valueText = Replace$(valueText, vbLf, " ")
    valueText = Replace$(valueText, vbTab, " ")
    valueText = Trim$(valueText)

    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
        valueText = Replace$(valueText, "  ", " ")
    Loop

    mp_NormalizeSearchToken = valueText
End Function

Private Function mp_IsLikelyFullPersonKey(ByVal searchKey As String) As Boolean
    Dim normalized As String
    Dim tokens As Variant
    Dim i As Long
    Dim tokenCount As Long

    normalized = mp_NormalizeSearchToken(searchKey)
    If Len(normalized) = 0 Then Exit Function

    tokens = Split(normalized, " ")
    For i = LBound(tokens) To UBound(tokens)
        If Len(Trim$(CStr(tokens(i)))) > 0 Then tokenCount = tokenCount + 1
    Next i

    mp_IsLikelyFullPersonKey = (tokenCount >= 2)
End Function

Private Function mp_RenderEventsNoData( _
    ByVal wsOut As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Variant, _
    ByVal headerRows As Collection, _
    ByVal resultFieldRanges As Collection, _
    ByVal resultTable As obj_ResultTable _
) As Long
    Dim outHeaderRow As Long
    Dim outDataRow As Long
    Dim fieldCount As Long
    Dim i As Long
    Dim headerValues() As Variant
    Dim fieldAlias As String

    outHeaderRow = rowIndex
    headerRows.Add outHeaderRow

    fieldCount = UBound(fields) - LBound(fields) + 1
    ReDim headerValues(1 To 1, 1 To fieldCount)

    For i = LBound(fields) To UBound(fields)
        fieldAlias = Trim$(CStr(fields(i)))
        headerValues(1, 1 + (i - LBound(fields))) = mp_GetFieldLabel(cfg, sourceAlias, tableAlias, fieldAlias)
    Next i

    wsOut.Range(wsOut.Cells(outHeaderRow, 1), wsOut.Cells(outHeaderRow, fieldCount)).Value = headerValues

    outDataRow = outHeaderRow + 1
    wsOut.Cells(outDataRow, 1).Value = "(no events found for this person)"
    mp_AddResultFieldRangesForFields resultFieldRanges, cfg, sourceAlias, tableAlias, fields, outHeaderRow, outDataRow
    mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, fields, outDataRow, outDataRow, visualRowStart:=outHeaderRow, visualRowEnd:=outDataRow

    mp_RenderEventsNoData = outDataRow + 1
End Function

Private Function mp_RenderStateNoData( _
    ByVal wsOut As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Variant, _
    ByVal headerRows As Collection, _
    ByVal resultFieldRanges As Collection, _
    ByVal resultTable As obj_ResultTable _
) As Long
    Dim outHeaderRow As Long
    Dim outDataRow As Long
    Dim fieldCount As Long
    Dim i As Long
    Dim headerValues() As Variant
    Dim fieldAlias As String

    outHeaderRow = rowIndex
    headerRows.Add outHeaderRow

    fieldCount = UBound(fields) - LBound(fields) + 1
    ReDim headerValues(1 To 1, 1 To fieldCount)

    For i = LBound(fields) To UBound(fields)
        fieldAlias = Trim$(CStr(fields(i)))
        headerValues(1, 1 + (i - LBound(fields))) = mp_GetFieldLabel(cfg, sourceAlias, tableAlias, fieldAlias)
    Next i

    wsOut.Range(wsOut.Cells(outHeaderRow, 1), wsOut.Cells(outHeaderRow, fieldCount)).Value = headerValues

    outDataRow = outHeaderRow + 1
    wsOut.Cells(outDataRow, 1).Value = "(no state found for this person)"
    mp_AddResultFieldRangesForFields resultFieldRanges, cfg, sourceAlias, tableAlias, fields, outHeaderRow, outDataRow
    mp_CaptureResultTableRowsFromOutput wsOut, resultTable, sourceAlias, tableAlias, fields, outDataRow, outDataRow, visualRowStart:=outHeaderRow, visualRowEnd:=outDataRow

    mp_RenderStateNoData = outDataRow + 1
End Function

Private Function mp_RenderStateCandidatesWarningBanner( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal searchKey As String, _
    ByVal candidateCount As Long, _
    ByVal pendingWarningBanners As Collection _
) As Long
    Dim bannerCols As Long
    Dim bannerRows As Long
    Dim bannerRangeAddress As String
    Dim titleText As String
    Dim messageText As String
    Dim entry As Object

    If startRow < 1 Then startRow = 1
    mp_GetWarningBannerDimensions bannerCols, bannerRows
    bannerRows = 1
    bannerRangeAddress = "A" & CStr(startRow) & ":" & mp_ToColumnLetter(bannerCols) & CStr(startRow + bannerRows - 1)

    titleText = "WARNING: Multiple candidates found"
    messageText = "Search key '" & searchKey & "' returned " & CStr(candidateCount) & " matches. Select the correct candidate from the list below, copy the full key value, paste it into the search field, and run search again."

    If Not pendingWarningBanners Is Nothing Then
        Set entry = CreateObject("Scripting.Dictionary")
        entry.CompareMode = 1
        entry("RangeAddress") = bannerRangeAddress
        entry("Title") = titleText
        entry("Message") = messageText
        pendingWarningBanners.Add entry
    End If

    ex_Messaging.m_RenderWarningBanner ws, messageText, titleText, bannerRangeAddress

    mp_RenderStateCandidatesWarningBanner = startRow + bannerRows + 1
End Function

Private Sub mp_RenderPendingWarningBanners(ByVal ws As Worksheet, ByVal pendingWarningBanners As Collection)
    Dim i As Long
    Dim entry As Object

    If ws Is Nothing Then Exit Sub
    If pendingWarningBanners Is Nothing Then Exit Sub

    For i = 1 To pendingWarningBanners.Count
        Set entry = pendingWarningBanners(i)
        If entry Is Nothing Then GoTo ContinueBanner
        ex_Messaging.m_RenderWarningBanner ws, CStr(entry("Message")), CStr(entry("Title")), CStr(entry("RangeAddress"))
ContinueBanner:
    Next i
End Sub

Private Sub mp_GetWarningBannerDimensions(ByRef outColumns As Long, ByRef outRows As Long)
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle

    If ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
        outColumns = bannerStyle.Columns
        outRows = bannerStyle.Rows
    ElseIf ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
        outColumns = bannerStyle.Columns
        outRows = bannerStyle.Rows
    End If

    If outColumns < 1 Then outColumns = 8
    If outRows < 1 Then outRows = 3
End Sub

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

Private Function mp_ToCellValue(ByVal valueIn As Variant) As Variant
    If IsError(valueIn) Then
        mp_ToCellValue = vbNullString
        Exit Function
    End If
    If IsNull(valueIn) Then
        mp_ToCellValue = vbNullString
        Exit Function
    End If

    mp_ToCellValue = valueIn
End Function

Private Function mp_GetWorkbookForSource(ByVal wbCache As Object, ByVal cfg As Object, ByVal sourceAlias As String) As Workbook
    Dim sourcePath As String
    Dim snapshotPath As String

    If wbCache.Exists(sourceAlias) Then
        Set mp_GetWorkbookForSource = wbCache(sourceAlias)
        Exit Function
    End If

    sourcePath = mp_GetResolvedSourcePath(cfg, sourceAlias)

    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 1360, "ex_PersonTimeline", "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)

    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=snapshotPath, ReadOnly:=True, UpdateLinks:=0)

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo 0

    wbCache.Add sourceAlias, wb
    Set mp_GetWorkbookForSource = wb

End Function

Private Sub mp_CloseWorkbooks(ByVal wbCache As Object)

    If wbCache Is Nothing Then Exit Sub

    On Error Resume Next
    Dim key As Variant
    For Each key In wbCache.Keys
        Dim wb As Workbook
        Set wb = wbCache(key)
        If Not wb Is Nothing Then
            wb.Close SaveChanges:=False
        End If
    Next key
    wbCache.RemoveAll
    On Error GoTo 0

End Sub

Private Function mp_GetCfgRequired(ByVal cfg As Object, ByVal keyName As String) As String

    If Not cfg.Exists(keyName) Then
        Err.Raise vbObjectError + 1370, "ex_PersonTimeline", "Missing config key: " & keyName
    End If

    Dim valueText As String
    valueText = Trim$(CStr(cfg(keyName)))

    If Len(valueText) = 0 Then
        Err.Raise vbObjectError + 1371, "ex_PersonTimeline", "Empty config value: " & keyName
    End If

    mp_GetCfgRequired = valueText

End Function

Private Function mp_GetCfgOptional(ByVal cfg As Object, ByVal keyName As String, ByVal defaultValue As String) As String

    If cfg.Exists(keyName) Then
        mp_GetCfgOptional = Trim$(CStr(cfg(keyName)))
    Else
        mp_GetCfgOptional = defaultValue
    End If

End Function

Private Function mp_GetListRequired(ByVal cfg As Object, ByVal keyName As String) As Variant

    Dim raw As String
    raw = mp_GetCfgRequired(cfg, keyName)
    mp_GetListRequired = mp_SplitList(raw)

    If mp_IsEmptyVariantArray(mp_GetListRequired) Then
        Err.Raise vbObjectError + 1380, "ex_PersonTimeline", "List is empty for config key: " & keyName
    End If

End Function

Private Function mp_GetOrderedFieldAliases(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As Variant
    Dim ordered As Variant
    ordered = mp_GetMapAliasesInConfigOrder(sourceAlias, tableAlias)

    If Not mp_IsEmptyVariantArray(ordered) Then
        mp_GetOrderedFieldAliases = ordered
        Exit Function
    End If

    mp_GetOrderedFieldAliases = mp_GetListRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].FieldsAliases")
End Function

Private Function mp_GetEffectiveFieldAliases(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As Variant
    Dim baseFields As Variant
    Dim dslVirtuals As Variant

    baseFields = mp_GetOrderedFieldAliases(cfg, sourceAlias, tableAlias)
    dslVirtuals = ex_FetchDslEngine.m_GetVirtualColumns(cfg, sourceAlias, tableAlias)

    If Not mp_IsEmptyVariantArray(dslVirtuals) Then
        mp_GetEffectiveFieldAliases = mp_AppendFieldAliases(baseFields, dslVirtuals)
        Exit Function
    End If

    mp_GetEffectiveFieldAliases = baseFields
End Function

Private Function mp_AppendFieldAliases(ByVal baseFields As Variant, ByVal appendFields As Variant) As Variant
    Dim result As Object
    Dim i As Long
    Dim token As String
    Dim arr() As String
    Dim key As Variant
    Dim idx As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not mp_IsEmptyVariantArray(baseFields) Then
        For i = LBound(baseFields) To UBound(baseFields)
            token = Trim$(CStr(baseFields(i)))
            If Len(token) > 0 Then result(token) = True
        Next i
    End If

    If Not mp_IsEmptyVariantArray(appendFields) Then
        For i = LBound(appendFields) To UBound(appendFields)
            token = Trim$(CStr(appendFields(i)))
            If Len(token) > 0 Then result(token) = True
        Next i
    End If

    If result.Count = 0 Then
        mp_AppendFieldAliases = Array()
        Exit Function
    End If

    ReDim arr(0 To result.Count - 1)
    idx = 0
    For Each key In result.Keys
        arr(idx) = CStr(key)
        idx = idx + 1
    Next key

    mp_AppendFieldAliases = arr
End Function

Private Function mp_IsVirtualFieldAlias(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldAlias As String) As Boolean
    mp_IsVirtualFieldAlias = ex_FetchDslEngine.m_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias)
End Function

Private Function mp_GetMapAliasesInConfigOrder(ByVal sourceAlias As String, ByVal tableAlias As String) As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim prefix As String
    Dim seen As Object
    Dim aliases() As String
    Dim count As Long
    Dim r As Long
    Dim markerText As String
    Dim keyText As String
    Dim suffix As String
    Dim closingPos As Long
    Dim fieldAlias As String

    On Error GoTo EH

    Set ws = ws_Dev
    Set tbl = ws.ListObjects(DEV_CONFIG_TABLE_NAME)
    If tbl Is Nothing Then
        mp_GetMapAliasesInConfigOrder = Array()
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        mp_GetMapAliasesInConfigOrder = Array()
        Exit Function
    End If

    Set dataRange = tbl.DataBodyRange
    prefix = LCase$(sourceAlias & ".Sheet[" & tableAlias & "].Map[")
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    count = 0

    For r = 1 To dataRange.Rows.Count
        markerText = Trim$(CStr(dataRange.Cells(r, DEV_COL_MARKER).Value))
        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then GoTo ContinueRow

        keyText = Trim$(CStr(dataRange.Cells(r, DEV_COL_KEY).Value))
        If Len(keyText) = 0 Then GoTo ContinueRow
        If LCase$(Left$(keyText, Len(prefix))) <> prefix Then GoTo ContinueRow

        suffix = Mid$(keyText, Len(prefix) + 1)
        closingPos = InStr(1, suffix, "]", vbBinaryCompare)
        If closingPos <= 1 Then GoTo ContinueRow
        If Len(Trim$(Mid$(suffix, closingPos + 1))) <> 0 Then GoTo ContinueRow

        fieldAlias = Trim$(Left$(suffix, closingPos - 1))
        If Len(fieldAlias) = 0 Then GoTo ContinueRow
        If seen.Exists(fieldAlias) Then GoTo ContinueRow

        seen.Add fieldAlias, True
        ReDim Preserve aliases(0 To count)
        aliases(count) = fieldAlias
        count = count + 1

ContinueRow:
    Next r

    If count = 0 Then
        mp_GetMapAliasesInConfigOrder = Array()
    Else
        mp_GetMapAliasesInConfigOrder = aliases
    End If
    Exit Function

EH:
    mp_GetMapAliasesInConfigOrder = Array()
End Function

Private Function mp_SplitList(ByVal raw As String) As Variant

    raw = Trim$(raw)
    If Len(raw) = 0 Then
        mp_SplitList = Array()
        Exit Function
    End If

    raw = Replace$(raw, ",", ";")

    Dim parts As Variant
    parts = Split(raw, ";")

    Dim count As Long
    count = 0

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        mp_SplitList = Array()
        Exit Function
    End If

    Dim out() As String
    ReDim out(0 To count - 1)

    Dim j As Long
    j = 0
    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then
            out(j) = token
            j = j + 1
        End If
    Next i

    mp_SplitList = out

End Function

Private Function mp_ArrayContainsText(ByVal values As Variant, ByVal needle As String) As Boolean

    If mp_IsEmptyVariantArray(values) Then Exit Function

    Dim i As Long
    For i = LBound(values) To UBound(values)
        If StrComp(Trim$(CStr(values(i))), Trim$(needle), vbTextCompare) = 0 Then
            mp_ArrayContainsText = True
            Exit Function
        End If
    Next i

End Function

Private Function mp_GetMappedSourceHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String

    Dim raw As String
    raw = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]")

    Dim p As Long
    p = InStr(1, raw, "|", vbBinaryCompare)

    If p > 0 Then
        mp_GetMappedSourceHeader = Trim$(Left$(raw, p - 1))
    Else
        mp_GetMappedSourceHeader = Trim$(raw)
    End If

    If Len(mp_GetMappedSourceHeader) = 0 Then
        Err.Raise vbObjectError + 1390, "ex_PersonTimeline", _
            "Mapped source header is empty for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If

End Function

Private Function mp_GetLabel( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String

    Dim raw As String
    raw = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]")

    Dim p As Long
    p = InStr(1, raw, "|", vbBinaryCompare)

    If p > 0 Then
        Dim lbl As String
        lbl = Trim$(Mid$(raw, p + 1))
        If Len(lbl) > 0 Then
            mp_GetLabel = lbl
            Exit Function
        End If
    End If

    mp_GetLabel = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldAlias)

End Function

Private Function mp_GetFieldLabel( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    If mp_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then
        mp_GetFieldLabel = fieldAlias
        Exit Function
    End If

    mp_GetFieldLabel = mp_GetLabel(cfg, sourceAlias, tableAlias, fieldAlias)
End Function

Private Function mp_TryGetTableColumnByFieldAlias( _
    ByVal lo As ListObject, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As Long

    Dim headerName As String
    headerName = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldAlias)

    mp_TryGetTableColumnByFieldAlias = mp_FindHeaderColumnInTable(lo, headerName)

End Function

Private Function mp_WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    sheetName = Trim$(sheetName)
    If Len(sheetName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    mp_WorksheetExists = Not ws Is Nothing
End Function

Private Function mp_CreateOrClearSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim usedRows As Long
    Dim usedCols As Long
    Dim clearRows As Long
    Dim clearCols As Long
    Dim clearRange As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ' Previous render may leave merged areas whose value is stored only in the top-left cell.
        ' Such merge extents are not reflected in last used column detection, so unmerge whole sheet first.
        On Error Resume Next
        ws.Cells.UnMerge
        On Error GoTo 0

        If ex_SheetStylesXmlProvider.m_GetUsedRangeSize(ws, usedRows, usedCols) Then
            clearRows = usedRows
            clearCols = usedCols
            If clearRows < ws.Rows.Count Then clearRows = clearRows + 2
            If clearCols < ws.Columns.Count Then clearCols = clearCols + 8
            Set clearRange = ws.Range(ws.Cells(1, 1), ws.Cells(clearRows, clearCols))

            On Error GoTo ClearFallback
            clearRange.UnMerge
            clearRange.Clear
            On Error GoTo 0
        Else
            ws.Cells(1, 1).Clear
        End If
    End If

    ws.Cells.NumberFormat = "@"

    Set mp_CreateOrClearSheet = ws
    Exit Function

ClearFallback:
    On Error Resume Next
    ws.Cells.Clear
    ws.Cells.NumberFormat = "@"
    Set mp_CreateOrClearSheet = ws
    On Error GoTo 0

End Function

Private Function mp_NormalizeHeader(ByVal s As String) As String
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, ChrW$(160), " ")
    ' ACE/OLEDB can expose dots in Excel headers as '#', e.g. "Вх. №" -> "Вх# №".
    s = Replace$(s, "#", ".")
    s = Replace$(s, ChrW$(&H2019), "'")
    s = Replace$(s, ChrW$(&H2BC), "'")
    s = Replace$(s, ChrW$(&H60), "'")
    s = Replace$(s, ChrW$(&HB4), "'")
    s = Replace$(s, "  ", " ")
    s = Replace$(s, "  ", " ")
    mp_NormalizeHeader = LCase$(Trim$(s))

End Function

Private Function mp_FindListObjectByName(ByVal wbSrc As Workbook, ByVal tableName As String) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wbSrc.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set mp_FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set mp_FindListObjectByName = Nothing

End Function

Private Function mp_FindHeaderColumnInTable(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim normalizedNeedle As String
    normalizedNeedle = mp_NormalizeHeader(headerName)

    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        If mp_NormalizeHeader(CStr(lo.HeaderRowRange.Cells(1, c).Value)) = normalizedNeedle Then
            mp_FindHeaderColumnInTable = c
            Exit Function
        End If
    Next c

    mp_FindHeaderColumnInTable = -1

End Function

Private Function mp_FindDataRowByKeyInTable(ByVal lo As ListObject, ByVal keyColIndex As Long, ByVal keyValue As String) As Long
    Dim matchedRows As Collection

    If lo.DataBodyRange Is Nothing Then
        mp_FindDataRowByKeyInTable = -1
        Exit Function
    End If

    Set matchedRows = mp_CollectMatchingRowsByKey(lo, keyColIndex, keyValue)
    If matchedRows.Count > 0 Then
        mp_FindDataRowByKeyInTable = CLng(matchedRows(1))
        Exit Function
    End If

    mp_FindDataRowByKeyInTable = -1

End Function

Private Function mp_CollectMatchingRowsByKey(ByVal lo As ListObject, ByVal keyColIndex As Long, ByVal keyValue As String) As Collection
    Dim matches As Collection
    Dim keyRange As Range
    Dim firstFound As Range
    Dim currentFound As Range
    Dim firstAddress As String
    Dim dataValues As Variant
    Dim rowCount As Long
    Dim r As Long

    Set matches = New Collection

    If lo Is Nothing Then
        Set mp_CollectMatchingRowsByKey = matches
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        Set mp_CollectMatchingRowsByKey = matches
        Exit Function
    End If
    If keyColIndex <= 0 Or keyColIndex > lo.ListColumns.Count Then
        Set mp_CollectMatchingRowsByKey = matches
        Exit Function
    End If

    Set keyRange = lo.ListColumns(keyColIndex).DataBodyRange
    If keyRange Is Nothing Then
        Set mp_CollectMatchingRowsByKey = matches
        Exit Function
    End If

    On Error Resume Next
    Set firstFound = keyRange.Find(What:=keyValue, After:=keyRange.Cells(keyRange.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    On Error GoTo 0

    If Not firstFound Is Nothing Then
        firstAddress = firstFound.Address
        Set currentFound = firstFound
        Do
            matches.Add (currentFound.Row - keyRange.Row + 1)
            Set currentFound = keyRange.FindNext(currentFound)
            If currentFound Is Nothing Then Exit Do
        Loop While currentFound.Address <> firstAddress
    End If

    If matches.Count > 0 Then
        Set mp_CollectMatchingRowsByKey = matches
        Exit Function
    End If

    dataValues = keyRange.Value2
    rowCount = UBound(dataValues, 1)
    For r = 1 To rowCount
        If StrComp(Trim$(CStr(dataValues(r, 1))), keyValue, vbTextCompare) = 0 Then
            matches.Add r
        End If
    Next r

    Set mp_CollectMatchingRowsByKey = matches
End Function

Private Sub mp_SortRangeByColumnIndex(ByVal ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, ByVal leftCol As Long, ByVal rightCol As Long, ByVal sortColRelative As Long)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))

    rng.Sort Key1:=ws.Cells(topRow + 1, leftCol + sortColRelative - 1), Order1:=xlAscending, Header:=xlYes

End Sub

Private Sub mp_NormalizeDateColumn(ByVal ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, ByVal colIndex As Long)
    Dim values As Variant
    Dim normalized() As Variant
    Dim r As Long
    Dim v As Variant
    Dim dt As Date
    Dim rowCount As Long

    If ws Is Nothing Then Exit Sub
    If topRow <= 0 Or bottomRow < topRow Then Exit Sub
    If colIndex <= 0 Then Exit Sub

    rowCount = bottomRow - topRow + 1
    If rowCount <= 0 Then Exit Sub

    values = ws.Range(ws.Cells(topRow, colIndex), ws.Cells(bottomRow, colIndex)).Value2
    ReDim normalized(1 To rowCount, 1 To 1)

    For r = 1 To rowCount
        v = values(r, 1)
        If mp_TryParseDate(v, dt) Then
            normalized(r, 1) = CDbl(dt)
        Else
            normalized(r, 1) = v
        End If
    Next r

    ws.Range(ws.Cells(topRow, colIndex), ws.Cells(bottomRow, colIndex)).Value2 = normalized
    ws.Range(ws.Cells(topRow, colIndex), ws.Cells(bottomRow, colIndex)).NumberFormat = "dd.mm.yyyy"
End Sub

Private Function mp_TryParseDate(ByVal valueIn As Variant, ByRef dateOut As Date) As Boolean
    Dim s As String
    Dim sep As String
    Dim parts As Variant
    Dim p1 As Long
    Dim p2 As Long
    Dim p3 As Long
    Dim d As Long
    Dim m As Long
    Dim y As Long

    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function

    s = Trim$(CStr(valueIn))
    If Len(s) = 0 Then
        If IsDate(valueIn) Then
            dateOut = CDate(valueIn)
            mp_TryParseDate = True
        End If
        Exit Function
    End If

    If InStr(1, s, ".", vbBinaryCompare) > 0 Then
        sep = "."
    ElseIf InStr(1, s, "/", vbBinaryCompare) > 0 Then
        sep = "/"
    ElseIf InStr(1, s, "-", vbBinaryCompare) > 0 Then
        sep = "-"
    Else
        If IsDate(valueIn) Then
            dateOut = CDate(valueIn)
            mp_TryParseDate = True
        End If
        Exit Function
    End If

    parts = Split(s, sep)
    If UBound(parts) - LBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function

    p1 = CLng(parts(0))
    p2 = CLng(parts(1))
    p3 = CLng(parts(2))

    If p1 > 31 Or p2 > 31 Then Exit Function

    If p3 < 100 Then
        If p3 <= 29 Then
            y = 2000 + p3
        Else
            y = 1900 + p3
        End If
    Else
        y = p3
    End If

    If sep = "." Then
        d = p1
        m = p2
    ElseIf sep = "/" Then
        m = p1
        d = p2
    Else
        If p1 > 12 And p2 <= 12 Then
            d = p1
            m = p2
        ElseIf p2 > 12 And p1 <= 12 Then
            m = p1
            d = p2
        Else
            d = p1
            m = p2
        End If
    End If

    On Error GoTo EH
    dateOut = DateSerial(y, m, d)
    mp_TryParseDate = True
    Exit Function

EH:
    mp_TryParseDate = False
End Function

Private Function mp_GetResolvedSourcePath(ByVal cfg As Object, ByVal sourceAlias As String) As String
    Dim sourcePrefix As String
    Dim fileKey As String
    Dim resolverKey As String
    Dim resolverArgsKey As String
    Dim rawPath As String
    Dim resolverName As String
    Dim resolverCallName As String
    Dim resolverArgs As String
    Dim resolvedValue As Variant
    Dim resolvedPath As String

    sourcePrefix = "Source." & Trim$(sourceAlias)
    fileKey = sourcePrefix & ".FilePath"
    resolverKey = sourcePrefix & ".FileResolver"
    resolverArgsKey = sourcePrefix & ".FileResolverArgs"

    rawPath = mp_GetCfgRequired(cfg, fileKey)
    resolverName = mp_GetCfgOptional(cfg, resolverKey, vbNullString)
    resolverArgs = mp_GetCfgOptional(cfg, resolverArgsKey, vbNullString)

    If Len(resolverName) = 0 Then
        If mp_HasPlaceholderTokens(rawPath) Then
            Err.Raise vbObjectError + 1762, "ex_PersonTimeline", _
                "Source path contains placeholders but no resolver is configured for key '" & fileKey & "'. " & _
                "Set '" & resolverKey & "' (for example: ex_SourceResolvers.m_ResolveLatestByDmyPattern)."
        End If

        mp_GetResolvedSourcePath = mp_ResolvePathLocal(rawPath)
        Exit Function
    End If

    If InStr(1, resolverName, "!", vbBinaryCompare) > 0 Then
        resolverCallName = resolverName
    Else
        resolverCallName = "'" & ThisWorkbook.Name & "'!" & resolverName
    End If

    On Error GoTo ResolverEH
    resolvedValue = Application.Run(resolverCallName, rawPath, resolverArgs)
    On Error GoTo 0

    resolvedPath = Trim$(CStr(resolvedValue))
    If Len(resolvedPath) = 0 Then
        Err.Raise vbObjectError + 1760, "ex_PersonTimeline", _
            "Source file resolver '" & resolverName & "' returned an empty path for key '" & fileKey & "'."
    End If

    mp_GetResolvedSourcePath = mp_ResolvePathLocal(resolvedPath)
    Exit Function

ResolverEH:
    Err.Raise vbObjectError + 1761, "ex_PersonTimeline", _
        "Source file resolver failed for key '" & fileKey & "' (resolver='" & resolverName & "'): " & Err.Description
End Function

Private Function mp_GetSourcePathSignatureValue(ByVal cfg As Object, ByVal sourceAlias As String) As String
    On Error GoTo EH

    mp_GetSourcePathSignatureValue = mp_GetResolvedSourcePath(cfg, sourceAlias)
    Exit Function

EH:
    mp_GetSourcePathSignatureValue = "#ERR:" & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function mp_HasPlaceholderTokens(ByVal valueText As String) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    mp_HasPlaceholderTokens = (InStr(1, normalized, "{", vbBinaryCompare) > 0) _
                              And (InStr(1, normalized, "}", vbBinaryCompare) > 0)
End Function

Private Function mp_ResolvePathLocal(ByVal inputPath As String) As String

    Dim basePath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then Exit Function

    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        mp_ResolvePathLocal = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_ResolvePathLocal = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    mp_ResolvePathLocal = basePath & inputPath

End Function

Private Function mp_IsEmptyVariantArray(ByVal v As Variant) As Boolean

    On Error GoTo EH

    If IsArray(v) = False Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    If UBound(v) < LBound(v) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    mp_IsEmptyVariantArray = False
    Exit Function

EH:
    mp_IsEmptyVariantArray = True

End Function
