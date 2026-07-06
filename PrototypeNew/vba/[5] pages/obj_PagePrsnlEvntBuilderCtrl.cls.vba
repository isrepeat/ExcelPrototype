VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePrsnlEvntBuilderCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PrsnlEvntBuilder.Controller"
Private Const CANDIDATE_TABLES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.EntityLookup.CandidateTables"
Private Const DUMMY_TABLES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.DummyTables"
Private Const HOTKEYS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Hotkeys"
Private Const SECTION_TYPES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.SectionTypes"
Private Const HOTKEY_ACTION_1 As String = "Action 1"
Private Const HOTKEY_ACTION_2 As String = "Action 2"
Private Const HOTKEY_SELECT_FORM_ROW As String = "Select Form Row"
Private Const EXPORT_CONFIG_PREFIX As String = "Export."
Private Const EXPORT_FILE_PATH_SUFFIX As String = ".FilePath"
Private Const EXPORT_CLASS_SUFFIX As String = ".ExporterClass"
Private Const EXPORT_SHEET_NAME_SUFFIX As String = ".SheetName"
Private Const EXPORT_RANGE_START_MARKER_SUFFIX As String = ".RangeStartMarker"
Private Const EXPORT_RANGE_END_MARKER_SUFFIX As String = ".RangeEndMarker"
Private Const EXPORT_ACTION_PREFIX As String = "Export "
Private Const DEFAULT_EXPORTER_CLASS As String = "obj_ExporterToDailyScope"
Private Const MAX_EXPORT_HOTKEYS As Long = 9
Private Const LOOKUP_CANDIDATES_CONTROL_NAME As String = "LookupCandidatesTable"
Private Const EVENT_DRAFT_FORM_CONTAINER_NAME As String = "EventDraftForm"
Private Const EVENT_DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"
Private Const EVENT_DRAFT_ORDER_NO_LABEL_CONTROL_NAME As String = "EventDraftOrderNoLabel"
Private Const EVENT_DRAFT_INCOMING_NO_HEADER_NAME As String = "Вх. №"
Private Const EVENT_DRAFT_ORDER_LABEL_PREFIX As String = "Наказ №: "
Private Const EXPORT_META_MANUAL_ORDER_NO_COLUMN_NAME As String = "meta_ManualOrderNo"
Private Const EXPORT_META_SECTION_TYPE_COLUMN_NAME As String = "meta_SectionType"
Private Const SECTION_TYPE_BUTTON_STYLE_NORMAL As String = "sectionTypeButton"
Private Const SECTION_TYPE_BUTTON_STYLE_SELECTED As String = "sectionTypeButtonSelected"
Private Const DRAFT_FIELD_HOSPITAL As String = "Hospital"
Private Const DRAFT_FIELD_HOSPITAL_SHORT As String = "HospitalShort"
Private Const DRAFT_FIELD_VACATION As String = "Vacation"
Private Const DRAFT_FIELD_RANK As String = "Rank"
Private Const DRAFT_FIELD_FIO As String = "Fio"
Private Const DRAFT_FIELD_IPN As String = "Ipn"
Private Const DRAFT_FIELD_POSITION_CODE As String = "PositionCode"
Private Const DRAFT_FIELD_POSITION_NAME As String = "PositionName"
Private Const DRAFT_FIELD_REPORT_TVO As String = "ReportTvo"
Private Const DRAFT_FIELD_REPORT_PERSON As String = "ReportPerson"
Private Const DRAFT_FIELD_INCOMING_NO As String = "IncomingNo"
Private Const DRAFT_FIELD_INCOMING_DATE As String = "IncomingDate"
Private Const DRAFT_FIELD_DOCUMENT_NOTE As String = "DocumentNote"
Private Const DRAFT_FIELD_DOC_NO As String = "DocNo"
Private Const DRAFT_FIELD_DOC_DATE As String = "DocDate"
Private Const DRAFT_FIELD_DURATION_DAYS As String = "DurationDays"
Private Const DRAFT_FIELD_DATE_FROM As String = "DateFrom"
Private Const DRAFT_FIELD_DATE_TO As String = "DateTo"
Private Const DRAFT_FIELD_VH_NO As String = "VhNo"
Private Const DRAFT_FIELD_VH_DATE As String = "VhDate"
Private Const DRAFT_FIELD_VLK_NO As String = "VlkNo"
Private Const DRAFT_FIELD_VLK_DATE As String = "VlkDate"

Private m_Page As obj_IPage
Private m_LookupFeature As obj_EntityLookupFeature
Private m_ExportAliases As Collection
Private m_ExporterClassByAlias As Object
Private m_ExportConfigTableByAlias As Object
Private m_SelectedSectionType As String
Private m_Data As obj_PrsnlEvntBuilderData
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get LookupFeature() As obj_EntityLookupFeature
    Set LookupFeature = m_LookupFeature
End Property

Public Property Get NonTreatmentFieldVisibilityState() As String
    NonTreatmentFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DATE_FROM)
End Property

Public Property Get HospitalFieldVisibilityState() As String
    HospitalFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_HOSPITAL)
End Property

Public Property Get HospitalShortFieldVisibilityState() As String
    HospitalShortFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_HOSPITAL_SHORT)
End Property

Public Property Get VacationFieldVisibilityState() As String
    VacationFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_VACATION)
End Property

Public Property Get RankFieldVisibilityState() As String
    RankFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_RANK)
End Property

Public Property Get FioFieldVisibilityState() As String
    FioFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_FIO)
End Property

Public Property Get IpnFieldVisibilityState() As String
    IpnFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_IPN)
End Property

Public Property Get PositionCodeFieldVisibilityState() As String
    PositionCodeFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_POSITION_CODE)
End Property

Public Property Get PositionNameFieldVisibilityState() As String
    PositionNameFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_POSITION_NAME)
End Property

Public Property Get ReportTvoFieldVisibilityState() As String
    ReportTvoFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_REPORT_TVO)
End Property

Public Property Get ReportPersonFieldVisibilityState() As String
    ReportPersonFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_REPORT_PERSON)
End Property

Public Property Get IncomingNoFieldVisibilityState() As String
    IncomingNoFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_INCOMING_NO)
End Property

Public Property Get IncomingDateFieldVisibilityState() As String
    IncomingDateFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_INCOMING_DATE)
End Property

Public Property Get DocumentNoteFieldVisibilityState() As String
    DocumentNoteFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DOCUMENT_NOTE)
End Property

Public Property Get DocNoFieldVisibilityState() As String
    DocNoFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DOC_NO)
End Property

Public Property Get DocDateFieldVisibilityState() As String
    DocDateFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DOC_DATE)
End Property

Public Property Get DurationDaysFieldVisibilityState() As String
    DurationDaysFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DURATION_DAYS)
End Property

Public Property Get DateFromFieldVisibilityState() As String
    DateFromFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DATE_FROM)
End Property

Public Property Get DateToFieldVisibilityState() As String
    DateToFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_DATE_TO)
End Property

Public Property Get VhNoFieldVisibilityState() As String
    VhNoFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_VH_NO)
End Property

Public Property Get VhDateFieldVisibilityState() As String
    VhDateFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_VH_DATE)
End Property

Public Property Get VlkNoFieldVisibilityState() As String
    VlkNoFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_VLK_NO)
End Property

Public Property Get VlkDateFieldVisibilityState() As String
    VlkDateFieldVisibilityState = private_GetDraftFieldVisibilityState(DRAFT_FIELD_VLK_DATE)
End Property

' //
' // API
' //
Public Function Initialize(ByVal page As Object) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePrsnlEvntBuilderCtrl.Initialize"
#End If
    Dim pageBase As obj_PageBase
    Dim pageInterface As obj_IPage

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PagePrsnlEvntBuilderCtrl initialization failed because page is not specified."
#End If
        Exit Function
    End If
    On Error Resume Next
    Set pageInterface = page
    On Error GoTo 0
    If pageInterface Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PagePrsnlEvntBuilderCtrl initialization failed because page does not implement obj_IPage."
#End If
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = pageInterface
    Set m_Data = New obj_PrsnlEvntBuilderData
    private_ResetExportSettings

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_LookupFeature = New obj_EntityLookupFeature
    If Not m_LookupFeature.Initialize( _
        pageInterface, _
        CANDIDATE_TABLES_RUNTIME_KEY, _
        "prsnlevntbuilder:entitylookup") Then Exit Function

    If Not private_RegisterSectionTypeOptions(False) Then Exit Function
    If Not private_RegisterDummyTables(False) Then Exit Function
    If Not private_EnsureHotkeyRows(False) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePrsnlEvntBuilderCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_LookupFeature Is Nothing Then m_LookupFeature.Dispose
    Set m_LookupFeature = Nothing
    Set m_Page = Nothing
    Set m_ExportAliases = Nothing
    Set m_ExporterClassByAlias = Nothing
    Set m_ExportConfigTableByAlias = Nothing
    Set m_Data = Nothing
    m_SelectedSectionType = VBA.vbNullString
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    UpdateData = m_LookupFeature.UpdateData(configControl)
    If Not UpdateData Then Exit Function
    If Not private_TryUpdateExportSettings(configControl) Then Exit Function
    UpdateData = True
End Function

Public Function PrepareRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.PrepareLookupRuntime(notifyChange) Then Exit Function
    If Not private_RegisterSectionTypeOptions(notifyChange) Then Exit Function
    If Not private_RegisterDummyTables(notifyChange) Then Exit Function
    If Not private_EnsureHotkeyRows(notifyChange) Then Exit Function
    PrepareRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    ClearLookupCandidates = m_LookupFeature.ClearLookupCandidates(renderNow)
End Function

Public Function RuntimeHandleHotkeyAction(ByVal actionId As Variant) As Boolean
    Dim pageBase As obj_PageBase
    Dim selectionObj As Object
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim actionText As String
    Dim cellValue As String

    ' Это page-specific action target для HotkeysControl.
    ' HotkeysControl передает только настроенный Action text; контроллер решает,
    ' что этот action значит на PrsnlEvntBuilder. Другие страницы могут переиспользовать
    ' HotkeysControl со своим actionMethod/dataContext.
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    Set selectionObj = Application.Selection
    If Not TypeOf selectionObj Is Range Then Exit Function

    Set targetCell = selectionObj.Cells(1, 1)
    If targetCell Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function

    actionText = VBA.Trim$(VBA.CStr(actionId))
    cellValue = VBA.CStr(targetCell.Value2)

    If private_IsExportAction(actionText) Then
        If Not private_TryExportDraftByAction(actionText) Then
            RuntimeHandleHotkeyAction = True
            Exit Function
        End If
        RuntimeHandleHotkeyAction = True
        Exit Function
    End If

    ' Page-specific actions branch by stable action ids and read current sheet state.
    Select Case VBA.LCase$(actionText)
        Case VBA.LCase$(HOTKEY_ACTION_1)
            If Not private_TryAcceptCandidateRowFromSelection(targetCell) Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If

        Case VBA.LCase$(HOTKEY_ACTION_2)
            targetCell.Interior.Color = VBA.RGB(126, 36, 121)
            targetCell.Font.Color = VBA.RGB(255, 255, 255)

        Case VBA.LCase$(HOTKEY_SELECT_FORM_ROW)
            If Not private_TrySelectScopedRowFromSelection(targetCell) Then Exit Function

        Case Else
            Exit Function
    End Select

    rt_Messaging.fn_ShowStatusBarSuccess actionText & ": " & targetCell.Address(False, False) & " = '" & cellValue & "'", 3
    RuntimeHandleHotkeyAction = True
End Function

Public Function OnSectionTypeButtonClick(Optional ByVal sectionTypeId As Variant) As Boolean
    Dim newSectionType As String
    Dim previousEnableEvents As Boolean
    Dim perfStart As Double
    Dim perfLast As Double

    perfStart = VBA.Timer
    perfLast = perfStart

    newSectionType = VBA.Trim$(VBA.CStr(sectionTypeId))
    If VBA.Len(newSectionType) = 0 Then Exit Function
    If VBA.StrComp(private_NormalizeText(newSectionType), private_NormalizeText(m_SelectedSectionType), vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        private_LogPerfStep "section-type-click:no-op", perfStart, perfLast, "sectionType='" & private_EscapeForLog(newSectionType) & "'"
#End If
        OnSectionTypeButtonClick = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-type-click:start", perfStart, perfLast, "from='" & private_EscapeForLog(m_SelectedSectionType) & "' to='" & private_EscapeForLog(newSectionType) & "'"
#End If

    m_SelectedSectionType = newSectionType
    If m_Page Is Nothing Then Exit Function
    If Not private_RegisterSectionTypeOptions(False) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-type-click:section-types-registered", perfStart, perfLast, "sectionType='" & private_EscapeForLog(m_SelectedSectionType) & "'"
#End If

    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-type-click:events-disabled", perfStart, perfLast, "previousEnableEvents=" & VBA.LCase$(VBA.CStr(previousEnableEvents))
#End If
    On Error GoTo EH

    OnSectionTypeButtonClick = rt_PageManager.fn_RenderPage(m_Page, "prsnlevntbuilder:section-type-changed")
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-type-click:render-returned", perfStart, perfLast, "ok=" & VBA.LCase$(VBA.CStr(OnSectionTypeButtonClick))
#End If

Cleanup:
    Application.EnableEvents = previousEnableEvents
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-type-click:done", perfStart, perfLast, "ok=" & VBA.LCase$(VBA.CStr(OnSectionTypeButtonClick))
#End If
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-type-click:error", perfStart, perfLast, "err='" & private_EscapeForLog(Err.Description) & "'"
#End If
    Resume Cleanup
End Function

Public Function SearchCandidates( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByRef outCandidateCount As Long, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    SearchCandidates = m_LookupFeature.SearchCandidates(lookupKey, queryText, outCandidateCount, notifyChange)
End Function

Public Function TryGetLookupKeys(ByRef outLookupKeys As Collection) As Boolean
    Set outLookupKeys = Nothing
    If m_LookupFeature Is Nothing Then Exit Function
    TryGetLookupKeys = m_LookupFeature.TryGetLookupKeys(outLookupKeys)
End Function

' //
' // Internal
' //
Private Function private_RegisterDummyTables(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim dummyTables As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set dummyTables = New Collection
    dummyTables.Add private_BuildDummyTable()

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(DUMMY_TABLES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(DUMMY_TABLES_RUNTIME_KEY), dummyTables, notifyChange) Then Exit Function

    private_RegisterDummyTables = True
End Function

Private Function private_RegisterSectionTypeOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim sectionTypes As Collection
    Dim sectionTypeOptions As Collection
    Dim perfStart As Double
    Dim perfLast As Double

    perfStart = VBA.Timer
    perfLast = perfStart

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Set m_Data = New obj_PrsnlEvntBuilderData
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-types:runtime-ready", perfStart, perfLast, "notifyChange=" & VBA.LCase$(VBA.CStr(notifyChange))
#End If

    Set sectionTypes = m_Data.SectionTypeNames
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-types:data-provider-ready", perfStart, perfLast
#End If
    If sectionTypes Is Nothing Then Exit Function
    If sectionTypes.Count = 0 Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-types:options-loaded", perfStart, perfLast, "count=" & VBA.CStr(sectionTypes.Count)
#End If
    If VBA.Len(VBA.Trim$(m_SelectedSectionType)) = 0 Then m_SelectedSectionType = VBA.Trim$(VBA.CStr(sectionTypes.Item(1)))
    If Not private_TryBuildSectionTypeButtonOptions(sectionTypes, sectionTypeOptions) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-types:button-options-built", perfStart, perfLast, "rows=" & VBA.CStr(sectionTypeOptions.Count)
#End If

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(SECTION_TYPES_RUNTIME_KEY)) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-types:runtime-source-removed", perfStart, perfLast
#End If
    If Not runtimeSources.SetItemsSource(VBA.LCase$(SECTION_TYPES_RUNTIME_KEY), sectionTypeOptions, notifyChange) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "section-types:runtime-source-set", perfStart, perfLast
#End If

    private_RegisterSectionTypeOptions = True
End Function

Private Function private_TryBuildSectionTypeButtonOptions( _
    ByVal sectionTypes As Collection, _
    ByRef outRows As Collection _
) As Boolean
    Dim sectionTypeObj As Variant
    Dim sectionTypeText As String
    Dim optionObj As obj_SelectOption
    Dim rowItems As Collection
    Dim rowObj As Object

    Set outRows = Nothing
    If sectionTypes Is Nothing Then Exit Function

    Set outRows = New Collection
    Set rowItems = Nothing
    For Each sectionTypeObj In sectionTypes
        sectionTypeText = VBA.Trim$(VBA.CStr(sectionTypeObj))
        If VBA.Len(sectionTypeText) = 0 Then GoTo ContinueSectionType

        If rowItems Is Nothing Then
            Set rowItems = New Collection
            Set rowObj = VBA.CreateObject("Scripting.Dictionary")
            rowObj.CompareMode = 1
            Set rowObj("Items") = rowItems
            outRows.Add rowObj
        ElseIf rowItems.Count >= 2 Then
            Set rowItems = New Collection
            Set rowObj = VBA.CreateObject("Scripting.Dictionary")
            rowObj.CompareMode = 1
            Set rowObj("Items") = rowItems
            outRows.Add rowObj
        End If

        Set optionObj = New obj_SelectOption
        optionObj.Caption = sectionTypeText
        optionObj.Id = sectionTypeText
        If VBA.StrComp(private_NormalizeText(sectionTypeText), private_NormalizeText(m_SelectedSectionType), VBA.vbTextCompare) = 0 Then
            optionObj.StyleName = SECTION_TYPE_BUTTON_STYLE_SELECTED
        Else
            optionObj.StyleName = SECTION_TYPE_BUTTON_STYLE_NORMAL
        End If
        rowItems.Add optionObj

ContinueSectionType:
    Next sectionTypeObj

    private_TryBuildSectionTypeButtonOptions = True
End Function

#If LOGGING_DEBUG_ENABLED Then
Private Sub private_LogPerfStep( _
    ByVal stepName As String, _
    ByVal startedAt As Double, _
    ByRef lastAt As Double, _
    Optional ByVal details As String = "" _
)
    Dim nowAt As Double
    Dim stepMs As Double
    Dim totalMs As Double
    Dim messageText As String

    nowAt = VBA.Timer
    stepMs = private_ElapsedMs(lastAt, nowAt)
    totalMs = private_ElapsedMs(startedAt, nowAt)
    lastAt = nowAt

    messageText = "perf:prsnlevntbuilder:" & stepName & _
        " stepMs=" & VBA.Format$(stepMs, "0.0") & _
        " totalMs=" & VBA.Format$(totalMs, "0.0")
    If VBA.Len(VBA.Trim$(details)) > 0 Then messageText = messageText & " " & details
    ex_Core.fn_Diagnostic_LogInfo messageText
End Sub

Private Function private_ElapsedMs(ByVal startedAt As Double, ByVal endedAt As Double) As Double
    If endedAt < startedAt Then endedAt = endedAt + 86400#
    private_ElapsedMs = (endedAt - startedAt) * 1000#
End Function

Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace$(VBA.Trim$(VBA.CStr(valueText)), "'", "''")
End Function
#End If

Private Function private_TryAcceptCandidateRowFromSelection(ByVal targetCell As Range) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim candidateRowsArea As Range
    Dim candidateRowRange As Range
    Dim draftValuesRange As Range
    Dim rowOffset As Long
    Dim sourceRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim previousEnableEvents As Boolean

    If targetCell Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function

    ' TableList already registers rendered data rows as controlPart=rows.
    ' This keeps the hotkey independent from hard-coded row/column numbers.
    If Not private_TryResolveCandidateRowsArea(ws, targetCell, candidateRowsArea) Then
        rt_Messaging.fn_ShowStatusBarWarning "Select a candidate row cell first.", 3
        Exit Function
    End If

    ' EventDraftValues is the named layout container for the form value row.
    ' The controller only needs this container range, not individual input names.
    If Not pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then
        rt_Messaging.fn_ShowStatusBarWarning "Event draft values container is not rendered.", 3
        Exit Function
    End If

    ' Convert the active cell position into the concrete rendered candidate row.
    ' The candidate rows area can move after rerender, so the relative row offset is used.
    rowOffset = targetCell.Row - candidateRowsArea.Row + 1
    If rowOffset <= 0 Or rowOffset > candidateRowsArea.Rows.Count Then Exit Function

    sourceRow = candidateRowsArea.Row + rowOffset - 1
    Set candidateRowRange = ws.Range( _
        ws.Cells(sourceRow, candidateRowsArea.Column), _
        ws.Cells(sourceRow, candidateRowsArea.Column + candidateRowsArea.Columns.Count - 1))

    ' LookupCandidates visually aligns result columns under the form columns.
    ' Therefore the safest generic mapping is the intersection of absolute Excel columns.
    firstCol = private_MaxLong(candidateRowRange.Column, draftValuesRange.Column)
    lastCol = private_MinLong( _
        candidateRowRange.Column + candidateRowRange.Columns.Count - 1, _
        draftValuesRange.Column + draftValuesRange.Columns.Count - 1)
    If lastCol < firstCol Then
        rt_Messaging.fn_ShowStatusBarWarning "Selected candidate row does not overlap the event form.", 3
        Exit Function
    End If

    ' Values are written directly into sheet cells. Disable events so this accept action
    ' does not recursively trigger input onChange/search/rerender for each copied cell.
    previousEnableEvents = Application.EnableEvents
    On Error GoTo RestoreEventsAndFail
    Application.EnableEvents = False
    For colIndex = firstCol To lastCol
        ws.Cells(draftValuesRange.Row, colIndex).Value2 = ws.Cells(sourceRow, colIndex).Value2
    Next colIndex
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0

    candidateRowRange.Select
    private_TryAcceptCandidateRowFromSelection = True
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
End Function

Private Function private_TrySelectScopedRowFromSelection(ByVal targetCell As Range) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim formRange As Range
    Dim scopedRow As Range
    Dim listObj As ListObject

    If targetCell Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function

    If pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_FORM_CONTAINER_NAME, formRange) Then
        If Not formRange Is Nothing Then
            Set scopedRow = Nothing
            On Error Resume Next
            If Not Application.Intersect(targetCell, formRange) Is Nothing Then
                Set scopedRow = Application.Intersect(targetCell.EntireRow, formRange)
            End If
            On Error GoTo 0
            If Not scopedRow Is Nothing Then
                scopedRow.Select
                private_TrySelectScopedRowFromSelection = True
                Exit Function
            End If
        End If
    End If

    For Each listObj In ws.ListObjects
        If listObj Is Nothing Then GoTo ContinueListObject
        If listObj.Range Is Nothing Then GoTo ContinueListObject
        Set scopedRow = Nothing
        On Error Resume Next
        If Not Application.Intersect(targetCell, listObj.Range) Is Nothing Then
            Set scopedRow = Application.Intersect(targetCell.EntireRow, listObj.Range)
        End If
        On Error GoTo 0
        If Not scopedRow Is Nothing Then
            scopedRow.Select
            private_TrySelectScopedRowFromSelection = True
            Exit Function
        End If

ContinueListObject:
    Next listObj

    targetCell.EntireRow.Select
    private_TrySelectScopedRowFromSelection = True
End Function

Private Function private_TryExportDraftByAction(ByVal actionId As String) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim exporter As obj_IDataExporter
    Dim exportAlias As String
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable

    If Not private_TryResolveExportAliasFromAction(actionId, exportAlias) Then Exit Function
        If Not private_TryGetExportSettings(exportAlias, exporterClassName, exportConfigTable) Then Exit Function
    If Not private_TryBuildDraftFormSourceTable(sourceTable) Then Exit Function

    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function

        If Not exporter.Export(sourceTable) Then Exit Function

    rt_Messaging.fn_ShowStatusBarSuccess EXPORT_ACTION_PREFIX & exportAlias & ": done", 3
    private_TryExportDraftByAction = True
End Function

Private Function private_TryUpdateExportSettings(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParserBase As obj_CfgParserBase
    Dim configEntries As Collection
    Dim cfgMap As Object

    private_ResetExportSettings
    If configControl Is Nothing Then
        private_TryUpdateExportSettings = True
        Exit Function
    End If

    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function

    Set cfgParserBase = New obj_CfgParserBase
    If Not cfgParserBase.Initialize(configTable) Then Exit Function
    If Not cfgParserBase.TryGetConfigEntries(configEntries) Then Exit Function
    If Not cfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then Exit Function

    If Not private_TryLoadExportSettings(configEntries, cfgParserBase, cfgMap) Then Exit Function
    private_TryUpdateExportSettings = True
End Function

Private Sub private_ResetExportSettings()
    Set m_ExportAliases = New Collection
    Set m_ExporterClassByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ExportConfigTableByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Function private_TryLoadExportSettings( _
    ByVal configEntries As Collection, _
    ByVal cfgParserBase As obj_CfgParserBase, _
    ByVal cfgMap As Object _
) As Boolean
    Dim entryObj As Variant
    Dim configEntry As obj_ConfigEntry
    Dim keyText As String
    Dim exportAlias As String
    Dim keySuffix As String
    Dim exporterClassName As String
    Dim targetWorkbookPath As String
    Dim targetSheetName As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim exportConfigTable As obj_ConfigTable
    Dim aliasObj As Variant

    If configEntries Is Nothing Then
        private_TryLoadExportSettings = True
        Exit Function
    End If
    If cfgParserBase Is Nothing Then Exit Function
    If cfgMap Is Nothing Then Exit Function

    For Each entryObj In configEntries
        If Not VBA.IsObject(entryObj) Then GoTo ContinueEntry
        Set configEntry = Nothing
        On Error Resume Next
        Set configEntry = entryObj
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueEntry

        keyText = VBA.Trim$(configEntry.Key)
        If Not private_TryParseExportConfigKey(keyText, exportAlias, keySuffix) Then GoTo ContinueEntry
        If Not private_ExportAliasExists(exportAlias) Then m_ExportAliases.Add exportAlias

ContinueEntry:
    Next entryObj

    For Each aliasObj In m_ExportAliases
        exportAlias = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(exportAlias) = 0 Then GoTo ContinueAlias

        exporterClassName = cfgParserBase.GetOptionalConfigValue( _
            cfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_CLASS_SUFFIX, _
            DEFAULT_EXPORTER_CLASS)
        exporterClassName = VBA.Trim$(exporterClassName)
        If VBA.Len(exporterClassName) = 0 Then exporterClassName = DEFAULT_EXPORTER_CLASS

        targetWorkbookPath = cfgParserBase.GetOptionalConfigValue( _
            cfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_FILE_PATH_SUFFIX, _
            VBA.vbNullString)
        targetSheetName = cfgParserBase.GetOptionalConfigValue( _
            cfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_SHEET_NAME_SUFFIX, _
            VBA.vbNullString)
        rangeStartMarker = cfgParserBase.GetOptionalConfigValue( _
            cfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_START_MARKER_SUFFIX, _
            VBA.vbNullString)
        rangeEndMarker = cfgParserBase.GetOptionalConfigValue( _
            cfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_END_MARKER_SUFFIX, _
            VBA.vbNullString)

        Set exportConfigTable = private_BuildExportConfigTable( _
            exportAlias, _
            exporterClassName, _
            targetWorkbookPath, _
            targetSheetName, _
            rangeStartMarker, _
            rangeEndMarker)
        If exportConfigTable Is Nothing Then Exit Function

        m_ExporterClassByAlias(exportAlias) = exporterClassName
        Set m_ExportConfigTableByAlias(exportAlias) = exportConfigTable

ContinueAlias:
    Next aliasObj

    private_TryLoadExportSettings = True
End Function

Private Function private_TryParseExportConfigKey( _
    ByVal keyText As String, _
    ByRef outExportAlias As String, _
    ByRef outKeySuffix As String _
) As Boolean
    Dim keyLower As String
    Dim suffixPos As Long

    outExportAlias = VBA.vbNullString
    outKeySuffix = VBA.vbNullString

    keyText = VBA.Trim$(keyText)
    keyLower = VBA.LCase$(keyText)
    If VBA.Left$(keyLower, VBA.Len(VBA.LCase$(EXPORT_CONFIG_PREFIX))) <> VBA.LCase$(EXPORT_CONFIG_PREFIX) Then Exit Function

    suffixPos = VBA.InStr(VBA.Len(EXPORT_CONFIG_PREFIX) + 1, keyText, ".", VBA.vbTextCompare)
    If suffixPos <= VBA.Len(EXPORT_CONFIG_PREFIX) + 1 Then Exit Function

    outKeySuffix = VBA.Mid$(keyText, suffixPos)
    If VBA.StrComp(outKeySuffix, EXPORT_FILE_PATH_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_CLASS_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_SHEET_NAME_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_RANGE_START_MARKER_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_RANGE_END_MARKER_SUFFIX, VBA.vbTextCompare) <> 0 Then Exit Function

    outExportAlias = VBA.Trim$(VBA.Mid$(keyText, VBA.Len(EXPORT_CONFIG_PREFIX) + 1, suffixPos - VBA.Len(EXPORT_CONFIG_PREFIX) - 1))
    If VBA.Len(outExportAlias) = 0 Then Exit Function

    private_TryParseExportConfigKey = True
End Function

Private Function private_ExportAliasExists(ByVal exportAlias As String) As Boolean
    Dim aliasObj As Variant

    If m_ExportAliases Is Nothing Then Exit Function
    exportAlias = VBA.Trim$(exportAlias)
    If VBA.Len(exportAlias) = 0 Then Exit Function

    For Each aliasObj In m_ExportAliases
        If VBA.StrComp(VBA.Trim$(VBA.CStr(aliasObj)), exportAlias, VBA.vbTextCompare) = 0 Then
            private_ExportAliasExists = True
            Exit Function
        End If
    Next aliasObj
End Function

Private Function private_IsExportAction(ByVal actionId As String) As Boolean
    actionId = VBA.Trim$(actionId)
    If VBA.Len(actionId) <= VBA.Len(EXPORT_ACTION_PREFIX) Then Exit Function
    private_IsExportAction = (VBA.Left$(VBA.LCase$(actionId), VBA.Len(VBA.LCase$(EXPORT_ACTION_PREFIX))) = VBA.LCase$(EXPORT_ACTION_PREFIX))
End Function

Private Function private_TryResolveExportAliasFromAction( _
    ByVal actionId As String, _
    ByRef outExportAlias As String _
) As Boolean
    Dim exportAlias As String

    outExportAlias = VBA.vbNullString
    If Not private_IsExportAction(actionId) Then Exit Function

    exportAlias = VBA.Trim$(VBA.Mid$(actionId, VBA.Len(EXPORT_ACTION_PREFIX) + 1))
    If VBA.Len(exportAlias) = 0 Then Exit Function
    If Not private_ExportAliasExists(exportAlias) Then
        VBA.MsgBox "PrototypeNew: export action is not configured: " & actionId, VBA.vbExclamation, "PrototypeNew / Data export"
        Exit Function
    End If

    outExportAlias = exportAlias
    private_TryResolveExportAliasFromAction = True
End Function

Private Function private_TryGetExportSettings( _
    ByVal exportAlias As String, _
    ByRef outExporterClassName As String, _
    ByRef outExportConfigTable As obj_ConfigTable _
) As Boolean
    outExporterClassName = VBA.vbNullString
    Set outExportConfigTable = Nothing

    exportAlias = VBA.Trim$(exportAlias)
    If VBA.Len(exportAlias) = 0 Then Exit Function
    If m_ExporterClassByAlias Is Nothing Then Exit Function
    If m_ExportConfigTableByAlias Is Nothing Then Exit Function
    If Not m_ExporterClassByAlias.Exists(exportAlias) Then Exit Function

    outExporterClassName = VBA.Trim$(VBA.CStr(m_ExporterClassByAlias(exportAlias)))
    If VBA.Len(outExporterClassName) = 0 Then outExporterClassName = DEFAULT_EXPORTER_CLASS
    If m_ExportConfigTableByAlias.Exists(exportAlias) Then Set outExportConfigTable = m_ExportConfigTableByAlias(exportAlias)
    If outExportConfigTable Is Nothing Then Exit Function

    private_TryGetExportSettings = True
End Function

Private Function private_TryCreateDataExporter( _
    ByVal exporterClassName As String, _
    ByVal exportConfigTable As obj_ConfigTable, _
    ByRef outExporter As obj_IDataExporter _
) As Boolean
    Dim exporterToDailyScope As obj_ExporterToDailyScope
    Dim exporterToMovement As obj_ExporterToMovement

    Set outExporter = Nothing
    exporterClassName = VBA.Trim$(exporterClassName)
    If VBA.Len(exporterClassName) = 0 Then exporterClassName = DEFAULT_EXPORTER_CLASS

    Select Case VBA.LCase$(exporterClassName)
        Case VBA.LCase$("obj_ExporterToDailyScope")
            Set exporterToDailyScope = New obj_ExporterToDailyScope
            If Not exporterToDailyScope.Initialize(exportConfigTable) Then Exit Function
            Set outExporter = exporterToDailyScope

        Case VBA.LCase$("obj_ExporterToMovement")
            Set exporterToMovement = New obj_ExporterToMovement
            If Not exporterToMovement.Initialize(exportConfigTable) Then Exit Function
            Set outExporter = exporterToMovement

        Case Else
            VBA.MsgBox "PrototypeNew: unsupported data exporter class: " & exporterClassName, VBA.vbExclamation, "PrototypeNew / Data export"
            Exit Function
    End Select

    private_TryCreateDataExporter = Not outExporter Is Nothing
End Function

Private Function private_BuildExportConfigTable( _
    ByVal exportAlias As String, _
    ByVal exporterClassName As String, _
    ByVal targetWorkbookPath As String, _
    ByVal targetSheetName As String, _
    ByVal rangeStartMarker As String, _
    ByVal rangeEndMarker As String _
) As obj_ConfigTable
    Dim configTable As obj_ConfigTable

    Set configTable = New obj_ConfigTable
    If Not configTable.Initialize() Then Exit Function

    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_CLASS_SUFFIX, exporterClassName) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_FILE_PATH_SUFFIX, targetWorkbookPath) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_SHEET_NAME_SUFFIX, targetSheetName) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_START_MARKER_SUFFIX, rangeStartMarker) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_END_MARKER_SUFFIX, rangeEndMarker) Then Exit Function

    Set private_BuildExportConfigTable = configTable
End Function

Private Function private_TryBuildDraftFormSourceTable(ByRef outTable As obj_TableDynamic) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim draftValuesRange As Range
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim colOffset As Long
    Dim sheetCol As Long
    Dim headerText As String
    Dim sectionTypeText As String
    Dim manualOrderNoText As String

    Set outTable = Nothing
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    If Not pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then
        rt_Messaging.fn_ShowStatusBarWarning "Event draft values container is not rendered.", 3
        Exit Function
    End If
    If draftValuesRange Is Nothing Then Exit Function
    If draftValuesRange.Row <= 1 Then
        rt_Messaging.fn_ShowStatusBarWarning "Event draft values container has no header row above it.", 3
        Exit Function
    End If

    Set sourceTable = New obj_TableDynamic
    sourceTable.SectionTitle = "PrsnlEvntBuilder draft form"
    Set sourceRow = New obj_Row

    For colOffset = 1 To draftValuesRange.Columns.Count
        sheetCol = draftValuesRange.Column + colOffset - 1
        headerText = private_ReadHeaderText(ws.Cells(draftValuesRange.Row - 1, sheetCol))
        If VBA.Len(headerText) = 0 Then headerText = "Column " & VBA.CStr(colOffset)
        If Not private_AddSourceColumn(sourceTable, headerText) Then Exit Function
        sourceRow.PushCellRaw ws.Cells(draftValuesRange.Row, sheetCol).Value2
    Next colOffset

    If Not private_TryGetSelectedSectionType(sectionTypeText) Then Exit Function
    manualOrderNoText = private_TryReadManualOrderNoValue(pageBase, ws)

    If Not private_AddSourceColumn(sourceTable, EXPORT_META_MANUAL_ORDER_NO_COLUMN_NAME) Then Exit Function
    sourceRow.PushCellRaw manualOrderNoText

    If Not private_AddSourceColumn(sourceTable, EXPORT_META_SECTION_TYPE_COLUMN_NAME) Then Exit Function
    sourceRow.PushCellRaw sectionTypeText

    If Not sourceTable.PushRow(sourceRow) Then Exit Function
    Set outTable = sourceTable
    private_TryBuildDraftFormSourceTable = True
End Function

Private Function private_TryReadManualOrderNoValue( _
    ByVal pageBase As obj_PageBase, _
    ByVal ws As Worksheet _
) As String
    Dim labelScope As Range
    Dim columnScope As Range

    If pageBase Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function

    Set labelScope = Nothing
    Set columnScope = Nothing
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope(ws, "label", EVENT_DRAFT_ORDER_NO_LABEL_CONTROL_NAME, "cell", labelScope, columnScope) Then Exit Function
    If labelScope Is Nothing Then Exit Function

    private_TryReadManualOrderNoValue = VBA.Trim$(VBA.CStr(labelScope.Cells(1, 1).Value2))
End Function

Private Function private_TryGetSelectedSectionType(ByRef outSectionType As String) As Boolean
    outSectionType = VBA.Trim$(m_SelectedSectionType)
    If VBA.Len(outSectionType) = 0 Then
        rt_Messaging.fn_ShowStatusBarWarning "Event section type is not selected.", 3
        Exit Function
    End If

    private_TryGetSelectedSectionType = True
End Function

Private Function private_GetDraftFieldVisibilityState(ByVal fieldKey As String) As String
    If private_IsDraftFieldVisible(fieldKey) Then
        private_GetDraftFieldVisibilityState = "visible"
    Else
        private_GetDraftFieldVisibilityState = "collapsed"
    End If
End Function

Private Function private_IsDraftFieldVisible(ByVal fieldKey As String) As Boolean
    Dim sectionTypeText As String
    Dim normalizedSectionType As String
    Dim fieldVisibilityFlags As Object

    fieldKey = VBA.Trim$(fieldKey)
    If VBA.Len(fieldKey) = 0 Then Exit Function

    If Not private_TryGetSelectedSectionType(sectionTypeText) Then
        private_IsDraftFieldVisible = True
        Exit Function
    End If

    normalizedSectionType = private_NormalizeText(sectionTypeText)
    Set fieldVisibilityFlags = private_BuildDraftFieldVisibilityFlags(normalizedSectionType)
    If fieldVisibilityFlags Is Nothing Then
        private_IsDraftFieldVisible = True
        Exit Function
    End If
    If Not fieldVisibilityFlags.Exists(fieldKey) Then
        private_IsDraftFieldVisible = True
        Exit Function
    End If

    private_IsDraftFieldVisible = VBA.CBool(fieldVisibilityFlags(fieldKey))
End Function

Private Function private_BuildDraftFieldVisibilityFlags(ByVal normalizedSectionType As String) As Object
    Dim flags As Object

    Set flags = private_CreateDraftFieldVisibilityFlags(True)
    If m_Data Is Nothing Then
        Set private_BuildDraftFieldVisibilityFlags = flags
        Exit Function
    End If

    normalizedSectionType = private_NormalizeText(normalizedSectionType)
    Select Case normalizedSectionType
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatment)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_HOSPITAL, _
                DRAFT_FIELD_HOSPITAL_SHORT, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE

        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentMedicalCompany)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE

        Case private_NormalizeText(m_Data.SectionTypeCloseFromAmbulatoryVlk)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE, _
                DRAFT_FIELD_VLK_NO, _
                DRAFT_FIELD_VLK_DATE

        Case private_NormalizeText(m_Data.SectionTypeToTreatment)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_HOSPITAL, _
                DRAFT_FIELD_HOSPITAL_SHORT, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE

        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromFamilyVacation)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_VH_NO, _
                DRAFT_FIELD_VH_DATE

        Case private_NormalizeText(m_Data.SectionTypeToAnnualVacationPart), _
             private_NormalizeText(m_Data.SectionTypeToFamilyVacation)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_VACATION, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DURATION_DAYS, _
                DRAFT_FIELD_DATE_FROM, _
                DRAFT_FIELD_DATE_TO, _
                DRAFT_FIELD_VH_NO, _
                DRAFT_FIELD_VH_DATE

        Case private_NormalizeText(m_Data.SectionTypeToTreatmentVacation)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_VACATION, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DURATION_DAYS, _
                DRAFT_FIELD_DATE_FROM, _
                DRAFT_FIELD_DATE_TO, _
                DRAFT_FIELD_VH_NO, _
                DRAFT_FIELD_VH_DATE, _
                DRAFT_FIELD_VLK_NO, _
                DRAFT_FIELD_VLK_DATE

        Case private_NormalizeText(m_Data.SectionTypeToTreatmentMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeToAmbulatoryVlk)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOC_NO

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentToTreatmentVacation)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_HOSPITAL, _
                DRAFT_FIELD_HOSPITAL_SHORT, _
                DRAFT_FIELD_VACATION, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE, _
                DRAFT_FIELD_DURATION_DAYS, _
                DRAFT_FIELD_DATE_FROM, _
                DRAFT_FIELD_DATE_TO, _
                DRAFT_FIELD_VH_NO, _
                DRAFT_FIELD_VH_DATE, _
                DRAFT_FIELD_VLK_NO, _
                DRAFT_FIELD_VLK_DATE

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatmentVacation)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_VACATION, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DURATION_DAYS, _
                DRAFT_FIELD_DATE_FROM, _
                DRAFT_FIELD_DATE_TO, _
                DRAFT_FIELD_VH_NO, _
                DRAFT_FIELD_VH_DATE, _
                DRAFT_FIELD_VLK_NO, _
                DRAFT_FIELD_VLK_DATE

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatment)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_HOSPITAL, _
                DRAFT_FIELD_HOSPITAL_SHORT, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE, _
                DRAFT_FIELD_DATE_FROM

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToVlk)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE, _
                DRAFT_FIELD_DATE_FROM

        Case private_NormalizeText(m_Data.SectionTypeTransferVlkToTreatmentVacation)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_VACATION, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DURATION_DAYS, _
                DRAFT_FIELD_DATE_FROM, _
                DRAFT_FIELD_DATE_TO, _
                DRAFT_FIELD_VH_NO, _
                DRAFT_FIELD_VH_DATE, _
                DRAFT_FIELD_VLK_NO, _
                DRAFT_FIELD_VLK_DATE

        Case private_NormalizeText(m_Data.SectionTypeTransferVlkToTreatment)
            private_SetAllDraftFieldVisibilityFlags flags, False
            private_SetDraftFieldsVisible flags, _
                DRAFT_FIELD_HOSPITAL, _
                DRAFT_FIELD_HOSPITAL_SHORT, _
                DRAFT_FIELD_RANK, _
                DRAFT_FIELD_FIO, _
                DRAFT_FIELD_IPN, _
                DRAFT_FIELD_POSITION_CODE, _
                DRAFT_FIELD_POSITION_NAME, _
                DRAFT_FIELD_REPORT_TVO, _
                DRAFT_FIELD_REPORT_PERSON, _
                DRAFT_FIELD_INCOMING_NO, _
                DRAFT_FIELD_INCOMING_DATE, _
                DRAFT_FIELD_DOCUMENT_NOTE, _
                DRAFT_FIELD_DOC_NO, _
                DRAFT_FIELD_DOC_DATE, _
                DRAFT_FIELD_VLK_NO, _
                DRAFT_FIELD_VLK_DATE

        Case Else
            private_SetAllDraftFieldVisibilityFlags flags, True
    End Select

    Set private_BuildDraftFieldVisibilityFlags = flags
End Function

Private Function private_CreateDraftFieldVisibilityFlags(ByVal defaultValue As Boolean) As Object
    Dim flags As Object

    Set flags = VBA.CreateObject("Scripting.Dictionary")
    flags.CompareMode = 1

    flags(DRAFT_FIELD_HOSPITAL) = defaultValue
    flags(DRAFT_FIELD_HOSPITAL_SHORT) = defaultValue
    flags(DRAFT_FIELD_VACATION) = defaultValue
    flags(DRAFT_FIELD_RANK) = defaultValue
    flags(DRAFT_FIELD_FIO) = defaultValue
    flags(DRAFT_FIELD_IPN) = defaultValue
    flags(DRAFT_FIELD_POSITION_CODE) = defaultValue
    flags(DRAFT_FIELD_POSITION_NAME) = defaultValue
    flags(DRAFT_FIELD_REPORT_TVO) = defaultValue
    flags(DRAFT_FIELD_REPORT_PERSON) = defaultValue
    flags(DRAFT_FIELD_INCOMING_NO) = defaultValue
    flags(DRAFT_FIELD_INCOMING_DATE) = defaultValue
    flags(DRAFT_FIELD_DOCUMENT_NOTE) = defaultValue
    flags(DRAFT_FIELD_DOC_NO) = defaultValue
    flags(DRAFT_FIELD_DOC_DATE) = defaultValue
    flags(DRAFT_FIELD_DURATION_DAYS) = defaultValue
    flags(DRAFT_FIELD_DATE_FROM) = defaultValue
    flags(DRAFT_FIELD_DATE_TO) = defaultValue
    flags(DRAFT_FIELD_VH_NO) = defaultValue
    flags(DRAFT_FIELD_VH_DATE) = defaultValue
    flags(DRAFT_FIELD_VLK_NO) = defaultValue
    flags(DRAFT_FIELD_VLK_DATE) = defaultValue

    Set private_CreateDraftFieldVisibilityFlags = flags
End Function

Private Sub private_SetAllDraftFieldVisibilityFlags(ByVal flags As Object, ByVal isVisible As Boolean)
    Dim fieldKey As Variant

    If flags Is Nothing Then Exit Sub
    For Each fieldKey In flags.Keys
        flags(fieldKey) = isVisible
    Next fieldKey
End Sub

Private Sub private_SetDraftFieldsVisible(ByVal flags As Object, ParamArray fieldKeys() As Variant)
    Dim fieldKey As Variant
    Dim fieldKeyText As String

    If flags Is Nothing Then Exit Sub
    For Each fieldKey In fieldKeys
        fieldKeyText = VBA.Trim$(VBA.CStr(fieldKey))
        If VBA.Len(fieldKeyText) > 0 Then flags(fieldKeyText) = True
    Next fieldKey
End Sub

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ".", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "/", " ")
    valueText = VBA.Replace(valueText, "(", " ")
    valueText = VBA.Replace(valueText, ")", " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeText = VBA.Trim$(valueText)
End Function

Private Function private_ReadHeaderText(ByVal headerCell As Range) As String
    Dim valueText As String

    If headerCell Is Nothing Then Exit Function

    On Error Resume Next
    If headerCell.MergeCells Then
        valueText = VBA.CStr(headerCell.MergeArea.Cells(1, 1).Value2)
    Else
        valueText = VBA.CStr(headerCell.Value2)
    End If
    On Error GoTo 0

    private_ReadHeaderText = VBA.Trim$(valueText)
End Function

Private Sub private_TryRefreshOrderNoLabel()
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim draftValuesRange As Range
    Dim orderNoText As String
    Dim labelScope As Range
    Dim columnScope As Range
    Dim labelText As String

    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Sub

    If Not pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then Exit Sub
    If draftValuesRange Is Nothing Then Exit Sub

    If Not private_TryReadDraftValueByHeader(ws, draftValuesRange, EVENT_DRAFT_INCOMING_NO_HEADER_NAME, orderNoText) Then
        orderNoText = VBA.vbNullString
    End If

    labelText = EVENT_DRAFT_ORDER_LABEL_PREFIX
    If VBA.Len(VBA.Trim$(orderNoText)) > 0 Then
        labelText = labelText & VBA.Trim$(orderNoText)
    Else
        labelText = labelText & "-"
    End If

    Set labelScope = Nothing
    Set columnScope = Nothing
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope(ws, "label", EVENT_DRAFT_ORDER_NO_LABEL_CONTROL_NAME, "cell", labelScope, columnScope) Then Exit Sub
    If labelScope Is Nothing Then Exit Sub

    labelScope.Value2 = labelText
End Sub

Private Function private_TryReadDraftValueByHeader( _
    ByVal ws As Worksheet, _
    ByVal draftValuesRange As Range, _
    ByVal headerName As String, _
    ByRef outValueText As String _
) As Boolean
    Dim colOffset As Long
    Dim sheetCol As Long
    Dim headerText As String

    outValueText = VBA.vbNullString
    If ws Is Nothing Then Exit Function
    If draftValuesRange Is Nothing Then Exit Function
    If draftValuesRange.Row <= 1 Then Exit Function

    headerName = VBA.Trim$(headerName)
    If VBA.Len(headerName) = 0 Then Exit Function

    For colOffset = 1 To draftValuesRange.Columns.Count
        sheetCol = draftValuesRange.Column + colOffset - 1
        headerText = private_ReadHeaderText(ws.Cells(draftValuesRange.Row - 1, sheetCol))
        If VBA.StrComp(VBA.Trim$(headerText), headerName, VBA.vbTextCompare) = 0 Then
            outValueText = VBA.Trim$(VBA.CStr(ws.Cells(draftValuesRange.Row, sheetCol).Value2))
            private_TryReadDraftValueByHeader = True
            Exit Function
        End If
    Next colOffset
End Function

Private Function private_TryResolveCandidateRowsArea( _
    ByVal ws As Worksheet, _
    ByVal targetCell As Range, _
    ByRef outRowsArea As Range _
) As Boolean
    Dim rowsScope As Range
    Dim columnScope As Range
    Dim area As Range

    Set outRowsArea = Nothing
    If ws Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function

    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, _
        "tablelist", _
        LOOKUP_CANDIDATES_CONTROL_NAME, _
        "rows", _
        rowsScope, _
        columnScope) Then Exit Function
    If rowsScope Is Nothing Then Exit Function

    For Each area In rowsScope.Areas
        If Not Application.Intersect(targetCell, area) Is Nothing Then
            Set outRowsArea = area
            private_TryResolveCandidateRowsArea = True
            Exit Function
        End If
    Next area
End Function

Private Function private_MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        private_MaxLong = leftValue
    Else
        private_MaxLong = rightValue
    End If
End Function

Private Function private_MinLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue <= rightValue Then
        private_MinLong = leftValue
    Else
        private_MinLong = rightValue
    End If
End Function

Private Function private_EnsureHotkeyRows(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim hotkeyRows As Collection
    Dim existingRows As Collection
    Dim hasChanges As Boolean

    ' Сеем default-строки хоткеев только когда page runtime source отсутствует/пустой.
    ' После Apply HotkeysControl пишет отредактированные строки обратно в тот же
    ' RuntimeItems key, поэтому rerender/PrepareRuntime не должны их перетирать.
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    If runtimeSources.TryGetItemsSourceByKey(VBA.LCase$(HOTKEYS_RUNTIME_KEY), existingRows, True) Then
        If Not existingRows Is Nothing Then
            If existingRows.Count > 0 Then
                Set hotkeyRows = existingRows
                If Not private_RemoveStaleExportHotkeyRows(hotkeyRows, hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_ACTION_1, "CTRL+ENTER", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_ACTION_2, "CTRL+SHIFT+R", hasChanges) Then Exit Function
                If Not private_EnsureExportHotkeyRows(hotkeyRows, hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE", hasChanges) Then Exit Function
                If hasChanges Then
                    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY)) Then Exit Function
                    If Not runtimeSources.SetItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, notifyChange) Then Exit Function
                End If
                private_EnsureHotkeyRows = True
                Exit Function
            End If
        End If
    Else
        Exit Function
    End If

    Set hotkeyRows = New Collection
    ' Defaults — это только стартовые данные страницы. Активными они становятся
    ' после render HotkeysControl и RuntimeRegisterBoundRows, где регистрируются
    ' routes для этой страницы.
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_ACTION_1, "CTRL+ENTER") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_ACTION_2, "CTRL+SHIFT+R") Then Exit Function
    If Not private_AddExportHotkeyRows(hotkeyRows) Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE") Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, notifyChange) Then Exit Function

    private_EnsureHotkeyRows = True
End Function

Private Function private_EnsureHotkeyRow( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String, _
    ByVal defaultHotkey As String, _
    ByRef ioHasChanges As Boolean _
) As Boolean
    If hotkeyRows Is Nothing Then Exit Function
    If Not private_HotkeyRowsContainAction(hotkeyRows, actionId) Then
        If Not private_AddHotkeyRow(hotkeyRows, actionId, defaultHotkey) Then Exit Function
        ioHasChanges = True
    End If

    private_EnsureHotkeyRow = True
End Function

Private Function private_EnsureExportHotkeyRows( _
    ByVal hotkeyRows As Collection, _
    ByRef ioHasChanges As Boolean _
) As Boolean
    Dim exportIndex As Long
    Dim aliasObj As Variant
    Dim exportAlias As String
    Dim actionId As String
    Dim defaultHotkey As String

    If hotkeyRows Is Nothing Then Exit Function
    If m_ExportAliases Is Nothing Then
        private_EnsureExportHotkeyRows = True
        Exit Function
    End If

    exportIndex = 0
    For Each aliasObj In m_ExportAliases
        exportAlias = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(exportAlias) = 0 Then GoTo ContinueAlias

        exportIndex = exportIndex + 1
        If exportIndex > MAX_EXPORT_HOTKEYS Then GoTo ContinueAlias

        actionId = private_BuildExportActionId(exportAlias)
        defaultHotkey = private_BuildExportDefaultHotkey(exportIndex)
        If Not private_EnsureHotkeyRow(hotkeyRows, actionId, defaultHotkey, ioHasChanges) Then Exit Function

ContinueAlias:
    Next aliasObj

    private_EnsureExportHotkeyRows = True
End Function

Private Function private_AddExportHotkeyRows(ByVal hotkeyRows As Collection) As Boolean
    Dim exportIndex As Long
    Dim aliasObj As Variant
    Dim exportAlias As String

    If hotkeyRows Is Nothing Then Exit Function
    If m_ExportAliases Is Nothing Then
        private_AddExportHotkeyRows = True
        Exit Function
    End If

    exportIndex = 0
    For Each aliasObj In m_ExportAliases
        exportAlias = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(exportAlias) = 0 Then GoTo ContinueAlias

        exportIndex = exportIndex + 1
        If exportIndex > MAX_EXPORT_HOTKEYS Then GoTo ContinueAlias

        If Not private_AddHotkeyRow( _
            hotkeyRows, _
            private_BuildExportActionId(exportAlias), _
            private_BuildExportDefaultHotkey(exportIndex)) Then Exit Function

ContinueAlias:
    Next aliasObj

    private_AddExportHotkeyRows = True
End Function

Private Function private_RemoveStaleExportHotkeyRows( _
    ByVal hotkeyRows As Collection, _
    ByRef ioHasChanges As Boolean _
) As Boolean
    Dim idx As Long
    Dim rowObj As Object
    Dim configEntry As obj_ConfigEntry
    Dim actionId As String

    If hotkeyRows Is Nothing Then Exit Function

    For idx = hotkeyRows.Count To 1 Step -1
        Set rowObj = Nothing
        Set configEntry = Nothing
        On Error Resume Next
        Set rowObj = hotkeyRows.Item(idx)
        Set configEntry = rowObj
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueRow

        actionId = VBA.Trim$(configEntry.Key)
        If VBA.Left$(VBA.LCase$(actionId), VBA.Len(VBA.LCase$(EXPORT_ACTION_PREFIX))) = VBA.LCase$(EXPORT_ACTION_PREFIX) Then
            If Not private_IsKnownExportAction(actionId) Then
                hotkeyRows.Remove idx
                ioHasChanges = True
            End If
        End If

ContinueRow:
    Next idx

    private_RemoveStaleExportHotkeyRows = True
End Function

Private Function private_IsKnownExportAction(ByVal actionId As String) As Boolean
    Dim exportAlias As String

    If Not private_IsExportAction(actionId) Then Exit Function
    exportAlias = VBA.Trim$(VBA.Mid$(VBA.Trim$(actionId), VBA.Len(EXPORT_ACTION_PREFIX) + 1))
    If VBA.Len(exportAlias) = 0 Then Exit Function

    private_IsKnownExportAction = private_ExportAliasExists(exportAlias)
End Function

Private Function private_BuildExportActionId(ByVal exportAlias As String) As String
    exportAlias = VBA.Trim$(exportAlias)
    If VBA.Len(exportAlias) = 0 Then Exit Function
    private_BuildExportActionId = EXPORT_ACTION_PREFIX & exportAlias
End Function

Private Function private_BuildExportDefaultHotkey(ByVal exportIndex As Long) As String
    If exportIndex <= 0 Or exportIndex > MAX_EXPORT_HOTKEYS Then Exit Function
    private_BuildExportDefaultHotkey = "CTRL+" & VBA.CStr(exportIndex)
End Function

Private Function private_HotkeyRowsContainAction( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String _
) As Boolean
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    actionId = VBA.Trim$(actionId)
    If VBA.Len(actionId) = 0 Then Exit Function

    For Each rowItem In hotkeyRows
        If Not VBA.IsObject(rowItem) Then GoTo ContinueRow
        Set configEntry = Nothing
        On Error Resume Next
        Set configEntry = rowItem
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueRow
        If VBA.StrComp(VBA.Trim$(configEntry.Key), actionId, VBA.vbTextCompare) = 0 Then
            private_HotkeyRowsContainAction = True
            Exit Function
        End If

ContinueRow:
    Next rowItem
End Function

Private Function private_AddHotkeyRow( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String, _
    ByVal defaultHotkey As String _
) As Boolean
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    Set configEntry = New obj_ConfigEntry
    configEntry.Attr = VBA.vbNullString
    configEntry.Key = VBA.Trim$(actionId)
    configEntry.Value = VBA.Trim$(defaultHotkey)
    hotkeyRows.Add configEntry
    private_AddHotkeyRow = True
End Function

Private Function private_BuildDummyTable() As obj_TableDynamic
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim colIndex As Long
    Dim rowIndex As Long

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = "PrsnlEvntBuilder dummy table"

    For colIndex = 1 To 6
        If Not private_AddColumn(tableObj, "Column " & VBA.CStr(colIndex)) Then Exit Function
    Next colIndex

    For rowIndex = 1 To 8
        Set rowObj = New obj_Row
        For colIndex = 1 To 6
            rowObj.PushCellRaw "R" & VBA.CStr(rowIndex) & "C" & VBA.CStr(colIndex)
        Next colIndex
        If Not tableObj.PushRow(rowObj) Then Exit Function
    Next rowIndex

    Set private_BuildDummyTable = tableObj
End Function

Private Function private_AddColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal columnName As String _
) As Boolean
    Dim colObj As obj_Column

    If tableObj Is Nothing Then Exit Function
    Set colObj = New obj_Column
    colObj.Name = VBA.Trim$(columnName)
    If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Column " & VBA.CStr(tableObj.ColumnCount + 1)
    colObj.Position = tableObj.ColumnCount + 1
    private_AddColumn = tableObj.PushColumn(colObj)
End Function

Private Function private_AddSourceColumn(ByVal tableObj As obj_TableDynamic, ByVal columnName As String) As Boolean
    Dim colObj As obj_Column

    If tableObj Is Nothing Then Exit Function
    Set colObj = New obj_Column
    colObj.Name = VBA.Trim$(columnName)
    If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Column " & VBA.CStr(tableObj.ColumnCount + 1)
    colObj.Position = tableObj.ColumnCount + 1
    private_AddSourceColumn = tableObj.PushColumn(colObj)
End Function
