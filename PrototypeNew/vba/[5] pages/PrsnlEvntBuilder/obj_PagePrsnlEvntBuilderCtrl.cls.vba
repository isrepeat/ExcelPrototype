VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePrsnlEvntBuilderCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' //
' // Flow "Формы экспорта"
' //
' // 1. Верхняя draft-форма всегда показывает текущий выбранный профиль:
' //    основной профиль события или meta-профиль.
' // 2. Кнопка Apply переносит текущую draft-форму в staging-состояние контроллера:
' //    - основной профиль создает/заменяет m_ExportMainTable;
' //    - meta-профиль добавляет новую таблицу в m_ExportMetaTables.
' // 3. UI рендерит staging-состояние как "Форму экспорта":
' //    RuntimeItems.PrsnlEvntBuilder.ExportForm.Main -> основная таблица;
' //    RuntimeItems.PrsnlEvntBuilder.ExportForm.Meta -> список meta-таблиц.
' //    Эти UI-таблицы намеренно чистые: в них нет служебных колонок вроде
' //    meta_ProfileType, ManualOrderNo или SectionType.
' // 4. При запуске экспорта контроллер сначала читает staging-состояние
' //    "Формы экспорта". Если Apply еще не нажимали и staging пустой,
' //    используется fallback: текущая draft-строка экспортируется как одна
' //    основная таблица без meta-таблиц.
' //    Контроллер не отдает UI-таблицы напрямую.
' //    private_TryBuildExportSourceTables собирает отдельный export-source:
' //    - Collection таблиц, где tables(1) = main, tables(2..n) = meta;
' //    - context-объект с общими значениями, например SectionType и ManualOrderNo.
' // 5. Для meta-таблиц только в export-source копии добавляется колонка
' //    meta_ProfileType. Это нужно экспортерам, чтобы понять тип meta-строки,
' //    но не засорять визуальную "Форму экспорта" на листе.
' // 6. При смене основного профиля контроллер очищает staging-состояние,
' //    чтобы meta-строки старого контекста не подтянулись к новому событию.
'
Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PrsnlEvntBuilder.Controller"
Private Const CANDIDATE_TABLES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.EntityLookup.CandidateTables"
Private Const DUMMY_TABLES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.DummyTables"
Private Const HOTKEYS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Hotkeys"
Private Const PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Profiles"
Private Const META_PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MetaProfiles"
Private Const EXPORT_FORM_MAIN_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ExportForm.Main"
Private Const EXPORT_FORM_META_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ExportForm.Meta"
Private Const HOTKEY_ACCEPT_CANDIDATE_ROW As String = "Accept Candidate Row"
Private Const HOTKEY_SELECT_FORM_ROW As String = "Select Form Row"
Private Const HOTKEY_APPLY_EXPORT_FORM As String = "Apply Export Form"
Private Const EXPORT_ACTION_PREFIX As String = "Export "
Private Const DEFAULT_EXPORTER_CLASS As String = "obj_PEB_ExptrDailyScope"
Private Const MAX_EXPORT_HOTKEYS As Long = 9
Private Const LOOKUP_CANDIDATES_CONTROL_NAME As String = "LookupCandidatesTable"
Private Const EVENT_DRAFT_FORM_CONTAINER_NAME As String = "EventDraftForm"
Private Const EVENT_DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"
Private Const EVENT_DRAFT_ORDER_NO_LABEL_CONTROL_NAME As String = "EventDraftOrderNoLabel"
Private Const EVENT_DRAFT_INCOMING_NO_HEADER_NAME As String = "Вх. №"
Private Const EVENT_DRAFT_ORDER_LABEL_PREFIX As String = "Наказ №: "
Private Const EXPORT_META_PROFILE_TYPE_COLUMN_NAME As String = "meta_ProfileType"
Private Const EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY As String = "ManualOrderNo"
Private Const EXPORT_CONTEXT_SECTION_TYPE_KEY As String = "SectionType"
Private Const EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY As String = "WordExportPreviewText"
Private Const PROFILE_BUTTON_STYLE_NORMAL As String = "profileButton"
Private Const PROFILE_BUTTON_STYLE_SELECTED As String = "profileButtonSelected"
Private Const META_PROFILE_BUTTON_STYLE_NORMAL As String = "metaProfileButton"
Private Const META_PROFILE_BUTTON_STYLE_SELECTED As String = "metaProfileButtonSelected"

Private m_Page As obj_IPage
Private m_LookupFeature As obj_EntityLookupFeature
Private m_ExportAliases As Collection
Private m_ExporterClassByAlias As Object
Private m_ExportConfigTableByAlias As Object
Private m_ProfileConfigTable As obj_ConfigTable
Private m_SourceColumnAliasByCaption As Object
Private m_SelectedProfile As String
' Состояние "Формы экспорта": одна основная таблица и ноль/несколько
' meta-таблиц, подготовленных кнопкой Apply перед передачей в экспортер.
Private m_SelectedMainProfile As String
Private m_ExportMainTable As obj_TableDynamic
Private m_ExportMetaTables As Collection
Private m_WordExportPreviewText As String
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

    If Not private_RegisterProfileOptions(False) Then Exit Function
    If Not private_RegisterMetaProfileOptions(False) Then Exit Function
    If Not private_RegisterExportFormTables(False) Then Exit Function
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
    m_SelectedProfile = VBA.vbNullString
    m_SelectedMainProfile = VBA.vbNullString
    Set m_ExportMainTable = Nothing
    Set m_ExportMetaTables = Nothing
    m_WordExportPreviewText = VBA.vbNullString
    On Error GoTo 0
End Sub

Public Property Get WordExportPreviewText() As String
    WordExportPreviewText = m_WordExportPreviewText
End Property

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    UpdateData = m_LookupFeature.UpdateData(configControl)
    If Not UpdateData Then Exit Function
    If Not private_TryUpdateProfilesProvider(configControl) Then
        UpdateData = False
        Exit Function
    End If
    If Not private_TryUpdateExportSettings(configControl) Then Exit Function
    If Not private_RegisterExportFormTables(False) Then Exit Function
    If Not private_EnsureHotkeyRows(False) Then Exit Function
    UpdateData = True
End Function

Public Function PrepareRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.PrepareLookupRuntime(notifyChange) Then Exit Function
    If Not private_RegisterProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterMetaProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterExportFormTables(notifyChange) Then Exit Function
    If Not private_RegisterDummyTables(notifyChange) Then Exit Function
    If Not private_EnsureHotkeyRows(notifyChange) Then Exit Function
    PrepareRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    ClearLookupCandidates = m_LookupFeature.ClearLookupCandidates(renderNow)
End Function

Public Function ResolveProfileVisibilityState(ByVal tagsText As String) As String
    Dim profileText As String

    ResolveProfileVisibilityState = "collapsed"
    If m_Data Is Nothing Then Exit Function
    If Not private_TryGetSelectedProfile(profileText) Then Exit Function

    ResolveProfileVisibilityState = m_Data.ResolveProfileVisibilityState(profileText, tagsText)
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
        Case VBA.LCase$(HOTKEY_ACCEPT_CANDIDATE_ROW)
            If Not private_TryAcceptCandidateRowFromSelection(targetCell) Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If

        Case VBA.LCase$(HOTKEY_SELECT_FORM_ROW)
            If Not private_TrySelectScopedRowFromSelection(targetCell) Then Exit Function

        Case VBA.LCase$(HOTKEY_APPLY_EXPORT_FORM)
            If Not Me.RuntimeApplyExportForm() Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If
            RuntimeHandleHotkeyAction = True
            Exit Function

        Case Else
            Exit Function
    End Select

    rt_Messaging.fn_ShowStatusBarSuccess actionText & ": " & targetCell.Address(False, False) & " = '" & cellValue & "'", 3
    RuntimeHandleHotkeyAction = True
End Function

Public Function OnProfileButtonClick(Optional ByVal profileId As Variant) As Boolean
    Dim newProfile As String
    Dim previousEnableEvents As Boolean
    Dim perfStart As Double
    Dim perfLast As Double

    perfStart = VBA.Timer
    perfLast = perfStart

    newProfile = VBA.Trim$(VBA.CStr(profileId))
    If VBA.Len(newProfile) = 0 Then Exit Function
    If VBA.StrComp(private_NormalizeText(newProfile), private_NormalizeText(m_SelectedProfile), vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        private_LogPerfStep "profile-click:no-op", perfStart, perfLast, "profile='" & private_EscapeForLog(newProfile) & "'"
#End If
        OnProfileButtonClick = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profile-click:start", perfStart, perfLast, "from='" & private_EscapeForLog(m_SelectedProfile) & "' to='" & private_EscapeForLog(newProfile) & "'"
#End If

    If private_IsMainProfile(newProfile) Then
        If VBA.Len(VBA.Trim$(m_SelectedMainProfile)) = 0 Then
            m_SelectedMainProfile = newProfile
        ElseIf VBA.StrComp(private_NormalizeText(newProfile), private_NormalizeText(m_SelectedMainProfile), VBA.vbTextCompare) <> 0 Then
            private_ClearExportFormState
            m_SelectedMainProfile = newProfile
        End If
    End If

    m_SelectedProfile = newProfile
    If m_Page Is Nothing Then Exit Function
    If Not private_RegisterProfileOptions(False) Then Exit Function
    If Not private_RegisterMetaProfileOptions(False) Then Exit Function
    If Not private_RegisterExportFormTables(False) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profile-click:profiles-registered", perfStart, perfLast, "profile='" & private_EscapeForLog(m_SelectedProfile) & "'"
#End If

    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profile-click:events-disabled", perfStart, perfLast, "previousEnableEvents=" & VBA.LCase$(VBA.CStr(previousEnableEvents))
#End If
    On Error GoTo EH

    OnProfileButtonClick = rt_PageManager.fn_RenderPage(m_Page, "prsnlevntbuilder:profile-changed")
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profile-click:render-returned", perfStart, perfLast, "ok=" & VBA.LCase$(VBA.CStr(OnProfileButtonClick))
#End If

Cleanup:
    Application.EnableEvents = previousEnableEvents
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profile-click:done", perfStart, perfLast, "ok=" & VBA.LCase$(VBA.CStr(OnProfileButtonClick))
#End If
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profile-click:error", perfStart, perfLast, "err='" & private_EscapeForLog(Err.Description) & "'"
#End If
    Resume Cleanup
End Function

Public Function OnMetaProfileButtonClick(Optional ByVal profileId As Variant) As Boolean
    OnMetaProfileButtonClick = Me.OnProfileButtonClick(profileId)
End Function

Public Function RuntimeApplyExportForm() As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim profileText As String
    Dim previousEnableEvents As Boolean

    ' Apply переносит текущую видимую draft-форму в "Форму экспорта".
    ' Основной профиль заменяет основную строку; meta-профиль дописывает meta-строку.
    If Not private_TryGetSelectedProfile(profileText) Then Exit Function
    If Not private_TryBuildDraftFormSourceTable(sourceTable, False) Then Exit Function
    If sourceTable Is Nothing Then Exit Function

    sourceTable.SectionTitle = profileText
    If private_IsMetaProfile(profileText) Then
        If m_ExportMainTable Is Nothing Then
            rt_Messaging.fn_ShowStatusBarWarning "Create main export row before adding meta rows.", 3
            Exit Function
        End If
        If m_ExportMetaTables Is Nothing Then Set m_ExportMetaTables = New Collection
        m_ExportMetaTables.Add sourceTable
    Else
        Set m_ExportMainTable = sourceTable
        m_SelectedMainProfile = profileText
    End If

    If Not private_RegisterExportFormTables(False) Then Exit Function

    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreEventsAndFail
    RuntimeApplyExportForm = rt_PageManager.fn_RenderPage(m_Page, "prsnlevntbuilder:export-form-applied")
    Application.EnableEvents = previousEnableEvents
    If RuntimeApplyExportForm Then rt_Messaging.fn_ShowStatusBarSuccess "Export form updated: " & profileText, 3
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
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

Private Function private_RegisterProfileOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim profiles As Collection
    Dim profileOptions As Collection
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
    private_LogPerfStep "profiles:runtime-ready", perfStart, perfLast, "notifyChange=" & VBA.LCase$(VBA.CStr(notifyChange))
#End If

    Set profiles = m_Data.ProfileNames
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profiles:data-provider-ready", perfStart, perfLast
#End If
    If profiles Is Nothing Then Exit Function
    If profiles.Count = 0 Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profiles:options-loaded", perfStart, perfLast, "count=" & VBA.CStr(profiles.Count)
#End If
    If VBA.Len(VBA.Trim$(m_SelectedProfile)) = 0 Then m_SelectedProfile = VBA.Trim$(VBA.CStr(profiles.Item(1)))
    If VBA.Len(VBA.Trim$(m_SelectedMainProfile)) = 0 Then m_SelectedMainProfile = VBA.Trim$(m_SelectedProfile)
    If Not private_TryBuildOptionButtonRows( _
        profiles, _
        m_SelectedMainProfile, _
        PROFILE_BUTTON_STYLE_NORMAL, _
        PROFILE_BUTTON_STYLE_SELECTED, _
        2, _
        profileOptions) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profiles:button-options-built", perfStart, perfLast, "rows=" & VBA.CStr(profileOptions.Count)
#End If

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(PROFILES_RUNTIME_KEY)) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profiles:runtime-source-removed", perfStart, perfLast
#End If
    If Not runtimeSources.SetItemsSource(VBA.LCase$(PROFILES_RUNTIME_KEY), profileOptions, notifyChange) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    private_LogPerfStep "profiles:runtime-source-set", perfStart, perfLast
#End If

    private_RegisterProfileOptions = True
End Function

Private Function private_RegisterMetaProfileOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim profiles As Collection
    Dim profileOptions As Collection

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Set m_Data = New obj_PrsnlEvntBuilderData
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set profiles = m_Data.MetaProfileNames
    If profiles Is Nothing Then Exit Function
    If Not private_TryBuildOptionButtonRows( _
        profiles, _
        m_SelectedProfile, _
        META_PROFILE_BUTTON_STYLE_NORMAL, _
        META_PROFILE_BUTTON_STYLE_SELECTED, _
        1, _
        profileOptions) Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(META_PROFILES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(META_PROFILES_RUNTIME_KEY), profileOptions, notifyChange) Then Exit Function

    private_RegisterMetaProfileOptions = True
End Function

Private Function private_RegisterExportFormTables(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim mainTables As Collection
    Dim metaTables As Collection
    Dim metaTableObj As Variant

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set mainTables = New Collection
    If Not m_ExportMainTable Is Nothing Then mainTables.Add m_ExportMainTable

    Set metaTables = New Collection
    If Not m_ExportMetaTables Is Nothing Then
        For Each metaTableObj In m_ExportMetaTables
            If IsObject(metaTableObj) Then metaTables.Add metaTableObj
        Next metaTableObj
    End If

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(EXPORT_FORM_MAIN_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(EXPORT_FORM_MAIN_RUNTIME_KEY), mainTables, notifyChange) Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(EXPORT_FORM_META_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(EXPORT_FORM_META_RUNTIME_KEY), metaTables, notifyChange) Then Exit Function

    private_RegisterExportFormTables = True
End Function

Private Function private_TryBuildOptionButtonRows( _
    ByVal profiles As Collection, _
    ByVal selectedProfileText As String, _
    ByVal normalStyleName As String, _
    ByVal selectedStyleName As String, _
    ByVal itemsPerRow As Long, _
    ByRef outRows As Collection _
) As Boolean
    Dim profileObj As Variant
    Dim profileText As String
    Dim optionObj As obj_SelectOption
    Dim rowItems As Collection
    Dim rowObj As Object

    Set outRows = Nothing
    If profiles Is Nothing Then Exit Function
    If itemsPerRow <= 0 Then itemsPerRow = 1

    Set outRows = New Collection
    Set rowItems = Nothing
    For Each profileObj In profiles
        profileText = VBA.Trim$(VBA.CStr(profileObj))
        If VBA.Len(profileText) = 0 Then GoTo ContinueProfile

        If rowItems Is Nothing Then
            Set rowItems = New Collection
            Set rowObj = VBA.CreateObject("Scripting.Dictionary")
            rowObj.CompareMode = 1
            Set rowObj("Items") = rowItems
            outRows.Add rowObj
        ElseIf rowItems.Count >= itemsPerRow Then
            Set rowItems = New Collection
            Set rowObj = VBA.CreateObject("Scripting.Dictionary")
            rowObj.CompareMode = 1
            Set rowObj("Items") = rowItems
            outRows.Add rowObj
        End If

        Set optionObj = New obj_SelectOption
        optionObj.Caption = profileText
        optionObj.Id = profileText
        If VBA.StrComp(private_NormalizeText(profileText), private_NormalizeText(selectedProfileText), VBA.vbTextCompare) = 0 Then
            optionObj.StyleName = selectedStyleName
        Else
            optionObj.StyleName = normalStyleName
        End If
        rowItems.Add optionObj

ContinueProfile:
    Next profileObj

    private_TryBuildOptionButtonRows = True
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
    Dim sourceTables As Collection
    Dim exportContext As Object
    Dim exporter As obj_IDataExporter
    Dim exportAlias As String
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable

    If Not private_TryResolveExportAliasFromAction(actionId, exportAlias) Then Exit Function
    If Not private_TryGetExportSettings(exportAlias, exporterClassName, exportConfigTable) Then Exit Function
    If Not private_TryBuildExportSourceTables(sourceTables, exportContext) Then Exit Function

    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function

    If Not exporter.Export(sourceTables, exportContext) Then Exit Function
    If Not private_TryCaptureWordExportPreview(exportContext) Then Exit Function

    rt_Messaging.fn_ShowStatusBarSuccess EXPORT_ACTION_PREFIX & exportAlias & ": done", 3
    private_TryExportDraftByAction = True
End Function

Private Function private_TryUpdateExportSettings(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParser As obj_PrsnlEvntBuilderCfgParser

    private_ResetExportSettings
    If configControl Is Nothing Then
        private_TryUpdateExportSettings = True
        Exit Function
    End If

    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function
    Set m_ProfileConfigTable = configTable

    Set cfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function
    If Not cfgParser.TryGetEntityLookupColumnAliasByCaption(m_SourceColumnAliasByCaption) Then Exit Function
    If Not cfgParser.TryGetExportSettings(m_ExportAliases, m_ExporterClassByAlias, m_ExportConfigTableByAlias) Then Exit Function

    private_TryUpdateExportSettings = True
End Function

Private Function private_TryUpdateProfilesProvider(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParser As obj_PrsnlEvntBuilderCfgParser
    Dim providerClassName As String

    If configControl Is Nothing Then Exit Function
    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function

    Set cfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function
    If Not cfgParser.TryGetProfilesProviderClass(providerClassName) Then Exit Function

    If Not private_TryCreateProfilesProvider(providerClassName) Then Exit Function
    private_TryUpdateProfilesProvider = True
End Function

Private Function private_TryCreateProfilesProvider(ByVal providerClassName As String) As Boolean
    providerClassName = VBA.Trim$(providerClassName)
    If VBA.Len(providerClassName) = 0 Then Exit Function

    Select Case VBA.LCase$(providerClassName)
        Case VBA.LCase$("obj_PrsnlEvntBuilderData")
            Set m_Data = New obj_PrsnlEvntBuilderData
        Case Else
            VBA.MsgBox "PrototypeNew: unsupported profiles provider class '" & providerClassName & "'.", VBA.vbExclamation, "PrototypeNew / PrsnlEvntBuilder"
            Exit Function
    End Select

    private_TryCreateProfilesProvider = Not m_Data Is Nothing
End Function

Private Sub private_ResetExportSettings()
    Set m_ExportAliases = New Collection
    Set m_ExporterClassByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ExportConfigTableByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ProfileConfigTable = Nothing
    Set m_SourceColumnAliasByCaption = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

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
    Dim exporterToDailyScope As obj_PEB_ExptrDailyScope
    Dim exporterToMovement As obj_PEB_ExptrMovement
    Dim exporterToWord As obj_PEB_ExptrWord

    Set outExporter = Nothing
    exporterClassName = VBA.Trim$(exporterClassName)
    If VBA.Len(exporterClassName) = 0 Then exporterClassName = DEFAULT_EXPORTER_CLASS

    Select Case VBA.LCase$(exporterClassName)
        Case VBA.LCase$("obj_PEB_ExptrDailyScope")
            Set exporterToDailyScope = New obj_PEB_ExptrDailyScope
            If Not exporterToDailyScope.Initialize(exportConfigTable, m_ProfileConfigTable) Then Exit Function
            Set outExporter = exporterToDailyScope

        Case VBA.LCase$("obj_PEB_ExptrMovement")
            Set exporterToMovement = New obj_PEB_ExptrMovement
            If Not exporterToMovement.Initialize(exportConfigTable, m_ProfileConfigTable) Then Exit Function
            Set outExporter = exporterToMovement

        Case VBA.LCase$("obj_PEB_ExptrWord")
            Set exporterToWord = New obj_PEB_ExptrWord
            If Not exporterToWord.Initialize(exportConfigTable, m_ProfileConfigTable) Then Exit Function
            Set outExporter = exporterToWord

        Case Else
            VBA.MsgBox "PrototypeNew: unsupported data exporter class: " & exporterClassName, VBA.vbExclamation, "PrototypeNew / Data export"
            Exit Function
    End Select

    private_TryCreateDataExporter = Not outExporter Is Nothing
End Function

Private Function private_TryCaptureWordExportPreview(ByVal exportContext As Object) As Boolean
    Dim previewText As String
    Dim previousEnableEvents As Boolean

    If exportContext Is Nothing Then
        private_TryCaptureWordExportPreview = True
        Exit Function
    End If

    On Error Resume Next
    If exportContext.Exists(EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY) Then
        previewText = VBA.CStr(exportContext(EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY))
    End If
    If Err.Number <> 0 Then
        Err.Clear
        previewText = VBA.CStr(VBA.CallByName(exportContext, EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY, VbGet))
    End If
    On Error GoTo 0

    If VBA.Len(VBA.Trim$(previewText)) = 0 Then
        private_TryCaptureWordExportPreview = True
        Exit Function
    End If

    m_WordExportPreviewText = previewText
    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreEventsAndFail
    private_TryCaptureWordExportPreview = rt_PageManager.fn_RenderPage(m_Page, "prsnlevntbuilder:word-preview-updated")
    Application.EnableEvents = previousEnableEvents
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
End Function

Private Function private_TryBuildExportSourceTables( _
    ByRef outTables As Collection, _
    ByRef outContext As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim mainTable As obj_TableDynamic
    Dim exportTable As obj_TableDynamic
    Dim metaSourceTable As obj_TableDynamic
    Dim metaTableObj As Variant
    Dim sectionTypeText As String
    Dim manualOrderNoText As String

    Set outTables = Nothing
    Set outContext = Nothing
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    ' Экспортеры получают чистый контракт:
    '   tables(1) = основная таблица экспорта
    '   tables(2..n) = meta-таблицы экспорта
    '   context = общие значения для всех таблиц, например номер приказа и тип секции.
    Set outTables = New Collection
    Set outContext = VBA.CreateObject("Scripting.Dictionary")
    outContext.CompareMode = 1

    sectionTypeText = VBA.Trim$(m_SelectedMainProfile)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(m_SelectedProfile)
    manualOrderNoText = private_TryReadManualOrderNoValue(pageBase, ws)

    outContext(EXPORT_CONTEXT_SECTION_TYPE_KEY) = sectionTypeText
    outContext(EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY) = manualOrderNoText

    If m_ExportMainTable Is Nothing Then
        ' Удобный shortcut для частого случая "одна строка без meta": CTRL+1/2/3
        ' может экспортировать текущую draft-строку даже без предварительного Apply.
        If Not private_TryBuildDraftFormSourceTable(mainTable, False) Then Exit Function
        If Not private_TryCloneSourceTableForExport(mainTable, VBA.vbNullString, exportTable) Then Exit Function
        outTables.Add exportTable
        rt_Messaging.fn_ShowStatusBarWarning "Export Form is empty. Export uses the current draft row.", 3
        private_TryBuildExportSourceTables = True
        Exit Function
    End If

    Set mainTable = m_ExportMainTable
    If Not private_TryCloneSourceTableForExport(mainTable, VBA.vbNullString, exportTable) Then Exit Function
    outTables.Add exportTable

    If Not m_ExportMetaTables Is Nothing Then
        For Each metaTableObj In m_ExportMetaTables
            If Not IsObject(metaTableObj) Then GoTo ContinueMetaTable
            Set metaSourceTable = Nothing
            On Error Resume Next
            Set metaSourceTable = metaTableObj
            On Error GoTo 0
            If metaSourceTable Is Nothing Then GoTo ContinueMetaTable
            If Not private_TryCloneSourceTableForExport(metaSourceTable, VBA.CStr(metaSourceTable.SectionTitle), exportTable) Then Exit Function
            outTables.Add exportTable
ContinueMetaTable:
        Next metaTableObj
    End If

    private_TryBuildExportSourceTables = True
End Function

Private Function private_TryCloneSourceTableForExport( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal metaProfileType As String, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim resultTable As obj_TableDynamic
    Dim sourceColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim resultRow As obj_Row
    Dim colIndex As Long
    Dim rowIndex As Long

    Set outTable = Nothing
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.ColumnCount <= 0 Then Exit Function

    ' UI-таблицы остаются чистыми. Только копии meta-таблиц для экспортера
    ' получают meta_ProfileType, чтобы экспортер мог понять тип meta-строки.
    Set resultTable = New obj_TableDynamic
    resultTable.SectionTitle = sourceTable.SectionTitle

    For colIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(colIndex)
        If sourceColumn Is Nothing Then Exit Function
        If Not resultTable.PushColumn(sourceColumn) Then Exit Function
    Next colIndex
    metaProfileType = VBA.Trim$(metaProfileType)
    If VBA.Len(metaProfileType) > 0 Then
        If Not private_AddSourceColumn(resultTable, EXPORT_META_PROFILE_TYPE_COLUMN_NAME) Then Exit Function
    End If

    For rowIndex = 1 To sourceTable.RowCount
        Set sourceRow = sourceTable.Rows.Item(rowIndex)
        If sourceRow Is Nothing Then Exit Function
        Set resultRow = New obj_Row
        For colIndex = 1 To sourceTable.ColumnCount
            resultRow.PushCellRaw sourceRow.GetCellValue(colIndex)
        Next colIndex
        If VBA.Len(metaProfileType) > 0 Then resultRow.PushCellRaw metaProfileType
        If Not resultTable.PushRow(resultRow) Then Exit Function
    Next rowIndex

    Set outTable = resultTable
    private_TryCloneSourceTableForExport = True
End Function

Private Function private_TryBuildDraftFormSourceTable( _
    ByRef outTable As obj_TableDynamic, _
    Optional ByVal includeBlankHeaderColumns As Boolean = False _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim draftValuesRange As Range
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim colOffset As Long
    Dim sheetCol As Long
    Dim headerText As String

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

    ' Собираем таблицу только из отрендеренных колонок draft-формы.
    ' Пустые заголовки по умолчанию пропускаем, чтобы не появлялись Column N.
    Set sourceTable = New obj_TableDynamic
    sourceTable.SectionTitle = "PrsnlEvntBuilder draft form"
    Set sourceRow = New obj_Row

    For colOffset = 1 To draftValuesRange.Columns.Count
        sheetCol = draftValuesRange.Column + colOffset - 1
        headerText = private_ReadHeaderText(ws.Cells(draftValuesRange.Row - 1, sheetCol))
        If VBA.Len(headerText) = 0 Then
            If Not includeBlankHeaderColumns Then GoTo ContinueDraftColumn
            headerText = "Column " & VBA.CStr(colOffset)
        End If
        If Not private_AddSourceColumn(sourceTable, headerText) Then Exit Function
        sourceRow.PushCellRaw ws.Cells(draftValuesRange.Row, sheetCol).Value2

ContinueDraftColumn:
    Next colOffset

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

Private Function private_TryGetSelectedProfile(ByRef outProfile As String) As Boolean
    outProfile = VBA.Trim$(m_SelectedProfile)
    If VBA.Len(outProfile) = 0 Then
        rt_Messaging.fn_ShowStatusBarWarning "Event profile is not selected.", 3
        Exit Function
    End If

    private_TryGetSelectedProfile = True
End Function

Private Function private_IsMetaProfile(ByVal profileText As String) As Boolean
    If m_Data Is Nothing Then Set m_Data = New obj_PrsnlEvntBuilderData
    private_IsMetaProfile = m_Data.IsMetaProfileName(profileText)
End Function

Private Function private_IsMainProfile(ByVal profileText As String) As Boolean
    profileText = VBA.Trim$(profileText)
    If VBA.Len(profileText) = 0 Then Exit Function
    private_IsMainProfile = Not private_IsMetaProfile(profileText)
End Function

Private Sub private_ClearExportFormState()
    Set m_ExportMainTable = Nothing
    Set m_ExportMetaTables = New Collection
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
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_ACCEPT_CANDIDATE_ROW, "CTRL+ENTER", hasChanges) Then Exit Function
                If Not private_EnsureExportHotkeyRows(hotkeyRows, hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_APPLY_EXPORT_FORM, "ALT+ARROWDOWN", hasChanges) Then Exit Function
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
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_ACCEPT_CANDIDATE_ROW, "CTRL+ENTER") Then Exit Function
    If Not private_AddExportHotkeyRows(hotkeyRows) Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_APPLY_EXPORT_FORM, "ALT+ARROWDOWN") Then Exit Function

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

Private Function private_RemoveHotkeyRowsByAction( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String, _
    ByRef ioHasChanges As Boolean _
) As Boolean
    Dim idx As Long
    Dim rowObj As Object
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    actionId = VBA.Trim$(actionId)
    If VBA.Len(actionId) = 0 Then
        private_RemoveHotkeyRowsByAction = True
        Exit Function
    End If

    For idx = hotkeyRows.Count To 1 Step -1
        Set rowObj = Nothing
        Set configEntry = Nothing
        On Error Resume Next
        Set rowObj = hotkeyRows.Item(idx)
        Set configEntry = rowObj
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueRow

        If VBA.StrComp(VBA.Trim$(configEntry.Key), actionId, VBA.vbTextCompare) = 0 Then
            hotkeyRows.Remove idx
            ioHasChanges = True
        End If

ContinueRow:
    Next idx

    private_RemoveHotkeyRowsByAction = True
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
    Dim columnAlias As String

    If tableObj Is Nothing Then Exit Function
    Set colObj = New obj_Column
    colObj.Name = VBA.Trim$(columnName)
    If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Column " & VBA.CStr(tableObj.ColumnCount + 1)
    colObj.Position = tableObj.ColumnCount + 1

    ' Export-form таблица строится из видимых заголовков листа. Чтобы WORD
    ' templates могли ссылаться на стабильные ключи ({FIO}, {IPN}, ...), рядом
    ' сохраняем alias из PrsnlEvntBuilderProfiles.xml.
    If Not m_SourceColumnAliasByCaption Is Nothing Then
        If m_SourceColumnAliasByCaption.Exists(colObj.Name) Then
            columnAlias = VBA.Trim$(VBA.CStr(m_SourceColumnAliasByCaption(colObj.Name)))
            If VBA.Len(columnAlias) > 0 Then
                If Not colObj.AddAlias(columnAlias) Then Exit Function
            End If
        End If
    End If

    private_AddSourceColumn = tableObj.PushColumn(colObj)
End Function
