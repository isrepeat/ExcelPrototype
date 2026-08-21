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
Private Const ORDER_HISTORY_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.OrderHistory"
Private Const VALIDATION_RESULTS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ValidationResults"
Private Const MEDICAL_REPORTS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MedicalReports"
Private Const MEDICAL_REPORTS_FILTERS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MedicalReportFilters"
Private Const MEDICAL_REPORTS_VIEW_CLASS_KEY As String = "PrsnlEvntBuilder.MedicalReportsViewClass"
Private Const EXPORT_EDITING_SUBMODE_CLASS_KEY As String = "PrsnlEvntBuilder.ExportEditingSubmodeClass"
Private Const VALIDATION_SUBMODE_CLASS_KEY As String = "PrsnlEvntBuilder.ValidationSubmodeClass"
Private Const BOTTOM_WORKSPACE_EDIT As String = "edit"
Private Const BOTTOM_WORKSPACE_VALIDATION As String = "validation"
Private Const BOTTOM_WORKSPACE_MEDICAL_REPORTS As String = "medical-reports"
Private Const HOTKEYS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Hotkeys"
Private Const PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Profiles"
Private Const ADDITIONAL_PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.AdditionalProfiles"
Private Const META_PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MetaProfiles"
Private Const EXPORT_FORM_MAIN_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ExportForm.Main"
Private Const EXPORT_FORM_META_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ExportForm.Meta"
Private Const MOVEMENT_HISTORY_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MovementHistory"
Private Const MOVEMENT_EVENTS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MovementEvents"
Private Const WORD_EVENTS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.WordEvents"
Private Const HOTKEY_ACCEPT_CANDIDATE_ROW As String = "Accept Candidate Row"
Private Const HOTKEY_REPORT_TVO_CANDIDATES As String = "Report TVO Candidates"
Private Const HOTKEY_SELECT_FORM_ROW As String = "Select Form Row"
Private Const HOTKEY_APPLY_EXPORT_FORM As String = "Apply Export Form"
Private Const HOTKEY_CLEAR_EXPORT_FORM As String = "Clear Export Form"
Private Const HOTKEY_DELETE_EXPORTED_EVENT As String = "Delete Exported Event"
Private Const HOTKEY_EXPORT_TO_WORD As String = "Export to WORD"
' Значение action id сохраняется для совместимости с существующими snapshot.
Private Const HOTKEY_EXPORT_MOVEMENT_WORD As String = "Word Movement Word"
Private Const EXPORT_ACTION_PREFIX As String = "Export "
Private Const MAX_EXPORT_HOTKEYS As Long = 9
Private Const LOOKUP_CANDIDATES_CONTROL_NAME As String = "LookupCandidatesTable"
Private Const FIO_LOOKUP_KEY As String = "op_FIO"
Private Const COMMANDER_LOOKUP_KEY As String = "op_Commander"
Private Const HOSPITAL_LOOKUP_KEY As String = "op_Hospital"
Private Const TO_HOSPITAL_LOOKUP_KEY As String = "op_ToHospital"
Private Const WORD_EXPORT_PANEL_CONTAINER_NAME As String = "WordExportPanel"
Private Const WORD_EXPORT_ACTIONS_CONTAINER_NAME As String = "WordExportActions"
Private Const WORD_EXPORT_PREVIEW_CONTROL_NAME As String = "WordExportPreview"
Private Const WORD_EXPORT_PREVIEW_BUTTON_SHAPE_NAME As String = "btn_UseWordPreview"
Private Const WORD_EXPORT_ACTIVE_BUTTON_SHAPE_NAME As String = "btn_UseWordPreviewActive"
Private Const WORD_BOOKMARKS_TOGGLE_CONTROL_NAME As String = "ToggleWordBookmarks"
Private Const EVENT_DRAFT_FORM_CONTAINER_NAME As String = "EventDraftForm"
Private Const EVENT_DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"
Private Const EVENT_DRAFT_ORDER_NO_CONTAINER_NAME As String = "EventDraftOrderNoValue"
Private Const EVENT_DRAFT_ORDER_YEAR_CONTAINER_NAME As String = "EventDraftOrderYearValue"
Private Const EVENT_EXPORT_MAIN_CONTROL_NAME As String = "EventExportMainTable"
Private Const EVENT_EXPORT_META_CONTROL_NAME As String = "EventExportMetaTables"
Private Const EXPORT_META_PROFILE_TYPE_COLUMN_NAME As String = "meta_ProfileType"
Private Const EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY As String = "ManualOrderNo"
Private Const EXPORT_CONTEXT_MANUAL_ORDER_YEAR_KEY As String = "ManualOrderYear"
Private Const EXPORT_CONTEXT_MANUAL_ORDER_DATE_SERIAL_KEY As String = "ManualOrderDateSerial"
Private Const EXPORT_CONTEXT_SECTION_TYPE_KEY As String = "SectionType"
Private Const EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY As String = "WordExportPreviewText"
Private Const EXPORT_CONTEXT_VALIDATE_MOVEMENT_KEY As String = "ValidateMovement"
Private Const EXPORT_CONTEXT_VALIDATE_WORD_KEY As String = "ValidateWord"
Private Const EXPORT_CONTEXT_REPORT_IS_TVO_KEY As String = "ReportIsTvo"
Private Const EXPORT_CONTEXT_MOVEMENT_PREVALIDATED_KEY As String = "MovementPrevalidated"
Private Const LOOKUP_MODE_CONTROL_NAME As String = "LookupMode"
Private Const TEMPORARY_PERSONNEL_CONTROL_NAME As String = "TemporaryPersonnel"
Private Const TEMPORARY_PERSONNEL_LOOKUP_KEY As String = "op_FIOTemporaryPersonnel"
Private Const TEMPORARY_IPN_MARKER As String = "ТП-ІПН"
Private Const TEMPORARY_POSITION_MARKER As String = "ТП-ПОСАДА"
Private Const MOVEMENT_HISTORY_TABLE_CONTROL_NAME As String = "MovementHistoryTable"
Private Const MOVEMENT_EVENTS_CONTROL_NAME As String = "MovementEventsMenu"
Private Const VALIDATION_RESULTS_CONTROL_NAME As String = "EmbeddedValidationResults"
Private Const MEDICAL_REPORTS_CONTROL_NAME As String = "EmbeddedMedicalReports"
Private Const MEDICAL_REPORTS_FILTER_INPUTS_CONTAINER As String = "MedicalReportsFilterInputs"
Private Const BOTTOM_WORKSPACE_CONTAINER_NAME As String = "BottomWorkspaceArea"
Private Const WORD_EVENTS_CONTROL_NAME As String = "WordEventsMenu"
Private Const MOVEMENT_HISTORY_LIMIT_INPUT_NAME As String = "MovementHistoryLimitInput"
Private Const ADDITIONAL_PROFILE_SELECT_CONTROL_NAME As String = "EventDraftAdditionalProfileSelect"
Private Const VALIDATE_MOVEMENT_CONTROL_NAME As String = "ValidateMovement"
Private Const VALIDATE_WORD_CONTROL_NAME As String = "ValidateWord"
Private Const PROFILE_BUTTON_TAG_ARRIVAL As String = "arrival"
Private Const PROFILE_BUTTON_TAG_DEPARTURE As String = "departure"
Private Const META_PROFILE_BUTTON_TAG_NORMAL As String = "meta"
Private Const PROFILE_BUTTON_STATE_SELECTED As String = "selected"
Private Const ABSENCE_DEPARTURE_LOOKBACK_DAYS As Long = 5
Private Const ABSENCE_DEPARTURE_LOOKAHEAD_DAYS As Long = 10
' Канонические алиасы полей draft-формы. Отображаемые Caption этих полей
' принадлежат конфигу и не должны использоваться в логике контроллера.
Private Const DRAFT_ALIAS_RANK As String = "_Rank"
Private Const DRAFT_ALIAS_FIO As String = "_FIO"
Private Const DRAFT_ALIAS_IPN As String = "_IPN"
Private Const DRAFT_ALIAS_POSITION_CODE As String = "_PositionCode"
Private Const DRAFT_ALIAS_POSITION_NAME As String = "_PositionName"
Private Const DRAFT_ALIAS_DESTINATION As String = "_Destination"
Private Const DRAFT_ALIAS_HOSPITAL As String = "_Hospital"
Private Const DRAFT_ALIAS_HOSPITAL_SHORT As String = "_HospitalShort"
Private Const DRAFT_ALIAS_TO_HOSPITAL_SHORT As String = "_ToHospitalShort"
Private Const DRAFT_ALIAS_REPORT_RANK As String = "_ReportRank"
Private Const DRAFT_ALIAS_REPORT_PERSON As String = "_ReportPerson"
Private Const DRAFT_ALIAS_REPORT_POSITION_CODE As String = "_ReportPositionCode"
Private Const DRAFT_ALIAS_INCOMING_NO As String = "_IncomingNo"
Private Const DRAFT_ALIAS_INCOMING_DATE As String = "_IncomingDate"
Private Const DRAFT_ALIAS_DOC_NO As String = "_DocNo"
Private Const DRAFT_ALIAS_DOC_DATE As String = "_DocDate"
Private Const DRAFT_ALIAS_DURATION_DAYS As String = "_DurationDays"
Private Const DRAFT_ALIAS_DATE_FROM As String = "_DateFrom"
Private Const ABSENCE_ORDER_DATE_ALIAS As String = "_AbsenceOrderDate"
Private Const DRAFT_ALIAS_VACATION_TICKET_NO As String = "_VacationTicketNo"
Private Const DRAFT_ALIAS_VACATION_TICKET_DATE As String = "_VacationTicketDate"
Private Const DRAFT_ALIAS_VLK_NO As String = "_VlkNo"
Private Const DRAFT_ALIAS_VLK_DATE As String = "_VlkDate"
Private Const EXPORT_SOURCE_IPN_COLUMN As String = "IPN"

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
Private m_ExportMainReportIsTvo As Boolean
Private m_ExportMetaTables As Collection
Private m_MovementHistoryTable As obj_TableDynamic
Private m_MovementEventIds As Collection
Private m_WordEventIds As Collection
Private m_MovementEventCaptions As Collection
Private m_WordEventCaptions As Collection
Private m_DeletedExportSnapshots As Object
Private m_WordExportPreviewText As String
Private m_IsWordPreviewExportMode As Boolean
Private m_AreWordBookmarksShownAsMarkers As Boolean
Private m_AreExportedEventsShown As Boolean
Private m_ExportCommonData As obj_PEB_ExptrCommonDataPrvdr
Private m_HasResolvedOrderPair As Boolean
Private m_ResolvedOrderNo As String
Private m_ResolvedOrderDate As Date
Private m_OrderHistoryItems As Collection
Private m_BottomWorkspaceMode As String
Private m_MedicalReportsScen As obj_PEB_MedicalReportsScen
Private m_ExportEditingScen As obj_PEB_ExportEditingScen
Private m_ExporterCfgDataProvider As obj_PEB_ExptrCfgDataPrvdr
Private m_CachedMovementExporter As obj_PEB_ExptrMovement
Private m_CachedWordExporter As obj_PEB_ExptrWord
Private m_HasPendingMovementReceipt As Boolean
Private m_PendingMovementReceiptIpn As String
Private m_PendingMovementReceiptSectionType As String
Private m_PendingMovementReceiptOrderNo As String
Private m_IsLookupEnabled As Boolean
Private m_IsTemporaryPersonnelEnabled As Boolean
Private m_IsMovementValidationEnabled As Boolean
Private m_IsWordValidationEnabled As Boolean
Private m_IsMovementHistoryEnabled As Boolean
Private m_DraftReportIsTvo As Boolean
Private m_ReportOwnPositionCode As String
Private m_IsReporterTvoCandidatesActive As Boolean
Private m_SuppressLookupSearch As Boolean
Private m_Data As obj_PrsnlEvntBuilderData
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
    Set m_Data = Nothing
    Set m_ExportCommonData = New obj_PEB_ExptrCommonDataPrvdr
    If Not m_ExportCommonData.Initialize() Then Exit Function
    private_ClearResolvedOrderPair
    m_BottomWorkspaceMode = VBA.vbNullString
    private_ResetExportSettings
    m_IsLookupEnabled = True
    m_IsTemporaryPersonnelEnabled = False
    m_IsMovementValidationEnabled = True
    m_IsWordValidationEnabled = False
    m_IsMovementHistoryEnabled = False
    m_AreWordBookmarksShownAsMarkers = False
    m_AreExportedEventsShown = False
    Set m_MovementEventIds = New Collection
    Set m_WordEventIds = New Collection
    Set m_MovementEventCaptions = New Collection
    Set m_WordEventCaptions = New Collection
    Set m_DeletedExportSnapshots = VBA.CreateObject("Scripting.Dictionary")
    m_DeletedExportSnapshots.CompareMode = VBA.vbTextCompare

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_LookupFeature = New obj_EntityLookupFeature
    If Not m_LookupFeature.Initialize( _
        pageInterface, _
        CANDIDATE_TABLES_RUNTIME_KEY, _
        "prsnlevntbuilder:entitylookup") Then Exit Function

    ' Profile-dependent runtime sources регистрируются после обязательного
    ' UpdateDataFromConfigTable, где фабрика создаёт настроенный provider.
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePrsnlEvntBuilderCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    ' Select хранит selectedId в CustomXMLPart и переживает удаление листа.
    ' Dispose страницы должен явно убрать это состояние, иначе после Clear Pages
    ' новая PEB-страница снова откроется на прежней дополнительной секции.
    private_ClearAdditionalProfilePersistedState
    If Not m_LookupFeature Is Nothing Then m_LookupFeature.Dispose
    Set m_LookupFeature = Nothing
    Set m_Page = Nothing
    Set m_ExportAliases = Nothing
    Set m_ExporterClassByAlias = Nothing
    Set m_ExportConfigTableByAlias = Nothing
    Set m_Data = Nothing
    If Not m_ExportCommonData Is Nothing Then m_ExportCommonData.Dispose
    Set m_ExportCommonData = Nothing
    m_HasResolvedOrderPair = False
    m_ResolvedOrderNo = VBA.vbNullString
    m_ResolvedOrderDate = 0
    Set m_OrderHistoryItems = Nothing
    ' Сначала освобождаем borrowers, затем общий config provider, которым
    ' WORD exporter может пользоваться без владения его lifetime.
    private_DisposeCachedExporters
    If Not m_ExporterCfgDataProvider Is Nothing Then m_ExporterCfgDataProvider.Dispose
    Set m_ExporterCfgDataProvider = Nothing
    If Not m_MedicalReportsScen Is Nothing Then m_MedicalReportsScen.Dispose
    Set m_MedicalReportsScen = Nothing
    If Not m_ExportEditingScen Is Nothing Then m_ExportEditingScen.Dispose
    Set m_ExportEditingScen = Nothing
    m_SelectedProfile = VBA.vbNullString
    m_SelectedMainProfile = VBA.vbNullString
    Set m_ExportMainTable = Nothing
    m_ExportMainReportIsTvo = False
    Set m_ExportMetaTables = Nothing
    Set m_MovementHistoryTable = Nothing
    Set m_MovementEventIds = Nothing
    Set m_WordEventIds = Nothing
    Set m_MovementEventCaptions = Nothing
    Set m_WordEventCaptions = Nothing
    Set m_DeletedExportSnapshots = Nothing
    m_WordExportPreviewText = VBA.vbNullString
    m_IsWordPreviewExportMode = False
    m_AreWordBookmarksShownAsMarkers = False
    m_AreExportedEventsShown = False
    m_IsMovementHistoryEnabled = False
    m_DraftReportIsTvo = False
    m_ReportOwnPositionCode = VBA.vbNullString
    On Error GoTo 0
End Sub

Private Sub private_ClearAdditionalProfilePersistedState()
    Dim pageBase As obj_PageBase
    Dim selectState As obj_SelectControlVMStatic
    Dim selectKey As String

    On Error GoTo EH

    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub
    If pageBase.Worksheet Is Nothing Then Exit Sub

    selectKey = VBA.LCase$(pageBase.Worksheet.Name & "|" & _
        ADDITIONAL_PROFILE_SELECT_CONTROL_NAME)
    Set selectState = New obj_SelectControlVMStatic
    If Not selectState.SetSelectedId(selectKey, VBA.vbNullString) Then
        ex_Core.fn_Diagnostic_LogError _
            "PrsnlEventBuilder: failed to clear persisted additional profile " & _
            "selection for key '" & selectKey & "'."
    End If
    Exit Sub

EH:
    ex_Core.fn_Diagnostic_LogError _
        "PrsnlEventBuilder: exception while clearing persisted additional " & _
        "profile selection: [" & VBA.CStr(Err.Number) & "] " & Err.Description
End Sub

Public Property Get WordExportPreviewText() As String
    WordExportPreviewText = m_WordExportPreviewText
End Property

Public Property Get HasResolvedOrderPair() As Boolean
    HasResolvedOrderPair = m_HasResolvedOrderPair
End Property

Public Property Get IsExportEditingWorkspaceVisible() As Boolean
    IsExportEditingWorkspaceVisible = (VBA.StrComp( _
        m_BottomWorkspaceMode, BOTTOM_WORKSPACE_EDIT, VBA.vbTextCompare) = 0)
End Property

Public Property Get IsValidationWorkspaceVisible() As Boolean
    IsValidationWorkspaceVisible = (VBA.StrComp( _
        m_BottomWorkspaceMode, BOTTOM_WORKSPACE_VALIDATION, VBA.vbTextCompare) = 0)
End Property

Public Property Get IsMedicalReportsWorkspaceVisible() As Boolean
    IsMedicalReportsWorkspaceVisible = (VBA.StrComp( _
        m_BottomWorkspaceMode, BOTTOM_WORKSPACE_MEDICAL_REPORTS, _
        VBA.vbTextCompare) = 0)
End Property

Public Property Get IsBottomWorkspaceVisible() As Boolean
    IsBottomWorkspaceVisible = (Me.IsExportEditingWorkspaceVisible Or _
        Me.IsValidationWorkspaceVisible Or _
        Me.IsMedicalReportsWorkspaceVisible)
End Property

Public Function ShowExportEditingWorkspace( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim previousWorkspaceMode As String

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Me.IsExportEditingWorkspaceVisible Then
        previousWorkspaceMode = m_BottomWorkspaceMode
        m_BottomWorkspaceMode = VBA.vbNullString
        private_DeleteBottomWorkspaceCommandShapes pageBase.Worksheet
        If Not pageBase.TryReflowLayoutContainer( _
            BOTTOM_WORKSPACE_CONTAINER_NAME) Then
            m_BottomWorkspaceMode = previousWorkspaceMode
            VBA.MsgBox "Не вдалося частково приховати таблицю редагування. " & _
                "Натисніть 'Update Sheet' для відновлення сторінки.", _
                VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
            Exit Function
        End If
        ShowExportEditingWorkspace = True
        Exit Function
    End If

    If Not private_TryEnsureExportEditingSubmode() Then Exit Function
    If Not m_ExportEditingScen.PrepareOpen() Then Exit Function

    previousWorkspaceMode = m_BottomWorkspaceMode
    m_BottomWorkspaceMode = BOTTOM_WORKSPACE_EDIT
    private_DeleteBottomWorkspaceCommandShapes pageBase.Worksheet
    If Not pageBase.TryReflowLayoutContainer( _
        BOTTOM_WORKSPACE_CONTAINER_NAME) Then
        m_BottomWorkspaceMode = previousWorkspaceMode
        VBA.MsgBox "Не вдалося частково відкрити таблицю редагування. " & _
            "Натисніть 'Update Sheet' для відновлення сторінки.", _
            VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
        Exit Function
    End If
    ShowExportEditingWorkspace = True
End Function

Public Function ShowValidationWorkspace( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pebMovementVldtnScen As obj_PEB_MovementVldtnScen
    Dim pageBase As obj_PageBase
    Dim previousWorkspaceMode As String

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Me.IsValidationWorkspaceVisible Then
        previousWorkspaceMode = m_BottomWorkspaceMode
        m_BottomWorkspaceMode = VBA.vbNullString
        private_DeleteBottomWorkspaceCommandShapes pageBase.Worksheet
        If Not pageBase.TryReflowLayoutContainer( _
            BOTTOM_WORKSPACE_CONTAINER_NAME) Then
            m_BottomWorkspaceMode = previousWorkspaceMode
            VBA.MsgBox "Не вдалося частково приховати таблицю валідації. " & _
                "Натисніть 'Update Sheet' для відновлення сторінки.", _
                VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
            Exit Function
        End If
        ShowValidationWorkspace = True
        Exit Function
    End If

    If Not m_HasResolvedOrderPair Then
        VBA.MsgBox "Спочатку прийміть номер і дату наказу.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Валідація"
        Exit Function
    End If
    If m_ProfileConfigTable Is Nothing Then Exit Function
    If Not private_TryCreateValidationSubmode(pebMovementVldtnScen) Then Exit Function
    If Not pebMovementVldtnScen.InitializeEmbedded( _
        m_Page, m_ProfileConfigTable, m_ResolvedOrderNo, _
        m_ResolvedOrderDate, VALIDATION_RESULTS_RUNTIME_KEY) Then Exit Function
    If Not pebMovementVldtnScen.RunEmbedded(False) Then Exit Function
    Set pebMovementVldtnScen = Nothing

    previousWorkspaceMode = m_BottomWorkspaceMode
    m_BottomWorkspaceMode = BOTTOM_WORKSPACE_VALIDATION
    private_DeleteBottomWorkspaceCommandShapes pageBase.Worksheet
    If Not pageBase.TryReflowLayoutContainer( _
        BOTTOM_WORKSPACE_CONTAINER_NAME) Then
        m_BottomWorkspaceMode = previousWorkspaceMode
        VBA.MsgBox "Не вдалося частково відкрити таблицю валідації. " & _
            "Натисніть 'Update Sheet' для відновлення сторінки.", _
            VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
        Exit Function
    End If
    ShowValidationWorkspace = True
End Function

Public Function ShowMedicalReportsWorkspace( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim previousWorkspaceMode As String

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Me.IsMedicalReportsWorkspaceVisible Then
        previousWorkspaceMode = m_BottomWorkspaceMode
        m_BottomWorkspaceMode = VBA.vbNullString
        private_DeleteBottomWorkspaceCommandShapes pageBase.Worksheet
        If Not pageBase.TryReflowLayoutContainer( _
            BOTTOM_WORKSPACE_CONTAINER_NAME) Then
            m_BottomWorkspaceMode = previousWorkspaceMode
            Exit Function
        End If
        ShowMedicalReportsWorkspace = True
        Exit Function
    End If

    If Not private_TryPrepareMedicalReports(False) Then Exit Function
    previousWorkspaceMode = m_BottomWorkspaceMode
    m_BottomWorkspaceMode = BOTTOM_WORKSPACE_MEDICAL_REPORTS
    private_DeleteBottomWorkspaceCommandShapes pageBase.Worksheet
    If Not pageBase.TryReflowLayoutContainer( _
        BOTTOM_WORKSPACE_CONTAINER_NAME) Then
        m_BottomWorkspaceMode = previousWorkspaceMode
        Exit Function
    End If
    ShowMedicalReportsWorkspace = True
End Function

Public Function RefreshBottomWorkspace( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pebMovementVldtnScen As obj_PEB_MovementVldtnScen
    Dim pageBase As obj_PageBase

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    If Me.IsExportEditingWorkspaceVisible Then
        If Not private_TryEnsureExportEditingSubmode() Then Exit Function
        If Not m_ExportEditingScen.Refresh() Then Exit Function
        RefreshBottomWorkspace = True
        Exit Function
    End If

    If Me.IsValidationWorkspaceVisible Then
        If m_ProfileConfigTable Is Nothing Then Exit Function
        If Not private_TryCreateValidationSubmode( _
            pebMovementVldtnScen) Then Exit Function
        If Not pebMovementVldtnScen.InitializeEmbedded( _
            m_Page, m_ProfileConfigTable, m_ResolvedOrderNo, _
            m_ResolvedOrderDate, VALIDATION_RESULTS_RUNTIME_KEY) Then Exit Function
        If Not pebMovementVldtnScen.RunEmbedded(False) Then Exit Function
        If Not pageBase.TryReflowControl(VALIDATION_RESULTS_CONTROL_NAME) Then
            VBA.MsgBox "Не вдалося частково оновити таблицю валідації.", _
                VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
            Exit Function
        End If
        RefreshBottomWorkspace = True
        Exit Function
    End If

    If Me.IsMedicalReportsWorkspaceVisible Then
        RefreshBottomWorkspace = Me.SearchMedicalReports()
    End If
End Function

Public Function TryLoadExportEditingEvents( _
    ByVal loadEvents As Boolean, _
    ByVal notifyChange As Boolean _
) As Boolean
    TryLoadExportEditingEvents = private_RegisterExportedEventMenus( _
        loadEvents, notifyChange)
End Function

Private Function private_TryEnsureExportEditingSubmode() As Boolean
    Dim prsnlEvntBuilderCfgParser As obj_PrsnlEvntBuilderCfgParser
    Dim submodeClassName As String

    If Not m_ExportEditingScen Is Nothing Then
        private_TryEnsureExportEditingSubmode = True
        Exit Function
    End If
    If m_ProfileConfigTable Is Nothing Then
        VBA.MsgBox "Відсутня конфігурація профілю для таблиці редагування.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Редагування РУХ / WORD"
        Exit Function
    End If
    Set prsnlEvntBuilderCfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not prsnlEvntBuilderCfgParser.Initialize( _
        m_ProfileConfigTable) Then Exit Function
    submodeClassName = VBA.Trim$(prsnlEvntBuilderCfgParser.GetOptionalValue( _
        EXPORT_EDITING_SUBMODE_CLASS_KEY, VBA.vbNullString))
    If VBA.Len(submodeClassName) = 0 Then
        VBA.MsgBox "Відсутній обов'язковий class таблиці редагування у ключі '" & _
            EXPORT_EDITING_SUBMODE_CLASS_KEY & "'.", VBA.vbExclamation, _
            "PrsnlEventBuilder / Редагування РУХ / WORD"
        Exit Function
    End If
    Select Case VBA.LCase$(submodeClassName)
        Case VBA.LCase$("obj_PEB_ExportEditingScen")
            Set m_ExportEditingScen = New obj_PEB_ExportEditingScen
        Case Else
            VBA.MsgBox "Непідтримуваний class таблиці редагування: '" & _
                submodeClassName & "'.", VBA.vbExclamation, _
                "PrsnlEventBuilder / Редагування РУХ / WORD"
            Exit Function
    End Select
    prsnlEvntBuilderCfgParser.Dispose
    Set prsnlEvntBuilderCfgParser = Nothing
    If Not m_ExportEditingScen.Initialize(m_Page, Me) Then
        Set m_ExportEditingScen = Nothing
        Exit Function
    End If
    private_TryEnsureExportEditingSubmode = True
End Function

Private Function private_TryCreateValidationSubmode( _
    ByRef outValidationSubmode As obj_PEB_MovementVldtnScen _
) As Boolean
    Dim prsnlEvntBuilderCfgParser As obj_PrsnlEvntBuilderCfgParser
    Dim submodeClassName As String

    Set outValidationSubmode = Nothing
    If m_ProfileConfigTable Is Nothing Then
        VBA.MsgBox "Відсутня конфігурація профілю для валідації.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Валідація"
        Exit Function
    End If
    Set prsnlEvntBuilderCfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not prsnlEvntBuilderCfgParser.Initialize( _
        m_ProfileConfigTable) Then Exit Function
    submodeClassName = VBA.Trim$(prsnlEvntBuilderCfgParser.GetOptionalValue( _
        VALIDATION_SUBMODE_CLASS_KEY, VBA.vbNullString))
    If VBA.Len(submodeClassName) = 0 Then
        VBA.MsgBox "Відсутній обов'язковий class валідації у ключі '" & _
            VALIDATION_SUBMODE_CLASS_KEY & "'.", VBA.vbExclamation, _
            "PrsnlEventBuilder / Валідація"
        Exit Function
    End If
    Select Case VBA.LCase$(submodeClassName)
        Case VBA.LCase$("obj_PEB_MovementVldtnScen")
            Set outValidationSubmode = New obj_PEB_MovementVldtnScen
        Case Else
            VBA.MsgBox "Непідтримуваний class валідації: '" & _
                submodeClassName & "'.", VBA.vbExclamation, _
                "PrsnlEventBuilder / Валідація"
            Exit Function
    End Select
    prsnlEvntBuilderCfgParser.Dispose
    Set prsnlEvntBuilderCfgParser = Nothing
    private_TryCreateValidationSubmode = True
End Function

Private Function private_TryPrepareMedicalReports( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim filterItems As Collection
    Dim emptyTables As Collection
    Dim prsnlEvntBuilderCfgParser As obj_PrsnlEvntBuilderCfgParser
    Dim viewClassName As String

    If m_ProfileConfigTable Is Nothing Then
        VBA.MsgBox "Відсутня конфігурація профілю для стройових записок.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Стройові записки"
        Exit Function
    End If

    If Not m_MedicalReportsScen Is Nothing Then m_MedicalReportsScen.Dispose
    Set m_MedicalReportsScen = Nothing
    Set prsnlEvntBuilderCfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not prsnlEvntBuilderCfgParser.Initialize( _
        m_ProfileConfigTable) Then Exit Function
    viewClassName = VBA.Trim$(prsnlEvntBuilderCfgParser.GetOptionalValue( _
        MEDICAL_REPORTS_VIEW_CLASS_KEY, VBA.vbNullString))
    If VBA.Len(viewClassName) = 0 Then
        VBA.MsgBox "Відсутній обов'язковий class підрежиму стройових " & _
            "записок у ключі '" & MEDICAL_REPORTS_VIEW_CLASS_KEY & "'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Стройові записки"
        Exit Function
    End If
    Select Case VBA.LCase$(VBA.Trim$(viewClassName))
        Case VBA.LCase$("obj_PEB_MedicalReportsScen")
            Set m_MedicalReportsScen = New obj_PEB_MedicalReportsScen
        Case Else
            VBA.MsgBox "Непідтримуваний class підрежиму стройових записок: '" & _
                viewClassName & "'.", VBA.vbExclamation, _
                "PrsnlEventBuilder / Стройові записки"
            Exit Function
    End Select
    prsnlEvntBuilderCfgParser.Dispose
    Set prsnlEvntBuilderCfgParser = Nothing

    If Not m_MedicalReportsScen.Initialize(m_ProfileConfigTable) Then Exit Function
    If Not m_MedicalReportsScen.TryGetFilterItems(filterItems) Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set emptyTables = New Collection
    If Not pageBase.RuntimeSources.SetItemsSource( _
        MEDICAL_REPORTS_FILTERS_RUNTIME_KEY, _
        filterItems, False) Then Exit Function
    If Not pageBase.RuntimeSources.SetItemsSource( _
        MEDICAL_REPORTS_RUNTIME_KEY, _
        emptyTables, notifyChange) Then Exit Function
    private_TryPrepareMedicalReports = True
End Function

Public Function SearchMedicalReports( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim filterValues As Object
    Dim reportTables As Collection

    If m_MedicalReportsScen Is Nothing Then
        If Not private_TryPrepareMedicalReports(False) Then Exit Function
    End If
    If Not private_TryReadMedicalReportFilters(filterValues) Then Exit Function

    ' Перед SQL-поиском освобождаем page-scoped engine PEB, чтобы два ADO
    ' connection не держали одновременно последний MedicalDaily-файл.
    private_DisposeCachedExporters
    If Not m_ExporterCfgDataProvider Is Nothing Then _
        m_ExporterCfgDataProvider.Dispose
    Set m_ExporterCfgDataProvider = Nothing
    ex_ExternalExcelSqlEngine.fn_ResetRuntimeCache

    If Not m_MedicalReportsScen.TryLoadTables( _
        filterValues, reportTables) Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetItemsSource( _
        MEDICAL_REPORTS_RUNTIME_KEY, reportTables, False) Then Exit Function
    If Not pageBase.TryReflowControl( _
        MEDICAL_REPORTS_CONTROL_NAME) Then Exit Function
    SearchMedicalReports = True
End Function

Private Function private_TryReadMedicalReportFilters( _
    ByRef outFilterValues As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim filterRange As Range
    Dim filterColumns As Collection
    Dim columnIndex As Long
    Dim cellValue As Variant
    Dim filterValue As String

    Set outFilterValues = Nothing
    If m_MedicalReportsScen Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange( _
        MEDICAL_REPORTS_FILTER_INPUTS_CONTAINER, filterRange) Then
        VBA.MsgBox "Не вдалося знайти рядок фільтрів стройових записок.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Стройові записки"
        Exit Function
    End If
    If Not m_MedicalReportsScen.TryGetFilterColumns( _
        filterColumns) Then Exit Function
    If filterRange.Columns.Count < filterColumns.Count Then
        VBA.MsgBox "Рядок фільтрів містить менше колонок, ніж налаштовано.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Стройові записки"
        Exit Function
    End If
    Set outFilterValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    For columnIndex = 1 To filterColumns.Count
        cellValue = filterRange.Cells(1, columnIndex).Value2
        filterValue = VBA.vbNullString
        If Not VBA.IsError(cellValue) And Not VBA.IsNull(cellValue) And _
            Not VBA.IsEmpty(cellValue) Then
            filterValue = VBA.Trim$(VBA.CStr(cellValue))
        End If
        outFilterValues(VBA.CStr(filterColumns.Item(columnIndex))) = _
            filterValue
    Next columnIndex
    private_TryReadMedicalReportFilters = True
End Function

Public Function OnAcceptOrderReferenceClick( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim numberOrDateText As String
    Dim orderYearText As String
    Dim resolvedOrderNo As String
    Dim resolvedOrderDate As Date
    Dim orderHistoryItems As Collection

    If m_Page Is Nothing Or m_ExportCommonData Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    ' Новая попытка всегда сначала инвалидирует прежнюю пару. Поэтому ошибка
    ' сопоставления не оставляет экспорту скрытое старое значение.
    private_ClearResolvedOrderPair
    numberOrDateText = private_TryReadManualOrderNoValue(pageBase, ws)
    orderYearText = private_TryReadManualOrderYearValue(pageBase, ws)
    If Not m_ExportCommonData.TryResolveOrderReference( _
        numberOrDateText, orderYearText, resolvedOrderNo, resolvedOrderDate) Then
        If Not private_EnsureOrderHistoryRuntime(False) Then Exit Function
        If Not rt_PageManager.fn_RenderPage( _
            m_Page, "prsnlevntbuilder:reject-order") Then Exit Function
        OnAcceptOrderReferenceClick = True
        Exit Function
    End If

    m_ResolvedOrderNo = resolvedOrderNo
    m_ResolvedOrderDate = resolvedOrderDate
    m_HasResolvedOrderPair = True
    If Not private_TryBuildOrderHistoryItems(orderHistoryItems) Then
        private_ClearResolvedOrderPair
        If Not private_EnsureOrderHistoryRuntime(False) Then Exit Function
        If Not rt_PageManager.fn_RenderPage( _
            m_Page, "prsnlevntbuilder:reject-order-history") Then Exit Function
        OnAcceptOrderReferenceClick = True
        Exit Function
    End If
    Set m_OrderHistoryItems = orderHistoryItems
    If Not private_EnsureOrderHistoryRuntime(False) Then Exit Function
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "prsnlevntbuilder:accept-order") Then Exit Function
    rt_Messaging.fn_ShowStatusBarSuccess _
        "Наказ № " & resolvedOrderNo & " від " & _
        VBA.Format$(resolvedOrderDate, "dd.mm.yyyy") & " прийнято.", 4
    OnAcceptOrderReferenceClick = True
End Function

Public Property Get IsWordPreviewExportMode() As Boolean
    IsWordPreviewExportMode = m_IsWordPreviewExportMode
End Property

Public Property Get AreWordBookmarksShownAsMarkers() As Boolean
    AreWordBookmarksShownAsMarkers = m_AreWordBookmarksShownAsMarkers
End Property

Public Property Get CanEnableWordPreviewExportMode() As Boolean
    CanEnableWordPreviewExportMode = _
        (VBA.Len(VBA.Trim$(m_WordExportPreviewText)) > 0 And _
         Not m_IsWordPreviewExportMode)
End Property

Public Function ResetWordPreviewExportMode( _
    Optional ByVal renderNow As Boolean = False _
) As Boolean
    Dim editedPreviewText As String
    Dim wasPreviewExportMode As Boolean

    wasPreviewExportMode = m_IsWordPreviewExportMode
    ' Перед выходом из edit mode сохраняем фактический текст Banner:
    ' пользователь мог изменить его непосредственно в ячейках листа.
    If wasPreviewExportMode Then
        If Not private_TryReadRenderedWordPreview(editedPreviewText) Then Exit Function
        If VBA.Len(VBA.Trim$(editedPreviewText)) > 0 Then
            m_WordExportPreviewText = editedPreviewText
        End If
    End If
    m_IsWordPreviewExportMode = False
    ' Если режим уже был выключен, layout и оформление менять не требуется.
    ' Это также не допускает второго reflow в ветках, которые после сброса
    ' вызывают ResetWordPreviewExportMode(True).
    If Not wasPreviewExportMode Then
        ResetWordPreviewExportMode = True
        Exit Function
    End If
    If m_Page Is Nothing Then Exit Function
    ' При фактическом сбросе обновление панели обязательно даже для False:
    ' это значение передают обработчики четырёх lookup-полей перед поиском,
    ' который перерисовывает только кандидатов и не затрагивает WORD actions.
    ResetWordPreviewExportMode = private_TryRenderWordPreviewMode( _
        "prsnlevntbuilder:word-preview-mode-reset")
End Function

Public Function ToggleWordPreviewExportMode( _
    Optional ByVal ignored As Variant _
) As Boolean
    If VBA.Len(VBA.Trim$(m_WordExportPreviewText)) = 0 Then
        VBA.MsgBox "Спочатку сформуйте WORD preview.", _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD preview"
        Exit Function
    End If
    If m_IsWordPreviewExportMode Then
        ToggleWordPreviewExportMode = Me.ResetWordPreviewExportMode(True)
        Exit Function
    End If
    m_IsWordPreviewExportMode = True
    If m_Page Is Nothing Then Exit Function
    ToggleWordPreviewExportMode = private_TryRenderWordPreviewMode( _
        "prsnlevntbuilder:word-preview-mode-toggled")
End Function

Private Function private_TryRenderWordPreviewMode( _
    ByVal reasonText As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim previousSuppressLookupSearch As Boolean

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    previousSuppressLookupSearch = m_SuppressLookupSearch
    On Error GoTo RestoreSuppression
    ' Сам Banner не рендерим повторно: его посимвольные цвета существуют только
    ' в Excel Range и потерялись бы при повторной записи обычной строки.
    ' Перестраиваем лишь колонку действий, затем меняем фон и рамку исходного
    ' диапазона preview без обращения к свойствам Font.
    m_SuppressLookupSearch = True
    ' Взаимоисключающие кнопки занимают один layout slot. Удаляем Shape ушедшего
    ' состояния явно: retained reflow не обязан удалять Shape collapsed-контрола,
    ' и иначе активное оформление остаётся поверх обычной кнопки.
    If m_IsWordPreviewExportMode Then
        private_DeleteGeneratedShapeIfExists _
            pageBase.Worksheet, WORD_EXPORT_PREVIEW_BUTTON_SHAPE_NAME
    Else
        private_DeleteGeneratedShapeIfExists _
            pageBase.Worksheet, WORD_EXPORT_ACTIVE_BUTTON_SHAPE_NAME
    End If
    private_TryRenderWordPreviewMode = pageBase.TryReflowLayoutContainer( _
        WORD_EXPORT_ACTIONS_CONTAINER_NAME)
    If private_TryRenderWordPreviewMode Then
        private_TryRenderWordPreviewMode = _
            private_TryApplyWordPreviewModeAppearance(pageBase.Worksheet)
    End If
    If Not private_TryRenderWordPreviewMode Then
        VBA.MsgBox _
            "PrototypeNew: failed to partially render WORD preview panel '" & _
            WORD_EXPORT_PANEL_CONTAINER_NAME & "' (" & reasonText & ").", _
            VBA.vbExclamation, _
            "PrototypeNew / WORD preview"
    End If

RestoreSuppression:
    m_SuppressLookupSearch = previousSuppressLookupSearch
End Function

Private Function private_TryApplyWordPreviewModeAppearance( _
    ByVal ws As Worksheet _
) As Boolean
    Dim messageRange As Range
    Dim columnScope As Range
    Dim backgroundColor As Long
    Dim borderColor As Long

    If ws Is Nothing Then Exit Function
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, _
        "banner", _
        WORD_EXPORT_PREVIEW_CONTROL_NAME, _
        "message", _
        messageRange, _
        columnScope) Then Exit Function

    If m_IsWordPreviewExportMode Then
        backgroundColor = VBA.RGB(140, 57, 15)
        borderColor = VBA.RGB(52, 211, 153)
    Else
        backgroundColor = VBA.RGB(37, 42, 49)
        borderColor = VBA.RGB(0, 0, 0)
    End If

    On Error GoTo EH
    messageRange.Interior.Color = backgroundColor
    messageRange.Borders.Color = borderColor
    messageRange.Borders.Weight = xlThin
    private_TryApplyWordPreviewModeAppearance = True
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "PrsnlEventBuilder: failed to update WORD preview mode appearance: " & _
        Err.Description
#End If
End Function

Private Sub private_DeleteGeneratedShapeIfExists( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String _
)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub private_DeleteBottomWorkspaceCommandShapes(ByVal ws As Worksheet)
    ' У обычного и активного состояния одна и та же позиция в layout. Перед
    ' частичным reflow удаляем старые Shapes, чтобы collapsed-вариант не
    ' оставался под вновь созданной кнопкой.
    private_DeleteGeneratedShapeIfExists ws, "btn_RefreshWorkspace"
    private_DeleteGeneratedShapeIfExists ws, "btn_ShowExportEditing"
    private_DeleteGeneratedShapeIfExists ws, "btn_ShowExportEditingActive"
    private_DeleteGeneratedShapeIfExists ws, "btn_ShowValidation"
    private_DeleteGeneratedShapeIfExists ws, "btn_ShowValidationActive"
    private_DeleteGeneratedShapeIfExists ws, "btn_ShowMedicalReports"
    private_DeleteGeneratedShapeIfExists ws, "btn_ShowMedicalReportsActive"
End Sub

Public Property Get IsLookupEnabled() As Boolean
    IsLookupEnabled = m_IsLookupEnabled
End Property

Public Property Get IsTemporaryPersonnelEnabled() As Boolean
    IsTemporaryPersonnelEnabled = m_IsTemporaryPersonnelEnabled
End Property

Public Function ToggleTemporaryPersonnel() As Boolean
    m_IsTemporaryPersonnelEnabled = Not m_IsTemporaryPersonnelEnabled
    If Not ex_ControlRefreshRuntime.fn_TryRefreshStaticControl( _
        TEMPORARY_PERSONNEL_CONTROL_NAME) Then
        m_IsTemporaryPersonnelEnabled = Not m_IsTemporaryPersonnelEnabled
        VBA.MsgBox "PrototypeNew: failed to refresh the temporary personnel button.", _
            VBA.vbExclamation, "PrsnlEventBuilder / EntityLookup"
        Exit Function
    End If
    ToggleTemporaryPersonnel = True
End Function

Public Property Get IsMovementValidationEnabled() As Boolean
    IsMovementValidationEnabled = m_IsMovementValidationEnabled
End Property

Public Property Get IsWordValidationEnabled() As Boolean
    IsWordValidationEnabled = m_IsWordValidationEnabled
End Property

Public Property Get IsMovementHistoryEnabled() As Boolean
    IsMovementHistoryEnabled = m_IsMovementHistoryEnabled
End Property

Public Function ToggleMovementHistory() As Boolean
    Dim previousEnabled As Boolean
    Dim previousEnableEvents As Boolean
    Dim renderSucceeded As Boolean

    If m_Page Is Nothing Then Exit Function
    previousEnabled = m_IsMovementHistoryEnabled
    m_IsMovementHistoryEnabled = Not previousEnabled
    If Not private_RegisterMovementHistoryTable(False) Then
        m_IsMovementHistoryEnabled = previousEnabled
        Exit Function
    End If

    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    renderSucceeded = rt_PageManager.fn_RenderPage( _
        m_Page, "prsnlevntbuilder:toggle-movement-history")

Cleanup:
    Application.EnableEvents = previousEnableEvents
    If Not renderSucceeded Then
        m_IsMovementHistoryEnabled = previousEnabled
        Exit Function
    End If
    ToggleMovementHistory = True
    Exit Function

EH:
    renderSucceeded = False
    Resume Cleanup
End Function

Public Function ToggleMovementValidation() As Boolean
    ToggleMovementValidation = private_ToggleExporterValidation( _
        m_IsMovementValidationEnabled, VALIDATE_MOVEMENT_CONTROL_NAME, "Movement")
End Function

Public Function ToggleWordValidation() As Boolean
    ToggleWordValidation = private_ToggleExporterValidation( _
        m_IsWordValidationEnabled, VALIDATE_WORD_CONTROL_NAME, "WORD")
End Function

Private Function private_ToggleExporterValidation( _
    ByRef validationEnabled As Boolean, _
    ByVal controlName As String, _
    ByVal exporterCaption As String _
) As Boolean
    validationEnabled = Not validationEnabled
    If Not ex_ControlRefreshRuntime.fn_TryRefreshStaticControl(controlName) Then
        validationEnabled = Not validationEnabled
        VBA.MsgBox "PrototypeNew: failed to refresh the " & exporterCaption & _
            " validation button.", VBA.vbExclamation, "PrototypeNew / Data export"
        Exit Function
    End If

    private_ToggleExporterValidation = True
    If validationEnabled Then
        rt_Messaging.fn_ShowStatusBarSuccess exporterCaption & " validation enabled", 3
    Else
        rt_Messaging.fn_ShowStatusBarWarning exporterCaption & " validation disabled", 3
    End If
End Function

Public Function ToggleLookupEnabled() As Boolean
    m_IsLookupEnabled = Not m_IsLookupEnabled
    ' Это только visual-state кнопки: страница и input-ячейки
    ' не должны перерисовываться и инициировать lookup callbacks.
    If Not ex_ControlRefreshRuntime.fn_TryRefreshStaticControl(LOOKUP_MODE_CONTROL_NAME) Then
        m_IsLookupEnabled = Not m_IsLookupEnabled
        VBA.MsgBox "PrototypeNew: failed to refresh the Lookup button.", VBA.vbExclamation, "PrototypeNew / EntityLookup"
        Exit Function
    End If

    ToggleLookupEnabled = True
    If m_IsLookupEnabled Then
        rt_Messaging.fn_ShowStatusBarSuccess "Lookup enabled", 3
    Else
        rt_Messaging.fn_ShowStatusBarWarning "Lookup disabled", 3
    End If
End Function

' Очищает поля, значение которых зависит от изменённого Lookup-ключа.
' Используется единый Select Case, чтобы новые связанные группы добавлялись
' без отдельных обработчиков Worksheet_Change для каждого профиля формы.
Public Function ClearDependentDraftFields(ByVal lookupKey As String) As Boolean
    Dim dependentAliases As Object
    Dim sectionText As String

    Set dependentAliases = ex_Helpers.fn_CreateDictionaryTextCompare()
    sectionText = VBA.Trim$(m_SelectedMainProfile)
    If VBA.Len(sectionText) = 0 Then
        If Not private_TryGetSelectedProfile(sectionText) Then Exit Function
    End If

    Select Case VBA.LCase$(VBA.Trim$(lookupKey))
        Case VBA.LCase$(HOSPITAL_LOOKUP_KEY)
            dependentAliases(DRAFT_ALIAS_HOSPITAL_SHORT) = True

        Case VBA.LCase$(TO_HOSPITAL_LOOKUP_KEY)
            dependentAliases(DRAFT_ALIAS_TO_HOSPITAL_SHORT) = True

        Case "op_fio"
            ' Правила намеренно выбираются по основной секции. Если сейчас открыт
            ' meta-профиль, m_SelectedMainProfile всё равно указывает на событие,
            ' к которому относится редактируемая форма.
            If Not private_AppendFioDependentAliases(sectionText, dependentAliases) Then Exit Function

        Case "op_commander"
            If Not private_AppendCommanderDependentAliases(sectionText, dependentAliases) Then Exit Function

        Case Else
            ClearDependentDraftFields = True
            Exit Function
    End Select

    If Not private_ClearVisibleDraftFieldsByAlias(dependentAliases) Then Exit Function
    If VBA.StrComp(VBA.Trim$(lookupKey), FIO_LOOKUP_KEY, VBA.vbTextCompare) = 0 Then
        If Not private_TryApplyMovementEnrichedFieldsAppearance() Then Exit Function
    End If
    ClearDependentDraftFields = True
End Function

Public Function ResetReporterTvoStateForLookupChange( _
    ByVal lookupKey As String _
) As Boolean
    ' Выбранная вручную должность ТВО относится только к текущему рапортующему.
    ' При изменении ФІО старый собственный код нельзя использовать как базовый
    ' для нового человека, даже если Lookup сейчас выключен.
    If VBA.StrComp( _
        VBA.Trim$(lookupKey), COMMANDER_LOOKUP_KEY, _
        VBA.vbTextCompare) <> 0 Then
        ResetReporterTvoStateForLookupChange = True
        Exit Function
    End If

    m_ReportOwnPositionCode = VBA.vbNullString
    m_DraftReportIsTvo = False
    m_IsReporterTvoCandidatesActive = False
    ResetReporterTvoStateForLookupChange = _
        private_TryApplyReporterPositionCodeAppearance()
End Function

Private Function private_AppendFioDependentAliases( _
    ByVal sectionText As String, _
    ByVal dependentAliases As Object _
) As Boolean
    Dim sectionKey As String

    If dependentAliases Is Nothing Then Exit Function
    If m_Data Is Nothing Then Exit Function
    sectionKey = private_NormalizeText(sectionText)

    If m_Data.IsMovementMirrorTransferSectionType(sectionText) Then
        private_AddStatusChangeFioDependentAliases dependentAliases
        private_AppendFioDependentAliases = True
        Exit Function
    End If

    Select Case sectionKey
        ' Сейчас все основные события получают персональные данные из одного
        ' op_FIO. Отдельный Select Case оставляет правила секционными: когда для
        ' конкретного события набор изменится, его следует вынести в отдельный Case.
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatment)
            private_AddStandardFioDependentAliases dependentAliases
            dependentAliases(DRAFT_ALIAS_HOSPITAL) = True
            dependentAliases(DRAFT_ALIAS_HOSPITAL_SHORT) = True

        Case private_NormalizeText(m_Data.SectionTypeCloseFromMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAmbulatoryVlk), _
             private_NormalizeText(m_Data.SectionTypeCloseFromExternalVlk)
            private_AddStandardFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToTreatment), _
             private_NormalizeText(m_Data.SectionTypeToMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeToAmbulatoryVlk)
            private_AddStandardFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeCloseFromAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromFamilyVacation)
            private_AddCloseVacationFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeCloseFromBusinessTrip)
            private_AddCloseBusinessTripFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToAnnualVacationPart), _
             private_NormalizeText(m_Data.SectionTypeToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeToMaternityLeave)
            private_AddToVacationFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToTreatmentVacation)
            private_AddToTreatmentVacationFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentToExternalVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToVlk)
            private_AddStandardFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferExternalVlkToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferExternalVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeToBusinessTrip), _
             private_NormalizeText(m_Data.SectionTypeToBusinessTripSzch)
            private_AddStandardFioDependentAliases dependentAliases

        Case Else
            VBA.MsgBox "PrototypeNew: dependent-field rules are missing for section '" & _
                sectionText & "' and lookup 'op_FIO'.", _
                VBA.vbExclamation, "PrototypeNew / PrsnlEvntBuilder"
            Exit Function
    End Select

    private_AppendFioDependentAliases = True
End Function

Private Function private_AppendCommanderDependentAliases( _
    ByVal sectionText As String, _
    ByVal dependentAliases As Object _
) As Boolean
    Dim sectionKey As String

    If dependentAliases Is Nothing Then Exit Function
    If m_Data Is Nothing Then Exit Function
    sectionKey = private_NormalizeText(sectionText)

    Select Case sectionKey
        ' Набор можно разделять по секциям независимо от правил op_FIO.
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatment), _
             private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAmbulatoryVlk), _
             private_NormalizeText(m_Data.SectionTypeCloseFromBusinessTrip), _
             private_NormalizeText(m_Data.SectionTypeCloseFromExternalVlk)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToTreatment), _
             private_NormalizeText(m_Data.SectionTypeToMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeToAnnualVacationPart), _
             private_NormalizeText(m_Data.SectionTypeToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeToMaternityLeave), _
             private_NormalizeText(m_Data.SectionTypeToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeToAmbulatoryVlk)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentToExternalVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToVlk)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferExternalVlkToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferExternalVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeToBusinessTrip), _
             private_NormalizeText(m_Data.SectionTypeToBusinessTripSzch)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case Else
            VBA.MsgBox "PrototypeNew: dependent-field rules are missing for section '" & _
                sectionText & "' and lookup 'op_Commander'.", _
                VBA.vbExclamation, "PrototypeNew / PrsnlEvntBuilder"
            Exit Function
    End Select

    private_AppendCommanderDependentAliases = True
End Function

Private Sub private_AddStandardFioDependentAliases(ByVal dependentAliases As Object)
    dependentAliases(DRAFT_ALIAS_RANK) = True
    dependentAliases(DRAFT_ALIAS_IPN) = True
    dependentAliases(DRAFT_ALIAS_POSITION_CODE) = True
    dependentAliases(DRAFT_ALIAS_POSITION_NAME) = True
    ' Эти реквизиты относятся к выбранному человеку. Метод очистки применяет
    ' aliases только к видимым полям, поэтому секции без документа не меняются.
    dependentAliases(DRAFT_ALIAS_DOC_NO) = True
    dependentAliases(DRAFT_ALIAS_DOC_DATE) = True
    dependentAliases(DRAFT_ALIAS_VLK_NO) = True
    dependentAliases(DRAFT_ALIAS_VLK_DATE) = True
End Sub

Private Sub private_AddStatusChangeFioDependentAliases( _
    ByVal dependentAliases As Object _
)
    ' Реквизиты смены статуса относятся к конкретному человеку и не должны
    ' переживать замену ПІБ даже при совпадении профиля формы.
    private_AddStandardFioDependentAliases dependentAliases
    dependentAliases(DRAFT_ALIAS_HOSPITAL) = True
    dependentAliases(DRAFT_ALIAS_HOSPITAL_SHORT) = True
    dependentAliases(DRAFT_ALIAS_INCOMING_NO) = True
    dependentAliases(DRAFT_ALIAS_INCOMING_DATE) = True
    dependentAliases(DRAFT_ALIAS_DOC_NO) = True
    dependentAliases(DRAFT_ALIAS_DOC_DATE) = True
    dependentAliases(DRAFT_ALIAS_VLK_NO) = True
    dependentAliases(DRAFT_ALIAS_VLK_DATE) = True
    dependentAliases(DRAFT_ALIAS_DATE_FROM) = True
End Sub

Private Sub private_AddToVacationFioDependentAliases(ByVal dependentAliases As Object)
    ' При смене человека данные оформляемого отпуска больше не относятся
    ' к выбранной строке, поэтому очищаем их вместе с персональными полями.
    private_AddStandardFioDependentAliases dependentAliases
    dependentAliases(DRAFT_ALIAS_DESTINATION) = True
    dependentAliases(DRAFT_ALIAS_INCOMING_NO) = True
    dependentAliases(DRAFT_ALIAS_INCOMING_DATE) = True
    dependentAliases(DRAFT_ALIAS_DURATION_DAYS) = True
    dependentAliases(DRAFT_ALIAS_VACATION_TICKET_NO) = True
End Sub

Private Sub private_AddCloseVacationFioDependentAliases(ByVal dependentAliases As Object)
    ' Для возврата из отпуска очищаются только реквизиты отпускного билета.
    private_AddStandardFioDependentAliases dependentAliases
    dependentAliases(DRAFT_ALIAS_VACATION_TICKET_NO) = True
    dependentAliases(DRAFT_ALIAS_VACATION_TICKET_DATE) = True
End Sub

Private Sub private_AddCloseBusinessTripFioDependentAliases( _
    ByVal dependentAliases As Object _
)
    private_AddCloseVacationFioDependentAliases dependentAliases
    dependentAliases(DRAFT_ALIAS_DESTINATION) = True
End Sub

Private Sub private_AddToTreatmentVacationFioDependentAliases(ByVal dependentAliases As Object)
    private_AddToVacationFioDependentAliases dependentAliases
    dependentAliases(DRAFT_ALIAS_VLK_NO) = True
    dependentAliases(DRAFT_ALIAS_VLK_DATE) = True
End Sub

Private Sub private_AddStandardCommanderDependentAliases(ByVal dependentAliases As Object)
    dependentAliases(DRAFT_ALIAS_REPORT_RANK) = True
    dependentAliases(DRAFT_ALIAS_REPORT_POSITION_CODE) = True
    dependentAliases(DRAFT_ALIAS_INCOMING_NO) = True
End Sub

Public Function UpdateDataFromConfigTable(ByVal configTable As obj_ConfigTable) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    If configTable Is Nothing Then Exit Function
    UpdateDataFromConfigTable = m_LookupFeature.UpdateDataFromConfigTable(configTable)
    If Not UpdateDataFromConfigTable Then Exit Function
    If Not private_TryUpdateProfilesProvider(configTable) Then
        UpdateDataFromConfigTable = False
        Exit Function
    End If
    If Not private_TryUpdateExportSettings(configTable) Then Exit Function
    If Not private_RegisterExportFormTables(False) Then Exit Function
    If Not private_RegisterMovementHistoryTable(False) Then Exit Function
    If Not private_RegisterExportedEventMenus(False, False) Then Exit Function
    If Not private_EnsureMedicalReportsRuntime(False) Then Exit Function
    If Not private_EnsureHotkeyRows(False) Then Exit Function
    UpdateDataFromConfigTable = True
End Function

Private Function private_ClearRenderedExportedEvents() As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim emptyTableItems As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set emptyTableItems = New Collection
    If Not runtimeSources.RemoveItemsSource( _
        VBA.LCase$(MOVEMENT_EVENTS_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource( _
        VBA.LCase$(MOVEMENT_EVENTS_RUNTIME_KEY), _
        emptyTableItems, False) Then Exit Function
    Set m_MovementEventIds = New Collection
    Set m_MovementEventCaptions = New Collection
    Set m_WordEventIds = New Collection
    Set m_WordEventCaptions = New Collection
    m_AreExportedEventsShown = False
    private_ClearRenderedExportedEvents = True
End Function

Public Function PrepareRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.PrepareLookupRuntime(notifyChange) Then Exit Function
    If Not private_RegisterProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterAdditionalProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterMetaProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterExportFormTables(notifyChange) Then Exit Function
    If Not private_RegisterMovementHistoryTable(notifyChange) Then Exit Function
    If Not private_RegisterExportedEventMenus(False, notifyChange) Then Exit Function
    If Not private_RegisterDummyTables(notifyChange) Then Exit Function
    If Not private_EnsureOrderHistoryRuntime(notifyChange) Then Exit Function
    If Not private_EnsureValidationResultsRuntime(notifyChange) Then Exit Function
    If Not private_EnsureMedicalReportsRuntime(notifyChange) Then Exit Function
    If Not private_EnsureHotkeyRows(notifyChange) Then Exit Function
    PrepareRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    m_IsReporterTvoCandidatesActive = False
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
    Dim cellAddress As String
    Dim shouldRefreshMovementHistory As Boolean
    Dim restoreHandled As Boolean

    On Error GoTo EH

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
    cellAddress = targetCell.Address(False, False)

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:hotkey-action:start action='" & private_EscapeForLog(actionText) & "' cell='" & private_EscapeForLog(cellAddress) & "'"
#End If

    ' This dedicated action also starts with "Export ", so handle it before
    ' generic "Export <alias>" actions. The button calls the same method.
    If VBA.StrComp(actionText, HOTKEY_EXPORT_TO_WORD, VBA.vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:hotkey-action:word-export-branch"
#End If
        Call private_TryExportWordToDocument
        RuntimeHandleHotkeyAction = True
        Exit Function
    End If

    If VBA.StrComp(actionText, HOTKEY_EXPORT_MOVEMENT_WORD, VBA.vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:hotkey-action:movement-word-branch"
#End If
        If Not private_TryExportMovementAndWord() Then
            RuntimeHandleHotkeyAction = True
            Exit Function
        End If
        RuntimeHandleHotkeyAction = True
        Exit Function
    End If

    If private_IsExportAction(actionText) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:hotkey-action:export-branch action='" & private_EscapeForLog(actionText) & "'"
#End If
        If Not private_TryExportDraftByAction(actionText) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "prsnlevntbuilder:hotkey-action:export-branch-failed action='" & private_EscapeForLog(actionText) & "'"
#End If
            RuntimeHandleHotkeyAction = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:hotkey-action:export-branch-done action='" & private_EscapeForLog(actionText) & "'"
#End If
        RuntimeHandleHotkeyAction = True
        Exit Function
    End If

    ' Page-specific actions branch by stable action ids and read current sheet state.
    Select Case VBA.LCase$(actionText)
        Case VBA.LCase$(HOTKEY_REPORT_TVO_CANDIDATES)
            If Not private_TryShowReporterTvoCandidates(targetCell) Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If
            RuntimeHandleHotkeyAction = True
            Exit Function

        Case VBA.LCase$(HOTKEY_ACCEPT_CANDIDATE_ROW)
            If Not private_TryRestoreSelectedDeletedEvent( _
                targetCell, restoreHandled) Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If
            If restoreHandled Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If
            ' Любой выбранный Lookup-кандидат меняет данные, из которых был
            ' построен WORD preview: ФИО, больницу, рапортующего или документ.
            If Not Me.ResetWordPreviewExportMode(False) Then Exit Function
            If m_IsMovementHistoryEnabled Then
                ' Movement относится только к основному человеку события.
                ' Hospital/Commander candidates тоже заполняют draft-форму
                ' через этот hotkey, но не должны обновлять историю по старому
                ' или не относящемуся к ним ИПН.
                shouldRefreshMovementHistory = _
                    private_IsFioCandidatesContext()
                ' Сначала применяем выбранного Lookup-кандидата к draft-форме.
                ' Историю Movement запрашиваем уже по записанному в форму ИПН,
                ' чтобы preview и экспортная строка относились к одному человеку.
                If Not private_TryAcceptCandidateRowFromSelection(targetCell) Then
                    RuntimeHandleHotkeyAction = True
                    Exit Function
                End If
                If shouldRefreshMovementHistory Then
                    If Not private_TryRefreshMovementHistory() Then
                        RuntimeHandleHotkeyAction = True
                        Exit Function
                    End If
                    ' Partial reflow переносит выделенный диапазон кандидата,
                    ' но Excel назначает ActiveCell его левой верхней ячейке.
                    ' Возвращаем активную колонку исходной ячейки уже внутри
                    ' диапазона с обновлёнными адресами.
                    If Not private_TryActivateColumnWithinSelection( _
                        targetCell.Column) Then
                        RuntimeHandleHotkeyAction = True
                        Exit Function
                    End If
                End If
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If
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

        Case VBA.LCase$(HOTKEY_CLEAR_EXPORT_FORM)
            If Not Me.RuntimeClearExportFormAndCandidates() Then
                RuntimeHandleHotkeyAction = True
                Exit Function
            End If
            RuntimeHandleHotkeyAction = True
            Exit Function

        Case VBA.LCase$(HOTKEY_DELETE_EXPORTED_EVENT)
            Call private_TryDeleteSelectedExportedEvent(targetCell)
            RuntimeHandleHotkeyAction = True
            Exit Function

        Case Else
            Exit Function
    End Select

    rt_Messaging.fn_ShowStatusBarSuccess actionText & ": " & targetCell.Address(False, False) & " = '" & cellValue & "'", 3
    RuntimeHandleHotkeyAction = True
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "prsnlevntbuilder:hotkey-action:error action='" & private_EscapeForLog(actionText) & "' cell='" & private_EscapeForLog(cellAddress) & "' errNo=" & VBA.CStr(Err.Number) & " err='" & private_EscapeForLog(Err.Description) & "'"
#End If
End Function

Private Function private_TryDeleteSelectedExportedEvent( _
    ByVal targetCell As Range _
) As Boolean
    Dim eventId As String
    Dim eventCaption As String
    Dim isMovementEvent As Boolean

    If Not private_TryResolveSelectedCombinedExportedEvent( _
        targetCell, isMovementEvent, eventId, eventCaption) Then
        rt_Messaging.fn_ShowStatusBarWarning _
            "Выберите строку Movement или WORD events и нажмите Ctrl+Delete.", 4
        Exit Function
    End If
    If VBA.StrComp(eventId, "NONE", VBA.vbTextCompare) = 0 Then
        If isMovementEvent Then
            rt_Messaging.fn_ShowStatusBarWarning _
                "В этой строке нет события Movement.", 4
        Else
            rt_Messaging.fn_ShowStatusBarWarning _
                "В этой строке нет события WORD.", 4
        End If
        private_TryDeleteSelectedExportedEvent = True
        Exit Function
    End If
    If private_IsDeletedExportSnapshot(isMovementEvent, eventId) Then
        rt_Messaging.fn_ShowStatusBarWarning _
            "Событие уже удалено. Выберите его красную часть строки и нажмите Ctrl+Enter для восстановления.", 5
        private_TryDeleteSelectedExportedEvent = True
        Exit Function
    End If

    If isMovementEvent Then
        private_TryDeleteSelectedExportedEvent = Me.OnMovementEventClick(eventId)
    Else
        private_TryDeleteSelectedExportedEvent = Me.OnWordEventClick(eventId)
    End If
End Function

Private Function private_StoreDeletedExportSnapshot( _
    ByVal isMovementEvent As Boolean, _
    ByVal eventId As String, _
    ByVal snapshot As Object _
) As Boolean
    Dim snapshotKey As String
    Dim eventIds As Collection
    Dim eventCaptions As Collection
    Dim eventIndex As Long
    Dim relatedSnapshot As Object
    Dim aliasSnapshot As Object
    Dim relatedEventId As String
    Dim aliasKey As String

    If snapshot Is Nothing Then Exit Function
    If m_DeletedExportSnapshots Is Nothing Then
        Set m_DeletedExportSnapshots = VBA.CreateObject("Scripting.Dictionary")
        m_DeletedExportSnapshots.CompareMode = VBA.vbTextCompare
    End If
    snapshotKey = private_DeletedExportSnapshotKey(isMovementEvent, eventId)
    If Not isMovementEvent Then
        If Not private_TranslateWordSnapshotPositions( _
            snapshot, False, VBA.vbNullString) Then Exit Function
    End If
    If isMovementEvent Then
        Set eventIds = m_MovementEventIds
        Set eventCaptions = m_MovementEventCaptions
    Else
        Set eventIds = m_WordEventIds
        Set eventCaptions = m_WordEventCaptions
    End If
    If Not eventIds Is Nothing And Not eventCaptions Is Nothing Then
        For eventIndex = 1 To eventIds.Count
            If VBA.StrComp(VBA.CStr(eventIds.Item(eventIndex)), _
                eventId, VBA.vbBinaryCompare) = 0 Then
                snapshot("DisplayCaption") = VBA.CStr(eventCaptions.Item(eventIndex))
                snapshot("OptionCaption") = private_BuildDeletedEventOptionCaption( _
                    isMovementEvent, eventIndex, eventId)
                If VBA.Len(VBA.CStr(snapshot("OptionCaption"))) = 0 Then
                    VBA.MsgBox "Не удалось сохранить структуру удаляемого " & _
                        "события '" & eventId & "'. Удаление остановлено.", _
                        VBA.vbExclamation, _
                        "PrsnlEventBuilder / exported events"
                    Exit Function
                End If
                Exit For
            End If
        Next eventIndex
    End If
    If m_DeletedExportSnapshots.Exists(snapshotKey) Then _
        m_DeletedExportSnapshots.Remove snapshotKey
    m_DeletedExportSnapshots.Add snapshotKey, snapshot
    If isMovementEvent And snapshot.Exists("RelatedSnapshot") Then
        Set relatedSnapshot = snapshot("RelatedSnapshot")
        relatedEventId = VBA.CStr(relatedSnapshot("EventId"))
        aliasKey = private_DeletedExportSnapshotKey(True, relatedEventId)
        Set aliasSnapshot = VBA.CreateObject("Scripting.Dictionary")
        aliasSnapshot("Kind") = "Movement"
        aliasSnapshot("EventId") = relatedEventId
        aliasSnapshot("IsAlias") = True
        aliasSnapshot.Add "RootSnapshot", snapshot
        For eventIndex = 1 To m_MovementEventIds.Count
            If VBA.StrComp(VBA.CStr(m_MovementEventIds.Item(eventIndex)), _
                relatedEventId, VBA.vbBinaryCompare) = 0 Then
                aliasSnapshot("DisplayCaption") = _
                    VBA.CStr(m_MovementEventCaptions.Item(eventIndex))
                aliasSnapshot("OptionCaption") = _
                    private_BuildDeletedEventOptionCaption( _
                        True, eventIndex, relatedEventId)
                Exit For
            End If
        Next eventIndex
        ' У объединённой строки смены статуса связанный eventId намеренно не
        ' присутствует в UI, поэтому отдельный alias для него не требуется.
        If aliasSnapshot.Exists("OptionCaption") Then
            If m_DeletedExportSnapshots.Exists(aliasKey) Then _
                m_DeletedExportSnapshots.Remove aliasKey
            m_DeletedExportSnapshots.Add aliasKey, aliasSnapshot
        End If
    End If
    private_StoreDeletedExportSnapshot = True
End Function

Private Function private_TranslateWordSnapshotPositions( _
    ByVal changedSnapshot As Object, _
    ByVal isRestore As Boolean, _
    ByVal excludedSnapshotKey As String _
) As Boolean
    Dim snapshotKey As Variant
    Dim existingSnapshot As Object
    Dim changedStart As Long
    Dim changedLength As Long
    Dim delta As Long

    If changedSnapshot Is Nothing Then Exit Function
    If Not changedSnapshot.Exists("TargetPath") Or _
        Not changedSnapshot.Exists("RangeStart") Or _
        Not changedSnapshot.Exists("RangeLength") Then Exit Function
    changedStart = VBA.CLng(changedSnapshot("RangeStart"))
    changedLength = VBA.CLng(changedSnapshot("RangeLength"))
    If isRestore Then
        delta = changedLength
    Else
        delta = -changedLength
    End If
    If m_DeletedExportSnapshots Is Nothing Then
        private_TranslateWordSnapshotPositions = True
        Exit Function
    End If
    For Each snapshotKey In m_DeletedExportSnapshots.Keys
        If VBA.StrComp(VBA.CStr(snapshotKey), excludedSnapshotKey, _
            VBA.vbBinaryCompare) = 0 Then GoTo ContinueSnapshot
        Set existingSnapshot = m_DeletedExportSnapshots(snapshotKey)
        If existingSnapshot.Exists("TargetPath") And _
            existingSnapshot.Exists("RangeStart") Then
            If VBA.StrComp(VBA.CStr(existingSnapshot("TargetPath")), _
                VBA.CStr(changedSnapshot("TargetPath")), VBA.vbTextCompare) = 0 Then
                If VBA.CLng(existingSnapshot("RangeStart")) >= changedStart Then _
                    existingSnapshot("RangeStart") = _
                        VBA.CLng(existingSnapshot("RangeStart")) + delta
            End If
        End If
ContinueSnapshot:
    Next snapshotKey
    private_TranslateWordSnapshotPositions = True
End Function

Private Function private_BuildDeletedEventOptionCaption( _
    ByVal isMovementEvent As Boolean, _
    ByVal rowIndex As Long, _
    ByVal eventId As String _
) As String
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim tableItems As Collection
    Dim eventsTable As obj_TableDynamic
    Dim eventRow As obj_Row
    Dim separatorIndex As Long
    Dim ipnText As String

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function
    If Not runtimeSources.TryGetItemsSourceByKey( _
        VBA.LCase$(MOVEMENT_EVENTS_RUNTIME_KEY), tableItems, True) Then Exit Function
    If tableItems Is Nothing Or tableItems.Count <> 1 Then Exit Function
    Set eventsTable = tableItems.Item(1)
    If eventsTable Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > eventsTable.RowCount Then Exit Function
    Set eventRow = eventsTable.Rows.Item(rowIndex)
    If eventRow Is Nothing Then Exit Function
    If isMovementEvent Then
        ' В option caption Movement сохраняется исходный контракт:
        ' Тип, ПІБ, Подія, ІПН. В UI ПІБ и Подія показаны в обратном порядке.
        private_BuildDeletedEventOptionCaption = _
            VBA.CStr(eventRow.GetCellValue(1)) & VBA.vbTab & _
            VBA.CStr(eventRow.GetCellValue(3)) & VBA.vbTab & _
            VBA.CStr(eventRow.GetCellValue(2)) & VBA.vbTab & _
            VBA.CStr(eventRow.GetCellValue(4))
    Else
        ' Подія WORD является вычисляемой UI-колонкой и в исходный caption
        ' не входит. ИПН восстанавливается из стабильного имени bookmark.
        separatorIndex = VBA.InStrRev(eventId, "_", -1, VBA.vbBinaryCompare)
        If separatorIndex <= 0 Or separatorIndex >= VBA.Len(eventId) Then Exit Function
        ipnText = VBA.Mid$(eventId, separatorIndex + 1)
        private_BuildDeletedEventOptionCaption = _
            VBA.CStr(eventRow.GetCellValue(5)) & VBA.vbTab & _
            VBA.CStr(eventRow.GetCellValue(7)) & VBA.vbTab & _
            VBA.CStr(eventRow.GetCellValue(8)) & VBA.vbTab & ipnText
    End If
End Function

Private Function private_DeletedExportSnapshotKey( _
    ByVal isMovementEvent As Boolean, _
    ByVal eventId As String _
) As String
    If isMovementEvent Then
        private_DeletedExportSnapshotKey = "Movement|" & eventId
    Else
        private_DeletedExportSnapshotKey = "Word|" & eventId
    End If
End Function

Private Function private_TryMarkDeletedExportEvent( _
    ByVal isMovementEvent As Boolean, _
    ByVal eventId As String, _
    ByVal isDeleted As Boolean, _
    Optional ByVal renderNow As Boolean = True _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim tableItems As Collection
    Dim eventsTable As obj_TableDynamic
    Dim eventRow As obj_Row
    Dim eventCell As obj_Cell
    Dim eventIds As Collection
    Dim rowIndex As Long
    Dim firstColumnIndex As Long
    Dim lastColumnIndex As Long
    Dim columnIndex As Long

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function
    If isMovementEvent Then
        Set eventIds = m_MovementEventIds
    Else
        Set eventIds = m_WordEventIds
    End If
    If eventIds Is Nothing Then Exit Function
    For rowIndex = 1 To eventIds.Count
        If VBA.StrComp(VBA.CStr(eventIds.Item(rowIndex)), _
            eventId, VBA.vbBinaryCompare) = 0 Then Exit For
    Next rowIndex
    ' У подтверждённой смены статуса связанное прибытие скрыто из UI: одна
    ' строка представляет обе операции. Отсутствующий alias не является ошибкой,
    ' но последний вызов всё равно должен применить накопленный reflow.
    If rowIndex > eventIds.Count Then
        If renderNow Then
            If Not pageBase.TryReflowControl(MOVEMENT_EVENTS_CONTROL_NAME) Then Exit Function
        End If
        private_TryMarkDeletedExportEvent = True
        Exit Function
    End If
    If Not runtimeSources.TryGetItemsSourceByKey( _
        VBA.LCase$(MOVEMENT_EVENTS_RUNTIME_KEY), tableItems, True) Then Exit Function
    If tableItems Is Nothing Or tableItems.Count <> 1 Then Exit Function
    Set eventsTable = tableItems.Item(1)
    If eventsTable Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > eventsTable.RowCount Then Exit Function
    Set eventRow = eventsTable.Rows.Item(rowIndex)
    If eventRow Is Nothing Then Exit Function
    If isMovementEvent Then
        firstColumnIndex = 1
        lastColumnIndex = 4
    Else
        firstColumnIndex = 5
        lastColumnIndex = 8
    End If
    For columnIndex = firstColumnIndex To lastColumnIndex
        Set eventCell = eventRow.Cells.Item(columnIndex)
        If eventCell Is Nothing Then Exit Function
        If isDeleted Then
            eventCell.Desc = "diff:deleted"
        Else
            eventCell.Desc = VBA.vbNullString
        End If
    Next columnIndex
    If Not renderNow Then
        private_TryMarkDeletedExportEvent = True
        Exit Function
    End If
    If Not pageBase.TryReflowControl(MOVEMENT_EVENTS_CONTROL_NAME) Then
        VBA.MsgBox "Не удалось частично обновить таблицу экспортированных событий. " & _
            "Нажмите 'Update Sheet' для восстановления страницы.", _
            VBA.vbExclamation, "PrsnlEventBuilder / deleted events"
        Exit Function
    End If
    private_TryMarkDeletedExportEvent = True
End Function

Private Function private_TryRestoreSelectedDeletedEvent( _
    ByVal targetCell As Range, _
    ByRef outHandled As Boolean _
) As Boolean
    Dim isMovementEvent As Boolean
    Dim eventId As String
    Dim eventCaption As String
    Dim snapshotKey As String
    Dim snapshot As Object
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim movementExporter As obj_PEB_ExptrMovement
    Dim wordExporter As obj_PEB_ExptrWord
    Dim relatedSnapshot As Object
    Dim primaryEventId As String
    Dim relatedEventId As String
    Dim primarySnapshotKey As String
    Dim relatedSnapshotKey As String

    outHandled = False
    If m_DeletedExportSnapshots Is Nothing Then
        private_TryRestoreSelectedDeletedEvent = True
        Exit Function
    End If
    If Not private_TryResolveSelectedCombinedExportedEvent( _
        targetCell, isMovementEvent, eventId, eventCaption) Then
        private_TryRestoreSelectedDeletedEvent = True
        Exit Function
    End If
    snapshotKey = private_DeletedExportSnapshotKey(isMovementEvent, eventId)
    If Not m_DeletedExportSnapshots.Exists(snapshotKey) Then
        private_TryRestoreSelectedDeletedEvent = True
        Exit Function
    End If
    outHandled = True
    Set snapshot = m_DeletedExportSnapshots(snapshotKey)
    If snapshot.Exists("IsAlias") Then
        If VBA.CBool(snapshot("IsAlias")) Then Set snapshot = snapshot("RootSnapshot")
    End If
    If isMovementEvent Then
        If Not private_TryGetExportSettings( _
            "Movement", exporterClassName, exportConfigTable) Then Exit Function
        If Not private_TryCreateDataExporter( _
            exporterClassName, exportConfigTable, exporter) Then Exit Function
        Set movementExporter = exporter
        If Not movementExporter.RestoreDeletedEvent(snapshot) Then Exit Function
    Else
        If Not private_TryGetExportSettings( _
            "Word", exporterClassName, exportConfigTable) Then Exit Function
        If Not private_TryCreateDataExporter( _
            exporterClassName, exportConfigTable, exporter) Then Exit Function
        Set wordExporter = exporter
        If Not wordExporter.RestoreDeletedRecord(snapshot) Then Exit Function
        If Not private_TranslateWordSnapshotPositions( _
            snapshot, True, snapshotKey) Then Exit Function
    End If
    If isMovementEvent Then
        primaryEventId = VBA.CStr(snapshot("EventId"))
        primarySnapshotKey = private_DeletedExportSnapshotKey(True, primaryEventId)
        If m_DeletedExportSnapshots.Exists(primarySnapshotKey) Then _
            m_DeletedExportSnapshots.Remove primarySnapshotKey
        If snapshot.Exists("RelatedSnapshot") Then
            Set relatedSnapshot = snapshot("RelatedSnapshot")
            relatedEventId = VBA.CStr(relatedSnapshot("EventId"))
            relatedSnapshotKey = private_DeletedExportSnapshotKey(True, relatedEventId)
            If m_DeletedExportSnapshots.Exists(relatedSnapshotKey) Then _
                m_DeletedExportSnapshots.Remove relatedSnapshotKey
        End If
        If Not private_TryMarkDeletedExportEvent( _
            True, primaryEventId, False, _
            VBA.Len(relatedEventId) = 0) Then Exit Function
        If VBA.Len(relatedEventId) > 0 Then
            If Not private_TryMarkDeletedExportEvent( _
                True, relatedEventId, False) Then Exit Function
        End If
    Else
        m_DeletedExportSnapshots.Remove snapshotKey
        If Not private_TryMarkDeletedExportEvent( _
            False, eventId, False) Then Exit Function
    End If
    rt_Messaging.fn_ShowStatusBarSuccess _
        "Событие восстановлено: " & eventCaption, 4
    private_TryRestoreSelectedDeletedEvent = True
End Function

Private Function private_TryResolveSelectedCombinedExportedEvent( _
    ByVal targetCell As Range, _
    ByRef outIsMovementEvent As Boolean, _
    ByRef outEventId As String, _
    ByRef outEventCaption As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim rowsScope As Range
    Dim columnScope As Range
    Dim rowsArea As Range
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim eventIds As Collection
    Dim eventCaptions As Collection

    outEventId = VBA.vbNullString
    outEventCaption = VBA.vbNullString
    If targetCell Is Nothing Or m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "tablelist", MOVEMENT_EVENTS_CONTROL_NAME, "rows", _
        rowsScope, columnScope) Then Exit Function
    If rowsScope Is Nothing Then Exit Function

    For Each rowsArea In rowsScope.Areas
        If Not Application.Intersect(targetCell, rowsArea) Is Nothing Then
            rowIndex = targetCell.Row - rowsArea.Row + 1
            columnIndex = targetCell.Column - rowsArea.Column + 1
            outIsMovementEvent = (columnIndex >= 1 And columnIndex <= 4)
            If outIsMovementEvent Then
                Set eventIds = m_MovementEventIds
                Set eventCaptions = m_MovementEventCaptions
            ElseIf columnIndex >= 5 And columnIndex <= 8 Then
                Set eventIds = m_WordEventIds
                Set eventCaptions = m_WordEventCaptions
            Else
                Exit Function
            End If
            If eventIds Is Nothing Or eventCaptions Is Nothing Then Exit Function
            If rowIndex <= 0 Or rowIndex > eventIds.Count Then Exit Function
            outEventId = VBA.CStr(eventIds.Item(rowIndex))
            outEventCaption = VBA.CStr(eventCaptions.Item(rowIndex))
            private_TryResolveSelectedCombinedExportedEvent = True
            Exit Function
        End If
    Next rowsArea
End Function

Private Function private_TryResolveSelectedExportedEvent( _
    ByVal targetCell As Range, _
    ByVal controlName As String, _
    ByVal eventIds As Collection, _
    ByVal eventCaptions As Collection, _
    ByRef outEventId As String, _
    ByRef outEventCaption As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim rowsScope As Range
    Dim columnScope As Range
    Dim rowsArea As Range
    Dim rowIndex As Long

    outEventId = VBA.vbNullString
    outEventCaption = VBA.vbNullString
    If targetCell Is Nothing Then Exit Function
    If eventIds Is Nothing Or eventCaptions Is Nothing Then Exit Function
    If eventIds.Count = 0 Or eventIds.Count <> eventCaptions.Count Then Exit Function
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "tablelist", controlName, "rows", rowsScope, columnScope) Then _
        Exit Function
    If rowsScope Is Nothing Then Exit Function

    For Each rowsArea In rowsScope.Areas
        If Not Application.Intersect(targetCell, rowsArea) Is Nothing Then
            rowIndex = targetCell.Row - rowsArea.Row + 1
            If rowIndex <= 0 Or rowIndex > eventIds.Count Then Exit Function
            outEventId = VBA.CStr(eventIds.Item(rowIndex))
            outEventCaption = VBA.CStr(eventCaptions.Item(rowIndex))
            private_TryResolveSelectedExportedEvent = True
            Exit Function
        End If
    Next rowsArea
End Function

Private Function private_TryExportMovementAndWord() As Boolean
    ' Единая команда экспортирует Movement, затем сразу пишет результат в WORD.
    ' Preview-команда CTRL+3 в эту последовательность больше не входит.
    ' Связь двух действий реализована на уровне экспортов Movement/WORD, поэтому
    ' тот же контракт работает и при их раздельном запуске через CTRL+2/CTRL+4.
    If Not private_TryExportDraftByAction( _
        private_BuildExportActionId("Movement")) Then Exit Function
    If Not private_TryExportWordToDocument() Then Exit Function

    private_TryExportMovementAndWord = True
End Function

Private Function private_IsFioCandidatesContext() As Boolean
    Dim entityLookupCfgParser As obj_EntityLookupCfgParser
    Dim lookupKey As String
    Dim candidateTable As obj_TableDynamic
    Dim searchColumnAlias As String

    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetActiveCandidatesContext( _
        entityLookupCfgParser, _
        lookupKey, _
        candidateTable, _
        searchColumnAlias) Then Exit Function

    private_IsFioCandidatesContext = _
        (VBA.StrComp( _
            VBA.Trim$(lookupKey), _
            FIO_LOOKUP_KEY, _
            VBA.vbTextCompare) = 0)
End Function

Private Function private_IsReporterTvoCandidatesContext() As Boolean
    private_IsReporterTvoCandidatesContext = _
        m_IsReporterTvoCandidatesActive
End Function

Private Function private_TryShowReporterTvoCandidates( _
    ByVal targetCell As Range _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim positionCodeRange As Range
    Dim candidates As obj_TableDynamic
    Dim currentPositionCode As String
    Dim lookupPositionCode As String

    m_IsReporterTvoCandidatesActive = False
    If targetCell Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryGetFirstLayoutTagRange( _
        DRAFT_ALIAS_REPORT_POSITION_CODE, positionCodeRange, "visible") Then Exit Function
    If positionCodeRange Is Nothing Then Exit Function
    If Application.Intersect(targetCell, positionCodeRange) Is Nothing Then
        VBA.MsgBox "Для вибору посади ТВО виділіть поле 'Код посади (рапорт)' " & _
            "та натисніть Ctrl+/.", VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
        Exit Function
    End If

    currentPositionCode = VBA.Trim$(VBA.CStr(positionCodeRange.Cells(1, 1).Value2))
    lookupPositionCode = currentPositionCode
    If m_DraftReportIsTvo And VBA.Len(VBA.Trim$(m_ReportOwnPositionCode)) > 0 Then
        lookupPositionCode = m_ReportOwnPositionCode
    End If
    If Not private_TryEnsureExporterCfgDataProvider() Then Exit Function
    If Not m_ExporterCfgDataProvider.TryGetReporterTvoPositionCandidates( _
        lookupPositionCode, candidates) Then Exit Function
    If Not private_UpdateLookupActiveFormColumns() Then Exit Function

    If Not m_DraftReportIsTvo Then m_ReportOwnPositionCode = currentPositionCode
    m_IsReporterTvoCandidatesActive = True
    If Not m_LookupFeature.ShowPreparedCandidates( _
        COMMANDER_LOOKUP_KEY, DRAFT_ALIAS_REPORT_PERSON, _
        "Посади ТВО", candidates, False) Then
        m_IsReporterTvoCandidatesActive = False
        Exit Function
    End If
    ' Полный render восстанавливает значения draft-формы и тем самым запускает
    ' обычный auto-search командира. Обновляем только таблицу кандидатов, чтобы
    ' подготовленный список должностей ТВО не был немедленно перезаписан.
    If Not pageBase.TryReflowControl(LOOKUP_CANDIDATES_CONTROL_NAME) Then
        m_IsReporterTvoCandidatesActive = False
        VBA.MsgBox "Не вдалося оновити список посад ТВО. " & _
            "Натисніть 'Update Sheet' для відновлення сторінки.", _
            VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
        Exit Function
    End If
    If Not private_TryApplyReporterTvoCandidatesAppearance(pageBase.Worksheet) Then Exit Function
    private_TryShowReporterTvoCandidates = True
End Function

Private Function private_TryReadDraftValueByAlias( _
    ByVal aliasText As String, _
    ByRef outValue As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim valueRange As Range

    outValue = VBA.vbNullString
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryGetFirstLayoutTagRange( _
        aliasText, valueRange, "visible") Then Exit Function
    If valueRange Is Nothing Then Exit Function
    outValue = VBA.Trim$(VBA.CStr(valueRange.Cells(1, 1).Value2))
    private_TryReadDraftValueByAlias = True
End Function

Private Function private_RefreshDraftReporterTvoFlag() As Boolean
    Dim currentPositionCode As String

    ' Ручной флаг ТВО является производным: он установлен только тогда,
    ' когда выбранная должность отличается от собственной должности человека.
    m_DraftReportIsTvo = False
    If VBA.Len(VBA.Trim$(m_ReportOwnPositionCode)) = 0 Then
        private_RefreshDraftReporterTvoFlag = True
        Exit Function
    End If
    If Not private_TryReadDraftValueByAlias( _
        DRAFT_ALIAS_REPORT_POSITION_CODE, currentPositionCode) Then Exit Function

    m_DraftReportIsTvo = _
        (VBA.StrComp( _
            private_NormalizeText(currentPositionCode), _
            private_NormalizeText(m_ReportOwnPositionCode), _
            VBA.vbTextCompare) <> 0)
    private_RefreshDraftReporterTvoFlag = True
End Function

Public Function ApplyDraftVisualState() As Boolean
    If Not private_RefreshDraftReporterTvoFlag() Then Exit Function
    ApplyDraftVisualState = private_TryApplyReporterPositionCodeAppearance()
End Function

Private Function private_TryApplyReporterPositionCodeAppearance() As Boolean
    Dim pageBase As obj_PageBase
    Dim positionCodeRange As Range

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    ' Поле присутствует не во всех профилях, поэтому скрытый tag — штатный случай.
    If Not pageBase.TryGetFirstLayoutTagRange( _
        DRAFT_ALIAS_REPORT_POSITION_CODE, positionCodeRange, "visible") Then
        private_TryApplyReporterPositionCodeAppearance = True
        Exit Function
    End If
    If positionCodeRange Is Nothing Then
        private_TryApplyReporterPositionCodeAppearance = True
        Exit Function
    End If

    On Error GoTo EH
    If m_DraftReportIsTvo Then
        positionCodeRange.Font.Color = VBA.RGB(255, 235, 59)
    Else
        positionCodeRange.Font.Color = VBA.RGB(240, 240, 240)
    End If
    private_TryApplyReporterPositionCodeAppearance = True
    Exit Function

EH:
    VBA.MsgBox "Не вдалося оновити колір поля 'Код посади (рапорт)': " & _
        Err.Description, VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
End Function

Private Function private_TryApplyMovementEnrichedFieldsAppearance() As Boolean
    Dim pageBase As obj_PageBase
    Dim fieldAliases As Collection
    Dim fieldAliasObj As Variant
    Dim fieldRange As Range
    Dim fieldValue As String

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set fieldAliases = New Collection

    If private_IsHospitalReturnProfile() Then
        fieldAliases.Add DRAFT_ALIAS_HOSPITAL
        fieldAliases.Add DRAFT_ALIAS_HOSPITAL_SHORT
    ElseIf private_IsVacationReturnProfile() Then
        fieldAliases.Add DRAFT_ALIAS_VACATION_TICKET_NO
        fieldAliases.Add DRAFT_ALIAS_VACATION_TICKET_DATE
    Else
        private_TryApplyMovementEnrichedFieldsAppearance = True
        Exit Function
    End If

    On Error GoTo EH
    For Each fieldAliasObj In fieldAliases
        Set fieldRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange( _
            VBA.CStr(fieldAliasObj), fieldRange, "visible") Then GoTo MissingField
        If fieldRange Is Nothing Then GoTo MissingField
        fieldValue = VBA.Trim$(VBA.CStr(fieldRange.Cells(1, 1).Value2))
        If VBA.Len(fieldValue) > 0 Then
            fieldRange.Font.Color = VBA.RGB(255, 235, 59)
        Else
            fieldRange.Font.Color = VBA.RGB(240, 240, 240)
        End If
    Next fieldAliasObj

    private_TryApplyMovementEnrichedFieldsAppearance = True
    Exit Function

MissingField:
    VBA.MsgBox "Не знайдено поле форми для підсвічування Movement enrichment: " & _
        VBA.CStr(fieldAliasObj), VBA.vbExclamation, _
        "PrsnlEventBuilder / Movement enrichment"
    Exit Function

EH:
    VBA.MsgBox "Не вдалося оновити колір полів Movement enrichment: " & _
        Err.Description, VBA.vbExclamation, _
        "PrsnlEventBuilder / Movement enrichment"
End Function

Private Function private_TryApplyReporterTvoCandidatesAppearance( _
    ByVal ws As Worksheet _
) As Boolean
    Dim partRange As Range
    Dim columnScope As Range
    Dim partName As String

    If ws Is Nothing Then Exit Function
    On Error GoTo EH

    partName = "section"
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "tablelist", LOOKUP_CANDIDATES_CONTROL_NAME, partName, _
        partRange, columnScope) Then GoTo MissingPart
    private_ApplyReporterTvoCandidatePartStyle _
        partRange, VBA.RGB(22, 13, 21), VBA.RGB(255, 226, 240), True

    Set partRange = Nothing
    Set columnScope = Nothing
    partName = "header"
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "tablelist", LOOKUP_CANDIDATES_CONTROL_NAME, partName, _
        partRange, columnScope) Then GoTo MissingPart
    private_ApplyReporterTvoCandidatePartStyle _
        partRange, VBA.RGB(74, 16, 47), VBA.RGB(255, 226, 240), True

    Set partRange = Nothing
    Set columnScope = Nothing
    partName = "rows"
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "tablelist", LOOKUP_CANDIDATES_CONTROL_NAME, partName, _
        partRange, columnScope) Then GoTo MissingPart
    private_ApplyReporterTvoCandidatePartStyle _
        partRange, VBA.RGB(122, 31, 82), VBA.RGB(255, 255, 255), False

    private_TryApplyReporterTvoCandidatesAppearance = True
    Exit Function

MissingPart:
    VBA.MsgBox "Не знайдено обов'язкову частину '" & partName & _
        "' таблиці кандидатів ТВО.", VBA.vbExclamation, _
        "PrsnlEventBuilder / ТВО"
    Exit Function

EH:
    VBA.MsgBox "Не вдалося застосувати стиль таблиці кандидатів ТВО: " & _
        Err.Description, VBA.vbExclamation, "PrsnlEventBuilder / ТВО"
End Function

Private Sub private_ApplyReporterTvoCandidatePartStyle( _
    ByVal targetRange As Range, _
    ByVal backgroundColor As Long, _
    ByVal fontColor As Long, _
    ByVal fontBold As Boolean _
)
    targetRange.Interior.Color = backgroundColor
    targetRange.Font.Color = fontColor
    targetRange.Font.Bold = fontBold
    targetRange.Borders.Color = VBA.RGB(0, 0, 0)
    targetRange.Borders.Weight = xlThin
End Sub

Public Function RuntimeClearExportFormAndCandidates() As Boolean
    Dim previousSuppressLookupSearch As Boolean
    Dim renderSucceeded As Boolean
    Dim activeProfileWasMeta As Boolean

    If m_Page Is Nothing Then Exit Function
    previousSuppressLookupSearch = m_SuppressLookupSearch
    On Error GoTo RestoreSuppression

    ' ALT+ARROWUP очищает draft и применённую форму, сворачивает временное
    ' рабочее состояние, но сохраняет session undo Movement/WORD-записей.
    If Not private_TryClearRenderedFormValues() Then Exit Function
    private_ClearExportFormState
    m_DraftReportIsTvo = False
    m_ReportOwnPositionCode = VBA.vbNullString
    m_WordExportPreviewText = VBA.vbNullString
    m_IsWordPreviewExportMode = False
    m_BottomWorkspaceMode = VBA.vbNullString

    activeProfileWasMeta = private_IsMetaProfile(m_SelectedProfile)
    If activeProfileWasMeta Then
        If VBA.Len(VBA.Trim$(m_SelectedMainProfile)) = 0 Then
            VBA.MsgBox "Не выбрана основная секция, к которой можно вернуться " & _
                "после Meta-профиля.", VBA.vbExclamation, _
                "PrsnlEventBuilder / clear workspace"
            Exit Function
        End If
        m_SelectedProfile = m_SelectedMainProfile
    End If

    If Not m_LookupFeature Is Nothing Then
        If Not m_LookupFeature.ClearLookupCandidates(False) Then Exit Function
    End If
    m_IsReporterTvoCandidatesActive = False
    m_IsMovementHistoryEnabled = False
    Set m_MovementHistoryTable = Nothing

    If Not private_RegisterExportFormTables(False) Then Exit Function
    If Not private_RegisterMovementHistoryTable(False) Then Exit Function
    If Not private_ClearRenderedExportedEvents() Then Exit Function
    If Not private_RegisterProfileOptions(False) Then Exit Function
    If Not private_RegisterAdditionalProfileOptions(False) Then Exit Function
    If Not private_RegisterMetaProfileOptions(False) Then Exit Function
    If activeProfileWasMeta Then
        If Not m_Data.IsAdditionalProfileName(m_SelectedProfile) Then
            If Not private_TryResetAdditionalProfileSelect() Then Exit Function
        End If
    End If

    ' Полный render восстанавливает draft-значения и может повторно инициировать
    ' lookup для заполненного ФИО. Команда очистки не является поиском, поэтому
    ' временно подавляем только SearchCandidates, не меняя режим Lookup.
    m_SuppressLookupSearch = True
    renderSucceeded = rt_PageManager.fn_RenderPage( _
        m_Page, "prsnlevntbuilder:clear-export-form-and-candidates")
    m_SuppressLookupSearch = previousSuppressLookupSearch
    RuntimeClearExportFormAndCandidates = renderSucceeded

    If RuntimeClearExportFormAndCandidates Then
        rt_Messaging.fn_ShowStatusBarSuccess _
            "Форма, preview и раскрытые таблицы очищены.", 3
    End If
    Exit Function

RestoreSuppression:
    m_SuppressLookupSearch = previousSuppressLookupSearch
End Function

Private Function private_TryClearRenderedFormValues() As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim draftValuesRange As Range
    Dim rowsRange As Range
    Dim columnScope As Range
    Dim previousEnableEvents As Boolean
    Dim controlNames As Variant
    Dim controlName As Variant

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange( _
        EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then Exit Function
    If draftValuesRange Is Nothing Then Exit Function

    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    draftValuesRange.ClearContents

    controlNames = VBA.Array( _
        EVENT_EXPORT_MAIN_CONTROL_NAME, EVENT_EXPORT_META_CONTROL_NAME)
    For Each controlName In controlNames
        Set rowsRange = Nothing
        Set columnScope = Nothing
        ' У пустой таблицы part=rows может отсутствовать — это штатное состояние.
        If ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
            ws, "tablelist", VBA.CStr(controlName), "rows", _
            rowsRange, columnScope) Then
            If Not rowsRange Is Nothing Then rowsRange.ClearContents
        End If
    Next controlName

    Application.EnableEvents = previousEnableEvents
    private_TryClearRenderedFormValues = True
    Exit Function
EH:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
    VBA.MsgBox "Не удалось полностью очистить форму: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "PrsnlEventBuilder / clear workspace"
End Function

Public Function OnExportToWordClick(Optional ByVal ignored As Variant) As Boolean
    If m_AreWordBookmarksShownAsMarkers Then
        VBA.MsgBox _
            "Сначала восстановите WORD-закладки из видимых маркеров.", _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD bookmarks"
        Exit Function
    End If
    OnExportToWordClick = private_TryExportWordToDocument()
End Function

Public Function OnToggleWordBookmarksClick( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim orderNoText As String
    Dim wordExporter As obj_PEB_ExptrWord
    Dim markersAreVisible As Boolean
    Dim convertedCount As Long

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", _
            VBA.vbExclamation, "PrototypeNew / WORD bookmarks"
        Exit Function
    End If
    If Not private_TryCreateDataExporter( _
        exporterClassName, exportConfigTable, exporter) Then Exit Function
    Set wordExporter = exporter
    If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function
    If Not wordExporter.ToggleSupportedBookmarks( _
        markersAreVisible, convertedCount, orderNoText) Then Exit Function

    m_AreWordBookmarksShownAsMarkers = markersAreVisible
    If Not ex_ControlRefreshRuntime.fn_TryRefreshStaticControl( _
        WORD_BOOKMARKS_TOGGLE_CONTROL_NAME) Then Exit Function

    If markersAreVisible Then
        rt_Messaging.fn_ShowStatusBarSuccess _
            "WORD bookmarks shown as markers: " & VBA.CStr(convertedCount), 4
    Else
        rt_Messaging.fn_ShowStatusBarSuccess _
            "WORD bookmarks restored: " & VBA.CStr(convertedCount), 4
    End If
    OnToggleWordBookmarksClick = True
End Function

Public Function OnRefreshExportedEventsClick( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pageBase As obj_PageBase

    If Not private_RegisterExportedEventMenus(True, False) Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryReflowControl(MOVEMENT_EVENTS_CONTROL_NAME) Then
        VBA.MsgBox "Не вдалося частково оновити таблицю редагування. " & _
            "Натисніть 'Update Sheet' для відновлення сторінки.", _
            VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
        Exit Function
    End If
    OnRefreshExportedEventsClick = True
End Function

Public Function OnMovementEventClick(Optional ByVal eventId As Variant) As Boolean
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim movementExporter As obj_PEB_ExptrMovement
    Dim orderNoText As String
    Dim deleteSnapshot As Object
    Dim relatedSnapshot As Object

    If VBA.StrComp(VBA.CStr(eventId), "NONE", VBA.vbTextCompare) = 0 Then
        OnMovementEventClick = True
        Exit Function
    End If
    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function
    If Not private_TryGetExportSettings( _
        "Movement", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Movement settings are missing.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement events"
        Exit Function
    End If
    If Not private_TryCreateDataExporter( _
        exporterClassName, exportConfigTable, exporter) Then Exit Function
    Set movementExporter = exporter
    If Not movementExporter.DeleteEventById( _
        VBA.CStr(eventId), orderNoText, deleteSnapshot) Then Exit Function
    If Not private_StoreDeletedExportSnapshot( _
        True, VBA.CStr(eventId), deleteSnapshot) Then Exit Function
    If Not private_TryMarkDeletedExportEvent( _
        True, VBA.CStr(eventId), True, _
        Not deleteSnapshot.Exists("RelatedSnapshot")) Then Exit Function
    If deleteSnapshot.Exists("RelatedSnapshot") Then
        Set relatedSnapshot = deleteSnapshot("RelatedSnapshot")
        If Not private_TryMarkDeletedExportEvent( _
            True, VBA.CStr(relatedSnapshot("EventId")), _
            True) Then Exit Function
    End If
    OnMovementEventClick = True
End Function

Public Function OnWordEventClick(Optional ByVal eventId As Variant) As Boolean
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim wordExporter As obj_PEB_ExptrWord
    Dim orderNoText As String
    Dim deleteSnapshot As Object

    If VBA.StrComp(VBA.CStr(eventId), "NONE", VBA.vbTextCompare) = 0 Then
        OnWordEventClick = True
        Exit Function
    End If
    If m_AreWordBookmarksShownAsMarkers Then
        VBA.MsgBox "Сначала восстановите WORD-закладки из видимых маркеров.", _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
        Exit Function
    End If
    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
        Exit Function
    End If
    If Not private_TryCreateDataExporter( _
        exporterClassName, exportConfigTable, exporter) Then Exit Function
    Set wordExporter = exporter
    If Not wordExporter.DeleteRecordBookmark( _
        VBA.CStr(eventId), orderNoText, deleteSnapshot) Then Exit Function
    If Not private_StoreDeletedExportSnapshot( _
        False, VBA.CStr(eventId), deleteSnapshot) Then Exit Function
    OnWordEventClick = private_TryMarkDeletedExportEvent( _
        False, VBA.CStr(eventId), True)
End Function

Public Function OnClearWordDocumentClick(Optional ByVal ignored As Variant) As Boolean
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim clearedBlockCount As Long
    Dim orderNoText As String
    Dim wordExporter As obj_PEB_ExptrWord

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function
    Set wordExporter = exporter
    If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function
    If Not wordExporter.RemoveResultDocumentAnchors( _
        clearedBlockCount, orderNoText) Then Exit Function

    rt_Messaging.fn_ShowStatusBarSuccess "WORD document: removed anchor blocks: " & VBA.CStr(clearedBlockCount), 3
    OnClearWordDocumentClick = True
End Function

Public Function OnUndoLastExportClick(Optional ByVal ignored As Variant) As Boolean
    Dim actionMatched As Boolean

    ' Кнопка использует общий Excel undo-stack, но откатывает только последнее
    ' действие экспорта PEB. Более новое действие другого типа не пропускается.
    If rt_UndoManager.fn_TryUndoLastByScopePrefix("PEB.Export.", actionMatched) Then
        rt_Messaging.fn_ShowStatusBarSuccess "Last export action was undone.", 3
        OnUndoLastExportClick = True
    ElseIf Not actionMatched Then
        rt_Messaging.fn_ShowStatusBarWarning _
            "There is no last PEB export action available to undo.", 4
    End If
End Function

Public Function OnDisconnectDataSourcesClick( _
    Optional ByVal ignored As Variant _
) As Boolean
    Dim pebExptrCommonDataPrvdr As obj_PEB_ExptrCommonDataPrvdr

    On Error GoTo EH

    ' Сначала освобождаем exporters-заёмщиков, затем профильный provider:
    ' он владеет собственным query engine и временным Movement snapshot.
    private_DisposeCachedExporters
    If Not m_ExporterCfgDataProvider Is Nothing Then
        m_ExporterCfgDataProvider.Dispose
    End If
    Set m_ExporterCfgDataProvider = Nothing

    ' Common provider также держит отдельный obj_ExtWorkbookQueryEngine.
    ' Новый экземпляр остаётся холодным: Initialize подготавливает runtime,
    ' а ADO-соединение будет создано только следующим фактическим запросом.
    If Not m_ExportCommonData Is Nothing Then m_ExportCommonData.Dispose
    Set m_ExportCommonData = Nothing

    ' Глобальный SQL engine не входит в lifecycle страницы и поэтому
    ' сбрасывается явно, как при выполнении команды Clear Pages.
    ex_ExternalExcelSqlEngine.fn_ResetRuntimeCache

    Set pebExptrCommonDataPrvdr = New obj_PEB_ExptrCommonDataPrvdr
    If Not pebExptrCommonDataPrvdr.Initialize() Then
        VBA.MsgBox _
            "З'єднання закрито, але не вдалося повторно підготувати " & _
            "provider спільних даних. Оновіть сторінку перед наступним запитом.", _
            VBA.vbExclamation, _
            "PrsnlEventBuilder / З'єднання"
        Exit Function
    End If
    Set m_ExportCommonData = pebExptrCommonDataPrvdr
    If m_HasResolvedOrderPair Then
        If Not m_ExportCommonData.SetResolvedOrderPair( _
            m_ResolvedOrderNo, m_ResolvedOrderDate) Then Exit Function
    End If

    rt_Messaging.fn_ShowStatusBarSuccess _
        "Усі з'єднання з зовнішніми таблицями розірвано.", 4
    OnDisconnectDataSourcesClick = True
    Exit Function

EH:
    VBA.MsgBox _
        "Не вдалося розірвати всі з'єднання з зовнішніми таблицями: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, _
        "PrsnlEventBuilder / З'єднання"
End Function

Private Function private_TryExportWordToDocument() As Boolean
    Dim sourceTables As Collection
    Dim exportContext As Object
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim editedPreviewText As String

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_TryBuildExportSourceTables(sourceTables, exportContext) Then Exit Function
    private_ApplyPendingMovementReceipt sourceTables, exportContext
    If m_IsWordPreviewExportMode Then
        If Not private_TryReadRenderedWordPreview(editedPreviewText) Then Exit Function
    End If
    If m_IsWordPreviewExportMode And VBA.Len(VBA.Trim$(editedPreviewText)) > 0 Then
        ' Источником становится фактический текст Banner на листе, а не
        ' сохранённая модель, но только после явного включения режима кнопкой.
        exportContext(EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY) = editedPreviewText
    End If
    exportContext("WriteToWord") = True
    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function
    If Not exporter.Export(sourceTables, exportContext) Then Exit Function
    private_ClearPendingMovementReceipt
    ' CTRL+4/button exports the already prepared logical result to WORD.
    ' Do not capture the preview here: that path calls RenderPage and a WORD
    ' write does not change any visible state on the Excel page.
    rt_Messaging.fn_ShowStatusBarSuccess "Export to WORD: done", 3
    private_TryExportWordToDocument = True
End Function

Private Function private_TryReadRenderedWordPreview( _
    ByRef outPreviewText As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim messageRange As Range
    Dim columnScope As Range

    outPreviewText = VBA.vbNullString

    ' Пустая модель означает, что Banner не отрендерен и control part
    ' отсутствует. В этом случае WORD exporter использует обычный шаблон.
    If VBA.Len(VBA.Trim$(m_WordExportPreviewText)) = 0 Then
        private_TryReadRenderedWordPreview = True
        Exit Function
    End If
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, _
        "banner", _
        WORD_EXPORT_PREVIEW_CONTROL_NAME, _
        "message", _
        messageRange, _
        columnScope) Then Exit Function
    If messageRange Is Nothing Then
        VBA.MsgBox _
            "Не вдалося знайти область тексту WORD preview на аркуші. " & _
            "Натисніть 'Update Sheet' та повторіть експорт.", _
            VBA.vbExclamation, _
            "PrsnlEventBuilder / WORD preview"
        Exit Function
    End If

    On Error GoTo EH
    ' Banner message является объединённым диапазоном; значение хранится
    ' в его верхней левой ячейке даже после ручного редактирования.
    outPreviewText = VBA.CStr(messageRange.Cells(1, 1).Value2)
    private_TryReadRenderedWordPreview = True
    Exit Function

EH:
    VBA.MsgBox _
        "Не вдалося прочитати текст WORD preview: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, _
        "PrsnlEventBuilder / WORD preview"
End Function

Public Function OnProfileButtonClick(Optional ByVal profileId As Variant) As Boolean
    Dim newProfile As String
    Dim previousEnableEvents As Boolean
    Dim draftValuesByAlias As Object
    newProfile = VBA.Trim$(VBA.CStr(profileId))
    If VBA.Len(newProfile) = 0 Then Exit Function
    If VBA.StrComp(private_NormalizeText(newProfile), private_NormalizeText(m_SelectedProfile), vbTextCompare) = 0 Then
        OnProfileButtonClick = True
        Exit Function
    End If

    If Not private_TryCaptureDraftValuesByAlias( _
        draftValuesByAlias) Then Exit Function


    If private_IsMainProfile(newProfile) Then
        If VBA.Len(VBA.Trim$(m_SelectedMainProfile)) = 0 Then
            m_SelectedMainProfile = newProfile
        ElseIf VBA.StrComp(private_NormalizeText(newProfile), private_NormalizeText(m_SelectedMainProfile), VBA.vbTextCompare) <> 0 Then
            private_ClearExportFormState
            m_SelectedMainProfile = newProfile
        End If
    End If

    ' Текст preview оставляем доступным для просмотра, но после смены секции
    ' экспорт снова обязан использовать актуальный шаблонный контекст.
    If Not Me.ResetWordPreviewExportMode(False) Then Exit Function
    m_SelectedProfile = newProfile
    If m_Page Is Nothing Then Exit Function
    If private_IsMainProfile(newProfile) And Not m_Data.IsAdditionalProfileName(newProfile) Then
        If Not private_TryResetAdditionalProfileSelect() Then Exit Function
    End If
    If Not private_RegisterProfileOptions(False) Then Exit Function
    If Not private_RegisterAdditionalProfileOptions(False) Then Exit Function
    If Not private_RegisterMetaProfileOptions(False) Then Exit Function
    If Not private_RegisterExportFormTables(False) Then Exit Function
    ' Проекция кандидатов относится к схеме формы профиля, где запускался поиск.
    ' Перед отрисовкой другого профиля очищаем её, иначе LookupCandidates попробует
    ' совместить старую проекцию с новым набором видимых полей.
    If Not m_LookupFeature Is Nothing Then
        If Not m_LookupFeature.ClearLookupCandidates(False) Then Exit Function
    End If

    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo EH

    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "prsnlevntbuilder:profile-changed") Then GoTo Cleanup
    If Not private_TryRestoreDraftValuesByAlias( _
        draftValuesByAlias) Then GoTo Cleanup
    OnProfileButtonClick = True

Cleanup:
    Application.EnableEvents = previousEnableEvents
    Exit Function

EH:
    Resume Cleanup
End Function

Private Function private_TryCaptureDraftValuesByAlias( _
    ByRef outValuesByAlias As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim draftValuesRange As Range
    Dim declaredAliases As Collection
    Dim aliasItem As Variant
    Dim aliasText As String
    Dim valueRange As Range

    Set outValuesByAlias = VBA.CreateObject("Scripting.Dictionary")
    outValuesByAlias.CompareMode = VBA.vbTextCompare
    If m_Page Is Nothing Or m_LookupFeature Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange( _
        EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then Exit Function
    If draftValuesRange Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetFormColumnKeys(declaredAliases) Then Exit Function
    If declaredAliases Is Nothing Then Exit Function

    For Each aliasItem In declaredAliases
        aliasText = VBA.Trim$(VBA.CStr(aliasItem))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueAlias
        Set valueRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange( _
            aliasText, valueRange, "visible") Then GoTo ContinueAlias
        If valueRange Is Nothing Then GoTo ContinueAlias
        If Application.Intersect(valueRange, draftValuesRange) Is Nothing Then _
            GoTo ContinueAlias
        outValuesByAlias(aliasText) = valueRange.Cells(1, 1).Value2
ContinueAlias:
    Next aliasItem
    private_TryCaptureDraftValuesByAlias = True
End Function

Private Function private_TryRestoreDraftValuesByAlias( _
    ByVal valuesByAlias As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim draftValuesRange As Range
    Dim aliasItem As Variant
    Dim valueRange As Range

    If valuesByAlias Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange( _
        EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then Exit Function
    If draftValuesRange Is Nothing Then Exit Function

    For Each aliasItem In valuesByAlias.Keys
        Set valueRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange( _
            VBA.CStr(aliasItem), valueRange, "visible") Then GoTo ContinueAlias
        If valueRange Is Nothing Then GoTo ContinueAlias
        If Application.Intersect(valueRange, draftValuesRange) Is Nothing Then _
            GoTo ContinueAlias
        valueRange.Cells(1, 1).Value2 = valuesByAlias(aliasItem)
ContinueAlias:
    Next aliasItem
    private_TryRestoreDraftValuesByAlias = True
End Function

Public Function OnMetaProfileButtonClick(Optional ByVal profileId As Variant) As Boolean
    OnMetaProfileButtonClick = Me.OnProfileButtonClick(profileId)
End Function

Public Function OnAdditionalProfileChanged(Optional ByVal ignored As Variant) As Boolean
    Dim pageBase As obj_PageBase
    Dim rawControl As Object
    Dim selectControl As obj_SelectControlVM
    Dim selectedProfile As String

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryGetRegisteredControlByName( _
        ADDITIONAL_PROFILE_SELECT_CONTROL_NAME, rawControl) Then Exit Function
    If rawControl Is Nothing Then Exit Function
    If Not TypeOf rawControl Is obj_SelectControlVM Then Exit Function

    Set selectControl = rawControl
    selectedProfile = VBA.Trim$(selectControl.GetSelectedId())
    ' Пустой ID — штатное placeholder-состояние Select, а не профиль.
    If VBA.Len(selectedProfile) = 0 Then
        OnAdditionalProfileChanged = True
        Exit Function
    End If
    OnAdditionalProfileChanged = Me.OnProfileButtonClick(selectedProfile)
End Function

Public Function RuntimeApplyExportForm() As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim profileText As String
    Dim previousEnableEvents As Boolean

    ' Apply переносит текущую видимую draft-форму в "Форму экспорта".
    ' Основной профиль заменяет основную строку; meta-профиль дописывает meta-строку.
    If Not private_RefreshDraftReporterTvoFlag() Then Exit Function
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
        m_ExportMainReportIsTvo = m_DraftReportIsTvo
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

Private Function private_TryRefreshMovementHistory() As Boolean
    Dim pageBase As obj_PageBase
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim ipnText As String
    Dim fioText As String
    Dim maxRows As Long
    Dim movementHistoryTable As obj_TableDynamic

    If Not m_IsMovementHistoryEnabled Then Exit Function
    If m_ProfileConfigTable Is Nothing Then
        VBA.MsgBox "Не завантажено конфігурацію профілю Movement.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement"
        Exit Function
    End If

    If Not private_TryBuildDraftFormSourceTable(sourceTable, False) Then Exit Function
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <> 1 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function
    If Not sourceRow.TryGetCellValueByColumn("ІПН", ipnText) Then Exit Function
    If Not sourceRow.TryGetCellValueByColumn("ПІБ", fioText) Then fioText = VBA.vbNullString

    ipnText = VBA.Trim$(ipnText)
    If VBA.Len(ipnText) = 0 Then
        VBA.MsgBox "Для перегляду історії руху заповніть поле 'ІПН'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement"
        Exit Function
    End If
    If Not private_TryReadMovementHistoryMaxRows(maxRows) Then Exit Function

    If Not private_TryEnsureExporterCfgDataProvider() Then Exit Function
    If Not m_ExporterCfgDataProvider.TryGetMovementHistoryByIpn( _
        ipnText, movementHistoryTable, maxRows, fioText) Then GoTo Cleanup

    Set m_MovementHistoryTable = movementHistoryTable
    If Not private_RegisterMovementHistoryTable(False) Then GoTo Cleanup

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then GoTo Cleanup
    ' Таблица уже существует в retained layout после включения checkbox.
    ' Partial reflow перерисовывает только её и сдвигает готовую draft-форму
    ' с кандидатами, не запуская полный render всей PEB-страницы.
    If Not pageBase.TryReflowControl( _
        MOVEMENT_HISTORY_TABLE_CONTROL_NAME) Then
        VBA.MsgBox "Не вдалося частково оновити таблицю історії руху. " & _
            "Натисніть 'Update Sheet' для відновлення сторінки.", _
            VBA.vbExclamation, "PrsnlEventBuilder / partial reflow"
        GoTo Cleanup
    End If

    If movementHistoryTable.RowCount = 0 Then
        rt_Messaging.fn_ShowStatusBarWarning _
            "Movement: події для ІПН " & ipnText & " не знайдено.", 3
    Else
        rt_Messaging.fn_ShowStatusBarSuccess _
            "Movement: знайдено подій: " & VBA.CStr(movementHistoryTable.RowCount), 3
    End If
    private_TryRefreshMovementHistory = True

Cleanup:
    Exit Function
End Function

Private Function private_TryReadMovementHistoryMaxRows( _
    ByRef outMaxRows As Long _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim rawControl As Object
    Dim movementHistoryLimitInput As obj_InputControlVM
    Dim rawValue As String
    Dim numericValue As Double

    outMaxRows = 0
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    If Not pageBase.TryGetRegisteredControlByName( _
        MOVEMENT_HISTORY_LIMIT_INPUT_NAME, rawControl) Then
        VBA.MsgBox "Не знайдено поле кількості записів Movement.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement"
        Exit Function
    End If
    If rawControl Is Nothing Then Exit Function
    If Not TypeOf rawControl Is obj_InputControlVM Then Exit Function
    Set movementHistoryLimitInput = rawControl
    If Not movementHistoryLimitInput.TryGetValue(rawValue) Then Exit Function

    rawValue = VBA.Trim$(rawValue)
    ' Пустое поле явно означает отсутствие SQL/result limit.
    If VBA.Len(rawValue) = 0 Then
        private_TryReadMovementHistoryMaxRows = True
        Exit Function
    End If

    If Not VBA.IsNumeric(rawValue) Then GoTo InvalidValue
    On Error GoTo InvalidValue
    numericValue = VBA.CDbl(rawValue)
    If numericValue < 1# Then GoTo InvalidValue
    If numericValue <> VBA.Fix(numericValue) Then GoTo InvalidValue
    If numericValue > 2147483647# Then GoTo InvalidValue
    outMaxRows = VBA.CLng(numericValue)
    private_TryReadMovementHistoryMaxRows = True
    Exit Function

InvalidValue:
    outMaxRows = 0
    VBA.MsgBox "Поле 'Кількість записів' має містити додатне ціле число " & _
        "або бути порожнім для завантаження всіх записів.", _
        VBA.vbExclamation, "PrsnlEventBuilder / Movement"
End Function

Public Function SearchCandidates( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByRef outCandidateCount As Long, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    Dim minAbsenceDepartureDate As Date
    Dim maxAbsenceDepartureDate As Date
    Dim absenceReferenceDate As Date
    Dim absenceSelectorImpl As obj_PEB_AbsenceCnddtSlctr
    Dim absenceSelector As obj_ILookupCandidateSelector

    outCandidateCount = 0
    If m_SuppressLookupSearch Then
        SearchCandidates = True
        Exit Function
    End If
    m_IsReporterTvoCandidatesActive = False
    If m_LookupFeature Is Nothing Then Exit Function
    If Not private_UpdateLookupActiveFormColumns() Then Exit Function
    If VBA.StrComp(VBA.Trim$(lookupKey), FIO_LOOKUP_KEY, VBA.vbTextCompare) = 0 And _
       private_IsVacationReturnProfile() Then
        If Not m_LookupFeature.SearchCandidates( _
            lookupKey, queryText, outCandidateCount, False) Then Exit Function
        If Not private_TryAppendTemporaryPersonnel(queryText) Then Exit Function
        If Not private_TryApplyMovementVacationTicketDefaults() Then
            If notifyChange Then
                If Not rt_PageManager.fn_RenderPage( _
                    m_Page, "prsnlevntbuilder:movement-ticket-enrichment-failed") Then Exit Function
            End If
            SearchCandidates = True
            Exit Function
        End If
        If notifyChange Then
            If Not rt_PageManager.fn_RenderPage( _
                m_Page, "prsnlevntbuilder:movement-vacation-ticket-defaults") Then Exit Function
        End If
        SearchCandidates = True
        Exit Function
    End If
    If VBA.StrComp(VBA.Trim$(lookupKey), FIO_LOOKUP_KEY, VBA.vbTextCompare) = 0 And _
       private_IsHospitalReturnProfile() Then
        If Not m_LookupFeature.SearchCandidates( _
            lookupKey, queryText, outCandidateCount, False) Then Exit Function
        If Not private_TryAppendTemporaryPersonnel(queryText) Then Exit Function
        If Not private_TryApplyMovementHospitalDefaults() Then
            If notifyChange Then
                If Not rt_PageManager.fn_RenderPage( _
                    m_Page, "prsnlevntbuilder:movement-hospital-enrichment-failed") Then Exit Function
            End If
            SearchCandidates = True
            Exit Function
        End If
        If notifyChange Then
            If Not rt_PageManager.fn_RenderPage( _
                m_Page, "prsnlevntbuilder:movement-hospital-defaults") Then Exit Function
        End If
        SearchCandidates = True
        Exit Function
    End If
    ' Расширение из нескольких источников включено только на этой странице.
    ' Универсальный EntityLookup не зависит от источников PrsnlEvntBuilder.
    If VBA.StrComp(VBA.Trim$(lookupKey), "op_FIO", VBA.vbTextCompare) = 0 Then
        If Not private_ShouldExtendAbsenceCandidates() Then
            SearchCandidates = m_LookupFeature.SearchCandidates( _
                lookupKey, queryText, outCandidateCount, notifyChange)
            Exit Function
        End If
        If Not private_TryResolveAbsenceDepartureDateRange( _
            minAbsenceDepartureDate, maxAbsenceDepartureDate, _
            absenceReferenceDate) Then Exit Function
        Set absenceSelectorImpl = New obj_PEB_AbsenceCnddtSlctr
        If Not absenceSelectorImpl.Initialize( _
            DRAFT_ALIAS_DATE_FROM, ABSENCE_ORDER_DATE_ALIAS, _
            minAbsenceDepartureDate, _
            maxAbsenceDepartureDate, absenceReferenceDate) Then Exit Function
        Set absenceSelector = absenceSelectorImpl
        If Not m_LookupFeature.SearchCandidates(lookupKey, queryText, outCandidateCount, False) Then Exit Function
        If Not private_TryAppendTemporaryPersonnel(queryText) Then Exit Function
        If Not m_LookupFeature.ExtendCandidates( _
            "op_FIOAbsenceExtension", _
            "_FIO", _
            queryText, _
            absenceSelector, _
            False) Then Exit Function
        If Not private_TryApplyAbsenceCandidateDefaults() Then Exit Function
        If notifyChange Then
            If Not rt_PageManager.fn_RenderPage( _
                m_Page, "prsnlevntbuilder:absence-candidates-enriched") Then Exit Function
        End If
        SearchCandidates = True
    Else
        If VBA.StrComp(VBA.Trim$(lookupKey), FIO_LOOKUP_KEY, VBA.vbTextCompare) = 0 Then
            If Not m_LookupFeature.SearchCandidates( _
                lookupKey, queryText, outCandidateCount, False) Then Exit Function
            If Not private_TryAppendTemporaryPersonnel(queryText) Then Exit Function
            If notifyChange Then
                If Not rt_PageManager.fn_RenderPage( _
                    m_Page, "prsnlevntbuilder:temporary-personnel-candidates") Then Exit Function
            End If
            SearchCandidates = True
        Else
            SearchCandidates = m_LookupFeature.SearchCandidates( _
                lookupKey, queryText, outCandidateCount, notifyChange)
        End If
    End If
End Function

Private Function private_TryAppendTemporaryPersonnel( _
    ByVal queryText As String _
) As Boolean
    Dim fixedValues As Object

    If Not m_IsTemporaryPersonnelEnabled Then
        private_TryAppendTemporaryPersonnel = True
        Exit Function
    End If
    Set fixedValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    fixedValues(DRAFT_ALIAS_IPN) = TEMPORARY_IPN_MARKER
    fixedValues(DRAFT_ALIAS_POSITION_CODE) = TEMPORARY_POSITION_MARKER
    private_TryAppendTemporaryPersonnel = m_LookupFeature.AppendCandidates( _
        TEMPORARY_PERSONNEL_LOOKUP_KEY, DRAFT_ALIAS_FIO, queryText, _
        fixedValues, False)
End Function

Private Function private_IsHospitalReturnProfile() As Boolean
    Dim sectionText As String
    Dim sectionKey As String

    If m_Data Is Nothing Then Exit Function
    sectionText = VBA.Trim$(m_SelectedMainProfile)
    If VBA.Len(sectionText) = 0 Then sectionText = VBA.Trim$(m_SelectedProfile)
    sectionKey = private_NormalizeText(sectionText)

    Select Case sectionKey
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatment), _
             private_NormalizeText( _
                m_Data.SectionTypeTransferTreatmentToTreatmentVacation)
            private_IsHospitalReturnProfile = True
    End Select
End Function

Private Function private_TryApplyMovementHospitalDefaults() As Boolean
    Dim entityLookupCfgParser As obj_EntityLookupCfgParser
    Dim candidateTable As obj_TableDynamic
    Dim candidateRow As obj_Row
    Dim lookupKey As String
    Dim searchColumnAlias As String
    Dim ipnColumnIndex As Long
    Dim fioColumnIndex As Long
    Dim hospitalColumnIndex As Long
    Dim hospitalShortColumnIndex As Long
    Dim ipnText As String
    Dim fioText As String
    Dim hospitalFound As Boolean
    Dim hospitalShortText As String
    Dim hospitalName As String
    Dim hospitalNameFound As Boolean
    Dim rowIndex As Long

    If m_LookupFeature Is Nothing Then Exit Function
    If Not private_TryEnsureExporterCfgDataProvider() Then Exit Function
    If m_ExporterCfgDataProvider.CommonData Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetActiveCandidatesContext( _
        entityLookupCfgParser, lookupKey, candidateTable, searchColumnAlias) Then Exit Function
    If candidateTable Is Nothing Then Exit Function
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_IPN, ipnColumnIndex) Then
        VBA.MsgBox "У таблиці кандидатів відсутня колонка '_IPN'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement enrichment"
        Exit Function
    End If
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_FIO, fioColumnIndex) Then Exit Function
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_HOSPITAL, hospitalColumnIndex) Then
        VBA.MsgBox "У таблиці кандидатів відсутня колонка '_Hospital'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement enrichment"
        Exit Function
    End If
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_HOSPITAL_SHORT, hospitalShortColumnIndex) Then
        VBA.MsgBox "У таблиці кандидатів відсутня колонка '_HospitalShort'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement enrichment"
        Exit Function
    End If

    For rowIndex = 1 To candidateTable.RowCount
        Set candidateRow = candidateTable.Rows.Item(rowIndex)
        If candidateRow Is Nothing Then Exit Function
        ipnText = VBA.Trim$(VBA.CStr(candidateRow.GetCellValue(ipnColumnIndex)))
        fioText = VBA.Trim$(VBA.CStr(candidateRow.GetCellValue(fioColumnIndex)))
        If VBA.Len(ipnText) = 0 Then GoTo ContinueRow
        If Not m_ExporterCfgDataProvider.TryGetLatestMovementDestination( _
            ipnText, hospitalFound, hospitalShortText, fioText) Then
            VBA.MsgBox "Не вдалося прочитати останнє місце призначення з " & _
                "Movement для ІПН '" & ipnText & "'.", VBA.vbExclamation, _
                "PrsnlEventBuilder / Movement enrichment"
            Exit Function
        End If
        If Not hospitalFound Then GoTo ContinueRow
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveHospitalName( _
            hospitalShortText, hospitalName, True, hospitalNameFound) Then Exit Function
        If Not hospitalNameFound Then GoTo ContinueRow
        If Not candidateRow.SetCellRaw( _
            hospitalColumnIndex, hospitalName) Then Exit Function
        If Not candidateRow.SetCellRaw( _
            hospitalShortColumnIndex, hospitalShortText) Then Exit Function
ContinueRow:
    Next rowIndex

    private_TryApplyMovementHospitalDefaults = True
End Function

Private Function private_IsVacationReturnProfile() As Boolean
    Dim sectionText As String
    Dim sectionKey As String

    If m_Data Is Nothing Then Exit Function
    sectionText = VBA.Trim$(m_SelectedMainProfile)
    If VBA.Len(sectionText) = 0 Then sectionText = VBA.Trim$(m_SelectedProfile)
    sectionKey = private_NormalizeText(sectionText)

    Select Case sectionKey
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromFamilyVacation)
            private_IsVacationReturnProfile = True
    End Select
End Function

Private Function private_TryApplyMovementVacationTicketDefaults() As Boolean
    Dim entityLookupCfgParser As obj_EntityLookupCfgParser
    Dim candidateTable As obj_TableDynamic
    Dim candidateRow As obj_Row
    Dim lookupKey As String
    Dim searchColumnAlias As String
    Dim ipnColumnIndex As Long
    Dim fioColumnIndex As Long
    Dim ticketNoColumnIndex As Long
    Dim ticketDateColumnIndex As Long
    Dim ipnText As String
    Dim fioText As String
    Dim ticketFound As Boolean
    Dim ticketNoText As String
    Dim departureOrderText As String
    Dim departureOrderDate As Date
    Dim departureOrderDateFound As Boolean
    Dim rowIndex As Long

    If m_LookupFeature Is Nothing Then Exit Function
    If Not private_TryEnsureExporterCfgDataProvider() Then Exit Function
    If m_ExporterCfgDataProvider.CommonData Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetActiveCandidatesContext( _
        entityLookupCfgParser, lookupKey, candidateTable, searchColumnAlias) Then Exit Function
    If candidateTable Is Nothing Then Exit Function
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_IPN, ipnColumnIndex) Then
        VBA.MsgBox "У таблиці кандидатів відсутня колонка '_IPN'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement enrichment"
        Exit Function
    End If
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_FIO, fioColumnIndex) Then Exit Function
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_VACATION_TICKET_NO, ticketNoColumnIndex) Then
        VBA.MsgBox "У таблиці кандидатів відсутня колонка '_VacationTicketNo'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement enrichment"
        Exit Function
    End If
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_VACATION_TICKET_DATE, ticketDateColumnIndex) Then
        VBA.MsgBox "У таблиці кандидатів відсутня колонка '_VacationTicketDate'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement enrichment"
        Exit Function
    End If

    For rowIndex = 1 To candidateTable.RowCount
        Set candidateRow = candidateTable.Rows.Item(rowIndex)
        If candidateRow Is Nothing Then Exit Function
        ipnText = VBA.Trim$(VBA.CStr(candidateRow.GetCellValue(ipnColumnIndex)))
        fioText = VBA.Trim$(VBA.CStr(candidateRow.GetCellValue(fioColumnIndex)))
        If VBA.Len(ipnText) = 0 Then GoTo ContinueRow
        If Not m_ExporterCfgDataProvider.TryGetLatestMovementVacationTicket( _
            ipnText, ticketFound, ticketNoText, departureOrderText, fioText) Then
            VBA.MsgBox "Не вдалося прочитати відпускний квиток з Movement " & _
                "для ІПН '" & ipnText & "'.", VBA.vbExclamation, _
                "PrsnlEventBuilder / Movement enrichment"
            Exit Function
        End If
        If Not ticketFound Then GoTo ContinueRow
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveOrderDateByNumber( _
            departureOrderText, departureOrderDate, True, _
            departureOrderDateFound) Then Exit Function
        If Not departureOrderDateFound Then GoTo ContinueRow
        If Not candidateRow.SetCellRaw(ticketNoColumnIndex, ticketNoText) Then Exit Function
        If Not candidateRow.SetCellRaw( _
            ticketDateColumnIndex, VBA.Format$(departureOrderDate, "dd.mm")) Then Exit Function
ContinueRow:
    Next rowIndex

    private_TryApplyMovementVacationTicketDefaults = True
End Function

Private Function private_ShouldExtendAbsenceCandidates() As Boolean
    Dim sectionText As String
    Dim sectionKey As String

    If m_Data Is Nothing Then Exit Function
    sectionText = VBA.Trim$(m_SelectedMainProfile)
    If VBA.Len(sectionText) = 0 Then sectionText = VBA.Trim$(m_SelectedProfile)
    sectionKey = private_NormalizeText(sectionText)

    ' ЕЖОС дополняет командировки, оформление нового выбытия в отпуск и смену
    ' статуса, связанную с отпуском. Чистые возвраты используют только ШПС.
    Select Case sectionKey
        Case private_NormalizeText(m_Data.SectionTypeToBusinessTrip), _
             private_NormalizeText(m_Data.SectionTypeToBusinessTripSzch), _
             private_NormalizeText(m_Data.SectionTypeToAnnualVacationPart), _
             private_NormalizeText(m_Data.SectionTypeToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferExternalVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferMedicalCompanyTreatmentVacationToTreatment)
            private_ShouldExtendAbsenceCandidates = True
    End Select
End Function

Private Function private_TryApplyAbsenceCandidateDefaults() As Boolean
    Dim entityLookupCfgParser As obj_EntityLookupCfgParser
    Dim candidateTable As obj_TableDynamic
    Dim candidateRow As obj_Row
    Dim lookupKey As String
    Dim searchColumnAlias As String
    Dim dateFromColumnIndex As Long
    Dim vacationTicketDateColumnIndex As Long
    Dim dateFromValue As Variant
    Dim orderDateText As String
    Dim rowIndex As Long

    If m_LookupFeature Is Nothing Then Exit Function
    If m_ExportCommonData Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetActiveCandidatesContext( _
        entityLookupCfgParser, lookupKey, candidateTable, searchColumnAlias) Then Exit Function
    If candidateTable Is Nothing Then Exit Function

    ' Не каждый профиль показывает эти поля. В таком профиле extension остаётся
    ' валидным, но вычисляемое значение просто некуда проецировать.
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_DATE_FROM, dateFromColumnIndex) Then
        private_TryApplyAbsenceCandidateDefaults = True
        Exit Function
    End If
    If Not candidateTable.TryGetColumnIndexByAlias( _
        DRAFT_ALIAS_VACATION_TICKET_DATE, vacationTicketDateColumnIndex) Then
        private_TryApplyAbsenceCandidateDefaults = True
        Exit Function
    End If

    If Not m_HasResolvedOrderPair Then Exit Function
    orderDateText = VBA.Format$(m_ResolvedOrderDate, "dd.mm")

    For rowIndex = 1 To candidateTable.RowCount
        Set candidateRow = candidateTable.Rows.Item(rowIndex)
        If candidateRow Is Nothing Then Exit Function
        dateFromValue = candidateRow.GetCellValue(dateFromColumnIndex)
        If VBA.Len(VBA.Trim$(VBA.CStr(dateFromValue))) > 0 Then
            If Not candidateRow.SetCellRaw( _
                vacationTicketDateColumnIndex, orderDateText) Then Exit Function
        End If
    Next rowIndex

    private_TryApplyAbsenceCandidateDefaults = True
End Function

Private Function private_TryResolveAbsenceDepartureDateRange( _
    ByRef outMinDate As Date, _
    ByRef outMaxDate As Date, _
    ByRef outReferenceDate As Date _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim orderNoText As String
    Dim orderDate As Date

    outMinDate = 0
    outMaxDate = 0
    outReferenceDate = 0
    If m_Page Is Nothing Then Exit Function
    If m_ExportCommonData Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    If Not private_TryResolveCurrentOrderPair( _
        pageBase, ws, orderNoText, orderDate) Then Exit Function

    outMinDate = VBA.DateAdd( _
        "d", -ABSENCE_DEPARTURE_LOOKBACK_DAYS, VBA.DateValue(orderDate))
    outMaxDate = VBA.DateAdd( _
        "d", ABSENCE_DEPARTURE_LOOKAHEAD_DAYS, VBA.DateValue(orderDate))
    outReferenceDate = VBA.DateValue(orderDate)
    private_TryResolveAbsenceDepartureDateRange = True
End Function

Private Function private_UpdateLookupActiveFormColumns() As Boolean
    Dim pageBase As obj_PageBase
    Dim declaredAliases As Collection
    Dim activeAliases As Collection
    Dim aliasObj As Variant
    Dim aliasText As String
    Dim targetRange As Range
    Dim aliases() As String
    Dim gridColumns() As Long
    Dim itemCount As Long
    Dim i As Long
    Dim j As Long
    Dim swapAlias As String
    Dim swapColumn As Long

    If m_Page Is Nothing Then Exit Function
    If m_LookupFeature Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetFormColumnKeys(declaredAliases) Then Exit Function
    If declaredAliases Is Nothing Then Exit Function

    ' Учитываем только видимые контролы. Их фактические колонки листа задают
    ' порядок полей формы для активного профиля и режима.
    For Each aliasObj In declaredAliases
        aliasText = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueAlias
        Set targetRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange(aliasText, targetRange, "visible") Then GoTo ContinueAlias
        If targetRange Is Nothing Then GoTo ContinueAlias

        itemCount = itemCount + 1
        ReDim Preserve aliases(1 To itemCount)
        ReDim Preserve gridColumns(1 To itemCount)
        aliases(itemCount) = aliasText
        gridColumns(itemCount) = targetRange.Column
ContinueAlias:
    Next aliasObj

    If itemCount <= 0 Then
        VBA.MsgBox "PrototypeNew: failed to resolve visible EntityLookup form columns for the active profile.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    ' Стабильно сортируем по фактической колонке Excel независимо от порядка в конфиге.
    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If gridColumns(j) < gridColumns(i) Then
                swapColumn = gridColumns(i)
                gridColumns(i) = gridColumns(j)
                gridColumns(j) = swapColumn
                swapAlias = aliases(i)
                aliases(i) = aliases(j)
                aliases(j) = swapAlias
            End If
        Next j
    Next i

    Set activeAliases = New Collection
    For i = 1 To itemCount
        activeAliases.Add aliases(i)
    Next i

    private_UpdateLookupActiveFormColumns = m_LookupFeature.UpdateActiveFormColumns(activeAliases)
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

Private Sub private_ClearResolvedOrderPair()
    m_HasResolvedOrderPair = False
    m_ResolvedOrderNo = VBA.vbNullString
    m_ResolvedOrderDate = 0
    Set m_OrderHistoryItems = New Collection
    m_BottomWorkspaceMode = VBA.vbNullString
    If Not m_ExportCommonData Is Nothing Then _
        m_ExportCommonData.ClearOrderReference
End Sub

Private Function private_EnsureOrderHistoryRuntime( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If m_OrderHistoryItems Is Nothing Then Set m_OrderHistoryItems = New Collection
    If Not pageBase.RuntimeSources.RemoveItemsSource( _
        VBA.LCase$(ORDER_HISTORY_RUNTIME_KEY)) Then Exit Function
    private_EnsureOrderHistoryRuntime = pageBase.RuntimeSources.SetItemsSource( _
        VBA.LCase$(ORDER_HISTORY_RUNTIME_KEY), m_OrderHistoryItems, notifyChange)
End Function

Private Function private_EnsureValidationResultsRuntime( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim existingItems As Collection
    Dim emptyItems As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.TryGetItemsSourceByKey( _
        VBA.LCase$(VALIDATION_RESULTS_RUNTIME_KEY), existingItems, True) Then Exit Function
    If Not existingItems Is Nothing Then
        private_EnsureValidationResultsRuntime = True
        Exit Function
    End If
    Set emptyItems = New Collection
    private_EnsureValidationResultsRuntime = pageBase.RuntimeSources.SetItemsSource( _
        VBA.LCase$(VALIDATION_RESULTS_RUNTIME_KEY), emptyItems, notifyChange)
End Function

Private Function private_EnsureMedicalReportsRuntime( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim existingItems As Collection
    Dim emptyItems As Collection
    Dim existingFilterItems As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.TryGetItemsSourceByKey( _
        VBA.LCase$(MEDICAL_REPORTS_RUNTIME_KEY), _
        existingItems, True) Then Exit Function
    If existingItems Is Nothing Then
        Set emptyItems = New Collection
        If Not pageBase.RuntimeSources.SetItemsSource( _
            VBA.LCase$(MEDICAL_REPORTS_RUNTIME_KEY), _
            emptyItems, notifyChange) Then Exit Function
    End If
    If Not pageBase.RuntimeSources.TryGetItemsSourceByKey( _
        VBA.LCase$(MEDICAL_REPORTS_FILTERS_RUNTIME_KEY), _
        existingFilterItems, True) Then Exit Function
    If existingFilterItems Is Nothing Then
        Set emptyItems = New Collection
        If Not pageBase.RuntimeSources.SetItemsSource( _
            VBA.LCase$(MEDICAL_REPORTS_FILTERS_RUNTIME_KEY), _
            emptyItems, notifyChange) Then Exit Function
    End If
    private_EnsureMedicalReportsRuntime = True
End Function

Private Function private_TryBuildOrderHistoryItems( _
    ByRef outItems As Collection _
) As Boolean
    Dim historyTable As obj_TableDynamic
    Dim columnObj As obj_Column
    Dim rowObj As obj_Row
    Dim cellObj As obj_Cell
    Dim currentOrderNumber As Long
    Dim listedOrderNo As String
    Dim listedOrderDate As Date
    Dim listedOrderFound As Boolean
    Dim orderYear As Long
    Dim offset As Long

    Set outItems = New Collection
    If Not m_HasResolvedOrderPair Then
        private_TryBuildOrderHistoryItems = True
        Exit Function
    End If

    Set historyTable = New obj_TableDynamic
    Set columnObj = New obj_Column
    columnObj.Name = "Номер наказу"
    columnObj.Position = 1
    If Not columnObj.AddAlias("Номер наказу") Then Exit Function
    If Not historyTable.PushColumn(columnObj) Then Exit Function
    Set columnObj = New obj_Column
    columnObj.Name = "Дата наказу"
    columnObj.Position = 2
    If Not columnObj.AddAlias("Дата наказу") Then Exit Function
    If Not historyTable.PushColumn(columnObj) Then Exit Function

    orderYear = VBA.Year(m_ResolvedOrderDate)
    If VBA.IsNumeric(m_ResolvedOrderNo) Then _
        currentOrderNumber = VBA.CLng(m_ResolvedOrderNo)
    For offset = 0 To 4
        If offset = 0 Then
            listedOrderNo = m_ResolvedOrderNo
            listedOrderDate = m_ResolvedOrderDate
            listedOrderFound = True
        ElseIf currentOrderNumber > offset Then
            listedOrderNo = VBA.CStr(currentOrderNumber - offset)
            listedOrderDate = 0
            listedOrderFound = False
            If Not m_ExportCommonData.TryResolveOrderDateByNumber( _
                listedOrderNo, listedOrderDate, True, listedOrderFound, _
                orderYear) Then Exit Function
        Else
            listedOrderFound = False
        End If

        If listedOrderFound Then
            Set rowObj = New obj_Row
            Set cellObj = New obj_Cell
            cellObj.Value = listedOrderNo
            If Not rowObj.PushCell(cellObj) Then Exit Function
            Set cellObj = New obj_Cell
            cellObj.Value = VBA.Format$(listedOrderDate, "dd.mm.yyyy")
            If Not rowObj.PushCell(cellObj) Then Exit Function
            If Not historyTable.PushRow(rowObj) Then Exit Function
        End If
    Next offset

    outItems.Add historyTable
    private_TryBuildOrderHistoryItems = True
End Function

Private Function private_RegisterExportedEventMenus( _
    ByVal loadEvents As Boolean, _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim movementEvents As Collection
    Dim wordEvents As Collection
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim movementExporter As obj_PEB_ExptrMovement
    Dim wordExporter As obj_PEB_ExptrWord
    Dim orderNoText As String
    Dim existingMovementEvents As Collection
    Dim movementEventCount As Long
    Dim wordEventCount As Long
    Dim exportedEventsTableItems As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function
    Set movementEvents = New Collection
    Set wordEvents = New Collection

    If Not loadEvents And Not m_AreExportedEventsShown Then
        If Not private_ClearRenderedExportedEvents() Then Exit Function
        private_RegisterExportedEventMenus = True
        Exit Function
    End If

    If loadEvents Then
        If Not private_TryEnsureModeConfigCurrent() Then Exit Function
        If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function

        If Not private_TryGetExportSettings( _
            "Movement", exporterClassName, exportConfigTable) Then
            VBA.MsgBox "PrototypeNew: Export.Movement settings are missing.", _
                VBA.vbExclamation, "PrsnlEventBuilder / exported events"
            Exit Function
        End If
        If Not private_TryCreateDataExporter( _
            exporterClassName, exportConfigTable, exporter) Then Exit Function
        Set movementExporter = exporter
        If Not movementExporter.GetEventsByOrderNo( _
            orderNoText, movementEvents) Then Exit Function

        Set exporter = Nothing
        If Not private_TryGetExportSettings( _
            "Word", exporterClassName, exportConfigTable) Then
            VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", _
                VBA.vbExclamation, "PrsnlEventBuilder / exported events"
            Exit Function
        End If
        If Not private_TryCreateDataExporter( _
            exporterClassName, exportConfigTable, exporter) Then Exit Function
        Set wordExporter = exporter
        If Not wordExporter.GetRecordBookmarks( _
            orderNoText, wordEvents) Then Exit Function

        If Not private_AppendDeletedExportSnapshots( _
            movementEvents, wordEvents) Then Exit Function

        movementEventCount = movementEvents.Count
        wordEventCount = wordEvents.Count
    End If

    If Not private_BuildCombinedExportedEventsTableItems( _
        movementEvents, wordEvents, exportedEventsTableItems, _
        m_MovementEventIds, m_MovementEventCaptions, _
        m_WordEventIds, m_WordEventCaptions) Then Exit Function

    If Not runtimeSources.TryGetItemsSourceByKey( _
        VBA.LCase$(MOVEMENT_EVENTS_RUNTIME_KEY), _
        existingMovementEvents, True) Then Exit Function
    If existingMovementEvents Is Nothing Then
        If Not runtimeSources.SetItemsSource( _
            VBA.LCase$(MOVEMENT_EVENTS_RUNTIME_KEY), _
            exportedEventsTableItems, notifyChange) Then Exit Function
    Else
        If Not private_SyncEventMenuItems( _
            existingMovementEvents, exportedEventsTableItems) Then Exit Function
    End If
    If loadEvents Then
        m_AreExportedEventsShown = True
        rt_Messaging.fn_ShowStatusBarSuccess _
            "Exported events: Movement=" & VBA.CStr(movementEventCount) & _
            "; WORD=" & VBA.CStr(wordEventCount), 5
    End If
    private_RegisterExportedEventMenus = True
End Function

Private Function private_AppendDeletedExportSnapshots( _
    ByVal movementEvents As Collection, _
    ByVal wordEvents As Collection _
) As Boolean
    Dim snapshotKey As Variant
    Dim snapshot As Object
    Dim targetEvents As Collection
    Dim optionObj As obj_SelectOption
    Dim optionItem As Variant
    Dim eventExists As Boolean
    Dim staleSnapshotKeys As Collection
    Dim staleSnapshotKey As Variant

    If movementEvents Is Nothing Or wordEvents Is Nothing Then Exit Function
    If m_DeletedExportSnapshots Is Nothing Then
        private_AppendDeletedExportSnapshots = True
        Exit Function
    End If
    Set staleSnapshotKeys = New Collection
    For Each snapshotKey In m_DeletedExportSnapshots.Keys
        Set snapshot = m_DeletedExportSnapshots(snapshotKey)
        If snapshot Is Nothing Then Exit Function
        If Not snapshot.Exists("Kind") Or _
            Not snapshot.Exists("EventId") Or _
            Not snapshot.Exists("OptionCaption") Then Exit Function
        If VBA.StrComp(VBA.CStr(snapshot("Kind")), _
            "Movement", VBA.vbTextCompare) = 0 Then
            Set targetEvents = movementEvents
        ElseIf VBA.StrComp(VBA.CStr(snapshot("Kind")), _
            "Word", VBA.vbTextCompare) = 0 Then
            Set targetEvents = wordEvents
        Else
            Exit Function
        End If
        eventExists = False
        For Each optionItem In targetEvents
            Set optionObj = optionItem
            If VBA.StrComp(optionObj.Id, VBA.CStr(snapshot("EventId")), _
                VBA.vbBinaryCompare) = 0 Then
                eventExists = True
                Exit For
            End If
        Next optionItem
        If Not eventExists Then
            Set optionObj = New obj_SelectOption
            optionObj.Id = VBA.CStr(snapshot("EventId"))
            optionObj.Caption = VBA.CStr(snapshot("OptionCaption"))
            targetEvents.Add optionObj
        Else
            ' WORD bookmark или Movement event снова существует физически:
            ' обычный повторный экспорт заменяет session undo и не должен
            ' оставаться красным tombstone.
            staleSnapshotKeys.Add VBA.CStr(snapshotKey)
        End If
    Next snapshotKey
    For Each staleSnapshotKey In staleSnapshotKeys
        If m_DeletedExportSnapshots.Exists(VBA.CStr(staleSnapshotKey)) Then _
            m_DeletedExportSnapshots.Remove VBA.CStr(staleSnapshotKey)
    Next staleSnapshotKey
    private_AppendDeletedExportSnapshots = True
End Function

Private Function private_BuildCombinedExportedEventsTableItems( _
    ByVal movementOptions As Collection, _
    ByVal wordOptions As Collection, _
    ByRef outTableItems As Collection, _
    ByRef outMovementIds As Collection, _
    ByRef outMovementCaptions As Collection, _
    ByRef outWordIds As Collection, _
    ByRef outWordCaptions As Collection _
) As Boolean
    Dim eventsTable As obj_TableDynamic
    Dim movementByIpn As Object
    Dim wordByIpn As Object
    Dim wordEventCache As Object
    Dim ipnOrder As Collection
    Dim sortedIpnOrder As Collection
    Dim optionItem As Variant
    Dim optionObj As obj_SelectOption
    Dim valueParts As Variant
    Dim ipnKey As String

    Set outTableItems = New Collection
    Set outMovementIds = New Collection
    Set outMovementCaptions = New Collection
    Set outWordIds = New Collection
    Set outWordCaptions = New Collection
    If movementOptions Is Nothing Or wordOptions Is Nothing Then Exit Function
    Set movementByIpn = VBA.CreateObject("Scripting.Dictionary")
    Set wordByIpn = VBA.CreateObject("Scripting.Dictionary")
    Set wordEventCache = VBA.CreateObject("Scripting.Dictionary")
    movementByIpn.CompareMode = VBA.vbTextCompare
    wordByIpn.CompareMode = VBA.vbTextCompare
    wordEventCache.CompareMode = VBA.vbTextCompare
    Set ipnOrder = New Collection

    For Each optionItem In movementOptions
        Set optionObj = optionItem
        valueParts = VBA.Split(optionObj.Caption, VBA.vbTab)
        If UBound(valueParts) <> 3 Then GoTo InvalidCaption
        ipnKey = VBA.Trim$(VBA.CStr(valueParts(3)))
        If VBA.Len(ipnKey) = 0 Then GoTo MissingMovementIpn
        If Not private_AddExportedEventToIpnGroup( _
            movementByIpn, ipnOrder, ipnKey, optionObj) Then Exit Function
    Next optionItem
    For Each optionItem In wordOptions
        Set optionObj = optionItem
        valueParts = VBA.Split(optionObj.Caption, VBA.vbTab)
        If UBound(valueParts) <> 3 Then GoTo InvalidCaption
        ipnKey = VBA.Trim$(VBA.CStr(valueParts(3)))
        If VBA.Len(ipnKey) = 0 Then GoTo MissingWordIpn
        If Not private_AddExportedEventToIpnGroup( _
            wordByIpn, ipnOrder, ipnKey, optionObj) Then Exit Function
    Next optionItem
    If Not private_TrySortExportedEventIpnGroups( _
        ipnOrder, movementByIpn, sortedIpnOrder) Then Exit Function

    Set eventsTable = New obj_TableDynamic
    eventsTable.SectionTitle = "Exported events — Movement (" & _
        VBA.CStr(movementOptions.Count) & ") / WORD (" & _
        VBA.CStr(wordOptions.Count) & ")"
    If Not private_AddCombinedExportedEventColumns(eventsTable) Then Exit Function
    For Each optionItem In sortedIpnOrder
        If Not private_AppendCombinedIpnRows( _
            VBA.CStr(optionItem), movementByIpn, wordByIpn, eventsTable, _
            wordEventCache, _
            outMovementIds, outMovementCaptions, _
            outWordIds, outWordCaptions) Then Exit Function
    Next optionItem
    outTableItems.Add eventsTable
    private_BuildCombinedExportedEventsTableItems = True
    Exit Function

InvalidCaption:
    VBA.MsgBox "Некорректная структура экспортированного события: " & _
        optionObj.Caption, VBA.vbExclamation, _
        "PrsnlEventBuilder / exported events"
    Exit Function
MissingMovementIpn:
    VBA.MsgBox "У события Movement отсутствует ИПН. Невозможно сгруппировать " & _
        "общую таблицу.", VBA.vbExclamation, _
        "PrsnlEventBuilder / exported events"
    Exit Function
MissingWordIpn:
    VBA.MsgBox "Не удалось извлечь ИПН из WORD-закладки '" & _
        optionObj.Id & "'. Невозможно сгруппировать общую таблицу.", _
        VBA.vbExclamation, "PrsnlEventBuilder / exported events"
End Function

Private Function private_TrySortExportedEventIpnGroups( _
    ByVal sourceIpnOrder As Collection, _
    ByVal movementByIpn As Object, _
    ByRef outSortedIpnOrder As Collection _
) As Boolean
    Dim ipnItem As Variant
    Dim ipnKey As String
    Dim groupSortRank As Long
    Dim passIndex As Long
    Dim movementItems As Collection

    Set outSortedIpnOrder = New Collection
    If sourceIpnOrder Is Nothing Or movementByIpn Is Nothing Then Exit Function
    ' Приоритет Movement: прибытия, выбытия, смены статуса.
    ' Исходный порядок людей внутри каждого класса сохраняется.
    For passIndex = 1 To 3
        For Each ipnItem In sourceIpnOrder
            ipnKey = VBA.CStr(ipnItem)
            If movementByIpn.Exists(ipnKey) Then
                Set movementItems = movementByIpn(ipnKey)
                If Not private_TryGetMovementGroupSortRank( _
                    movementItems, groupSortRank) Then Exit Function
                If groupSortRank = passIndex Then outSortedIpnOrder.Add ipnKey
            End If
        Next ipnItem
    Next passIndex
    ' WORD-пункты без соответствующего Movement не теряются и идут последними.
    For Each ipnItem In sourceIpnOrder
        ipnKey = VBA.CStr(ipnItem)
        If Not movementByIpn.Exists(ipnKey) Then _
            outSortedIpnOrder.Add ipnKey
    Next ipnItem
    private_TrySortExportedEventIpnGroups = True
End Function

Private Function private_TryGetMovementGroupSortRank( _
    ByVal movementItems As Collection, _
    ByRef outSortRank As Long _
) As Boolean
    Dim optionItem As Variant
    Dim optionObj As obj_SelectOption
    Dim valueParts As Variant
    Dim itemSortRank As Long

    outSortRank = 0
    If movementItems Is Nothing Or movementItems.Count = 0 Then Exit Function
    For Each optionItem In movementItems
        Set optionObj = optionItem
        valueParts = VBA.Split(optionObj.Caption, VBA.vbTab)
        If UBound(valueParts) < 0 Then Exit Function
        If Not private_TryGetMovementDirectionSortRank( _
            VBA.CStr(valueParts(0)), itemSortRank) Then Exit Function
        If outSortRank = 0 Or itemSortRank < outSortRank Then _
            outSortRank = itemSortRank
    Next optionItem
    private_TryGetMovementGroupSortRank = (outSortRank > 0)
End Function

Private Function private_TryGetMovementDirectionSortRank( _
    ByVal directionText As String, _
    ByRef outSortRank As Long _
) As Boolean
    outSortRank = 0
    Select Case VBA.LCase$(VBA.Trim$(directionText))
        Case VBA.LCase$("Прибуття")
            outSortRank = 1
        Case VBA.LCase$("Вибуття")
            outSortRank = 2
        Case VBA.LCase$("Зміна статусу")
            outSortRank = 3
        Case Else
            VBA.MsgBox "Неизвестный тип события Movement: " & directionText, _
                VBA.vbExclamation, "PrsnlEventBuilder / exported events"
            Exit Function
    End Select
    private_TryGetMovementDirectionSortRank = True
End Function

Private Function private_IpnGroupHasMovementDirection( _
    ByVal movementByIpn As Object, _
    ByVal ipnKey As String, _
    ByVal expectedDirection As String _
) As Boolean
    Dim movementItems As Collection
    Dim optionItem As Variant
    Dim optionObj As obj_SelectOption
    Dim valueParts As Variant

    If movementByIpn Is Nothing Then Exit Function
    If Not movementByIpn.Exists(ipnKey) Then Exit Function
    Set movementItems = movementByIpn(ipnKey)
    For Each optionItem In movementItems
        Set optionObj = optionItem
        valueParts = VBA.Split(optionObj.Caption, VBA.vbTab)
        If UBound(valueParts) >= 0 Then
            If VBA.StrComp(VBA.Trim$(VBA.CStr(valueParts(0))), _
                expectedDirection, VBA.vbTextCompare) = 0 Then
                private_IpnGroupHasMovementDirection = True
                Exit Function
            End If
        End If
    Next optionItem
End Function

Private Function private_AddExportedEventToIpnGroup( _
    ByVal groupsByIpn As Object, _
    ByVal ipnOrder As Collection, _
    ByVal ipnText As String, _
    ByVal optionObj As obj_SelectOption _
) As Boolean
    Dim groupItems As Collection
    Dim groupKey As String

    If groupsByIpn Is Nothing Or ipnOrder Is Nothing Or optionObj Is Nothing Then Exit Function
    groupKey = VBA.LCase$(VBA.Trim$(ipnText))
    If VBA.Len(groupKey) = 0 Then Exit Function
    If Not groupsByIpn.Exists(groupKey) Then
        Set groupItems = New Collection
        groupsByIpn.Add groupKey, groupItems
        If Not private_CollectionContainsText(ipnOrder, groupKey) Then ipnOrder.Add groupKey
    Else
        Set groupItems = groupsByIpn(groupKey)
    End If
    groupItems.Add optionObj
    private_AddExportedEventToIpnGroup = True
End Function

Private Function private_CollectionContainsText( _
    ByVal sourceItems As Collection, ByVal expectedText As String _
) As Boolean
    Dim itemValue As Variant
    For Each itemValue In sourceItems
        If VBA.StrComp(VBA.CStr(itemValue), expectedText, VBA.vbTextCompare) = 0 Then
            private_CollectionContainsText = True
            Exit Function
        End If
    Next itemValue
End Function

Private Function private_AddCombinedExportedEventColumns( _
    ByVal eventsTable As obj_TableDynamic _
) As Boolean
    Dim columnNames As Variant
    Dim columnAliases As Variant
    Dim columnIndex As Long
    Dim eventColumn As obj_Column

    columnNames = VBA.Array("Тип", "Подія", "ПІБ", "ІПН", _
        "WORD", "Подія WORD", "Bookmark", "Текст пункту")
    columnAliases = VBA.Array("movement.type", "movement.description", _
        "movement.person", "movement.ipn", "word.source", _
        "word.event", "word.bookmark", "word.description")
    For columnIndex = LBound(columnNames) To UBound(columnNames)
        Set eventColumn = New obj_Column
        eventColumn.Name = VBA.CStr(columnNames(columnIndex))
        If Not eventColumn.AddAlias(VBA.CStr(columnAliases(columnIndex))) Then Exit Function
        If Not eventsTable.PushColumn(eventColumn) Then Exit Function
    Next columnIndex
    private_AddCombinedExportedEventColumns = True
End Function

Private Function private_AppendCombinedIpnRows( _
    ByVal ipnKey As String, _
    ByVal movementByIpn As Object, _
    ByVal wordByIpn As Object, _
    ByVal eventsTable As obj_TableDynamic, _
    ByVal wordEventCache As Object, _
    ByVal movementIds As Collection, _
    ByVal movementCaptions As Collection, _
    ByVal wordIds As Collection, _
    ByVal wordCaptions As Collection _
) As Boolean
    Dim movementItems As Collection
    Dim sortedMovementItems As Collection
    Dim wordItems As Collection
    Dim movementOption As obj_SelectOption
    Dim wordOption As obj_SelectOption
    Dim movementParts As Variant
    Dim wordParts As Variant
    Dim eventRow As obj_Row
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim wordOrder As Collection
    Dim wordItemIndex As Long
    Dim wordEventText As String

    Set movementItems = New Collection
    Set wordItems = New Collection
    If movementByIpn.Exists(ipnKey) Then Set movementItems = movementByIpn(ipnKey)
    If wordByIpn.Exists(ipnKey) Then Set wordItems = wordByIpn(ipnKey)
    If Not private_TrySortMovementGroupItems( _
        movementItems, sortedMovementItems) Then Exit Function
    If Not private_TryBuildAlignedWordItemOrder( _
        sortedMovementItems, wordItems, wordOrder) Then Exit Function
    rowCount = wordOrder.Count
    For rowIndex = 1 To rowCount
        Set eventRow = New obj_Row
        If rowIndex <= sortedMovementItems.Count Then
            Set movementOption = sortedMovementItems.Item(rowIndex)
            movementParts = VBA.Split(movementOption.Caption, VBA.vbTab)
            eventRow.PushCellRaw VBA.CStr(movementParts(0))
            eventRow.PushCellRaw VBA.CStr(movementParts(2))
            eventRow.PushCellRaw VBA.CStr(movementParts(1))
            eventRow.PushCellRaw VBA.CStr(movementParts(3))
            movementIds.Add movementOption.Id
            movementCaptions.Add VBA.Replace(movementOption.Caption, VBA.vbTab, " — ")
            If private_IsDeletedExportSnapshot(True, movementOption.Id) Then _
                private_SetRowCellDesc eventRow, 1, 4, "diff:deleted"
        Else
            eventRow.PushCellRaw VBA.vbNullString
            eventRow.PushCellRaw VBA.vbNullString
            eventRow.PushCellRaw VBA.vbNullString
            eventRow.PushCellRaw VBA.vbNullString
            movementIds.Add "NONE"
            movementCaptions.Add VBA.vbNullString
        End If
        wordItemIndex = VBA.CLng(wordOrder.Item(rowIndex))
        If wordItemIndex > 0 Then
            Set wordOption = wordItems.Item(wordItemIndex)
            wordParts = VBA.Split(wordOption.Caption, VBA.vbTab)
            If Not private_TryResolveWordEventText( _
                wordOption.Id, wordEventCache, wordEventText) Then Exit Function
            eventRow.PushCellRaw VBA.CStr(wordParts(0))
            eventRow.PushCellRaw wordEventText
            eventRow.PushCellRaw VBA.CStr(wordParts(1))
            eventRow.PushCellRaw VBA.CStr(wordParts(2))
            wordIds.Add wordOption.Id
            wordCaptions.Add VBA.Replace(wordOption.Caption, VBA.vbTab, " — ")
            If private_IsDeletedExportSnapshot(False, wordOption.Id) Then _
                private_SetRowCellDesc eventRow, 5, 8, "diff:deleted"
        Else
            eventRow.PushCellRaw VBA.vbNullString
            eventRow.PushCellRaw VBA.vbNullString
            eventRow.PushCellRaw VBA.vbNullString
            eventRow.PushCellRaw VBA.vbNullString
            wordIds.Add "NONE"
            wordCaptions.Add VBA.vbNullString
        End If
        If Not eventsTable.PushRow(eventRow) Then Exit Function
    Next rowIndex
    private_AppendCombinedIpnRows = True
End Function

Private Function private_TryResolveWordEventText( _
    ByVal bookmarkName As String, _
    ByVal wordEventCache As Object, _
    ByRef outEventText As String _
) As Boolean
    Dim bookmarkParts As Variant
    Dim separatorIndex As Long
    Dim templateId As String
    Dim cacheKey As String

    outEventText = VBA.vbNullString
    bookmarkName = VBA.Trim$(bookmarkName)
    If VBA.Len(bookmarkName) = 0 Or m_Data Is Nothing Then Exit Function
    bookmarkParts = VBA.Split(bookmarkName, "_")
    If UBound(bookmarkParts) >= 2 And _
        VBA.StrComp(VBA.CStr(bookmarkParts(0)), "PEB", _
            VBA.vbTextCompare) = 0 And _
        VBA.Len(VBA.CStr(bookmarkParts(1))) = 2 Then
        If m_CachedWordExporter Is Nothing Then GoTo MissingExporter
        cacheKey = VBA.UCase$(VBA.CStr(bookmarkParts(1)))
        If Not wordEventCache Is Nothing Then
            If wordEventCache.Exists(cacheKey) Then
                outEventText = VBA.CStr(wordEventCache(cacheKey))
                private_TryResolveWordEventText = True
                Exit Function
            End If
        End If
        If Not m_CachedWordExporter.TryGetTemplateNameByHash( _
            VBA.CStr(bookmarkParts(1)), templateId) Then Exit Function
    Else
        separatorIndex = VBA.InStrRev(bookmarkName, "_", -1, VBA.vbBinaryCompare)
        If separatorIndex <= 5 Or _
            VBA.StrComp(VBA.Left$(bookmarkName, 4), "PEB_", _
                VBA.vbTextCompare) <> 0 Then GoTo InvalidBookmark
        templateId = VBA.Mid$(bookmarkName, 5, separatorIndex - 5)
        cacheKey = "LEGACY:" & VBA.LCase$(templateId)
    End If
    If m_Data.TryResolveSectionTypeByWordTemplateId( _
        templateId, outEventText) Then
        If Not wordEventCache Is Nothing Then _
            wordEventCache(cacheKey) = outEventText
        private_TryResolveWordEventText = True
        Exit Function
    End If
    VBA.MsgBox "Для WORD-шаблона '" & templateId & _
        "' не найден тип события.", VBA.vbExclamation, _
        "PrsnlEventBuilder / WORD events"
    Exit Function
MissingExporter:
    VBA.MsgBox "WORD exporter не инициализирован. Невозможно определить " & _
        "событие по закладке '" & bookmarkName & "'.", VBA.vbExclamation, _
        "PrsnlEventBuilder / WORD events"
    Exit Function
InvalidBookmark:
    VBA.MsgBox "Некорректное имя WORD-закладки: '" & bookmarkName & "'.", _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
End Function

Private Function private_TryBuildAlignedWordItemOrder( _
    ByVal movementItems As Collection, _
    ByVal wordItems As Collection, _
    ByRef outWordOrder As Collection _
) As Boolean
    Dim usedWordIndexes As Object
    Dim movementItem As Variant
    Dim movementOption As obj_SelectOption
    Dim movementParts As Variant
    Dim templateId As String
    Dim wordIndex As Long
    Dim matchedWordIndex As Long

    Set outWordOrder = New Collection
    If movementItems Is Nothing Or wordItems Is Nothing Then Exit Function
    Set usedWordIndexes = VBA.CreateObject("Scripting.Dictionary")
    For Each movementItem In movementItems
        Set movementOption = movementItem
        movementParts = VBA.Split(movementOption.Caption, VBA.vbTab)
        If UBound(movementParts) <> 3 Then Exit Function
        matchedWordIndex = 0
        templateId = VBA.vbNullString
        ' Для единственной пары одного ИПН дополнительная семантическая
        ' развязка не нужна и не должна разносить записи по двум строкам.
        If movementItems.Count = 1 And wordItems.Count = 1 Then
            matchedWordIndex = 1
            usedWordIndexes.Add "1", True
            GoTo AddMatchedWordIndex
        End If
        If Not private_TryResolveWordTemplateForMovementOption( _
            VBA.CStr(movementParts(0)), VBA.CStr(movementParts(2)), _
            templateId) Then templateId = VBA.vbNullString
        If VBA.Len(templateId) > 0 Then
            For wordIndex = 1 To wordItems.Count
                If Not usedWordIndexes.Exists(VBA.CStr(wordIndex)) Then
                    If private_WordOptionMatchesTemplate( _
                        wordItems.Item(wordIndex), templateId) Then
                        matchedWordIndex = wordIndex
                        usedWordIndexes.Add VBA.CStr(wordIndex), True
                        Exit For
                    End If
                End If
            Next wordIndex
        End If
AddMatchedWordIndex:
        outWordOrder.Add matchedWordIndex
    Next movementItem
    ' WORD без соответствующего Movement сохраняются отдельными строками.
    For wordIndex = 1 To wordItems.Count
        If Not usedWordIndexes.Exists(VBA.CStr(wordIndex)) Then _
            outWordOrder.Add wordIndex
    Next wordIndex
    private_TryBuildAlignedWordItemOrder = True
End Function

Private Function private_TryResolveWordTemplateForMovementOption( _
    ByVal directionText As String, _
    ByVal movementEventText As String, _
    ByRef outTemplateId As String _
) As Boolean
    Dim sectionTypes As Collection
    Dim sectionTypeItem As Variant
    Dim sectionTypeText As String
    Dim mappedEventText As String
    Dim previousEventText As String
    Dim movementParts As Variant
    Dim sectionParts As Variant

    outTemplateId = VBA.vbNullString
    If m_Data Is Nothing Then Exit Function
    Set sectionTypes = m_Data.SectionTypeNames
    If sectionTypes Is Nothing Then Exit Function
    For Each sectionTypeItem In sectionTypes
        sectionTypeText = VBA.CStr(sectionTypeItem)
        mappedEventText = VBA.vbNullString
        previousEventText = VBA.vbNullString
        Select Case VBA.LCase$(VBA.Trim$(directionText))
            Case VBA.LCase$("Прибуття")
                If m_Data.TryGetRequiredPreviousMovementEvent( _
                    sectionTypeText, previousEventText) Then
                    If VBA.StrComp(VBA.Trim$(previousEventText), _
                        VBA.Trim$(movementEventText), VBA.vbTextCompare) <> 0 Then _
                        GoTo ContinueSection
                Else
                    GoTo ContinueSection
                End If
            Case VBA.LCase$("Вибуття")
                If m_Data.TryMapMovementSectionTypeToEventText( _
                    sectionTypeText, mappedEventText) Then
                    If VBA.StrComp(VBA.Trim$(mappedEventText), _
                        VBA.Trim$(movementEventText), VBA.vbTextCompare) <> 0 Then _
                        GoTo ContinueSection
                Else
                    GoTo ContinueSection
                End If
            Case VBA.LCase$("Зміна статусу")
                If Not m_Data.IsMovementMirrorTransferSectionType( _
                    sectionTypeText) Then GoTo ContinueSection
                movementParts = VBA.Split(movementEventText, "=>")
                sectionParts = VBA.Split(sectionTypeText, "=>")
                If UBound(movementParts) <> 1 Or UBound(sectionParts) <> 1 Then _
                    GoTo ContinueSection
                If VBA.StrComp(private_NormalizeMovementStatusName( _
                    VBA.CStr(movementParts(0))), _
                    private_NormalizeMovementStatusName( _
                    VBA.CStr(sectionParts(0))), VBA.vbTextCompare) <> 0 Then _
                    GoTo ContinueSection
                If VBA.StrComp(private_NormalizeMovementStatusName( _
                    VBA.CStr(movementParts(1))), _
                    private_NormalizeMovementStatusName( _
                    VBA.CStr(sectionParts(1))), VBA.vbTextCompare) <> 0 Then _
                    GoTo ContinueSection
            Case Else
                Exit Function
        End Select
        If m_Data.TryResolveWordTemplateId(sectionTypeText, outTemplateId) Then
            private_TryResolveWordTemplateForMovementOption = True
            Exit Function
        End If
ContinueSection:
    Next sectionTypeItem
End Function

Private Function private_NormalizeMovementStatusName( _
    ByVal statusText As String _
) As String
    statusText = VBA.LCase$(VBA.Trim$(statusText))
    statusText = VBA.Replace(statusText, "стаціонарне лікування", "лікування")
    Do While VBA.InStr(1, statusText, "  ", VBA.vbBinaryCompare) > 0
        statusText = VBA.Replace(statusText, "  ", " ")
    Loop
    private_NormalizeMovementStatusName = VBA.Trim$(statusText)
End Function

Private Function private_WordOptionMatchesTemplate( _
    ByVal wordOption As obj_SelectOption, _
    ByVal expectedTemplateId As String _
) As Boolean
    Dim bookmarkName As String
    Dim separatorIndex As Long
    Dim bookmarkTemplateId As String
    Dim expectedTemplateHash As String
    Dim bookmarkParts As Variant

    If wordOption Is Nothing Then Exit Function
    bookmarkName = VBA.CStr(wordOption.Id)
    bookmarkParts = VBA.Split(bookmarkName, "_")
    If UBound(bookmarkParts) >= 2 Then
        If VBA.StrComp(VBA.CStr(bookmarkParts(0)), "PEB", _
            VBA.vbTextCompare) = 0 And _
            VBA.Len(VBA.CStr(bookmarkParts(1))) = 2 Then
            bookmarkTemplateId = VBA.CStr(bookmarkParts(1))
            If m_CachedWordExporter Is Nothing Then Exit Function
            If Not m_CachedWordExporter.TryGetTemplateHashByName( _
                expectedTemplateId, expectedTemplateHash) Then Exit Function
            private_WordOptionMatchesTemplate = (VBA.StrComp( _
                bookmarkTemplateId, expectedTemplateHash, _
                VBA.vbTextCompare) = 0)
            Exit Function
        End If
    End If
    separatorIndex = VBA.InStrRev(bookmarkName, "_", -1, VBA.vbBinaryCompare)
    If separatorIndex <= 5 Then Exit Function
    If VBA.StrComp(VBA.Left$(bookmarkName, 4), "PEB_", _
        VBA.vbTextCompare) <> 0 Then Exit Function
    bookmarkTemplateId = VBA.Mid$(bookmarkName, 5, separatorIndex - 5)
    If VBA.Len(bookmarkTemplateId) = 0 Then Exit Function
    ' Имя bookmark ограничено Word и длинный templateId обрезается перед ИПН.
    private_WordOptionMatchesTemplate = (VBA.StrComp( _
        VBA.Left$(expectedTemplateId, VBA.Len(bookmarkTemplateId)), _
        bookmarkTemplateId, VBA.vbTextCompare) = 0)
End Function

Private Function private_IsDeletedExportSnapshot( _
    ByVal isMovementEvent As Boolean, _
    ByVal eventId As String _
) As Boolean
    If m_DeletedExportSnapshots Is Nothing Then Exit Function
    private_IsDeletedExportSnapshot = m_DeletedExportSnapshots.Exists( _
        private_DeletedExportSnapshotKey(isMovementEvent, eventId))
End Function

Private Sub private_SetRowCellDesc( _
    ByVal eventRow As obj_Row, _
    ByVal firstColumnIndex As Long, _
    ByVal lastColumnIndex As Long, _
    ByVal descText As String _
)
    Dim columnIndex As Long
    Dim eventCell As obj_Cell

    If eventRow Is Nothing Then Exit Sub
    For columnIndex = firstColumnIndex To lastColumnIndex
        Set eventCell = eventRow.Cells.Item(columnIndex)
        If eventCell Is Nothing Then Exit Sub
        eventCell.Desc = descText
    Next columnIndex
End Sub

Private Function private_TrySortMovementGroupItems( _
    ByVal sourceItems As Collection, _
    ByRef outSortedItems As Collection _
) As Boolean
    Dim optionItem As Variant
    Dim optionObj As obj_SelectOption
    Dim valueParts As Variant
    Dim itemSortRank As Long
    Dim passIndex As Long

    Set outSortedItems = New Collection
    If sourceItems Is Nothing Then Exit Function
    ' Исходный порядок внутри каждого из трёх направлений сохраняется.
    For passIndex = 1 To 3
        For Each optionItem In sourceItems
            Set optionObj = optionItem
            valueParts = VBA.Split(optionObj.Caption, VBA.vbTab)
            If UBound(valueParts) < 0 Then Exit Function
            If Not private_TryGetMovementDirectionSortRank( _
                VBA.CStr(valueParts(0)), itemSortRank) Then Exit Function
            If itemSortRank = passIndex Then outSortedItems.Add optionObj
        Next optionItem
    Next passIndex
    private_TrySortMovementGroupItems = True
End Function

Private Function private_BuildExportedEventsTableItems( _
    ByVal sectionName As String, _
    ByVal eventOptions As Collection, _
    ByRef outTableItems As Collection, _
    ByRef outEventIds As Collection, _
    ByRef outEventCaptions As Collection _
) As Boolean
    Dim eventsTable As obj_TableDynamic
    Dim eventColumn As obj_Column
    Dim eventRow As obj_Row
    Dim optionObj As obj_SelectOption
    Dim optionItem As Variant
    Dim eventCount As Long
    Dim valueParts As Variant
    Dim columnNames As Variant
    Dim columnAliases As Variant
    Dim columnIndex As Long
    Dim expectedPartCount As Long

    Set outTableItems = New Collection
    Set outEventIds = New Collection
    Set outEventCaptions = New Collection
    If eventOptions Is Nothing Then Exit Function
    eventCount = eventOptions.Count

    Set eventsTable = New obj_TableDynamic
    eventsTable.SectionTitle = sectionName & " (" & VBA.CStr(eventCount) & ")"
    If VBA.InStr(1, sectionName, "Movement", VBA.vbTextCompare) > 0 Then
        columnNames = VBA.Array("Тип", "ПІБ", "Подія", "ІПН")
        columnAliases = VBA.Array("event.type", "event.person", _
            "event.description", "event.ipn")
    Else
        columnNames = VBA.Array("Джерело", "Bookmark", "Текст пункту")
        columnAliases = VBA.Array("event.source", "event.bookmark", _
            "event.description")
    End If
    expectedPartCount = UBound(columnNames) - LBound(columnNames) + 1
    For columnIndex = LBound(columnNames) To UBound(columnNames)
        Set eventColumn = New obj_Column
        eventColumn.Name = VBA.CStr(columnNames(columnIndex))
        If Not eventColumn.AddAlias(VBA.CStr( _
            columnAliases(columnIndex))) Then Exit Function
        If Not eventsTable.PushColumn(eventColumn) Then Exit Function
    Next columnIndex
    For Each optionItem In eventOptions
        Set optionObj = optionItem
        valueParts = VBA.Split(optionObj.Caption, VBA.vbTab)
        If UBound(valueParts) - LBound(valueParts) + 1 <> expectedPartCount Then
            VBA.MsgBox "Некорректная структура экспортированного события: " & _
                optionObj.Caption, VBA.vbExclamation, _
                "PrsnlEventBuilder / exported events"
            Exit Function
        End If
        Set eventRow = New obj_Row
        For columnIndex = LBound(valueParts) To UBound(valueParts)
            eventRow.PushCellRaw VBA.CStr(valueParts(columnIndex))
        Next columnIndex
        If Not eventsTable.PushRow(eventRow) Then Exit Function
        outEventIds.Add optionObj.Id
        outEventCaptions.Add VBA.Replace(optionObj.Caption, VBA.vbTab, " — ")
    Next optionItem
    outTableItems.Add eventsTable
    private_BuildExportedEventsTableItems = True
End Function

Private Function private_SyncEventMenuItems( _
    ByVal targetItems As Collection, _
    ByVal sourceItems As Collection _
) As Boolean
    Dim itemObj As Variant

    If targetItems Is Nothing Then Exit Function
    If sourceItems Is Nothing Then Exit Function
    Do While targetItems.Count > 0
        targetItems.Remove targetItems.Count
    Loop
    For Each itemObj In sourceItems
        targetItems.Add itemObj
    Next itemObj
    private_SyncEventMenuItems = True
End Function

Private Function private_RegisterProfileOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim profiles As Collection
    Dim profileOptions As Collection

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set profiles = m_Data.PrimaryProfileNames
    If profiles Is Nothing Then Exit Function
    If profiles.Count = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_SelectedProfile)) = 0 Then m_SelectedProfile = VBA.Trim$(VBA.CStr(profiles.Item(1)))
    If VBA.Len(VBA.Trim$(m_SelectedMainProfile)) = 0 Then m_SelectedMainProfile = VBA.Trim$(m_SelectedProfile)
    If Not private_TryBuildOptionButtonRows( _
        profiles, _
        m_SelectedMainProfile, _
        PROFILE_BUTTON_TAG_ARRIVAL, _
        PROFILE_BUTTON_STATE_SELECTED, _
        True, _
        2, _
        profileOptions) Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(PROFILES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(PROFILES_RUNTIME_KEY), profileOptions, notifyChange) Then Exit Function

    private_RegisterProfileOptions = True
End Function

Private Function private_RegisterAdditionalProfileOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim profiles As Collection
    Dim profileOptions As Collection
    Dim profileObj As Variant
    Dim profileText As String
    Dim optionObj As obj_SelectOption

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set profiles = m_Data.AdditionalProfileNames
    If profiles Is Nothing Then Exit Function
    Set profileOptions = New Collection

    For Each profileObj In profiles
        profileText = VBA.Trim$(VBA.CStr(profileObj))
        If VBA.Len(profileText) = 0 Then GoTo ContinueProfile
        Set optionObj = New obj_SelectOption
        optionObj.Caption = profileText
        optionObj.Id = profileText
        profileOptions.Add optionObj
ContinueProfile:
    Next profileObj
    If profileOptions.Count = 0 Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(ADDITIONAL_PROFILES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource( _
        VBA.LCase$(ADDITIONAL_PROFILES_RUNTIME_KEY), profileOptions, notifyChange) Then Exit Function

    private_RegisterAdditionalProfileOptions = True
End Function

Private Function private_TrySyncAdditionalProfileSelectState() As Boolean
    Dim pageBase As obj_PageBase
    Dim selectState As obj_SelectControlVMStatic
    Dim selectKey As String
    Dim selectedId As String

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.Worksheet Is Nothing Then Exit Function

    If m_Data.IsAdditionalProfileName(m_SelectedProfile) Then
        selectedId = VBA.Trim$(m_SelectedProfile)
    Else
        selectedId = VBA.vbNullString
    End If
    selectKey = VBA.LCase$(pageBase.Worksheet.Name & "|" & _
        ADDITIONAL_PROFILE_SELECT_CONTROL_NAME)
    Set selectState = New obj_SelectControlVMStatic
    private_TrySyncAdditionalProfileSelectState = _
        selectState.SetSelectedId(selectKey, selectedId)
End Function

Private Function private_TryResetAdditionalProfileSelect() As Boolean
    Dim pageBase As obj_PageBase
    Dim rawControl As Object
    Dim selectControl As obj_SelectControlVM

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    Set rawControl = Nothing
    If Not pageBase.TryGetRegisteredControlByName( _
        ADDITIONAL_PROFILE_SELECT_CONTROL_NAME, rawControl) Then Exit Function
    If rawControl Is Nothing Then Exit Function
    If Not TypeOf rawControl Is obj_SelectControlVM Then Exit Function

    Set selectControl = rawControl
    private_TryResetAdditionalProfileSelect = selectControl.ResetSelection()
End Function

Private Function private_RegisterMetaProfileOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim profiles As Collection
    Dim profileOptions As Collection

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set profiles = m_Data.MetaProfileNames
    If profiles Is Nothing Then Exit Function
    If Not private_TryBuildOptionButtonRows( _
        profiles, _
        m_SelectedProfile, _
        META_PROFILE_BUTTON_TAG_NORMAL, _
        PROFILE_BUTTON_STATE_SELECTED, _
        False, _
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

Private Function private_RegisterMovementHistoryTable( _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim movementHistoryTables As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set movementHistoryTables = New Collection
    If m_IsMovementHistoryEnabled Then
        If Not m_MovementHistoryTable Is Nothing Then
            movementHistoryTables.Add m_MovementHistoryTable
        End If
    End If

    If Not runtimeSources.RemoveItemsSource( _
        VBA.LCase$(MOVEMENT_HISTORY_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource( _
        VBA.LCase$(MOVEMENT_HISTORY_RUNTIME_KEY), _
        movementHistoryTables, _
        notifyChange) Then Exit Function

    private_RegisterMovementHistoryTable = True
End Function

Private Function private_TryBuildOptionButtonRows( _
    ByVal profiles As Collection, _
    ByVal selectedProfileText As String, _
    ByVal normalTagName As String, _
    ByVal selectedStateName As String, _
    ByVal useMovementDirectionTags As Boolean, _
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
            If Not optionObj.SetState(selectedStateName, True) Then Exit Function
        Else
            ' Контроллер сообщает только семантический тег; оформление задаётся в XML.
            If useMovementDirectionTags And Not m_Data.IsMovementClosingSectionType(profileText) Then
                If Not optionObj.AddTag(PROFILE_BUTTON_TAG_DEPARTURE) Then Exit Function
            Else
                If Not optionObj.AddTag(normalTagName) Then Exit Function
            End If
        End If
        rowItems.Add optionObj

ContinueProfile:
    Next profileObj

    private_TryBuildOptionButtonRows = True
End Function

#If LOGGING_DEBUG_ENABLED Then


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
    Dim candidateTable As obj_TableDynamic
    Dim candidateRow As obj_Row
    Dim formAliases As Collection
    Dim aliasObj As Variant
    Dim aliasText As String
    Dim candidateColumnIndex As Long
    Dim candidateValue As String
    Dim valueRange As Range
    Dim ignoredCfgParser As obj_EntityLookupCfgParser
    Dim ignoredLookupKey As String
    Dim ignoredSearchAlias As String
    Dim previousEnableEvents As Boolean
    Dim isReporterTvoCandidate As Boolean
    Dim isFioCandidate As Boolean

    If targetCell Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not (targetCell.Worksheet Is ws) Then Exit Function
    isReporterTvoCandidate = private_IsReporterTvoCandidatesContext()
    isFioCandidate = private_IsFioCandidatesContext()

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

    ' Берём выбранную строку из модели, а не из физических колонок листа.
    ' Динамический layout может скрывать поля и сдвигать candidate-table.
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetActiveCandidatesContext( _
        ignoredCfgParser, ignoredLookupKey, candidateTable, _
        ignoredSearchAlias) Then Exit Function
    If candidateTable Is Nothing Then Exit Function
    If rowOffset > candidateTable.RowCount Then Exit Function
    Set candidateRow = candidateTable.Rows.Item(rowOffset)
    If candidateRow Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetFormColumnKeys(formAliases) Then Exit Function
    If formAliases Is Nothing Then Exit Function

    ' Values are written directly into sheet cells. Disable events so this accept action
    ' does not recursively trigger input onChange/search/rerender for each copied cell.
    previousEnableEvents = Application.EnableEvents
    On Error GoTo RestoreEventsAndFail
    Application.EnableEvents = False
    For Each aliasObj In formAliases
        aliasText = VBA.Trim$(VBA.CStr(aliasObj))
        candidateColumnIndex = 0
        ' Не все поля формы входят в схему таблицы кандидатов.
        ' Сначала проверяем наличие колонки без диагностического MsgBox.
        If Not candidateRow.TryGetColumnIndex( _
            aliasText, candidateColumnIndex) Then GoTo ContinueAlias
        candidateValue = candidateRow.GetCellValue(candidateColumnIndex)
        If VBA.Len(VBA.Trim$(candidateValue)) = 0 Then GoTo ContinueAlias
        Set valueRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange( _
            aliasText, valueRange, "visible") Then GoTo ContinueAlias
        If valueRange Is Nothing Then GoTo ContinueAlias
        If Application.Intersect(valueRange, draftValuesRange) Is Nothing Then _
            GoTo ContinueAlias
        valueRange.Cells(1, 1).Value2 = candidateValue
ContinueAlias:
    Next aliasObj
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0

    If isReporterTvoCandidate Then
        m_IsReporterTvoCandidatesActive = False
    End If

    If Not private_RefreshDraftReporterTvoFlag() Then Exit Function
    If Not private_TryApplyReporterPositionCodeAppearance() Then Exit Function
    If isFioCandidate Then
        If Not private_TryApplyMovementEnrichedFieldsAppearance() Then Exit Function
    End If

    candidateRowRange.Select
    ' Range.Select по умолчанию делает активной крайнюю левую ячейку.
    ' Activate внутри уже выделенной строки сохраняет весь Selection и задаёт
    ' естественную стартовую позицию для следующего нажатия стрелки.
    targetCell.Activate
    private_TryAcceptCandidateRowFromSelection = True
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
End Function

Private Function private_TryActivateColumnWithinSelection( _
    ByVal preferredColumn As Long _
) As Boolean
    Dim selectionObj As Object
    Dim selectedRange As Range
    Dim activeColumnRange As Range

    On Error GoTo EH
    Set selectionObj = Application.Selection
    If Not TypeOf selectionObj Is Range Then
        VBA.MsgBox _
            "PrototypeNew: candidate selection is unavailable after layout reflow.", _
            VBA.vbExclamation, _
            "PrototypeNew / candidate selection"
        Exit Function
    End If
    Set selectedRange = selectionObj
    Set activeColumnRange = Application.Intersect( _
        selectedRange, _
        selectedRange.Worksheet.Columns(preferredColumn))
    If activeColumnRange Is Nothing Then
        VBA.MsgBox _
            "PrototypeNew: the previously active candidate column is outside the restored selection.", _
            VBA.vbExclamation, _
            "PrototypeNew / candidate selection"
        Exit Function
    End If

    activeColumnRange.Cells(1, 1).Activate
    private_TryActivateColumnWithinSelection = True
    Exit Function

EH:
    VBA.MsgBox _
        "PrototypeNew: failed to restore the active candidate cell after layout reflow." & _
        VBA.vbCrLf & "Error: " & Err.Description, _
        VBA.vbExclamation, _
        "PrototypeNew / candidate selection"
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
    Dim isMovementExport As Boolean

    On Error GoTo EH

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:start action='" & private_EscapeForLog(actionId) & "'"
#End If

    If Not private_TryResolveExportAliasFromAction(actionId, exportAlias) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:alias alias='" & private_EscapeForLog(exportAlias) & "'"
#End If
    If Not private_TryGetExportSettings(exportAlias, exporterClassName, exportConfigTable) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:settings class='" & private_EscapeForLog(exporterClassName) & "'"
#End If
    If Not private_TryBuildExportSourceTables(sourceTables, exportContext) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:source-ready tables=" & VBA.CStr(sourceTables.Count)
#End If

    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:exporter-ready type='" & private_EscapeForLog(VBA.TypeName(exporter)) & "'"
#End If

    isMovementExport = _
        (VBA.StrComp(exportAlias, "Movement", VBA.vbTextCompare) = 0)
    If isMovementExport Then
        If Not private_ConfirmPartialExport(sourceTables, "Movement") Then Exit Function
        private_ClearPendingMovementReceipt
    ElseIf VBA.StrComp(exportAlias, "Word", VBA.vbTextCompare) = 0 Then
        private_ApplyPendingMovementReceipt sourceTables, exportContext
    End If

    If Not exporter.Export(sourceTables, exportContext) Then Exit Function
    If isMovementExport Then
        If Not private_TryCaptureMovementReceipt( _
            sourceTables, exportContext) Then
            VBA.MsgBox "Movement export completed, but its validation receipt " & _
                "could not be created for the subsequent WORD export." & _
                VBA.vbCrLf & "WORD will use standalone Movement validation.", _
                VBA.vbExclamation, "PrototypeNew / Movement export"
            Exit Function
        End If
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:exporter-done alias='" & private_EscapeForLog(exportAlias) & "'"
#End If
    If Not private_TryCaptureWordExportPreview(exportContext) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:export-action:preview-captured alias='" & private_EscapeForLog(exportAlias) & "'"
#End If

    rt_Messaging.fn_ShowStatusBarSuccess EXPORT_ACTION_PREFIX & exportAlias & ": done", 3
    private_TryExportDraftByAction = True
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "prsnlevntbuilder:export-action:error action='" & private_EscapeForLog(actionId) & "' alias='" & private_EscapeForLog(exportAlias) & "' class='" & private_EscapeForLog(exporterClassName) & "' errNo=" & VBA.CStr(Err.Number) & " err='" & private_EscapeForLog(Err.Description) & "'"
#End If
End Function

Private Function private_ConfirmPartialExport( _
    ByVal sourceTables As Collection, _
    ByVal exportCaption As String _
) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim ipnText As String
    Dim positionCodeText As String

    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function
    Set sourceTable = sourceTables.Item(1)
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function
    If Not sourceRow.TryGetCellValueByColumn( _
        EXPORT_SOURCE_IPN_COLUMN, ipnText) Then Exit Function
    If Not sourceRow.TryGetCellValueByColumn( _
        "PositionCode", positionCodeText) Then positionCodeText = VBA.vbNullString

    If VBA.StrComp(VBA.Trim$(ipnText), TEMPORARY_IPN_MARKER, _
       VBA.vbTextCompare) <> 0 And _
       VBA.StrComp(VBA.Trim$(positionCodeText), TEMPORARY_POSITION_MARKER, _
       VBA.vbTextCompare) <> 0 Then
        private_ConfirmPartialExport = True
        Exit Function
    End If

    private_ConfirmPartialExport = (VBA.MsgBox( _
        "Для тимчасово прибулого відсутні ІПН та код посади." & _
        VBA.vbCrLf & "Частина перевірок і пошуку відмінків буде виконана " & _
        "за даними форми; проблемні фрази позначаються жовтим у preview." & _
        VBA.vbCrLf & VBA.vbCrLf & "Продовжити експорт " & exportCaption & "?", _
        VBA.vbQuestion Or VBA.vbYesNo Or VBA.vbDefaultButton2, _
        "PrsnlEventBuilder / Неповні дані") = VBA.vbYes)
End Function

Private Function private_TryEnsureModeConfigCurrent() As Boolean
    Dim builderPage As obj_PagePrsnlEvntBuilder

    If m_Page Is Nothing Then Exit Function
    If Not TypeOf m_Page Is obj_PagePrsnlEvntBuilder Then Exit Function
    Set builderPage = m_Page
    private_TryEnsureModeConfigCurrent = builderPage.EnsureModeConfigCurrent()
End Function

Private Function private_TryUpdateExportSettings(ByVal configTable As obj_ConfigTable) As Boolean
    Dim prsnlEvntBuilderCfgParser As obj_PrsnlEvntBuilderCfgParser

    private_ResetExportSettings
    If Not m_ExportEditingScen Is Nothing Then m_ExportEditingScen.Dispose
    Set m_ExportEditingScen = Nothing
    If Not m_MedicalReportsScen Is Nothing Then m_MedicalReportsScen.Dispose
    Set m_MedicalReportsScen = Nothing
    If configTable Is Nothing Then
        private_TryUpdateExportSettings = True
        Exit Function
    End If
    Set m_ProfileConfigTable = configTable

    Set prsnlEvntBuilderCfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not prsnlEvntBuilderCfgParser.Initialize(configTable) Then Exit Function
    If Not prsnlEvntBuilderCfgParser.TryGetEntityLookupColumnAliasByCaption(m_SourceColumnAliasByCaption) Then Exit Function
    If Not prsnlEvntBuilderCfgParser.TryGetExportSettings(m_ExportAliases, m_ExporterClassByAlias, m_ExportConfigTableByAlias) Then Exit Function

    ' Один provider живёт вместе с текущей конфигурацией страницы. Благодаря
    ' этому принадлежащий ему QueryEngine сохраняет ADO-соединение с Movement
    ' между последовательными запросами истории по Ctrl+Enter.
    Set m_ExporterCfgDataProvider = New obj_PEB_ExptrCfgDataPrvdr
    If Not m_ExporterCfgDataProvider.Initialize(configTable) Then
        Set m_ExporterCfgDataProvider = Nothing
        Exit Function
    End If

    private_TryUpdateExportSettings = True
End Function

Private Function private_TryUpdateProfilesProvider(ByVal configTable As obj_ConfigTable) As Boolean
    Dim prsnlEvntBuilderCfgParser As obj_PrsnlEvntBuilderCfgParser
    Dim providerClassName As String

    If configTable Is Nothing Then Exit Function

    Set prsnlEvntBuilderCfgParser = New obj_PrsnlEvntBuilderCfgParser
    If Not prsnlEvntBuilderCfgParser.Initialize(configTable) Then Exit Function
    If Not prsnlEvntBuilderCfgParser.TryGetProfilesProviderClass(providerClassName) Then Exit Function

    If Not private_TryCreateProfilesProvider(providerClassName) Then Exit Function
    private_TryUpdateProfilesProvider = True
End Function

Private Function private_TryCreateProfilesProvider(ByVal providerClassName As String) As Boolean
    Dim pebFactory As obj_PEB_Factory

    providerClassName = VBA.Trim$(providerClassName)
    If VBA.Len(providerClassName) = 0 Then Exit Function
    Set pebFactory = New obj_PEB_Factory
    private_TryCreateProfilesProvider = _
        pebFactory.TryCreateProfilesProvider(providerClassName, m_Data)
End Function

Private Sub private_ResetExportSettings()
    ' Exporters own profile-backed lookup providers; keep them warm between
    ' exports and invalidate them only when configuration is rebuilt.
    private_ClearPendingMovementReceipt
    private_DisposeCachedExporters
    If Not m_ExporterCfgDataProvider Is Nothing Then m_ExporterCfgDataProvider.Dispose
    Set m_ExporterCfgDataProvider = Nothing
    Set m_ExportAliases = New Collection
    Set m_ExporterClassByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ExportConfigTableByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ProfileConfigTable = Nothing
    Set m_SourceColumnAliasByCaption = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Sub private_ClearPendingMovementReceipt()
    m_HasPendingMovementReceipt = False
    m_PendingMovementReceiptIpn = VBA.vbNullString
    m_PendingMovementReceiptSectionType = VBA.vbNullString
    m_PendingMovementReceiptOrderNo = VBA.vbNullString
End Sub

Private Function private_TryCaptureMovementReceipt( _
    ByVal sourceTables As Collection, _
    ByVal exportContext As Object _
) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim ipnText As String
    Dim sectionTypeText As String
    Dim orderNoText As String

    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function
    Set sourceTable = sourceTables.Item(1)
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function
    If Not sourceRow.TryGetCellValueByColumn( _
        EXPORT_SOURCE_IPN_COLUMN, ipnText) Then Exit Function
    If exportContext Is Nothing Then Exit Function
    If Not exportContext.Exists(EXPORT_CONTEXT_SECTION_TYPE_KEY) Then Exit Function
    sectionTypeText = VBA.CStr( _
        exportContext(EXPORT_CONTEXT_SECTION_TYPE_KEY))
    If Not exportContext.Exists(EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY) Then Exit Function
    orderNoText = VBA.CStr( _
        exportContext(EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY))

    ipnText = private_NormalizeText(ipnText)
    sectionTypeText = private_NormalizeText(sectionTypeText)
    orderNoText = private_NormalizeText(orderNoText)
    If VBA.Len(ipnText) = 0 Or VBA.Len(sectionTypeText) = 0 Or _
        VBA.Len(orderNoText) = 0 Then Exit Function

    m_PendingMovementReceiptIpn = ipnText
    m_PendingMovementReceiptSectionType = sectionTypeText
    m_PendingMovementReceiptOrderNo = orderNoText
    m_HasPendingMovementReceipt = True
    private_TryCaptureMovementReceipt = True
End Function

Private Sub private_ApplyPendingMovementReceipt( _
    ByVal sourceTables As Collection, _
    ByVal exportContext As Object _
)
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim ipnText As String
    Dim sectionTypeText As String
    Dim orderNoText As String

    If Not m_HasPendingMovementReceipt Then Exit Sub
    If sourceTables Is Nothing Then Exit Sub
    If sourceTables.Count <= 0 Then Exit Sub
    Set sourceTable = sourceTables.Item(1)
    If sourceTable Is Nothing Then Exit Sub
    If sourceTable.RowCount <= 0 Then Exit Sub
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Sub
    If Not sourceRow.TryGetCellValueByColumn( _
        EXPORT_SOURCE_IPN_COLUMN, ipnText) Then Exit Sub
    If exportContext Is Nothing Then Exit Sub
    If Not exportContext.Exists(EXPORT_CONTEXT_SECTION_TYPE_KEY) Then Exit Sub
    sectionTypeText = VBA.CStr( _
        exportContext(EXPORT_CONTEXT_SECTION_TYPE_KEY))
    If Not exportContext.Exists(EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY) Then Exit Sub
    orderNoText = VBA.CStr( _
        exportContext(EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY))

    If VBA.StrComp( _
        private_NormalizeText(ipnText), _
        m_PendingMovementReceiptIpn, _
        VBA.vbTextCompare) <> 0 Then Exit Sub
    If VBA.StrComp( _
        private_NormalizeText(sectionTypeText), _
        m_PendingMovementReceiptSectionType, _
        VBA.vbTextCompare) <> 0 Then Exit Sub
    If VBA.StrComp( _
        private_NormalizeText(orderNoText), _
        m_PendingMovementReceiptOrderNo, _
        VBA.vbTextCompare) <> 0 Then Exit Sub

    exportContext(EXPORT_CONTEXT_MOVEMENT_PREVALIDATED_KEY) = True
End Sub

Private Sub private_DisposeCachedExporters()
    ' Cached exporters освобождаем при смене profile config или закрытии
    ' страницы. WORD exporter может ссылаться на общий config provider, но
    ' его Dispose в таком случае только отпускает ссылку и не закрывает provider.
    On Error Resume Next
    If Not m_CachedMovementExporter Is Nothing Then m_CachedMovementExporter.Dispose
    If Not m_CachedWordExporter Is Nothing Then m_CachedWordExporter.Dispose
    Set m_CachedMovementExporter = Nothing
    Set m_CachedWordExporter = Nothing
    On Error GoTo 0
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
    If VBA.Len(outExporterClassName) = 0 Then
        ' Пустой class не подменяется другой реализацией: такой fallback мог
        ' незаметно направить данные в exporter другого назначения.
        VBA.MsgBox "PrototypeNew: exporter class is empty for alias '" & exportAlias & "'.", _
            VBA.vbExclamation, "PrototypeNew / Data export"
        Exit Function
    End If
    If m_ExportConfigTableByAlias.Exists(exportAlias) Then Set outExportConfigTable = m_ExportConfigTableByAlias(exportAlias)
    If outExportConfigTable Is Nothing Then Exit Function

    private_TryGetExportSettings = True
End Function

Private Function private_TryCreateDataExporter( _
    ByVal exporterClassName As String, _
    ByVal exportConfigTable As obj_ConfigTable, _
    ByRef outExporter As obj_IDataExporter _
) As Boolean
    Dim pebFactory As obj_PEB_Factory

    Set outExporter = Nothing
    exporterClassName = VBA.Trim$(exporterClassName)
    If VBA.Len(exporterClassName) = 0 Then
        VBA.MsgBox "PrototypeNew: exporter class is not specified.", _
            VBA.vbExclamation, "PrototypeNew / Data export"
        Exit Function
    End If

    Select Case VBA.LCase$(exporterClassName)
        Case VBA.LCase$("obj_PEB_ExptrMovement")
            If Not m_CachedMovementExporter Is Nothing Then
                Set outExporter = m_CachedMovementExporter
                private_TryCreateDataExporter = True
                Exit Function
            End If
        Case VBA.LCase$("obj_PEB_ExptrWord")
            If Not m_CachedWordExporter Is Nothing Then
                Set outExporter = m_CachedWordExporter
                private_TryCreateDataExporter = True
                Exit Function
            End If
    End Select

    If Not private_TryEnsureExporterCfgDataProvider() Then Exit Function
    Set pebFactory = New obj_PEB_Factory
    private_TryCreateDataExporter = pebFactory.TryCreateDataExporter( _
        exporterClassName, exportConfigTable, m_ProfileConfigTable, _
        m_ExporterCfgDataProvider, outExporter)
    If Not private_TryCreateDataExporter Then Exit Function
    If TypeOf outExporter Is obj_PEB_ExptrMovement Then
        Set m_CachedMovementExporter = outExporter
    ElseIf TypeOf outExporter Is obj_PEB_ExptrWord Then
        Set m_CachedWordExporter = outExporter
    End If
End Function

Private Function private_TryEnsureExporterCfgDataProvider() As Boolean
    Dim exporterCfgDataProvider As obj_PEB_ExptrCfgDataPrvdr

    If Not m_ExporterCfgDataProvider Is Nothing Then
        private_TryEnsureExporterCfgDataProvider = True
        Exit Function
    End If
    If m_ProfileConfigTable Is Nothing Then
        VBA.MsgBox _
            "Не завантажено конфігурацію профілю для підключення " & _
            "до Personnel та Movement.", _
            VBA.vbExclamation, _
            "PrsnlEventBuilder / З'єднання"
        Exit Function
    End If

    ' Після ручного disconnect сторінка та її config залишаються живими.
    ' Provider відновлюється ліниво перед першим запитом або WORD-дією,
    ' тому сама команда розриву не відкриває нове ADO-з'єднання одразу.
    Set exporterCfgDataProvider = New obj_PEB_ExptrCfgDataPrvdr
    If Not exporterCfgDataProvider.Initialize(m_ProfileConfigTable) Then
        exporterCfgDataProvider.Dispose
        VBA.MsgBox _
            "Не вдалося повторно підготувати provider конфігурації " & _
            "для Personnel та Movement.", _
            VBA.vbExclamation, _
            "PrsnlEventBuilder / З'єднання"
        Exit Function
    End If
    Set m_ExporterCfgDataProvider = exporterCfgDataProvider
    private_TryEnsureExporterCfgDataProvider = True
End Function

Private Function private_TryCaptureWordExportPreview(ByVal exportContext As Object) As Boolean
    Dim previewText As String
    Dim pageBase As obj_PageBase

    On Error GoTo EH

    If exportContext Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:preview-capture:context-empty"
#End If
        private_TryCaptureWordExportPreview = True
        Exit Function
    End If

    On Error Resume Next
    If exportContext.Exists(EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY) Then
        previewText = VBA.CStr(exportContext(EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY))
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:preview-capture:context-exists-hit"
#End If
    End If
    If Err.Number <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogWarning "prsnlevntbuilder:preview-capture:context-exists-error errNo=" & VBA.CStr(Err.Number) & " err='" & private_EscapeForLog(Err.Description) & "'"
#End If
        Err.Clear
        previewText = VBA.CStr(VBA.CallByName(exportContext, EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY, VbGet))
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:preview-capture:callbyname-hit"
#End If
    End If
    On Error GoTo 0

    If VBA.Len(VBA.Trim$(previewText)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogWarning "prsnlevntbuilder:preview-capture:preview-empty"
#End If
        private_TryCaptureWordExportPreview = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:preview-capture:preview-ready len=" & VBA.CStr(VBA.Len(previewText))
#End If
    m_WordExportPreviewText = previewText
    ' Даже новое preview становится источником экспорта только после отдельного
    ' подтверждения пользователя кнопкой рядом с Banner.
    m_IsWordPreviewExportMode = False
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    ' Обновляем минимальный boundary. Retained layout распространяет разницу
    ' высоты по предкам и переносит только следующий vertical flow.
    private_TryCaptureWordExportPreview = pageBase.TryReflowLayoutContainer( _
        WORD_EXPORT_PANEL_CONTAINER_NAME)
    If Not private_TryCaptureWordExportPreview Then
        VBA.MsgBox _
            "PrototypeNew: failed to partially render WORD preview panel '" & _
            WORD_EXPORT_PANEL_CONTAINER_NAME & "'.", _
            VBA.vbExclamation, _
            "PrototypeNew / WORD preview"
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "prsnlevntbuilder:preview-capture:partial-container-returned ok=" & VBA.LCase$(VBA.CStr(private_TryCaptureWordExportPreview))
#End If
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "prsnlevntbuilder:preview-capture:error errNo=" & VBA.CStr(Err.Number) & " err='" & private_EscapeForLog(Err.Description) & "'"
#End If
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
    Dim manualOrderDate As Date
    Dim manualOrderYear As Long

    Set outTables = Nothing
    Set outContext = Nothing
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If m_ExportMainTable Is Nothing Then
        If Not private_RefreshDraftReporterTvoFlag() Then Exit Function
    End If

    ' Экспортеры получают чистый контракт:
    '   tables(1) = основная таблица экспорта
    '   tables(2..n) = meta-таблицы экспорта
    '   context = общие значения для всех таблиц, например номер приказа и тип секции.
    Set outTables = New Collection
    Set outContext = VBA.CreateObject("Scripting.Dictionary")
    outContext.CompareMode = 1

    sectionTypeText = VBA.Trim$(m_SelectedMainProfile)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(m_SelectedProfile)
    If Not private_TryResolveCurrentOrderPair( _
        pageBase, ws, manualOrderNoText, manualOrderDate) Then Exit Function
    manualOrderYear = VBA.Year(manualOrderDate)

    outContext(EXPORT_CONTEXT_SECTION_TYPE_KEY) = sectionTypeText
    outContext(EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY) = manualOrderNoText
    outContext(EXPORT_CONTEXT_MANUAL_ORDER_YEAR_KEY) = VBA.CStr(manualOrderYear)
    outContext(EXPORT_CONTEXT_MANUAL_ORDER_DATE_SERIAL_KEY) = _
        VBA.CStr(VBA.CDbl(manualOrderDate))
    outContext(EXPORT_CONTEXT_VALIDATE_MOVEMENT_KEY) = m_IsMovementValidationEnabled
    outContext(EXPORT_CONTEXT_VALIDATE_WORD_KEY) = m_IsWordValidationEnabled
    If m_ExportMainTable Is Nothing Then
        outContext(EXPORT_CONTEXT_REPORT_IS_TVO_KEY) = m_DraftReportIsTvo
    Else
        outContext(EXPORT_CONTEXT_REPORT_IS_TVO_KEY) = m_ExportMainReportIsTvo
    End If
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
    Dim declaredAliases As Collection
    Dim aliasObj As Variant
    Dim aliases() As String
    Dim gridColumns() As Long
    Dim itemCount As Long
    Dim i As Long
    Dim j As Long
    Dim swapAlias As String
    Dim swapColumn As Long
    Dim valueRange As Range
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

    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetFormColumnKeys(declaredAliases) Then Exit Function
    If declaredAliases Is Nothing Then Exit Function

    ' Динамический layout может сдвигать видимые поля после смены профиля.
    ' Поэтому связываем значение с колонкой по runtime-тегу контрола, а не по
    ' одинаковому смещению внутри диапазонов заголовков и значений.
    For Each aliasObj In declaredAliases
        Set valueRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange( _
            VBA.CStr(aliasObj), valueRange, "visible") Then GoTo ContinueAlias
        If valueRange Is Nothing Then GoTo ContinueAlias
        If Application.Intersect(valueRange, draftValuesRange) Is Nothing Then
            GoTo ContinueAlias
        End If

        itemCount = itemCount + 1
        ReDim Preserve aliases(1 To itemCount)
        ReDim Preserve gridColumns(1 To itemCount)
        aliases(itemCount) = VBA.CStr(aliasObj)
        gridColumns(itemCount) = valueRange.Column
ContinueAlias:
    Next aliasObj
    If itemCount <= 0 Then
        VBA.MsgBox "PrototypeNew: active draft form has no registered visible fields.", _
            VBA.vbExclamation, "PrototypeNew / Export"
        Exit Function
    End If

    ' Таблица должна повторять фактический порядок полей на листе.
    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If gridColumns(j) < gridColumns(i) Then
                swapColumn = gridColumns(i)
                gridColumns(i) = gridColumns(j)
                gridColumns(j) = swapColumn
                swapAlias = aliases(i)
                aliases(i) = aliases(j)
                aliases(j) = swapAlias
            End If
        Next j
    Next i

    Set sourceTable = New obj_TableDynamic
    sourceTable.SectionTitle = "PrsnlEvntBuilder draft form"
    Set sourceRow = New obj_Row

    For i = 1 To itemCount
        Set valueRange = Nothing
        If Not pageBase.TryGetFirstLayoutTagRange( _
            aliases(i), valueRange, "visible") Then Exit Function
        If valueRange Is Nothing Then Exit Function
        headerText = private_ReadHeaderText( _
            ws.Cells(valueRange.Row - 1, valueRange.Column))
        If VBA.Len(headerText) = 0 Then
            If Not includeBlankHeaderColumns Then GoTo ContinueDraftColumn
            headerText = aliases(i)
        End If
        If Not private_AddSourceColumn( _
            sourceTable, headerText, aliases(i)) Then Exit Function
        ' Draft form является текстовым UI. Экспортная sourceTable должна
        ' получать именно отображаемую строку, а не типизированный Value2:
        ' иначе Excel превращает, например, номер документа 198/07 в дату.
        sourceRow.PushCellRaw valueRange.Cells(1, 1).Text

ContinueDraftColumn:
    Next i

    If Not sourceTable.PushRow(sourceRow) Then Exit Function
    Set outTable = sourceTable
    private_TryBuildDraftFormSourceTable = True
End Function

Private Function private_TryReadManualOrderNoValue( _
    ByVal pageBase As obj_PageBase, _
    ByVal ws As Worksheet _
) As String
    Dim orderNoScope As Range

    If pageBase Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function

    ' Номер приказа является редактируемым значением layout-контейнера.
    ' Читаем тот же контейнер, который obj_PagePrsnlEvntBuilder сохраняет и
    ' восстанавливает при rerender. ControlParts registry относится к деталям
    ' текущего render-а и может быть уже перестроен, хотя значение в ячейке
    ' визуально осталось; из-за этого прежний lookup иногда возвращал пусто.
    Set orderNoScope = Nothing
    If Not pageBase.TryGetLayoutContainerRange( _
        EVENT_DRAFT_ORDER_NO_CONTAINER_NAME, _
        orderNoScope) Then Exit Function
    If orderNoScope Is Nothing Then Exit Function

    ' Text сохраняет введённую дату в отображаемом полном формате. Value2
    ' превратил бы её в serial number Excel и resolver принял бы дату за номер.
    private_TryReadManualOrderNoValue = _
        VBA.Trim$(VBA.CStr(orderNoScope.Cells(1, 1).Text))
End Function

Private Function private_TryReadManualOrderYearValue( _
    ByVal pageBase As obj_PageBase, _
    ByVal ws As Worksheet _
) As String
    Dim orderYearScope As Range

    If pageBase Is Nothing Then Exit Function
    If ws Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange( _
        EVENT_DRAFT_ORDER_YEAR_CONTAINER_NAME, orderYearScope) Then Exit Function
    If orderYearScope Is Nothing Then Exit Function
    private_TryReadManualOrderYearValue = _
        VBA.Trim$(VBA.CStr(orderYearScope.Cells(1, 1).Text))
End Function

Private Function private_TryResolveCurrentOrderPair( _
    ByVal pageBase As obj_PageBase, _
    ByVal ws As Worksheet, _
    ByRef outOrderNo As String, _
    ByRef outOrderDate As Date _
) As Boolean
    outOrderNo = VBA.vbNullString
    outOrderDate = 0
    If Not m_HasResolvedOrderPair Then Exit Function
    outOrderNo = m_ResolvedOrderNo
    outOrderDate = m_ResolvedOrderDate
    private_TryResolveCurrentOrderPair = True
End Function

Private Function private_TryGetCurrentManualOrderNo(ByRef outOrderNo As String) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim orderDate As Date

    outOrderNo = VBA.vbNullString
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    private_TryGetCurrentManualOrderNo = private_TryResolveCurrentOrderPair( _
        pageBase, ws, outOrderNo, orderDate)
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
    If m_Data Is Nothing Then Exit Function
    private_IsMetaProfile = m_Data.IsMetaProfileName(profileText)
End Function

Private Function private_IsMainProfile(ByVal profileText As String) As Boolean
    profileText = VBA.Trim$(profileText)
    If VBA.Len(profileText) = 0 Then Exit Function
    private_IsMainProfile = Not private_IsMetaProfile(profileText)
End Function

Private Sub private_ClearExportFormState()
    Set m_ExportMainTable = Nothing
    m_ExportMainReportIsTvo = False
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

Private Function private_ClearVisibleDraftFieldsByAlias(ByVal aliasesToClear As Object) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim draftValuesRange As Range
    Dim tagEntries As Collection
    Dim tagEntryObj As Variant
    Dim tagEntry As Object
    Dim tagText As String
    Dim tagRange As Range
    Dim clearRange As Range
    Dim previousEnableEvents As Boolean

    If aliasesToClear Is Nothing Then Exit Function
    If aliasesToClear.Count = 0 Then
        private_ClearVisibleDraftFieldsByAlias = True
        Exit Function
    End If
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then Exit Function
    If draftValuesRange Is Nothing Then Exit Function

    Set tagEntries = Nothing
    If Not pageBase.TryGetLayoutTagEntriesInRange(draftValuesRange, tagEntries, "visible") Then Exit Function
    If tagEntries Is Nothing Then
        private_ClearVisibleDraftFieldsByAlias = True
        Exit Function
    End If

    previousEnableEvents = Application.EnableEvents
    On Error GoTo RestoreEventsAndFail
    Application.EnableEvents = False

    ' Layout registry связывает канонический алиас (_Rank, _IPN...) с текущим
    ' физическим адресом. Поэтому Caption и позиция колонки здесь не участвуют.
    For Each tagEntryObj In tagEntries
        If Not VBA.IsObject(tagEntryObj) Then GoTo ContinueTag
        Set tagEntry = tagEntryObj
        If tagEntry Is Nothing Then GoTo ContinueTag
        If Not tagEntry.Exists("Tag") Then GoTo ContinueTag

        tagText = VBA.Trim$(VBA.CStr(tagEntry("Tag")))
        If Not aliasesToClear.Exists(tagText) Then GoTo ContinueTag

        Set tagRange = Nothing
        Set clearRange = Nothing
        Set tagRange = ws.Range( _
            ws.Cells(VBA.CLng(tagEntry("RowStart")), VBA.CLng(tagEntry("ColStart"))), _
            ws.Cells(VBA.CLng(tagEntry("RowEnd")), VBA.CLng(tagEntry("ColEnd"))))
        Set clearRange = Application.Intersect(tagRange, draftValuesRange)
        If clearRange Is Nothing Then GoTo ContinueTag
        clearRange.ClearContents

ContinueTag:
    Next tagEntryObj

    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
    private_ClearVisibleDraftFieldsByAlias = True
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
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
                If Not private_ClearHotkeyAssignment(hotkeyRows, "CTRL+1", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_ACCEPT_CANDIDATE_ROW, "CTRL+ENTER", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow( _
                    hotkeyRows, HOTKEY_REPORT_TVO_CANDIDATES, _
                    "CTRL+/", hasChanges, True) Then Exit Function
                If Not private_EnsureExportHotkeyRows(hotkeyRows, hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_EXPORT_MOVEMENT_WORD, "CTRL+E", hasChanges, True) Then Exit Function
                ' CTRL+4, как и CTRL+3 для WORD preview, является системным
                ' контрактом страницы и не зависит от порядка Export aliases.
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_EXPORT_TO_WORD, "CTRL+4", hasChanges, True) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_APPLY_EXPORT_FORM, "ALT+ARROWDOWN", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_CLEAR_EXPORT_FORM, "ALT+ARROWUP", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, _
                    HOTKEY_DELETE_EXPORTED_EVENT, "CTRL+DELETE", _
                    hasChanges, True) Then Exit Function
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
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_REPORT_TVO_CANDIDATES, "CTRL+/") Then Exit Function
    If Not private_AddExportHotkeyRows(hotkeyRows) Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_EXPORT_MOVEMENT_WORD, "CTRL+E") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_EXPORT_TO_WORD, "CTRL+4") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_APPLY_EXPORT_FORM, "ALT+ARROWDOWN") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_CLEAR_EXPORT_FORM, "ALT+ARROWUP") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, _
        HOTKEY_DELETE_EXPORTED_EVENT, "CTRL+DELETE") Then Exit Function

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, notifyChange) Then Exit Function

    private_EnsureHotkeyRows = True
End Function

Private Function private_ClearHotkeyAssignment( _
    ByVal hotkeyRows As Collection, _
    ByVal disabledHotkey As String, _
    ByRef ioHasChanges As Boolean _
) As Boolean
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    disabledHotkey = VBA.UCase$(VBA.Replace$(VBA.Trim$(disabledHotkey), " ", VBA.vbNullString))

    For Each rowItem In hotkeyRows
        Set configEntry = Nothing
        On Error Resume Next
        Set configEntry = rowItem
        On Error GoTo 0
        If Not configEntry Is Nothing Then
            If VBA.StrComp( _
                VBA.UCase$(VBA.Replace$(VBA.Trim$(configEntry.Value), " ", VBA.vbNullString)), _
                disabledHotkey, _
                VBA.vbBinaryCompare) = 0 Then
                configEntry.Value = VBA.vbNullString
                ioHasChanges = True
            End If
        End If
    Next rowItem

    private_ClearHotkeyAssignment = True
End Function

Private Function private_EnsureHotkeyRow( _
    ByVal hotkeyRows As Collection, _
    ByVal actionId As String, _
    ByVal defaultHotkey As String, _
    ByRef ioHasChanges As Boolean, _
    Optional ByVal enforceDefault As Boolean = False _
) As Boolean
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry
    Dim targetEntry As obj_ConfigEntry
    Dim normalizedDefault As String

    If hotkeyRows Is Nothing Then Exit Function
    If Not private_HotkeyRowsContainAction(hotkeyRows, actionId) Then
        If Not private_AddHotkeyRow(hotkeyRows, actionId, defaultHotkey) Then Exit Function
        ioHasChanges = True
    End If
    If enforceDefault Then
        normalizedDefault = VBA.UCase$(VBA.Replace$(VBA.Trim$(defaultHotkey), " ", VBA.vbNullString))

        ' Сначала освобождаем обязательный hotkey у другой action-строки.
        ' Затем обновляем нужную строку. Так runtime автоматически мигрирует
        ' старое Export Word = CTRL+2, появившееся из-за ordinal mapping.
        For Each rowItem In hotkeyRows
            Set configEntry = Nothing
            On Error Resume Next
            Set configEntry = rowItem
            On Error GoTo 0
            If Not configEntry Is Nothing Then
                If VBA.StrComp(VBA.Trim$(configEntry.Key), actionId, VBA.vbTextCompare) = 0 Then
                    Set targetEntry = configEntry
                ElseIf VBA.StrComp( _
                    VBA.UCase$(VBA.Replace$(VBA.Trim$(configEntry.Value), " ", VBA.vbNullString)), _
                    normalizedDefault, _
                    VBA.vbBinaryCompare) = 0 Then
                    configEntry.Value = VBA.vbNullString
                    ioHasChanges = True
                End If
            End If
        Next rowItem

        If Not targetEntry Is Nothing Then
            If VBA.StrComp( _
                VBA.UCase$(VBA.Replace$(VBA.Trim$(targetEntry.Value), " ", VBA.vbNullString)), _
                normalizedDefault, _
                VBA.vbBinaryCompare) <> 0 Then
                targetEntry.Value = defaultHotkey
                ioHasChanges = True
            End If
        End If
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
        defaultHotkey = private_BuildExportDefaultHotkey(exportIndex, exportAlias)
        If Not private_EnsureHotkeyRow( _
            hotkeyRows, actionId, defaultHotkey, ioHasChanges, _
            private_IsRequiredExportHotkeyAlias(exportAlias)) Then Exit Function

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
            private_BuildExportDefaultHotkey(exportIndex, exportAlias)) Then Exit Function

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

Private Function private_BuildExportDefaultHotkey( _
    ByVal exportIndex As Long, _
    Optional ByVal exportAlias As String = "" _
) As String
    If exportIndex <= 0 Or exportIndex > MAX_EXPORT_HOTKEYS Then Exit Function

    ' Системные экспортеры имеют стабильные shortcuts. Раньше hotkey зависел
    ' от ordinal alias: после отключения Daily aliases Word стал вторым и
    ' автоматически переехал с CTRL+3 на CTRL+2.
    Select Case VBA.LCase$(VBA.Trim$(exportAlias))
        Case "movement"
            private_BuildExportDefaultHotkey = "CTRL+2"
            Exit Function
        Case "word"
            private_BuildExportDefaultHotkey = "CTRL+3"
            Exit Function
    End Select

    ' CTRL+1 временно зарезервирован и не должен запускать PEB export.
    If exportIndex = 1 Then Exit Function
    private_BuildExportDefaultHotkey = "CTRL+" & VBA.CStr(exportIndex)
End Function

Private Function private_IsRequiredExportHotkeyAlias(ByVal exportAlias As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(exportAlias))
        Case "movement", "word"
            private_IsRequiredExportHotkeyAlias = True
    End Select
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

Private Function private_AddSourceColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal columnName As String, _
    Optional ByVal explicitColumnAlias As String = VBA.vbNullString _
) As Boolean
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
    columnAlias = VBA.Trim$(explicitColumnAlias)
    If VBA.Len(columnAlias) = 0 And Not m_SourceColumnAliasByCaption Is Nothing Then
        If m_SourceColumnAliasByCaption.Exists(colObj.Name) Then
            columnAlias = VBA.Trim$(VBA.CStr(m_SourceColumnAliasByCaption(colObj.Name)))
        End If
    End If
    If VBA.Len(columnAlias) > 0 Then
        If Not colObj.AddAlias(columnAlias) Then Exit Function
        ' Целевые алиасы EntityLookup имеют префикс "_", отличающий их от
        ' алиасов внешних источников. Старый алиас сохраняем для совместимости
        ' с существующими экспортёрами и шаблонами.
        If VBA.Left$(columnAlias, 1) = "_" And VBA.Len(columnAlias) > 1 Then
            If Not colObj.AddAlias(VBA.Mid$(columnAlias, 2)) Then Exit Function
        End If
    End If

    private_AddSourceColumn = tableObj.PushColumn(colObj)
End Function
