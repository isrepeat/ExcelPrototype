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
Private Const ADDITIONAL_PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.AdditionalProfiles"
Private Const META_PROFILES_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MetaProfiles"
Private Const EXPORT_FORM_MAIN_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ExportForm.Main"
Private Const EXPORT_FORM_META_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.ExportForm.Meta"
Private Const MOVEMENT_HISTORY_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.MovementHistory"
Private Const HOTKEY_ACCEPT_CANDIDATE_ROW As String = "Accept Candidate Row"
Private Const HOTKEY_SELECT_FORM_ROW As String = "Select Form Row"
Private Const HOTKEY_APPLY_EXPORT_FORM As String = "Apply Export Form"
Private Const HOTKEY_CLEAR_EXPORT_FORM As String = "Clear Export Form"
Private Const HOTKEY_EXPORT_TO_WORD As String = "Export to WORD"
Private Const EXPORT_ACTION_PREFIX As String = "Export "
Private Const MAX_EXPORT_HOTKEYS As Long = 9
Private Const LOOKUP_CANDIDATES_CONTROL_NAME As String = "LookupCandidatesTable"
Private Const FIO_LOOKUP_KEY As String = "op_FIO"
Private Const WORD_EXPORT_PANEL_CONTAINER_NAME As String = "WordExportPanel"
Private Const EVENT_DRAFT_FORM_CONTAINER_NAME As String = "EventDraftForm"
Private Const EVENT_DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"
Private Const EVENT_DRAFT_ORDER_NO_CONTAINER_NAME As String = "EventDraftOrderNoValue"
Private Const EXPORT_META_PROFILE_TYPE_COLUMN_NAME As String = "meta_ProfileType"
Private Const EXPORT_CONTEXT_MANUAL_ORDER_NO_KEY As String = "ManualOrderNo"
Private Const EXPORT_CONTEXT_SECTION_TYPE_KEY As String = "SectionType"
Private Const EXPORT_CONTEXT_WORD_PREVIEW_TEXT_KEY As String = "WordExportPreviewText"
Private Const EXPORT_CONTEXT_VALIDATE_DAILY_SCOPE_KEY As String = "ValidateDailyScope"
Private Const EXPORT_CONTEXT_VALIDATE_MOVEMENT_KEY As String = "ValidateMovement"
Private Const EXPORT_CONTEXT_VALIDATE_WORD_KEY As String = "ValidateWord"
Private Const LOOKUP_MODE_CONTROL_NAME As String = "LookupMode"
Private Const MOVEMENT_HISTORY_TABLE_CONTROL_NAME As String = "MovementHistoryTable"
Private Const MOVEMENT_HISTORY_LIMIT_INPUT_NAME As String = "MovementHistoryLimitInput"
Private Const ADDITIONAL_PROFILE_SELECT_CONTROL_NAME As String = "EventDraftAdditionalProfileSelect"
Private Const VALIDATE_DAILY_SCOPE_CONTROL_NAME As String = "ValidateDailyScope"
Private Const VALIDATE_MOVEMENT_CONTROL_NAME As String = "ValidateMovement"
Private Const VALIDATE_WORD_CONTROL_NAME As String = "ValidateWord"
Private Const PROFILE_BUTTON_TAG_ARRIVAL As String = "arrival"
Private Const PROFILE_BUTTON_TAG_DEPARTURE As String = "departure"
Private Const META_PROFILE_BUTTON_TAG_NORMAL As String = "meta"
Private Const PROFILE_BUTTON_STATE_SELECTED As String = "selected"
Private Const ABSENCE_DEPARTURE_LOOKBACK_DAYS As Long = 5
Private Const ABSENCE_DEPARTURE_LOOKAHEAD_DAYS As Long = 10
Private Const HOSPITALS_SOURCE_PATH_CONFIG_KEY As String = "Source.Hospitals.FilePath"
Private Const INSTITUTIONS_SHEET_NAME As String = "Лікувальні Заклади"
Private Const INSTITUTIONS_HEADER_ROW As Long = 3
Private Const INSTITUTIONS_FIRST_DATA_ROW As Long = 5
Private Const INSTITUTION_CODE_INPUT_NAME As String = "InstitutionCodeInput"
Private Const INSTITUTION_NAME_INPUT_NAME As String = "InstitutionNameInput"
Private Const INSTITUTION_REGION_INPUT_NAME As String = "InstitutionRegionInput"
Private Const INSTITUTION_GENITIVE_INPUT_NAME As String = "InstitutionGenitiveInput"
Private Const INSTITUTION_ACCUSATIVE_INPUT_NAME As String = "InstitutionAccusativeInput"
Private Const INSTITUTION_DATIVE_INPUT_NAME As String = "InstitutionDativeInput"
' Канонические алиасы полей draft-формы. Отображаемые Caption этих полей
' принадлежат конфигу и не должны использоваться в логике контроллера.
Private Const DRAFT_ALIAS_RANK As String = "_Rank"
Private Const DRAFT_ALIAS_IPN As String = "_IPN"
Private Const DRAFT_ALIAS_POSITION_CODE As String = "_PositionCode"
Private Const DRAFT_ALIAS_POSITION_NAME As String = "_PositionName"
Private Const DRAFT_ALIAS_DESTINATION As String = "_Destination"
Private Const DRAFT_ALIAS_REPORT_RANK As String = "_ReportRank"
Private Const DRAFT_ALIAS_REPORT_POSITION_CODE As String = "_ReportPositionCode"
Private Const DRAFT_ALIAS_INCOMING_NO As String = "_IncomingNo"
Private Const DRAFT_ALIAS_INCOMING_DATE As String = "_IncomingDate"
Private Const DRAFT_ALIAS_DURATION_DAYS As String = "_DurationDays"
Private Const DRAFT_ALIAS_VACATION_TICKET_NO As String = "_VacationTicketNo"
Private Const DRAFT_ALIAS_VACATION_TICKET_DATE As String = "_VacationTicketDate"
Private Const DRAFT_ALIAS_VLK_NO As String = "_VlkNo"
Private Const DRAFT_ALIAS_VLK_DATE As String = "_VlkDate"

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
Private m_MovementHistoryTable As obj_TableDynamic
Private m_WordExportPreviewText As String
Private m_ExportCommonData As obj_PEB_ExptrCommonDataPrvdr
Private m_ExporterCfgDataProvider As obj_PEB_ExptrCfgDataPrvdr
Private m_CachedDailyScopeExporter As obj_PEB_ExptrDailyScope
Private m_CachedMovementExporter As obj_PEB_ExptrMovement
Private m_CachedWordExporter As obj_PEB_ExptrWord
Private m_IsLookupEnabled As Boolean
Private m_IsDailyScopeValidationEnabled As Boolean
Private m_IsMovementValidationEnabled As Boolean
Private m_IsWordValidationEnabled As Boolean
Private m_IsInstitutionsAppendFormChecked As Boolean
Private m_IsMovementHistoryEnabled As Boolean
Private m_SuppressLookupSearch As Boolean
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
    Set m_ExportCommonData = New obj_PEB_ExptrCommonDataPrvdr
    If Not m_ExportCommonData.Initialize() Then Exit Function
    private_ResetExportSettings
    m_IsLookupEnabled = True
    m_IsDailyScopeValidationEnabled = True
    m_IsMovementValidationEnabled = True
    m_IsWordValidationEnabled = False
    m_IsInstitutionsAppendFormChecked = False
    m_IsMovementHistoryEnabled = False

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_LookupFeature = New obj_EntityLookupFeature
    If Not m_LookupFeature.Initialize( _
        pageInterface, _
        CANDIDATE_TABLES_RUNTIME_KEY, _
        "prsnlevntbuilder:entitylookup") Then Exit Function

    If Not private_RegisterProfileOptions(False) Then Exit Function
    ' Select хранит runtime selectedId между page renders. При первом открытии
    ' синхронизируем его с фактически выбранным профилем, чтобы устаревший
    ' дополнительный пункт не подменял placeholder основной секции.
    If Not private_TrySyncAdditionalProfileSelectState() Then Exit Function
    If Not private_RegisterAdditionalProfileOptions(False) Then Exit Function
    If Not private_RegisterMetaProfileOptions(False) Then Exit Function
    If Not private_RegisterExportFormTables(False) Then Exit Function
    If Not private_RegisterMovementHistoryTable(False) Then Exit Function
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
    If Not m_ExportCommonData Is Nothing Then m_ExportCommonData.Dispose
    Set m_ExportCommonData = Nothing
    ' Сначала освобождаем borrowers, затем общий config provider, которым
    ' WORD exporter может пользоваться без владения его lifetime.
    private_DisposeCachedExporters
    If Not m_ExporterCfgDataProvider Is Nothing Then m_ExporterCfgDataProvider.Dispose
    Set m_ExporterCfgDataProvider = Nothing
    m_SelectedProfile = VBA.vbNullString
    m_SelectedMainProfile = VBA.vbNullString
    Set m_ExportMainTable = Nothing
    Set m_ExportMetaTables = Nothing
    Set m_MovementHistoryTable = Nothing
    m_WordExportPreviewText = VBA.vbNullString
    m_IsInstitutionsAppendFormChecked = False
    m_IsMovementHistoryEnabled = False
    On Error GoTo 0
End Sub

Public Property Get WordExportPreviewText() As String
    WordExportPreviewText = m_WordExportPreviewText
End Property

Public Property Get IsLookupEnabled() As Boolean
    IsLookupEnabled = m_IsLookupEnabled
End Property

Public Property Get IsDailyScopeValidationEnabled() As Boolean
    IsDailyScopeValidationEnabled = m_IsDailyScopeValidationEnabled
End Property

Public Property Get IsMovementValidationEnabled() As Boolean
    IsMovementValidationEnabled = m_IsMovementValidationEnabled
End Property

Public Property Get IsWordValidationEnabled() As Boolean
    IsWordValidationEnabled = m_IsWordValidationEnabled
End Property

Public Property Get IsInstitutionsAppendFormChecked() As Boolean
    IsInstitutionsAppendFormChecked = m_IsInstitutionsAppendFormChecked
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

Public Function ToggleInstitutionsAppendForm() As Boolean
    Dim previousChecked As Boolean
    Dim previousEnableEvents As Boolean
    Dim renderSucceeded As Boolean

    If m_Page Is Nothing Then Exit Function
    previousChecked = m_IsInstitutionsAppendFormChecked
    m_IsInstitutionsAppendFormChecked = Not previousChecked

    ' Checkbox управляет целым layout-контейнером. Collapsed меняет высоту
    ' страницы, поэтому требуется page render, а не refresh одной кнопки.
    '
    ' Как и при переключении секций, на время полного render отключаем события.
    ' Иначе восстановление значений формы после layout вызывает Worksheet_Change,
    ' повторную dispatch-обработку и заметно замедляет простой UI toggle.
    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    renderSucceeded = rt_PageManager.fn_RenderPage( _
        m_Page, "prsnlevntbuilder:toggle-institutions-append-form")

Cleanup:
    Application.EnableEvents = previousEnableEvents
    If Not renderSucceeded Then
        m_IsInstitutionsAppendFormChecked = previousChecked
        Exit Function
    End If

    ToggleInstitutionsAppendForm = True
    Exit Function

EH:
    renderSucceeded = False
    Resume Cleanup
End Function

Public Function ToggleDailyScopeValidation() As Boolean
    ToggleDailyScopeValidation = private_ToggleExporterValidation( _
        m_IsDailyScopeValidationEnabled, VALIDATE_DAILY_SCOPE_CONTROL_NAME, "DailyScope")
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
    ClearDependentDraftFields = True
End Function

Private Function private_AppendFioDependentAliases( _
    ByVal sectionText As String, _
    ByVal dependentAliases As Object _
) As Boolean
    Dim sectionKey As String

    If dependentAliases Is Nothing Then Exit Function
    If m_Data Is Nothing Then Set m_Data = New obj_PrsnlEvntBuilderData
    sectionKey = private_NormalizeText(sectionText)

    Select Case sectionKey
        ' Сейчас все основные события получают персональные данные из одного
        ' op_FIO. Отдельный Select Case оставляет правила секционными: когда для
        ' конкретного события набор изменится, его следует вынести в отдельный Case.
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatment), _
             private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAmbulatoryVlk), _
             private_NormalizeText(m_Data.SectionTypeCloseFromStationaryVlk)
            private_AddStandardFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToTreatment), _
             private_NormalizeText(m_Data.SectionTypeToTreatmentMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeToAmbulatoryVlk)
            private_AddStandardFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeCloseFromAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromFamilyVacation)
            private_AddCloseVacationFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToAnnualVacationPart), _
             private_NormalizeText(m_Data.SectionTypeToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeToMaternityLeave)
            private_AddToVacationFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToTreatmentVacation)
            private_AddToTreatmentVacationFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentToStationaryVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToVlk)
            private_AddStandardFioDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferStationaryVlkToTreatmentVacation), _
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
    If m_Data Is Nothing Then Set m_Data = New obj_PrsnlEvntBuilderData
    sectionKey = private_NormalizeText(sectionText)

    Select Case sectionKey
        ' Набор можно разделять по секциям независимо от правил op_FIO.
        Case private_NormalizeText(m_Data.SectionTypeCloseFromTreatment), _
             private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeCloseFromTreatmentMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeCloseFromAmbulatoryVlk), _
             private_NormalizeText(m_Data.SectionTypeCloseFromStationaryVlk)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeToTreatment), _
             private_NormalizeText(m_Data.SectionTypeToAnnualVacationPart), _
             private_NormalizeText(m_Data.SectionTypeToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeToMaternityLeave), _
             private_NormalizeText(m_Data.SectionTypeToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeToTreatmentMedicalCompany), _
             private_NormalizeText(m_Data.SectionTypeToAmbulatoryVlk)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferTreatmentToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentToStationaryVlk), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAnnualVacationToFamilyVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferFamilyVacationToAnnualVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferTreatmentVacationToVlk)
            private_AddStandardCommanderDependentAliases dependentAliases

        Case private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatment), _
             private_NormalizeText(m_Data.SectionTypeTransferAmbulatoryVlkToTreatmentVacation), _
             private_NormalizeText(m_Data.SectionTypeTransferStationaryVlkToTreatmentVacation), _
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
    If Not private_EnsureHotkeyRows(False) Then Exit Function
    UpdateDataFromConfigTable = True
End Function

Public Function PrepareRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.PrepareLookupRuntime(notifyChange) Then Exit Function
    If Not private_RegisterProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterAdditionalProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterMetaProfileOptions(notifyChange) Then Exit Function
    If Not private_RegisterExportFormTables(notifyChange) Then Exit Function
    If Not private_RegisterMovementHistoryTable(notifyChange) Then Exit Function
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
    Dim cellAddress As String
    Dim shouldRefreshMovementHistory As Boolean

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
        Case VBA.LCase$(HOTKEY_ACCEPT_CANDIDATE_ROW)
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

Public Function RuntimeClearExportFormAndCandidates() As Boolean
    Dim previousSuppressLookupSearch As Boolean
    Dim renderSucceeded As Boolean

    If m_Page Is Nothing Then Exit Function
    previousSuppressLookupSearch = m_SuppressLookupSearch
    On Error GoTo RestoreSuppression

    ' Верхняя draft-форма остаётся заполненной. Команда очищает только staging,
    ' сформированный через Apply, его WORD preview и текущих Lookup-кандидатов.
    private_ClearExportFormState
    m_WordExportPreviewText = VBA.vbNullString

    If Not m_LookupFeature Is Nothing Then
        If Not m_LookupFeature.ClearLookupCandidates(False) Then Exit Function
    End If
    If Not private_RegisterExportFormTables(False) Then Exit Function

    ' Полный render восстанавливает draft-значения и может повторно инициировать
    ' lookup для заполненного ФИО. Команда очистки не является поиском, поэтому
    ' временно подавляем только SearchCandidates, не меняя режим Lookup.
    m_SuppressLookupSearch = True
    renderSucceeded = rt_PageManager.fn_RenderPage( _
        m_Page, "prsnlevntbuilder:clear-export-form-and-candidates")
    m_SuppressLookupSearch = previousSuppressLookupSearch
    RuntimeClearExportFormAndCandidates = renderSucceeded

    If RuntimeClearExportFormAndCandidates Then
        rt_Messaging.fn_ShowStatusBarSuccess "Export form and candidates cleared.", 3
    End If
    Exit Function

RestoreSuppression:
    m_SuppressLookupSearch = previousSuppressLookupSearch
End Function

Public Function OnExportToWordClick(Optional ByVal ignored As Variant) As Boolean
    OnExportToWordClick = private_TryExportWordToDocument()
End Function

Public Function OnClearWordDocumentClick(Optional ByVal ignored As Variant) As Boolean
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim clearedBlockCount As Long
    Dim orderNoText As String

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If VBA.StrComp(exporterClassName, "obj_PEB_ExptrWord", VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "PrototypeNew: WORD document actions require obj_PEB_ExptrWord, configured: " & exporterClassName, VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function
    If m_CachedWordExporter Is Nothing Then Exit Function
    If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function
    If Not m_CachedWordExporter.RemoveResultDocumentAnchors( _
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

Public Function OnRegroupWordHospitalPointsClick(Optional ByVal ignored As Variant) As Boolean
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable
    Dim regroupedPointCount As Long
    Dim orderNoText As String

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    ' Кнопка использует тот же кэшированный WORD exporter, что экспорт и
    ' удаление якорей, поэтому путь result-документа определяется единообразно.
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If VBA.StrComp(exporterClassName, "obj_PEB_ExptrWord", VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "PrototypeNew: WORD document actions require obj_PEB_ExptrWord, configured: " & _
            exporterClassName, VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function
    If m_CachedWordExporter Is Nothing Then Exit Function
    If Not private_TryGetCurrentManualOrderNo(orderNoText) Then Exit Function
    If Not m_CachedWordExporter.RegroupResultDocumentHospitalPoints( _
        regroupedPointCount, orderNoText) Then Exit Function

    rt_Messaging.fn_ShowStatusBarSuccess _
        "WORD document: grouped hospital points: " & VBA.CStr(regroupedPointCount), 3
    OnRegroupWordHospitalPointsClick = True
End Function

' Добавляет одну строку в справочник лечебных учреждений текущего профиля.
' Пустой регион и падежные формы допустимы: часть записей справочника
' заполняется постепенно, но ключ и полное название нужны обязательно.
Public Function OnAppendInstitutionClick(Optional ByVal ignored As Variant) As Boolean
    Dim fieldNames As Variant
    Dim values(1 To 6) As Variant
    Dim sourcePath As String
    Dim targetRow As Long
    Dim operationStage As String

    On Error GoTo EH
    operationStage = "read-form"
    fieldNames = Array( _
        INSTITUTION_CODE_INPUT_NAME, _
        INSTITUTION_NAME_INPUT_NAME, _
        INSTITUTION_REGION_INPUT_NAME, _
        INSTITUTION_GENITIVE_INPUT_NAME, _
        INSTITUTION_ACCUSATIVE_INPUT_NAME, _
        INSTITUTION_DATIVE_INPUT_NAME)

    If Not private_TryReadInstitutionForm(fieldNames, values) Then Exit Function
    If VBA.Len(VBA.Trim$(VBA.CStr(values(1)))) = 0 Then
        VBA.MsgBox "Заповніть поле 'Позначення'.", VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(VBA.CStr(values(2)))) = 0 Then
        VBA.MsgBox "Заповніть поле 'Назва'.", VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        Exit Function
    End If

    operationStage = "resolve-source"
    If Not private_TryResolveHospitalsSourcePath(sourcePath) Then Exit Function
    operationStage = "append-row"
    If Not private_TryAppendInstitutionRow(sourcePath, values, targetRow) Then Exit Function

    ' С этого момента запись уже сохранена на диске. Сразу фиксируем успешный
    ' результат, чтобы сбой очистки UI или обновления provider не провоцировал
    ' пользователя повторно добавить ту же строку.
    OnAppendInstitutionClick = True
    VBA.MsgBox "Лікувальний заклад додано в рядок " & VBA.CStr(targetRow) & ".", _
        VBA.vbInformation, "PrsnlEventBuilder / Установи"

    operationStage = "clear-form"
    If Not private_TryClearInstitutionForm(fieldNames) Then
        VBA.MsgBox "Рядок збережено, але не вдалося очистити поля форми.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
    End If

    ' Provider держит соединения со справочниками. После записи создаём его
    ' заново, чтобы следующий экспорт гарантированно увидел новую строку.
    If Not m_ExportCommonData Is Nothing Then m_ExportCommonData.Dispose
    Set m_ExportCommonData = New obj_PEB_ExptrCommonDataPrvdr
    operationStage = "refresh-common-data"
    If Not m_ExportCommonData.Initialize() Then
        VBA.MsgBox "Рядок збережено, але не вдалося оновити кеш довідника. " & _
            "Перезапустіть сторінку перед експортом.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
    End If
    Exit Function

EH:
    VBA.MsgBox "Не вдалося додати лікувальний заклад." & VBA.vbCrLf & _
        "Етап: " & operationStage & VBA.vbCrLf & _
        "Помилка: " & Err.Description, _
        VBA.vbExclamation, "PrsnlEventBuilder / Установи"
End Function

Private Function private_TryExportWordToDocument() As Boolean
    Dim sourceTables As Collection
    Dim exportContext As Object
    Dim exporter As obj_IDataExporter
    Dim exporterClassName As String
    Dim exportConfigTable As obj_ConfigTable

    If Not private_TryEnsureModeConfigCurrent() Then Exit Function
    If Not private_TryGetExportSettings("Word", exporterClassName, exportConfigTable) Then
        VBA.MsgBox "PrototypeNew: Export.Word settings are missing.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_TryBuildExportSourceTables(sourceTables, exportContext) Then Exit Function
    exportContext("WriteToWord") = True
    If Not private_TryCreateDataExporter(exporterClassName, exportConfigTable, exporter) Then Exit Function
    If Not exporter.Export(sourceTables, exportContext) Then Exit Function
    ' CTRL+4/button exports the already prepared logical result to WORD.
    ' Do not capture the preview here: that path calls RenderPage and a WORD
    ' write does not change any visible state on the Excel page.
    rt_Messaging.fn_ShowStatusBarSuccess "Export to WORD: done", 3
    private_TryExportWordToDocument = True
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

    ' Preview относится к конкретной секции и собранному для нее export context.
    ' После фактической смены main/meta-профиля старый текст больше не валиден.
    ' Пустое значение заставляет visibility binding скрыть banner и WORD-кнопку;
    ' новое preview появится только после следующего запуска CTRL+3.
    m_WordExportPreviewText = VBA.vbNullString
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

Private Function private_TryRefreshMovementHistory() As Boolean
    Dim pageBase As obj_PageBase
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim ipnText As String
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

    ipnText = VBA.Trim$(ipnText)
    If VBA.Len(ipnText) = 0 Then
        VBA.MsgBox "Для перегляду історії руху заповніть поле 'ІПН'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Movement"
        Exit Function
    End If
    If Not private_TryReadMovementHistoryMaxRows(maxRows) Then Exit Function

    If m_ExporterCfgDataProvider Is Nothing Then
        Set m_ExporterCfgDataProvider = New obj_PEB_ExptrCfgDataPrvdr
        If Not m_ExporterCfgDataProvider.Initialize(m_ProfileConfigTable) Then
            Set m_ExporterCfgDataProvider = Nothing
            Exit Function
        End If
    End If
    If Not m_ExporterCfgDataProvider.TryGetMovementHistoryByIpn( _
        ipnText, movementHistoryTable, maxRows) Then GoTo Cleanup

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
    Dim absenceSelectorImpl As obj_PEB_AbsenceCnddtSlctr
    Dim absenceSelector As obj_ILookupCandidateSelector

    outCandidateCount = 0
    If m_SuppressLookupSearch Then
        SearchCandidates = True
        Exit Function
    End If
    If m_LookupFeature Is Nothing Then Exit Function
    If Not private_UpdateLookupActiveFormColumns() Then Exit Function
    ' Расширение из нескольких источников включено только на этой странице.
    ' Универсальный EntityLookup не зависит от источников PrsnlEvntBuilder.
    If VBA.StrComp(VBA.Trim$(lookupKey), "op_FIO", VBA.vbTextCompare) = 0 Then
        If Not private_TryResolveAbsenceDepartureDateRange( _
            minAbsenceDepartureDate, maxAbsenceDepartureDate) Then Exit Function
        Set absenceSelectorImpl = New obj_PEB_AbsenceCnddtSlctr
        If Not absenceSelectorImpl.Initialize( _
            "DepartureDate", minAbsenceDepartureDate, maxAbsenceDepartureDate) Then Exit Function
        Set absenceSelector = absenceSelectorImpl
        If Not m_LookupFeature.SearchCandidates(lookupKey, queryText, outCandidateCount, False) Then Exit Function
        If Not m_LookupFeature.ExtendCandidates( _
            "op_FIOAbsenceExtension", _
            "_FIO", _
            queryText, _
            absenceSelector, _
            notifyChange) Then Exit Function
        SearchCandidates = True
    Else
        SearchCandidates = m_LookupFeature.SearchCandidates(lookupKey, queryText, outCandidateCount, notifyChange)
    End If
End Function

Private Function private_TryResolveAbsenceDepartureDateRange( _
    ByRef outMinDate As Date, _
    ByRef outMaxDate As Date _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim orderNoText As String

    outMinDate = 0
    outMaxDate = 0
    If m_Page Is Nothing Then Exit Function
    If m_ExportCommonData Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    orderNoText = private_TryReadManualOrderNoValue(pageBase, ws)
    If VBA.Len(VBA.Trim$(orderNoText)) = 0 Then
        VBA.MsgBox "PrototypeNew: enter the current order number before searching FIO candidates. " & _
            "The ЕЖОС candidate must have a departure date from 5 days before " & _
            "the order date through 10 days after it.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If Not m_ExportCommonData.SetOrderNo(orderNoText) Then Exit Function
    If Not m_ExportCommonData.HasOrderDate Then
        VBA.MsgBox "PrototypeNew: order date was not found for order number '" & orderNoText & _
            "'. ЕЖОС candidate selection was stopped.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    outMinDate = VBA.DateAdd( _
        "d", -ABSENCE_DEPARTURE_LOOKBACK_DAYS, VBA.DateValue(m_ExportCommonData.OrderDate))
    outMaxDate = VBA.DateAdd( _
        "d", ABSENCE_DEPARTURE_LOOKAHEAD_DAYS, VBA.DateValue(m_ExportCommonData.OrderDate))
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

    Set profiles = m_Data.PrimaryProfileNames
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
        PROFILE_BUTTON_TAG_ARRIVAL, _
        PROFILE_BUTTON_STATE_SELECTED, _
        True, _
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

Private Function private_RegisterAdditionalProfileOptions(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim profiles As Collection
    Dim profileOptions As Collection
    Dim profileObj As Variant
    Dim profileText As String
    Dim optionObj As obj_SelectOption

    If m_Page Is Nothing Then Exit Function
    If m_Data Is Nothing Then Set m_Data = New obj_PrsnlEvntBuilderData
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
    Dim primaryCandidateColumnCount As Long
    Dim primaryCandidateLastCol As Long
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

    ' LookupCandidates выравнивает результат под колонками формы. После основного
    ' зелёного диапазона могут идти служебные колонки, которые нельзя переносить.
    If m_LookupFeature Is Nothing Then Exit Function
    If Not m_LookupFeature.TryGetPrimaryCandidateColumnCount(primaryCandidateColumnCount) Then Exit Function
    primaryCandidateLastCol = candidateRowRange.Column + primaryCandidateColumnCount - 1

    firstCol = private_MaxLong(candidateRowRange.Column, draftValuesRange.Column)
    lastCol = private_MinLong( _
        primaryCandidateLastCol, _
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
        ' Пустые ячейки кандидата обозначают разрыв или отсутствие данных,
        ' поэтому они не должны очищать уже заполненные значения формы.
        If VBA.Len(VBA.Trim$(VBA.CStr(ws.Cells(sourceRow, colIndex).Value2))) > 0 Then
            ws.Cells(draftValuesRange.Row, colIndex).Value2 = ws.Cells(sourceRow, colIndex).Value2
        End If
    Next colIndex
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0

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

    If Not exporter.Export(sourceTables, exportContext) Then Exit Function
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
    ' Exporters own profile-backed lookup providers; keep them warm between
    ' exports and invalidate them only when configuration is rebuilt.
    private_DisposeCachedExporters
    If Not m_ExporterCfgDataProvider Is Nothing Then m_ExporterCfgDataProvider.Dispose
    Set m_ExporterCfgDataProvider = Nothing
    Set m_ExportAliases = New Collection
    Set m_ExporterClassByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ExportConfigTableByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set m_ProfileConfigTable = Nothing
    Set m_SourceColumnAliasByCaption = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Sub private_DisposeCachedExporters()
    ' Cached exporters освобождаем при смене profile config или закрытии
    ' страницы. WORD exporter может ссылаться на общий config provider, но
    ' его Dispose в таком случае только отпускает ссылку и не закрывает provider.
    On Error Resume Next
    If Not m_CachedDailyScopeExporter Is Nothing Then m_CachedDailyScopeExporter.Dispose
    If Not m_CachedMovementExporter Is Nothing Then m_CachedMovementExporter.Dispose
    If Not m_CachedWordExporter Is Nothing Then m_CachedWordExporter.Dispose
    Set m_CachedDailyScopeExporter = Nothing
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
        ' Пустой class больше не подменяется DailyScope: такой fallback мог
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
    Dim exporterToDailyScope As obj_PEB_ExptrDailyScope
    Dim exporterToMovement As obj_PEB_ExptrMovement
    Dim exporterToWord As obj_PEB_ExptrWord

    Set outExporter = Nothing
    exporterClassName = VBA.Trim$(exporterClassName)
    If VBA.Len(exporterClassName) = 0 Then
        VBA.MsgBox "PrototypeNew: exporter class is not specified.", _
            VBA.vbExclamation, "PrototypeNew / Data export"
        Exit Function
    End If

    Select Case VBA.LCase$(exporterClassName)
        Case VBA.LCase$("obj_PEB_ExptrDailyScope")
            If Not m_CachedDailyScopeExporter Is Nothing Then
                Set outExporter = m_CachedDailyScopeExporter
                private_TryCreateDataExporter = True
                Exit Function
            End If
            Set exporterToDailyScope = New obj_PEB_ExptrDailyScope
            If Not exporterToDailyScope.Initialize(exportConfigTable, m_ProfileConfigTable) Then Exit Function
            Set outExporter = exporterToDailyScope
            Set m_CachedDailyScopeExporter = exporterToDailyScope

        Case VBA.LCase$("obj_PEB_ExptrMovement")
            If Not m_CachedMovementExporter Is Nothing Then
                Set outExporter = m_CachedMovementExporter
                private_TryCreateDataExporter = True
                Exit Function
            End If
            Set exporterToMovement = New obj_PEB_ExptrMovement
            If Not exporterToMovement.Initialize(exportConfigTable, m_ProfileConfigTable) Then Exit Function
            Set outExporter = exporterToMovement
            Set m_CachedMovementExporter = exporterToMovement

        Case VBA.LCase$("obj_PEB_ExptrWord")
            If Not m_CachedWordExporter Is Nothing Then
                Set outExporter = m_CachedWordExporter
                private_TryCreateDataExporter = True
                Exit Function
            End If
            If m_ExporterCfgDataProvider Is Nothing Then
                VBA.MsgBox _
                    "PrototypeNew: exporter configuration data provider is not initialized.", _
                    VBA.vbExclamation, _
                    "PrototypeNew / WORD export"
                Exit Function
            End If
            Set exporterToWord = New obj_PEB_ExptrWord
            ' WORD preview и «Історія руху» должны использовать один provider:
            ' его QueryEngine владеет единственным ADO handle к Movement snapshot.
            If Not exporterToWord.Initialize( _
                exportConfigTable, _
                m_ProfileConfigTable, _
                m_ExporterCfgDataProvider) Then Exit Function
            Set outExporter = exporterToWord
            Set m_CachedWordExporter = exporterToWord

        Case Else
            VBA.MsgBox "PrototypeNew: unsupported data exporter class: " & exporterClassName, VBA.vbExclamation, "PrototypeNew / Data export"
            Exit Function
    End Select

    private_TryCreateDataExporter = Not outExporter Is Nothing
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
    outContext(EXPORT_CONTEXT_VALIDATE_DAILY_SCOPE_KEY) = m_IsDailyScopeValidationEnabled
    outContext(EXPORT_CONTEXT_VALIDATE_MOVEMENT_KEY) = m_IsMovementValidationEnabled
    outContext(EXPORT_CONTEXT_VALIDATE_WORD_KEY) = m_IsWordValidationEnabled
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
        sourceRow.PushCellRaw valueRange.Cells(1, 1).Value2

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

    private_TryReadManualOrderNoValue = VBA.Trim$(VBA.CStr(orderNoScope.Cells(1, 1).Value2))
End Function

Private Function private_TryGetCurrentManualOrderNo(ByRef outOrderNo As String) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet

    outOrderNo = VBA.vbNullString
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    outOrderNo = private_TryReadManualOrderNoValue(pageBase, ws)
    If VBA.Len(VBA.Trim$(outOrderNo)) = 0 Then
        VBA.MsgBox "PrototypeNew: enter the order number before working with the WORD document.", _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    private_TryGetCurrentManualOrderNo = True
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

Private Function private_TryReadInstitutionForm( _
    ByVal fieldNames As Variant, _
    ByRef outValues() As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim rawControl As Object
    Dim inputControl As obj_InputControlVM
    Dim fieldIndex As Long
    Dim fieldValue As String

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    For fieldIndex = LBound(fieldNames) To UBound(fieldNames)
        Set rawControl = Nothing
        Set inputControl = Nothing
        If Not pageBase.TryGetRegisteredControlByName( _
            VBA.CStr(fieldNames(fieldIndex)), rawControl) Then
            VBA.MsgBox "Не знайдено поле форми: " & VBA.CStr(fieldNames(fieldIndex)), _
                VBA.vbExclamation, "PrsnlEventBuilder / Установи"
            Exit Function
        End If
        If rawControl Is Nothing Then Exit Function
        If Not TypeOf rawControl Is obj_InputControlVM Then Exit Function
        Set inputControl = rawControl
        fieldValue = VBA.vbNullString
        If Not inputControl.TryGetValue(fieldValue) Then Exit Function
        outValues(fieldIndex + 1) = fieldValue
    Next fieldIndex

    private_TryReadInstitutionForm = True
End Function

Private Function private_TryClearInstitutionForm(ByVal fieldNames As Variant) As Boolean
    Dim pageBase As obj_PageBase
    Dim rawControl As Object
    Dim inputControl As obj_InputControlVM
    Dim fieldIndex As Long
    Dim previousEnableEvents As Boolean

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    previousEnableEvents = Application.EnableEvents
    On Error GoTo RestoreEventsAndFail
    Application.EnableEvents = False
    For fieldIndex = LBound(fieldNames) To UBound(fieldNames)
        Set rawControl = Nothing
        Set inputControl = Nothing
        If Not pageBase.TryGetRegisteredControlByName( _
            VBA.CStr(fieldNames(fieldIndex)), rawControl) Then GoTo RestoreEventsAndFail
        If rawControl Is Nothing Then GoTo RestoreEventsAndFail
        If Not TypeOf rawControl Is obj_InputControlVM Then GoTo RestoreEventsAndFail
        Set inputControl = rawControl
        If Not inputControl.ClearValue() Then GoTo RestoreEventsAndFail
    Next fieldIndex
    Application.EnableEvents = previousEnableEvents
    private_TryClearInstitutionForm = True
    Exit Function

RestoreEventsAndFail:
    On Error Resume Next
    Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
End Function

Private Function private_TryResolveHospitalsSourcePath(ByRef outSourcePath As String) As Boolean
    Dim cfgParserBase As obj_CfgParserBase
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim configuredPath As String

    outSourcePath = VBA.vbNullString
    If m_ProfileConfigTable Is Nothing Then
        VBA.MsgBox "Не завантажено конфігурацію поточного профілю.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        Exit Function
    End If

    Set cfgParserBase = New obj_CfgParserBase
    If Not cfgParserBase.Initialize(m_ProfileConfigTable) Then Exit Function
    If Not cfgParserBase.TryGetConfigEntries(configEntries) Then Exit Function
    If Not cfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then Exit Function
    If Not cfgParserBase.TryGetRequiredConfigValue( _
        cfgMap, HOSPITALS_SOURCE_PATH_CONFIG_KEY, configuredPath) Then
        VBA.MsgBox "У профілі не задано " & HOSPITALS_SOURCE_PATH_CONFIG_KEY & ".", _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        Exit Function
    End If

    If private_IsAbsoluteWorkbookPath(configuredPath) Then
        outSourcePath = configuredPath
    Else
        outSourcePath = ex_XmlCore.fn_CombineBasePath(ThisWorkbook, configuredPath)
    End If
    If VBA.Len(VBA.Trim$(outSourcePath)) = 0 Or VBA.Len(VBA.Dir$(outSourcePath)) = 0 Then
        VBA.MsgBox "Файл Установи не знайдено:" & VBA.vbCrLf & outSourcePath, _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        outSourcePath = VBA.vbNullString
        Exit Function
    End If

    private_TryResolveHospitalsSourcePath = True
End Function

Private Function private_TryAppendInstitutionRow( _
    ByVal sourcePath As String, _
    ByRef values() As Variant, _
    ByRef outTargetRow As Long _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim lastCell As Range
    Dim sourceFormatRow As Long
    Dim expectedHeaders As Variant
    Dim headerIndex As Long
    Dim rowValues(1 To 1, 1 To 6) As Variant

    On Error GoTo EH
    Set wb = private_FindOpenWorkbookByPath(sourcePath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(sourcePath)
        openedHere = True
    End If
    If wb Is Nothing Then GoTo EH
    If wb.ReadOnly Then
        VBA.MsgBox "Файл відкрито лише для читання:" & VBA.vbCrLf & sourcePath, _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        GoTo Cleanup
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(INSTITUTIONS_SHEET_NAME)
    On Error GoTo EH
    If ws Is Nothing Then
        VBA.MsgBox "У файлі немає аркуша '" & INSTITUTIONS_SHEET_NAME & "'.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Установи"
        GoTo Cleanup
    End If

    expectedHeaders = Array("Позначення", "Назва", "Регіон", "Родовий", "Знахідний", "Давальний")
    For headerIndex = 0 To 5
        If VBA.StrComp( _
            VBA.Trim$(VBA.CStr(ws.Cells(INSTITUTIONS_HEADER_ROW, headerIndex + 1).Value2)), _
            VBA.CStr(expectedHeaders(headerIndex)), VBA.vbTextCompare) <> 0 Then
            VBA.MsgBox "Неочікуваний заголовок у " & ws.Cells(INSTITUTIONS_HEADER_ROW, headerIndex + 1).Address(False, False) & _
                ". Очікується '" & VBA.CStr(expectedHeaders(headerIndex)) & "'.", _
                VBA.vbExclamation, "PrsnlEventBuilder / Установи"
            GoTo Cleanup
        End If
    Next headerIndex

    Set lastCell = ws.Range("A:F").Find( _
        What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastCell Is Nothing Then
        outTargetRow = INSTITUTIONS_FIRST_DATA_ROW
    Else
        outTargetRow = private_MaxLong(INSTITUTIONS_FIRST_DATA_ROW, lastCell.Row + 1)
    End If

    sourceFormatRow = outTargetRow - 1
    If sourceFormatRow >= INSTITUTIONS_FIRST_DATA_ROW Then
        ws.Range(ws.Cells(sourceFormatRow, 1), ws.Cells(sourceFormatRow, 6)).Copy
        ws.Range(ws.Cells(outTargetRow, 1), ws.Cells(outTargetRow, 6)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        ws.Rows(outTargetRow).RowHeight = ws.Rows(sourceFormatRow).RowHeight
    End If

    For headerIndex = 1 To 6
        rowValues(1, headerIndex) = values(headerIndex)
    Next headerIndex
    ws.Range(ws.Cells(outTargetRow, 1), ws.Cells(outTargetRow, 6)).NumberFormat = "@"
    ws.Range(ws.Cells(outTargetRow, 1), ws.Cells(outTargetRow, 6)).Value2 = rowValues
    wb.Save
    private_TryAppendInstitutionRow = True

Cleanup:
    On Error Resume Next
    Application.CutCopyMode = False
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

EH:
    VBA.MsgBox "Не вдалося додати рядок у файл Установи:" & VBA.vbCrLf & Err.Description, _
        VBA.vbExclamation, "PrsnlEventBuilder / Установи"
    Resume Cleanup
End Function

Private Function private_FindOpenWorkbookByPath(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If VBA.StrComp(VBA.Trim$(wb.FullName), VBA.Trim$(workbookPath), VBA.vbTextCompare) = 0 Then
            Set private_FindOpenWorkbookByPath = wb
            Exit Function
        End If
    Next wb
End Function

Private Function private_IsAbsoluteWorkbookPath(ByVal workbookPath As String) As Boolean
    workbookPath = VBA.Trim$(workbookPath)
    If VBA.Len(workbookPath) >= 3 Then
        If VBA.Mid$(workbookPath, 2, 2) = ":\" Then
            private_IsAbsoluteWorkbookPath = True
            Exit Function
        End If
    End If
    private_IsAbsoluteWorkbookPath = (VBA.Left$(workbookPath, 2) = "\\")
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
                If Not private_EnsureExportHotkeyRows(hotkeyRows, hasChanges) Then Exit Function
                ' CTRL+4, как и CTRL+3 для WORD preview, является системным
                ' контрактом страницы и не зависит от порядка Export aliases.
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_EXPORT_TO_WORD, "CTRL+4", hasChanges, True) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_APPLY_EXPORT_FORM, "ALT+ARROWDOWN", hasChanges) Then Exit Function
                If Not private_EnsureHotkeyRow(hotkeyRows, HOTKEY_CLEAR_EXPORT_FORM, "ALT+ARROWUP", hasChanges) Then Exit Function
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
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_EXPORT_TO_WORD, "CTRL+4") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_SELECT_FORM_ROW, "SHIFT+SPACE") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_APPLY_EXPORT_FORM, "ALT+ARROWDOWN") Then Exit Function
    If Not private_AddHotkeyRow(hotkeyRows, HOTKEY_CLEAR_EXPORT_FORM, "ALT+ARROWUP") Then Exit Function

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
