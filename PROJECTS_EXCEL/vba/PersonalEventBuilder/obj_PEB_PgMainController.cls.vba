VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_lookupProfile As obj_LookupProfile
Private m_lookupService As obj_LookupService
Private m_lookupResult As obj_LookupResult
Private m_messages As Object
Private m_candidateEnricher As obj_PEB_CandidateEnricher
Private m_orderResolver As obj_OrderResolver
Private m_ordersByCaption As Object
Private m_pendingOrder As obj_LookupRecord
Private m_acceptedOrderNumber As String
Private m_acceptedOrderDate As Variant
Private m_modeFields As Collection
Private m_displayFields As Collection
Private m_viewportStart As Long
Private m_hasViewport As Boolean
Private m_modeId As String
Private m_modeDirection As String
Private m_eventType As String
Private m_modeConfig As Object
Private m_reporterFields As Collection
Private m_reporterActive As Boolean
Private m_hospitalActive As Boolean
Private m_documentActive As Boolean
Private m_documentFields As Collection
Private m_documentProfile As obj_LookupProfile
Private m_documentService As obj_LookupService
Private m_hospitalFields As Collection
Private m_hospitalProfile As obj_LookupProfile
Private m_hospitalService As obj_LookupService
Private m_exportService As obj_PEB_ExportService
Private m_environmentSwitchPending As Boolean
Private m_testBusy As Boolean
Private m_testScenario As String
Private m_testStepOrder As Long
Private m_historyRows As Object
Private m_historyResult As obj_LookupResult
Private m_editHistoryRow As Object
Private m_editHistoryStamp As String
Private m_editHistoryAutoFill As Variant
Private m_historyOrderNumber As String
Private m_historyOrderDate As String
Private m_conflicts As Collection
Private m_extensions As obj_PEB_FormExtensions

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    Dim messageKey As Variant
    Dim messageText As String
    Dim diagnostic As String
    Dim defaultMode As String
    Dim emptyHistory As New Collection

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    Set m_extensions = New obj_PEB_FormExtensions
    If Not m_extensions.Initialize(pageBase) Then Exit Function
    If m_pageBase Is Nothing Then
        ex_WindowsUi.fn_ShowMessage "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set m_messages = VBA.CreateObject("Scripting.Dictionary")
    For Each messageKey In VBA.Array("PersonnelError", "PersonnelMinimum", "PersonnelEmpty", _
            "PersonnelCount", "PersonnelMore", "PersonnelSelected", "PersonnelActionError", _
            "OrderReady", "OrderEmpty", "OrderFoundOne", "OrderFoundMany", "OrderChooseRequired", _
            "OrderChoiceFormat", "OrderAccepted", "OrderAppliedRefreshError", _
            "PersonnelManualSelected", "AutoFillEnabled", "AutoFillDisabled")
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text." & messageKey, messageText) Then
            ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text." & messageKey, vbExclamation, "Configuration"
            Exit Function
        End If
        m_messages(messageKey) = messageText
    Next messageKey
    Set m_lookupProfile = New obj_LookupProfile
    Set m_ordersByCaption = VBA.CreateObject("Scripting.Dictionary")
    Set m_orderResolver = New obj_OrderResolver
    If Not m_orderResolver.Initialize("PersonalEventBuilder::orders", "PersonalEventBuilder::text", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not m_lookupProfile.Initialize("PersonalEventBuilder::lookup.Personnel", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If VBA.Len(m_lookupProfile.Source.RangeAddress) = 0 Or VBA.LCase$(m_lookupProfile.Source.RangeAddress) = "auto" Then
        ex_WindowsUi.fn_ShowMessage "Explicit configuration range required: PersonalEventBuilder::lookup.Personnel.source.range", vbExclamation, "Configuration"
        Exit Function
    End If
    Set m_lookupService = New obj_LookupService
    If Not m_lookupService.Initialize(m_lookupProfile) Then Exit Function
    Set m_candidateEnricher = New obj_PEB_CandidateEnricher
    If Not m_candidateEnricher.Initialize(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If Not m_lookupService.TrySearch(vbNullString, m_lookupResult, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", m_lookupResult) Then Exit Function
    If Not private_CaptureModeConfig(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not private_ReadModeConfig("modes.default", defaultMode, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not private_PublishExtendedModes(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Event modes"
        Exit Function
    End If
    If Not private_LoadMode(defaultMode, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not private_PublishHistory(emptyHistory, VBA.vbNullString, diagnostic) Then
        private_SetActionStatus diagnostic, True
        Exit Function
    End If
    If Not private_PublishConflicts(False) Then Exit Function
    Initialize = True
    m_isInitialized = Initialize
    If Not Me.SearchOrdersHandler() Then
        Me.Dispose
        Initialize = False
    End If
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_extensions Is Nothing Then m_extensions.Dispose
    Set m_extensions = Nothing
    If Not m_exportService Is Nothing Then m_exportService.Dispose
    Set m_exportService = Nothing
    If Not m_lookupService Is Nothing Then m_lookupService.Dispose
    If Not m_lookupResult Is Nothing Then m_lookupResult.Dispose
    If Not m_lookupProfile Is Nothing Then m_lookupProfile.Dispose
    If Not m_candidateEnricher Is Nothing Then m_candidateEnricher.Dispose
    If Not m_orderResolver Is Nothing Then m_orderResolver.Dispose
    Set m_orderResolver = Nothing
    Set m_ordersByCaption = Nothing
    Set m_pendingOrder = Nothing
    m_acceptedOrderNumber = vbNullString
    m_acceptedOrderDate = Empty
    Set m_candidateEnricher = Nothing
    Set m_messages = Nothing
    Set m_lookupService = Nothing
    Set m_lookupProfile = Nothing
    Set m_lookupResult = Nothing
    Set m_modeFields = Nothing
    Set m_displayFields = Nothing
    Set m_modeConfig = Nothing
    Set m_reporterFields = Nothing
    If Not m_hospitalService Is Nothing Then m_hospitalService.Dispose
    If Not m_hospitalProfile Is Nothing Then m_hospitalProfile.Dispose
    Set m_hospitalService = Nothing
    Set m_hospitalProfile = Nothing
    Set m_hospitalFields = Nothing
    If Not m_documentService Is Nothing Then m_documentService.Dispose
    If Not m_documentProfile Is Nothing Then m_documentProfile.Dispose
    Set m_documentService = Nothing
    Set m_documentProfile = Nothing
    Set m_documentFields = Nothing
    m_hospitalActive = False
    m_documentActive = False
    m_reporterActive = False
    Set m_historyRows = Nothing
    Set m_historyResult = Nothing
    Set m_editHistoryRow = Nothing
    m_editHistoryStamp = VBA.vbNullString
    Set m_pageBase = Nothing
    Set m_conflicts = Nothing
End Sub

Public Function ReportNameChangedHandler() As Boolean
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    If Not VBA.CBool(autoFill) Then
        ReportNameChangedHandler = True
        Exit Function
    End If
    ReportNameChangedHandler = private_SearchSecondary(True)
End Function

Public Function HospitalChangedHandler() As Boolean
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    If Not VBA.CBool(autoFill) Then
        HospitalChangedHandler = True
        Exit Function
    End If
    HospitalChangedHandler = private_SearchSecondary(True, True)
End Function

Public Function DocumentChangedHandler() As Boolean
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    If Not VBA.CBool(autoFill) Then
        DocumentChangedHandler = True
        Exit Function
    End If
    DocumentChangedHandler = private_SearchSecondary(True, False, True, True)
End Function

Public Function PersonNameChangedHandler() As Boolean
    Dim diagnostic As String
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    ' Manual entry must not trigger lookup, dependency clearing or candidate updates.
    If Not VBA.CBool(autoFill) Then
        PersonNameChangedHandler = True
        Exit Function
    End If
    If Not private_ClearPersonnelGroup(False, VBA.vbNullString, diagnostic) Then
        private_SetActionStatus diagnostic, True
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    PersonNameChangedHandler = Me.SearchPersonnelHandler()
End Function

Public Function ScrollFormLeftHandler() As Boolean
    ScrollFormLeftHandler = private_ScrollForm(-1)
End Function

Public Function ScrollFormRightHandler() As Boolean
    ScrollFormRightHandler = private_ScrollForm(1)
End Function

Public Function ResetFormHandler() As Boolean
    Dim updates As Object
    Dim wordContext As Object
    Dim fields As String
    Dim fieldName As Variant
    Dim diagnostic As String
    Dim records As New Collection
    Dim result As obj_LookupResult

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_ReadModeConfig("forms.styledFields", fields, diagnostic) Then GoTo Failed
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each fieldName In VBA.Split(fields, "|")
        updates("Form." & VBA.CStr(fieldName)) = VBA.vbNullString
    Next fieldName
    updates("Form.WordPreviewText") = VBA.vbNullString
    updates("Form.ActionStatus") = VBA.vbNullString
    updates("Text.PersonnelStatus") = VBA.vbNullString
    Set result = New obj_LookupResult
    If Not result.Initialize(records, m_modeFields, False, diagnostic) Then GoTo Failed
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    Set wordContext = VBA.CreateObject("Scripting.Dictionary")
    If Not m_pageBase.BindingContext.SetObject("Data", "WordContext", wordContext) Then
        Err.Raise 5, , "Cannot clear Word context."
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
        Err.Raise 5, , "Cannot clear candidate selection."
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then
        Err.Raise 5, , "Cannot clear candidates."
    End If
    Set m_lookupResult = result
    m_reporterActive = False
    m_hospitalActive = False
    m_documentActive = False
    If Not private_EndHistoryEdit() Then Exit Function
    If Not private_PublishHistory(records, VBA.vbNullString, diagnostic) Then GoTo Failed
    m_historyOrderNumber = VBA.vbNullString
    m_historyOrderDate = VBA.vbNullString
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Collapsed") Then Exit Function
    If Not m_extensions.Reset(diagnostic) Then GoTo Failed
    If Not m_extensions.Build(ex_PEB_EventRoutes.fn_Value(m_modeId, "form"), diagnostic) Then GoTo Failed
    ResetFormHandler = m_pageBase.UpdatePage()
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
End Function

Public Function SearchPersonnelHandler() As Boolean
    SearchPersonnelHandler = private_SearchPersonnel(True)
End Function

Public Function SelectPersonnelHandler(ByVal selected As Object) As Boolean
    Dim record As obj_LookupRecord
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim startedAt As Double
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If selected Is Nothing Then Exit Function
    If Not TypeOf selected Is obj_LookupRecord Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    If Not VBA.CBool(autoFill) Then
        SelectPersonnelHandler = private_SetActionStatus(m_messages("PersonnelManualSelected"), False)
        Exit Function
    End If
    previousScreenUpdating = Application.ScreenUpdating
    startedAt = VBA.Timer
    On Error GoTo EH
    Application.ScreenUpdating = False
    Set record = selected
    If Not private_ApplyModeRecord(record, diagnostic) Then
        private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        GoTo CleanExit
    End If
    SelectPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", m_messages("PersonnelSelected"))
    If SelectPersonnelHandler Then SelectPersonnelHandler = private_SetActionStatus(m_messages("PersonnelSelected"), False)
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    ex_Core.fn_Diagnostic_WritePerf "Personnel.ApplySelection", startedAt
    Exit Function
EH:
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", VBA.Err.Description), True
    Resume CleanExit
End Function

Public Function SearchOrdersHandler() As Boolean
    Dim query As Variant
    Dim year As Variant
    Dim matches As Collection
    Dim displayMatches As Collection
    Dim options As Collection
    Dim todayOptions As Collection
    Dim record As obj_LookupRecord
    Dim caption As String
    Dim diagnostic As String
    Dim status As String
    Dim orderDate As Variant
    Dim isError As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If Not private_TryReadFormValue("OrderQuery", query) Then Exit Function
    If Not private_TryReadFormValue("OrderYear", year) Then Exit Function
    Set m_pendingOrder = Nothing
    If Not m_pageBase.BindingContext.SetValue("Form", "OrderChoice", vbNullString) Then Exit Function

    ' Dropdown contents depend only on today and neighboring rows.
    If Not m_orderResolver.TryGetNearby(displayMatches, diagnostic) Then GoTo Failed
    Set m_ordersByCaption = VBA.CreateObject("Scripting.Dictionary")
    Set options = New Collection
    Set todayOptions = New Collection
    For Each record In displayMatches
        caption = private_OrderText(record, "OrderChoiceFormat")
        options.Add caption
        Set m_ordersByCaption(caption) = record
        If record.TryGetValue("OrderDate", orderDate) Then
            If VBA.CDate(orderDate) = VBA.Date Then todayOptions.Add caption
        End If
    Next record
    If Not m_pageBase.BindingContext.SetObject("Data", "OrderOptions", options) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "TodayOrderOptions", todayOptions) Then Exit Function
    ex_Core.fn_Diagnostic_WriteLog "ORDER_DROPDOWN_NEARBY | Items=" & VBA.CStr(options.Count)

    ' Input resolution never changes the dropdown contents.
    If Not m_orderResolver.TrySearch(VBA.CStr(query), VBA.CStr(year), matches, diagnostic) Then GoTo Failed
    If VBA.Len(VBA.Trim$(VBA.CStr(query))) = 0 Then
        status = m_messages("OrderReady")
    ElseIf matches.Count = 1 Then
        Set m_pendingOrder = matches(1)
        If Not m_pageBase.BindingContext.SetValue("Form", "OrderChoice", private_OrderText(m_pendingOrder, "OrderChoiceFormat")) Then Exit Function
        SearchOrdersHandler = private_ApplyResolvedOrder()
        Exit Function
    ElseIf matches.Count > 1 Then
        status = VBA.Replace$(m_messages("OrderFoundMany"), "{count}", VBA.CStr(matches.Count))
    Else
        status = m_messages("OrderEmpty")
        isError = True
    End If
    SearchOrdersHandler = private_SetActionStatus(status, isError)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    SearchOrdersHandler = private_SetActionStatus(diagnostic, True)
End Function

Public Function OrderQueryChangedHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ' Typed input and dropdown values use the same resolver.
    OrderQueryChangedHandler = Me.SearchOrdersHandler()
End Function


Private Function private_ApplyResolvedOrder() As Boolean
    Dim number As Variant
    Dim orderDate As Variant
    Dim updates As Object
    Dim diagnostic As String
    Dim acceptedText As String
    Dim selectedValue As Variant
    Dim selectedObject As Object
    Dim isObject As Boolean
    Dim selectedKey As String
    Dim selectedRecord As obj_LookupRecord
    Dim record As obj_LookupRecord

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_pendingOrder Is Nothing Then
        private_ApplyResolvedOrder = private_SetActionStatus(m_messages("OrderChooseRequired"), True)
        Exit Function
    End If
    On Error GoTo EH
    If Not m_pendingOrder.TryGetValue("OrderNumber", number) Then Exit Function
    If Not m_pendingOrder.TryGetValue("OrderDate", orderDate) Then Exit Function
    acceptedText = private_OrderText(m_pendingOrder, "OrderAccepted")
    If VBA.Len(m_acceptedOrderNumber) > 0 Then
        If m_acceptedOrderNumber = VBA.CStr(number) And VBA.CDate(m_acceptedOrderDate) = VBA.CDate(orderDate) Then
            private_ApplyResolvedOrder = private_SetActionStatus(acceptedText, False)
            Exit Function
        End If
    End If
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.AcceptedOrderNumber") = number
    updates("Form.AcceptedOrderDate") = orderDate
    updates("Form.AcceptedOrderText") = acceptedText

    ' Apply only a uniquely resolved pair; invalid and ambiguous input keeps the previous pair.
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    m_acceptedOrderNumber = VBA.CStr(number)
    m_acceptedOrderDate = orderDate
    If m_reporterActive Or m_hospitalActive Or m_documentActive Then
        If Not private_SearchSecondary(False, m_hospitalActive, False, m_documentActive) Then
            private_ApplyResolvedOrder = private_SetActionStatus(m_messages("OrderAppliedRefreshError"), True)
            Exit Function
        End If
        private_ApplyResolvedOrder = private_SetActionStatus(acceptedText, False)
        Exit Function
    End If

    ' Refresh derived tickets and restore the previously selected personnel record.
    If m_pageBase.BindingContext.TryGetValue("Data", "SelectedPersonnel", selectedValue, selectedObject, isObject) Then
        If isObject Then
            If TypeOf selectedObject Is obj_LookupRecord Then
                Set selectedRecord = selectedObject
                selectedKey = selectedRecord.Key
            End If
        End If
    End If
    If Not private_SearchPersonnel(False) Then
        private_ApplyResolvedOrder = private_SetActionStatus(m_messages("OrderAppliedRefreshError"), True)
        Exit Function
    End If
    If VBA.Len(selectedKey) > 0 Then
        For Each record In m_lookupResult.Records
            If record.Key = selectedKey Then
                If Not m_pageBase.BindingContext.SetObject("Data", "SelectedPersonnel", record) Then Exit Function
                If Not Me.SelectPersonnelHandler(record) Then Exit Function
                Exit For
            End If
        Next record
    End If
    private_ApplyResolvedOrder = private_SetActionStatus(acceptedText, False)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    private_ApplyResolvedOrder = private_SetActionStatus(diagnostic, True)
End Function

Public Function RefreshPersonnelHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_exportService Is Nothing Then m_exportService.Dispose
    Set m_exportService = Nothing
    ex_TableDataRuntime.fn_ResetCaches
    m_lookupService.InvalidateCache
    m_candidateEnricher.InvalidateCache
    m_orderResolver.InvalidateCache
    If Not m_hospitalService Is Nothing Then m_hospitalService.InvalidateCache
    If Not m_documentService Is Nothing Then m_documentService.InvalidateCache
    m_extensions.InvalidateCache
    If m_reporterActive Or m_hospitalActive Or m_documentActive Then
        RefreshPersonnelHandler = private_SearchSecondary(False, m_hospitalActive, True, m_documentActive)
        Exit Function
    End If
    RefreshPersonnelHandler = Me.SearchPersonnelHandler()
End Function

Public Function DummyMode11Handler() As Boolean
    DummyMode11Handler = True
End Function

Public Function DummyMode12Handler() As Boolean
    DummyMode12Handler = True
End Function

Public Function DummyMode13Handler() As Boolean
    DummyMode13Handler = True
End Function

Public Function DummyMode14Handler() As Boolean
    DummyMode14Handler = True
End Function

Public Function DummyMode21Handler() As Boolean
    DummyMode21Handler = True
End Function

Public Function DummyMode22Handler() As Boolean
    DummyMode22Handler = True
End Function

Public Function DummyMode23Handler() As Boolean
    DummyMode23Handler = True
End Function

Public Function DummyMode24Handler() As Boolean
    DummyMode24Handler = True
End Function

Public Function AutoFillPersonnelChangedHandler() As Boolean
    Dim autoFill As Variant
    Dim selectedValue As Variant
    Dim selected As Object
    Dim isObject As Boolean
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    If Not private_UpdateReadOnly(m_modeId, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If VBA.CBool(autoFill) Then
        If Not m_pageBase.BindingContext.TryGetValue("Data", "SelectedPersonnel", selectedValue, selected, isObject) Then Exit Function
        If isObject Then
            If Not Me.SelectPersonnelHandler(selected) Then Exit Function
        Else
            If Not private_SetActionStatus(m_messages("AutoFillEnabled"), False) Then Exit Function
        End If
    Else
        If Not private_SetActionStatus(m_messages("AutoFillDisabled"), False) Then Exit Function
    End If
    AutoFillPersonnelChangedHandler = m_pageBase.FlushUpdates()
End Function

Public Function TryRefreshSourceLegend(ByRef diagnostic As String) As Boolean
    Dim sourceAliases As String
    Dim dateFormat As String
    Dim entryTemplate As String
    Dim unavailableText As String
    Dim title As String
    Dim sourceAlias As Variant
    Dim sourcePath As String
    Dim sourceCaption As String
    Dim modifiedText As String
    Dim entryText As String
    Dim fileSystem As Object
    Dim sourceFile As Object

    On Error GoTo EH
    If m_pageBase Is Nothing Or m_isDisposed Then Exit Function
    If Not private_ReadModeConfig("legend.sources", sourceAliases, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("legend.dateFormat", dateFormat, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("text.SourceLegendEntry", entryTemplate, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("text.SourceLegendUnavailable", unavailableText, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("text.SourceLegendTitle", title, diagnostic) Then Exit Function
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not m_pageBase.BindingContext.SetValue("Text", "SourceLegendTitle", title) Then Exit Function
    For Each sourceAlias In VBA.Split(sourceAliases, "|")
        If Not private_ReadModeConfig("legend." & sourceAlias & ".entryTemplate", entryTemplate, diagnostic) Then Exit Function
        If Not private_ReadModeConfig("lookup." & sourceAlias & ".source.path", sourcePath, diagnostic) Then Exit Function
        If VBA.Mid$(sourcePath, 2, 1) <> ":" And VBA.Left$(sourcePath, 2) <> "\\" Then
            sourcePath = fileSystem.BuildPath(ThisWorkbook.Path, sourcePath)
        End If
        sourcePath = fileSystem.GetAbsolutePathName(sourcePath)
        sourceCaption = fileSystem.GetFileName(sourcePath)
        modifiedText = unavailableText
        If fileSystem.FileExists(sourcePath) Then
            Set sourceFile = fileSystem.GetFile(sourcePath)
            modifiedText = VBA.Format$(sourceFile.DateLastModified, dateFormat)
        End If
        entryText = VBA.Replace$(entryTemplate, "{source}", sourceCaption)
        entryText = VBA.Replace$(entryText, "{modified}", modifiedText)
        If Not m_pageBase.BindingContext.SetValue("Text", "SourceLegend" & sourceAlias & "Entry", entryText) Then Exit Function
    Next sourceAlias
    TryRefreshSourceLegend = True
    Exit Function
EH:
    diagnostic = "Source legend: " & VBA.Err.Description
End Function

Public Function FromOutpatientHandler() As Boolean
    FromOutpatientHandler = private_SwitchMode("FromOutpatientExaminationOrVlk")
End Function

Public Function FromMedicalCompanyHandler() As Boolean
    FromMedicalCompanyHandler = private_SwitchMode("FromMedicalCompanyForTreatment")
End Function

Public Function FromBusinessTripHandler() As Boolean
    FromBusinessTripHandler = private_SwitchMode("FromBusinessTrip")
End Function

Public Function ToOutpatientHandler() As Boolean
    ToOutpatientHandler = private_SwitchMode("ToOutpatientExaminationOrVlk")
End Function

Public Function ToMedicalCompanyHandler() As Boolean
    ToMedicalCompanyHandler = private_SwitchMode("ToMedicalCompanyForTreatment")
End Function

Public Function ToBusinessTripHandler() As Boolean
    ToBusinessTripHandler = private_SwitchMode("ToBusinessTrip")
End Function

Public Function OutpatientToMedicalLeaveHandler() As Boolean
    OutpatientToMedicalLeaveHandler = private_SwitchMode("MoveFromOutpatientVlkToVacationForCuring")
End Function

Public Function MedCompanyToMedicalLeaveHandler() As Boolean
    MedCompanyToMedicalLeaveHandler = private_SwitchMode("MoveFromMedicalCompanyForTreatmentToVacationForCuring")
End Function

Public Function ToBusinessTripAwolHandler() As Boolean
    ToBusinessTripAwolHandler = private_SwitchMode("MoveFromAwolToBusinessTrip")
End Function

Public Function AnnualToMedicalLeaveHandler() As Boolean
    AnnualToMedicalLeaveHandler = private_SwitchMode("MoveFromAnnualVacationToVacationForCuring")
End Function

Public Function FamilyToMedicalLeaveHandler() As Boolean
    FamilyToMedicalLeaveHandler = private_SwitchMode("MoveFromFamilyVacationToVacationForCuring")
End Function

Public Function MedicalLeaveToHospitalHandler() As Boolean
    MedicalLeaveToHospitalHandler = private_SwitchMode("MoveFromVacationForCuringToHospital")
End Function

Public Function MedicalLeaveToAnnualHandler() As Boolean
    MedicalLeaveToAnnualHandler = private_SwitchMode("MoveFromVacationForCuringToAnnualVacation")
End Function

Public Function ExtendedModeChangedHandler() As Boolean
    Dim selected As Variant
    Dim eventId As Variant
    Dim diagnostic As String

    On Error GoTo EH_EXTENDED
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("ExtendedMode", selected) Then Exit Function
    If VBA.Len(VBA.CStr(selected)) = 0 Then
        ExtendedModeChangedHandler = True
        Exit Function
    End If
    For Each eventId In ex_PEB_EventRoutes.fn_Ids()
        If ex_PEB_EventRoutes.fn_Value(VBA.CStr(eventId), "visible") = "false" Then
            If ex_PEB_EventRoutes.fn_Value(VBA.CStr(eventId), "caption") = VBA.CStr(selected) Then
                ExtendedModeChangedHandler = private_SwitchMode(VBA.CStr(eventId))
                Exit Function
            End If
        End If
    Next eventId
    Err.Raise 5, , "Unknown extended event caption: " & VBA.CStr(selected)
EH_EXTENDED:
    diagnostic = VBA.Err.Description
    ExtendedModeChangedHandler = private_SetActionStatus(diagnostic, True)
End Function

Public Function HospitalToLeaveHandler() As Boolean
    HospitalToLeaveHandler = private_SwitchMode("MoveFromHospitalToVacationForCuring")
End Function

Public Function MedicalLeaveExtendHandler() As Boolean
    MedicalLeaveExtendHandler = private_SwitchMode("MoveFromVacationForCuringToVacationForCuring")
End Function

Public Function FromHospitalHandler() As Boolean
    FromHospitalHandler = private_SwitchMode("FromHospital")
End Function

Public Function ReinitializePageHandler() As Boolean
    Dim diagnostic As String
    Dim status As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If ex_PEB_Testing.fn_IsRunning() Then
        ReinitializePageHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If m_environmentSwitchPending Then
        ReinitializePageHandler = True
        Exit Function
    End If
    On Error GoTo EH_REINITIALIZE
    status = ex_PEB_Export.fn_Config("text.ReinitializePagePending")
    If Not ex_ConfigRefresh.fn_TryUpdate(diagnostic) Then GoTo FailedReinitialize
    ' Recreate the page after the current command has returned.
    If Not ex_AppCommands.fn_RequestReinitialize() Then
        Err.Raise 5, , "Cannot schedule PEB page reinitialization."
    End If
    m_environmentSwitchPending = True
    ReinitializePageHandler = private_SetActionStatus(status, False)
    Exit Function
EH_REINITIALIZE:
    diagnostic = VBA.Err.Description
FailedReinitialize:
    ReinitializePageHandler = private_SetActionStatus(diagnostic, True)
End Function

Public Function EnvironmentSwitchHandler() As Boolean
    Dim environment As String
    Dim target As String
    Dim configPath As String
    Dim diagnostic As String
    Dim status As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If ex_PEB_Testing.fn_IsRunning() Then
        EnvironmentSwitchHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If m_environmentSwitchPending Then
        EnvironmentSwitchHandler = True
        Exit Function
    End If
    On Error GoTo EH_ENVIRONMENT
    environment = ex_PEB_Export.fn_Config("environment.name")
    Select Case environment
        Case "PROD"
            target = "TEST"
        Case "TEST"
            target = "PROD"
        Case Else
            Err.Raise 5, , "Unknown PEB environment: " & environment
    End Select
    configPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("environment.config." & target))
    status = VBA.Replace$(ex_PEB_Export.fn_Config("text.EnvironmentSwitchPending"), "{environment}", target)
    If Not ex_ConfigRefresh.fn_TryUpdate(diagnostic, configPath) Then GoTo FailedEnvironment
    m_environmentSwitchPending = True
    ' Reinitialize after this handler returns; never dispose an executing UI control.
    If Not ex_AppCommands.fn_RequestReinitialize() Then
        Err.Raise 5, , "Cannot schedule PEB environment initialization."
    End If
    EnvironmentSwitchHandler = private_SetActionStatus(status, False)
    Exit Function
EH_ENVIRONMENT:
    diagnostic = VBA.Err.Description
    m_environmentSwitchPending = False
FailedEnvironment:
    private_SetActionStatus diagnostic, True
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "PersonalEventBuilder"
End Function

Public Function PreviewWordHandler() As Boolean
    PreviewWordHandler = private_Export("PreviewWord")
End Function

Public Function SaveHistoryHandler() As Boolean
    SaveHistoryHandler = private_Export("History")
End Function

Public Function EventVersionsHandler() As Boolean
    Dim diagnostic As String
    Dim identifier As String

    On Error GoTo EH_VERSIONS
    If m_editHistoryRow Is Nothing Then
        EventVersionsHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.HistorySelectRequired"), True)
        Exit Function
    End If
    identifier = m_editHistoryRow("RowId")
    If Not ex_PEB_History.fn_SetStatus(m_editHistoryRow, m_editHistoryRow("Status"), "", m_editHistoryStamp, diagnostic) Then
        GoTo FailedVersions
    End If
    If Not private_RefreshHistory(identifier, diagnostic) Then
        GoTo FailedVersions
    End If
    If Not ex_PEB_HistoryRevisions.fn_Show(identifier, diagnostic) Then
        GoTo FailedVersions
    End If
    EventVersionsHandler = True
    Exit Function
FailedVersions:
    EventVersionsHandler = private_SetActionStatus(diagnostic, True)
    Exit Function
EH_VERSIONS:
    diagnostic = VBA.Err.Description
    Resume FailedVersions
End Function

Public Function CancelEventHandler() As Boolean
    CancelEventHandler = private_SetEventStatus("Cancelled")
End Function

Public Function RestoreEventHandler() As Boolean
    RestoreEventHandler = private_SetEventStatus("Active")
End Function

Public Function ShowHistoryHandler() As Boolean
    Dim diagnostic As String
    Dim record As obj_LookupRecord

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If ex_PEB_Testing.fn_IsRunning() Then Exit Function
    If Not private_RefreshHistory(VBA.vbNullString, diagnostic) Then GoTo Failed
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Visible") Then Exit Function
    If Not m_editHistoryRow Is Nothing Then
        For Each record In m_historyResult.Records
            If record.Key = m_editHistoryRow("RowId") Then
                If Not m_pageBase.BindingContext.SetObject("Data", "SelectedHistory", record) Then Exit Function
                Exit For
            End If
        Next record
    End If
    If Not private_SetActionStatus(ex_PEB_Export.fn_Config("text.HistorySelectHint"), False) Then Exit Function
    ShowHistoryHandler = m_pageBase.FlushUpdates()
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    private_SetActionStatus diagnostic, True
End Function

Public Function CloseCandidatesHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Collapsed") Then Exit Function
    CloseCandidatesHandler = m_pageBase.FlushUpdates()
End Function

Public Function CloseHistoryHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then Exit Function
    CloseHistoryHandler = m_pageBase.FlushUpdates()
End Function

Public Function SelectHistoryHandler(ByVal selected As Object) As Boolean
    Dim record As obj_LookupRecord
    Dim row As Object
    Dim updates As Object
    Dim wordContext As Object
    Dim field As Variant
    Dim diagnostic As String
    Dim mode As String
    Dim scalar As Variant
    Dim previousScreenUpdating As Boolean
    Dim candidates As obj_LookupResult
    Dim reporterActive As Boolean
    Dim hospitalActive As Boolean
    Dim documentActive As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    If selected Is Nothing Then Exit Function
    If Not TypeOf selected Is obj_LookupRecord Then Exit Function
    Set record = selected
    If m_historyRows Is Nothing Then Exit Function
    If Not m_historyRows.Exists(record.Key) Then Exit Function
    If Not private_TryReadFormValue("AcceptedOrderNumber", scalar) Then Exit Function
    If VBA.CStr(scalar) <> m_historyOrderNumber Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryListStale")
    If Not private_TryReadFormValue("AcceptedOrderDate", scalar) Then Exit Function
    If VBA.IsNumeric(scalar) And VBA.VarType(scalar) <> VBA.vbString Then scalar = VBA.CDate(scalar)
    If VBA.Format$(ex_PEB_History.fn_Date(scalar), "dd.mm.yyyy") <> m_historyOrderDate Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryListStale")
    End If
    Set row = m_historyRows(record.Key)
    If Not private_EndHistoryEdit() Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", scalar) Then Exit Function
    m_editHistoryAutoFill = scalar
    Application.ScreenUpdating = False
    Set candidates = m_lookupResult
    reporterActive = m_reporterActive
    hospitalActive = m_hospitalActive
    documentActive = m_documentActive
    Set m_lookupResult = Nothing
    m_modeId = VBA.vbNullString
    mode = row("EventId")
    If Not private_LoadMode(mode, diagnostic) Then GoTo Failed
    Set m_lookupResult = candidates
    m_reporterActive = reporterActive
    m_hospitalActive = hospitalActive
    m_documentActive = documentActive
    If Not candidates Is Nothing Then
        If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", candidates) Then GoTo Failed
    End If
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.fields"), "|")
        If row.Exists(field) Then updates("Form." & VBA.CStr(field)) = row(field)
    Next field
    updates("Form.AutoFillPersonnel") = False
    updates("Form.WordPreviewText") = VBA.vbNullString
    updates("Text.SaveHistory") = ex_PEB_Export.fn_Config("text.UpdateHistory")
    updates("Text.HistoryEditMode") = VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryEditing"), "{id}", record.Key)
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    Set wordContext = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.scalarInputs"), "|")
        If field <> "DurationDays" And row.Exists(field) Then wordContext(field) = row(field)
    Next field
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.booleanInputs"), "|")
        Select Case VBA.LCase$(VBA.Trim$(VBA.CStr(row(field))))
            Case "true"
                wordContext(field) = True
            Case "false"
                wordContext(field) = False
            Case Else
                Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError027"), "{value}", VBA.CStr(field))
        End Select
    Next field
    If Not m_pageBase.BindingContext.SetObject("Data", "WordContext", wordContext) Then GoTo Failed
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then GoTo Failed
    If Not m_pageBase.BindingContext.SetObject("Data", "SelectedHistory", record) Then GoTo Failed
    m_acceptedOrderNumber = row("AcceptedOrderNumber")
    m_acceptedOrderDate = ex_PEB_History.fn_Date(row("AcceptedOrderDate"))
    Set m_editHistoryRow = row
    m_editHistoryStamp = ex_PEB_History.fn_Stamp(row)
    If Not m_extensions.TryLoad(row("FormExtensionRows"), diagnostic) Then GoTo Failed
    If Not m_extensions.Build(ex_PEB_EventRoutes.fn_Value(m_modeId, "form"), diagnostic) Then GoTo Failed
    If Not m_pageBase.UpdatePage() Then Exit Function
    SelectHistoryHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.HistoryEditHint"), False)
    GoTo CleanupHistorySelection
EH:
    diagnostic = VBA.Err.Description
Failed:
    private_EndHistoryEdit
    private_SetActionStatus diagnostic, True
CleanupHistorySelection:
    Application.ScreenUpdating = previousScreenUpdating
End Function

Public Function ExportWordHandler() As Boolean
    Dim enabled As String
    Dim diagnostic As String
    Dim events As Collection
    Dim exporter As obj_PEB_ExptrWord
    Dim activeEvents As Collection
    Dim journal As Collection
    Dim historyItem As Object
    Dim exportItem As Object
    Dim journalDiagnostic As String
    Dim scalar As Variant
    Dim objectValue As Object
    Dim isObject As Boolean
    Dim orderNumber As String
    Dim orderDate As Variant
    Dim output As String
    Dim totalStartedAt As Double
    Dim startedAt As Double
    Dim succeeded As Boolean

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    enabled = ex_PEB_Export.fn_Config("export.Word.enabled")
    totalStartedAt = VBA.Timer
    If enabled = "false" Then
        ExportWordHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.ExportWordUnavailable"), True)
        Exit Function
    End If
    If enabled <> "true" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError045")
    End If
    If ex_PEB_Testing.fn_IsRunning() Then
        ExportWordHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If Not m_pageBase.BindingContext.TryGetValue("Form", "AcceptedOrderNumber", scalar, objectValue, isObject) Or isObject Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError046")
    End If
    orderNumber = VBA.CStr(scalar)
    If Not m_pageBase.BindingContext.TryGetValue("Form", "AcceptedOrderDate", scalar, objectValue, isObject) Or isObject Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError047")
    End If
    orderDate = scalar
    If VBA.IsNumeric(orderDate) And VBA.VarType(orderDate) <> VBA.vbString Then
        orderDate = VBA.CDate(orderDate)
    End If
    startedAt = VBA.Timer
    succeeded = ex_PEB_History.fn_ReadOrder(orderNumber, orderDate, events, diagnostic)
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.ReadEvents", startedAt
    If Not succeeded Then GoTo FailedHistoryWord
    startedAt = VBA.Timer
    Set activeEvents = New Collection
    Set journal = New Collection
    For Each historyItem In events
        Set exportItem = VBA.CreateObject("Scripting.Dictionary")
        exportItem.Add "RowId", historyItem("RowId")
        exportItem.Add "Revision", historyItem("Revision")
        exportItem.Add "Status", "CancelledOmitted"
        exportItem.Add "HistoryStamp", ex_PEB_MovementLinks.fn_EventStamp(historyItem)
        If historyItem("Status") = "Active" Then
            activeEvents.Add historyItem
            exportItem("Status") = "Generated"
        End If
        journal.Add exportItem
    Next historyItem
    If activeEvents.Count = 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryCancelledWord")
    End If
    Set events = activeEvents
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.PrepareJournal", startedAt
    Set exporter = New obj_PEB_ExptrWord
    If Not exporter.Initialize() Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError048")
    If Not exporter.TryGenerateHistory(events, output, diagnostic) Then GoTo FailedHistoryWord
    startedAt = VBA.Timer
    succeeded = ex_PEB_HistoryRevisions.fn_RecordExport(journal, output, "Word", journalDiagnostic)
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.RecordJournal", startedAt
    If Not succeeded Then
        ExportWordHandler = private_SetActionStatus(VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryJournalFailed"), "{details}", journalDiagnostic), True)
        GoTo CleanupHistoryWord
    End If
    startedAt = VBA.Timer
    succeeded = private_RefreshExportHistory(journalDiagnostic)
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.RefreshHistory", startedAt
    If Not succeeded Then
        ExportWordHandler = private_SetActionStatus(VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryRefreshAfterExport"), "{details}", journalDiagnostic), True)
        GoTo CleanupHistoryWord
    End If
    ExportWordHandler = private_SetActionStatus(VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryWordSuccess"), "{path}", output), False)
    GoTo CleanupHistoryWord
EH:
    diagnostic = VBA.Err.Description
FailedHistoryWord:
    private_SetActionStatus diagnostic, True
CleanupHistoryWord:
    If Not exporter Is Nothing Then exporter.Dispose
    If totalStartedAt > 0 Then
        ex_Core.fn_Diagnostic_WritePerf "Word.Export.TotalIncludingJournalAndRefresh", totalStartedAt
    End If
    If Not ex_Core.fn_Diagnostic_Flush(True) Then
        Debug.Print "Word command diagnostic flush failed."
    End If
End Function

Public Function ExportMovementHandler() As Boolean
    Dim enabled As String
    Dim diagnostic As String
    Dim scalar As Variant
    Dim objectValue As Object
    Dim isObject As Boolean
    Dim orderNumber As String
    Dim orderDate As Variant
    Dim conflicts As Collection
    Dim exported As Boolean

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    enabled = ex_PEB_Export.fn_Config("export.Movement.enabled")
    If enabled = "false" Then
        ExportMovementHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.ExportMovementUnavailable"), True)
        Exit Function
    End If
    If enabled <> "true" Then
        ex_PEB_MovementLinks.fn_Fail "InvalidEnabledSetting", enabled
    End If
    If ex_PEB_Testing.fn_IsRunning() Then
        ExportMovementHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If Not m_pageBase.BindingContext.TryGetValue("Form", "AcceptedOrderNumber", scalar, objectValue, isObject) Or isObject Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError046")
    End If
    orderNumber = VBA.CStr(scalar)
    If Not m_pageBase.BindingContext.TryGetValue("Form", "AcceptedOrderDate", scalar, objectValue, isObject) Or isObject Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError047")
    End If
    orderDate = scalar
    If VBA.IsNumeric(orderDate) And VBA.VarType(orderDate) <> VBA.vbString Then orderDate = VBA.CDate(orderDate)
    exported = ex_PEB_MovementBatch.fn_ExportOrder(orderNumber, orderDate, diagnostic, m_conflicts, conflicts)
    Set m_conflicts = conflicts
    If exported Then
        Set m_conflicts = Nothing
    End If
    If Not private_PublishConflicts() Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ConflictUiError")
    End If
    If Not exported Then
        GoTo FailedMovement
    End If
    If Not private_RefreshExportHistory(diagnostic) Then
        ExportMovementHandler = private_SetActionStatus(VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryRefreshAfterExport"), "{details}", diagnostic), True)
        Exit Function
    End If
    ExportMovementHandler = private_SetActionStatus(diagnostic, False)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
FailedMovement:
    private_SetActionStatus diagnostic, True
End Function

Public Function FromAnnualMainVacationHandler() As Boolean
    FromAnnualMainVacationHandler = private_SwitchMode("FromAnnualMainVacation")
End Function

Public Function SelectConflictHandler(ByVal selected As Object) As Boolean
    Dim record As obj_LookupRecord
    Dim item As Object
    Dim column As Variant
    Dim count As Long
    Dim message As String

    On Error GoTo EH_CONFLICT
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If selected Is Nothing Or m_conflicts Is Nothing Then Exit Function
    If Not TypeOf selected Is obj_LookupRecord Then Exit Function
    Set record = selected
    If Not private_TryReadFormValue("ConflictColumn", column) Then Exit Function
    If column <> 2 And column <> 3 Then Exit Function
    Set item = m_conflicts(VBA.CLng(record.Key))
    item("Choice") = "Movement"
    If column = 3 Then
        item("Choice") = "EventsHistory"
    End If
    For Each item In m_conflicts
        If VBA.Len(item("Choice")) > 0 Then
            count = count + 1
        End If
    Next item
    message = ex_PEB_Export.fn_Config("text.ConflictProgress")
    message = VBA.Replace$(message, "{resolved}", VBA.CStr(count))
    message = VBA.Replace$(message, "{total}", VBA.CStr(m_conflicts.Count))
    SelectConflictHandler = private_SetActionStatus(message, False)
    Exit Function
EH_CONFLICT:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function FromVacationForCuringHandler() As Boolean
    FromVacationForCuringHandler = private_SwitchMode("FromVacationForCuring")
End Function

Public Function FromVacationForFamilyHandler() As Boolean
    FromVacationForFamilyHandler = private_SwitchMode("FromVacationForFamily")
End Function

Public Function ToHospitalHandler() As Boolean
    ToHospitalHandler = private_SwitchMode("ToHospital")
End Function

Public Function ToAnnualMainVacationHandler() As Boolean
    ToAnnualMainVacationHandler = private_SwitchMode("ToAnnualMainVacation")
End Function

Public Function ToVacationForCuringHandler() As Boolean
    ToVacationForCuringHandler = private_SwitchMode("ToVacationForCuring")
End Function

Public Function ToVacationForFamilyHandler() As Boolean
    ToVacationForFamilyHandler = private_SwitchMode("ToVacationForFamily")
End Function

Public Function LoadTestHandler() As Boolean
    LoadTestHandler = private_LoadTest(False)
End Function

Public Function NextTestHandler() As Boolean
    NextTestHandler = private_LoadTest(True)
End Function

' //
' // Private
' //
Private Function private_LoadTest(ByVal nextStep As Boolean) As Boolean
    Dim pebTestScenarios As obj_PEB_TestScenarios
    Dim updates As Object
    Dim snapshot As Object
    Dim formValues As Object
    Dim source As Variant
    Dim oldResult As obj_LookupResult
    Dim emptyResult As obj_LookupResult
    Dim records As New Collection
    Dim scalar As Variant
    Dim isObject As Boolean
    Dim key As Variant
    Dim scenarioId As String
    Dim stepOrder As Long
    Dim modeId As String
    Dim oldMode As String
    Dim oldDirection As String
    Dim oldEventType As String
    Dim oldOrder As String
    Dim oldOrderDate As Variant
    Dim oldReporter As Boolean
    Dim oldHospital As Boolean
    Dim oldDocument As Boolean
    Dim oldFields As Collection
    Dim oldDisplayFields As Collection
    Dim status As String
    Dim diagnostic As String
    Dim rollbackDiagnostic As String
    Dim formChanged As Boolean

    If Not m_isInitialized Or m_isDisposed Or m_testBusy Or m_environmentSwitchPending Then Exit Function
    If ex_PEB_Testing.fn_IsRunning() Then
        private_LoadTest = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    On Error GoTo EH
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    If Not private_TryReadFormValue("TestScenario", scalar) Then Err.Raise 5, , "Cannot read Form.TestScenario."
    scenarioId = VBA.Trim$(VBA.CStr(scalar))
    stepOrder = 1
    If nextStep Then
        If scenarioId <> m_testScenario Or m_testStepOrder < 1 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestLoadFirst")
        stepOrder = m_testStepOrder + 1
    End If
    m_testBusy = True
    Set pebTestScenarios = New obj_PEB_TestScenarios
    If Not pebTestScenarios.Initialize(diagnostic) Then GoTo Failed
    If Not pebTestScenarios.TryPrepare(scenarioId, stepOrder, Not nextStep, modeId, updates, status, diagnostic) Then GoTo Failed

    ' Snapshot the form before publishing any mode or input changes.
    Set snapshot = VBA.CreateObject("Scripting.Dictionary")
    For Each source In VBA.Array("Form", "Text", "Settings")
        If Not m_pageBase.BindingContext.TryGetValue(VBA.CStr(source), VBA.vbNullString, scalar, formValues, isObject) Then
            Err.Raise 5, , "Cannot snapshot PEB bindings: " & source
        End If
        For Each key In formValues.Keys
            If VBA.IsObject(formValues(key)) Then Err.Raise 5, , "Scalar binding value required: " & key
            snapshot(source & "." & key) = formValues(key)
        Next key
    Next source
    Set oldResult = m_lookupResult
    Set oldFields = m_modeFields
    Set oldDisplayFields = m_displayFields
    oldMode = m_modeId
    oldDirection = m_modeDirection
    oldEventType = m_eventType
    oldOrder = m_acceptedOrderNumber
    oldOrderDate = m_acceptedOrderDate
    oldReporter = m_reporterActive
    oldHospital = m_hospitalActive
    oldDocument = m_documentActive
    formChanged = True

    ' Empty candidates without querying personnel or secondary directories.
    Set emptyResult = New obj_LookupResult
    If Not emptyResult.Initialize(records, m_modeFields, False, diagnostic) Then GoTo Failed
    Set m_lookupResult = emptyResult
    m_modeId = VBA.vbNullString
    m_reporterActive = False
    m_hospitalActive = False
    m_documentActive = False
    If Not private_LoadMode(modeId, diagnostic) Then GoTo Failed
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    m_acceptedOrderNumber = updates("Form.AcceptedOrderNumber")
    m_acceptedOrderDate = updates("Form.AcceptedOrderDate")
    If Not m_exportService Is Nothing Then m_exportService.Dispose
    Set m_exportService = Nothing
    ex_TableDataRuntime.fn_ResetCaches
    m_lookupService.InvalidateCache
    m_candidateEnricher.InvalidateCache
    m_orderResolver.InvalidateCache
    If Not m_hospitalService Is Nothing Then m_hospitalService.InvalidateCache
    If Not m_documentService Is Nothing Then m_documentService.InvalidateCache
    m_extensions.InvalidateCache
    If Not m_pageBase.FlushUpdates() Then Err.Raise 5, , "Cannot render the loaded scenario."
    If Not pebTestScenarios.TryCommit(diagnostic) Then GoTo Failed
    m_testScenario = scenarioId
    m_testStepOrder = stepOrder
    Set m_pendingOrder = Nothing
    private_LoadTest = True
Cleanup:
    If Not pebTestScenarios Is Nothing Then pebTestScenarios.Dispose
    m_testBusy = False
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    If VBA.Len(diagnostic) = 0 Then diagnostic = "Scenario loading failed."
    If formChanged Then
        ' Discard unsaved history and restore form values and the previous candidate projection.
        On Error Resume Next
        m_modeId = oldMode
        m_modeDirection = oldDirection
        m_eventType = oldEventType
        Set m_modeFields = oldFields
        Set m_displayFields = oldDisplayFields
        Set m_lookupResult = oldResult
        m_reporterActive = oldReporter
        m_hospitalActive = oldHospital
        m_documentActive = oldDocument
        m_acceptedOrderNumber = oldOrder
        m_acceptedOrderDate = oldOrderDate
        If Not m_pageBase.BindingContext.TryApplyValues(snapshot, rollbackDiagnostic) Then
            diagnostic = diagnostic & " Restore form: " & rollbackDiagnostic
        End If
        If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", oldResult) Then
            diagnostic = diagnostic & " Cannot restore candidates."
        End If
        If Not private_PublishViewport(rollbackDiagnostic) Then
            diagnostic = diagnostic & " Cannot restore form fields."
        End If
        If Not m_pageBase.FlushUpdates() Then diagnostic = diagnostic & " Cannot render restored form."
        If VBA.Err.Number <> 0 Then diagnostic = diagnostic & " Restore: " & VBA.Err.Description
        On Error GoTo 0
    End If
    If Not pebTestScenarios Is Nothing Then pebTestScenarios.Dispose
    m_testBusy = False
    status = VBA.Replace$(ex_PEB_Export.fn_Config("text.TestFailed"), "{error}", diagnostic)
    private_SetActionStatus status, True
    ex_WindowsUi.fn_ShowMessage status, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    GoTo Cleanup
End Function

Private Function private_SearchPersonnel(ByVal showCandidates As Boolean) As Boolean
    Dim text As Variant
    Dim formKind As Variant
    Dim keyField As String
    Dim diagnostic As String
    Dim result As obj_LookupResult
    Dim status As String
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_extensions.HideCandidates() Then Exit Function
    If m_reporterActive Or m_hospitalActive Or m_documentActive Then
        If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    m_reporterActive = False
    m_hospitalActive = False
    m_documentActive = False
    If Not private_TryReadFormValue("ModeForm", formKind) Then Exit Function
    If Not private_ReadModeConfig("forms." & VBA.CStr(formKind) & ".lookup.key", keyField, diagnostic) Then GoTo Failure
    If Not private_TryReadFormValue(keyField, text) Then Exit Function
    If Not m_lookupService.TrySearch(VBA.CStr(text), result, diagnostic) Then
        If Not m_lookupService.TrySearch(vbNullString, result, status) Then GoTo Failure
        If Not private_TryEnrichResult(result, status) Then GoTo Failure
        Set m_lookupResult = result
        m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", result
        m_pageBase.BindingContext.SetValue "Text", "PersonnelStatus", m_messages("PersonnelError")
        GoTo Failure
    End If
    If Not private_TryEnrichResult(result, diagnostic) Then
        If m_lookupService.TrySearch(vbNullString, result, status) Then
            If Not private_TryEnrichResult(result, status) Then GoTo Failure
            Set m_lookupResult = result
            m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", result
        End If
        GoTo Failure
    End If
    Set m_lookupResult = result
    ex_Core.fn_Diagnostic_WriteLog "LOOKUP_STAGE_STARTED | Name=PublishCandidates"
    startedAt = VBA.Timer
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then Exit Function
    ex_Core.fn_Diagnostic_WritePerf "Lookup.PublishCandidates | Rows=" & VBA.CStr(result.RowCount), startedAt
    If showCandidates Then
        If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Visible") Then
            Exit Function
        End If
    End If
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(text))) < m_lookupProfile.MinChars Then
        status = VBA.Replace$(m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(m_lookupProfile.MinChars))
    ElseIf result.RowCount = 0 Then
        status = m_messages("PersonnelEmpty")
    Else
        status = VBA.Replace$(m_messages("PersonnelCount"), "{count}", VBA.CStr(result.RowCount))
        If result.HasMore Then status = VBA.Replace$(m_messages("PersonnelMore"), "{count}", VBA.CStr(result.RowCount))
    End If
    private_SearchPersonnel = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", status)
    If private_SearchPersonnel Then
        private_SearchPersonnel = private_SetActionStatus(status, False)
    End If
    Exit Function
Failure:
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
End Function

Private Function private_SearchSecondary( _
    ByVal clearDependents As Boolean, _
    Optional ByVal hospital As Boolean = False, _
    Optional ByVal showCandidates As Boolean = True, _
    Optional ByVal document As Boolean = False _
) As Boolean
    Dim key As String
    Dim text As Variant
    Dim result As obj_LookupResult
    Dim projected As obj_LookupResult
    Dim column As Object
    Dim updates As Object
    Dim diagnostic As String
    Dim status As String
    Dim records As Collection
    Dim fields As Collection
    Dim profile As obj_LookupProfile
    Dim service As obj_LookupService
    Dim group As String

    On Error GoTo EH
    group = "reporter"
    If Not m_extensions.HideCandidates() Then Exit Function
    If hospital Then group = "hospital"
    If document Then group = "document"
    If Not private_SecondaryFields(group, fields, diagnostic) Then
        GoTo Failed
    End If
    If document Then
        If m_documentService Is Nothing Then
            Set m_documentProfile = New obj_LookupProfile
            If Not m_documentProfile.Initialize("PersonalEventBuilder::lookup.Document", diagnostic) Then GoTo Failed
            Set m_documentService = New obj_LookupService
            If Not m_documentService.Initialize(m_documentProfile) Then
                Err.Raise 5, , "Cannot initialize document lookup service."
            End If
        End If
        Set profile = m_documentProfile
        Set service = m_documentService
    ElseIf hospital Then
        If m_hospitalService Is Nothing Then
            Set m_hospitalProfile = New obj_LookupProfile
            If Not m_hospitalProfile.Initialize("PersonalEventBuilder::lookup.Hospital", diagnostic) Then GoTo Failed
            Set m_hospitalService = New obj_LookupService
            If Not m_hospitalService.Initialize(m_hospitalProfile) Then
                Err.Raise 5, , "Cannot initialize hospital lookup service."
            End If
        End If
        Set profile = m_hospitalProfile
        Set service = m_hospitalService
    Else
        Set profile = m_lookupProfile
        Set service = m_lookupService
    End If
    If Not private_ReadModeConfig(group & ".lookup.key", key, diagnostic) Then
        GoTo Failed
    End If
    If Not private_TryReadFormValue(key, text) Then
        Err.Raise 5, , "Secondary lookup key binding not found: " & key
    End If

    ' Secondary lookups share the candidate area without changing primary personnel fields.
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
        Err.Raise 5, , "Cannot clear the previous candidate selection."
    End If
    If clearDependents Then
        For Each column In fields
            If column("formField") <> key Then
                updates("Form." & column("formField")) = VBA.vbNullString
            End If
        Next column
    End If
    If updates.Count > 0 Then
        If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then
            GoTo Failed
        End If
    End If
    m_reporterActive = Not hospital And Not document
    m_hospitalActive = hospital
    m_documentActive = document

    ' Search the selected directory without enriching secondary records with absence data.
    If Not service.TrySearch(VBA.CStr(text), result, diagnostic) Then
        GoTo Failed
    End If
    If m_reporterActive Then
        If Not private_EnrichReporters(result.Records, diagnostic) Then GoTo Failed
    End If
    Set projected = New obj_LookupResult
    If Not projected.Initialize(result.Records, fields, result.HasMore, diagnostic) Then
        GoTo Failed
    End If
    Set m_lookupResult = projected
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", projected) Then
        Err.Raise 5, , "Cannot publish secondary candidates."
    End If
    If showCandidates Then
        If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Visible") Then
            Exit Function
        End If
    End If
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(text))) < profile.MinChars Then
        status = VBA.Replace$(m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(profile.MinChars))
    ElseIf projected.RowCount = 0 Then
        status = m_messages("PersonnelEmpty")
    ElseIf projected.HasMore Then
        status = VBA.Replace$(m_messages("PersonnelMore"), "{count}", VBA.CStr(projected.RowCount))
    Else
        status = VBA.Replace$(m_messages("PersonnelCount"), "{count}", VBA.CStr(projected.RowCount))
    End If
    If Not private_SetActionStatus(status, False) Then
        Exit Function
    End If
    ' Publishing candidates requests a layout refresh after the handler returns.
    ' A full page rebuild must not dispose a control while its handler is executing.
    private_SearchSecondary = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    ' On failure, clear stale records from the previous selection context.
    If Not fields Is Nothing Then
        Set records = New Collection
        Set projected = New obj_LookupResult
        If projected.Initialize(records, fields, False, status) Then
            Set m_lookupResult = projected
            m_reporterActive = Not hospital And Not document
            m_hospitalActive = hospital
            m_documentActive = document
            m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", projected
        End If
    End If
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "PersonalEventBuilder"
End Function

Private Function private_EnrichReporters( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim values As Object
    Dim dateText As String

    On Error GoTo EH_REPORT_DATE
    If VBA.InStr(1, "|" & ex_PEB_Export.fn_Config("lookup.IncomingDate.orderDateModes") & "|", "|" & m_modeId & "|", VBA.vbTextCompare) > 0 Then
        If VBA.Len(ex_PEB_Export.fn_Text(m_acceptedOrderDate)) > 0 Then
            dateText = VBA.Format$(ex_PEB_History.fn_Date(m_acceptedOrderDate), "dd.mm.yyyy")
        End If
    End If
    If Not ex_PEB_Export.fn_TryApplyReporterPositions(records, diagnostic) Then Exit Function
    For Each record In records
        Set values = VBA.CreateObject("Scripting.Dictionary")
        values("Event.IncomingDate") = dateText
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, values) Then Err.Raise 5, , "Cannot initialize reporter order date."
        If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot merge reporter order date."
    Next record
    private_EnrichReporters = True
    Exit Function
EH_REPORT_DATE:
    diagnostic = VBA.Err.Description
End Function

Private Function private_SecondaryFields( _
    ByVal group As String, _
    ByRef result As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim countText As String
    Dim index As Long
    Dim prefix As String
    Dim value As String
    Dim column As Object
    Dim fields As Collection
    Dim caption As String

    If group = "document" Then
        Set result = m_documentFields
    ElseIf group = "hospital" Then
        Set result = m_hospitalFields
    Else
        Set result = m_reporterFields
    End If
    If Not result Is Nothing Then
        private_SecondaryFields = True
        Exit Function
    End If
    If Not private_ReadModeConfig(group & ".fields.count", countText, diagnostic) Then
        Exit Function
    End If
    If Not VBA.IsNumeric(countText) Then
        Err.Raise 5, , "Secondary field count must be numeric."
    End If
    If VBA.CDbl(countText) < 1 Or VBA.CDbl(countText) > 16 Or VBA.CDbl(countText) <> VBA.Fix(VBA.CDbl(countText)) Then
        Err.Raise 5, , "Secondary field count must be an integer from 1 to 16."
    End If
    Set fields = New Collection
    For index = 1 To VBA.CLng(countText)
        prefix = group & ".fields." & VBA.CStr(index)
        Set column = VBA.CreateObject("Scripting.Dictionary")
        If Not private_ReadModeConfig(prefix & ".formField", value, diagnostic) Then
            Exit Function
        End If
        column("formField") = value
        If Not private_ReadModeConfig(prefix & ".record", value, diagnostic) Then
            Exit Function
        End If
        column("name") = value
        If Not private_ReadModeConfig(prefix & ".captionKey", value, diagnostic) Then
            Exit Function
        End If
        If Not ex_Core.fn_TryGetWorkbookConfigValue(value, caption) Then
            diagnostic = "Required secondary caption not found: " & value
            Exit Function
        End If
        column("caption") = caption
        If Not private_ReadModeConfig(prefix & ".sourceAlias", value, diagnostic) Then
            Exit Function
        End If
        column("sourceAlias") = value
        column("apply") = True
        fields.Add column
    Next index
    If group = "document" Then
        Set m_documentFields = fields
    ElseIf group = "hospital" Then
        Set m_hospitalFields = fields
    Else
        Set m_reporterFields = fields
    End If
    Set result = fields
    private_SecondaryFields = True
End Function

Private Function private_RefreshExportHistory(ByRef diagnostic As String) As Boolean
    Dim identifier As String
    Dim message As String

    On Error GoTo EH_REFRESH
    private_RefreshExportHistory = True
    If m_historyRows Is Nothing Then
        Exit Function
    End If
    If m_historyRows.Count = 0 Then
        Exit Function
    End If
    If Not m_editHistoryRow Is Nothing Then
        identifier = m_editHistoryRow("RowId")
    End If
    private_RefreshExportHistory = private_RefreshHistory(identifier, message)
    If Not private_RefreshExportHistory Then
        diagnostic = message
    End If
    Exit Function
EH_REFRESH:
    diagnostic = VBA.Err.Description
    private_RefreshExportHistory = False
End Function

Private Function private_SetEventStatus(ByVal status As String) As Boolean
    Dim reason As String
    Dim diagnostic As String
    Dim identifier As String

    On Error GoTo EH_STATUS
    If m_editHistoryRow Is Nothing Then
        private_SetEventStatus = private_SetActionStatus(ex_PEB_Export.fn_Config("text.HistorySelectRequired"), True)
        Exit Function
    End If
    reason = VBA.InputBox(ex_PEB_Export.fn_Config("text.HistoryReasonPrompt"))
    If VBA.Len(VBA.Trim$(reason)) = 0 Then
        Exit Function
    End If
    identifier = m_editHistoryRow("RowId")
    If Not ex_PEB_History.fn_SetStatus(m_editHistoryRow, status, reason, m_editHistoryStamp, diagnostic) Then
        GoTo FailedStatus
    End If
    If Not private_RefreshHistory(identifier, diagnostic) Then
        GoTo FailedStatus
    End If
    private_SetEventStatus = private_SetActionStatus(ex_PEB_Export.fn_Config("text.HistoryStatusSaved"), False)
    Exit Function
FailedStatus:
    private_SetEventStatus = private_SetActionStatus(diagnostic, True)
    Exit Function
EH_STATUS:
    diagnostic = VBA.Err.Description
    Resume FailedStatus
End Function

Private Function private_EndHistoryEdit() As Boolean
    Dim updates As Object
    Dim diagnostic As String

    Set m_editHistoryRow = Nothing
    m_editHistoryStamp = VBA.vbNullString
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Text.SaveHistory") = ex_PEB_Export.fn_Config("text.SaveHistory")
    updates("Text.HistoryEditMode") = ex_PEB_Export.fn_Config("text.HistoryEditMode")
    If Not VBA.IsEmpty(m_editHistoryAutoFill) Then updates("Form.AutoFillPersonnel") = m_editHistoryAutoFill
    m_editHistoryAutoFill = Empty
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then Exit Function
    private_EndHistoryEdit = m_pageBase.BindingContext.SetValue("Data", "SelectedHistory", VBA.vbNullString)
End Function

Private Function private_RefreshHistory( _
    ByVal selectedId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim events As Collection
    Dim emptyRows As New Collection
    Dim orderNumber As Variant
    Dim orderDate As Variant
    Dim failure As String

    If Not private_TryReadFormValue("AcceptedOrderNumber", orderNumber) Then Exit Function
    If Not private_TryReadFormValue("AcceptedOrderDate", orderDate) Then Exit Function
    If VBA.IsNumeric(orderDate) And VBA.VarType(orderDate) <> VBA.vbString Then orderDate = VBA.CDate(orderDate)
    If Not ex_PEB_History.fn_ReadOrder(VBA.CStr(orderNumber), orderDate, events, diagnostic) Then
        failure = diagnostic
        private_PublishHistory emptyRows, VBA.vbNullString, diagnostic
        diagnostic = failure
        Exit Function
    End If
    If Not private_PublishHistory(events, selectedId, diagnostic) Then Exit Function
    m_historyOrderNumber = VBA.CStr(orderNumber)
    m_historyOrderDate = VBA.Format$(ex_PEB_History.fn_Date(orderDate), "dd.mm.yyyy")
    private_RefreshHistory = True
End Function

Private Function private_PublishConflicts(Optional ByVal flush As Boolean = True) As Boolean
    Dim records As New Collection
    Dim columns As New Collection
    Dim result As obj_LookupResult
    Dim record As obj_LookupRecord
    Dim item As Object
    Dim fields As Object
    Dim column As Object
    Dim key As Variant
    Dim index As Long
    Dim diagnostic As String
    Dim visibility As String

    visibility = "Collapsed"
    If Not m_conflicts Is Nothing Then
        For Each item In m_conflicts
            index = index + 1
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            fields("Field") = item("EventRowId") & " / " & item("Side") & " / " & item("Field")
            fields("MovementValue") = item("MovementValue")
            fields("GeneratedValue") = item("GeneratedValue")
            If item("Choice") = "Movement" Then
                fields("MovementValue") = ex_PEB_Export.fn_Config("text.ConflictSelectedPrefix") & fields("MovementValue")
            ElseIf item("Choice") = "EventsHistory" Then
                fields("GeneratedValue") = ex_PEB_Export.fn_Config("text.ConflictSelectedPrefix") & fields("GeneratedValue")
            End If
            Set record = New obj_LookupRecord
            If Not record.Initialize(VBA.CStr(index), fields) Then Exit Function
            records.Add record
        Next item
        If m_conflicts.Count > 0 Then
            visibility = "Visible"
        End If
    End If
    For Each key In VBA.Array("Field", "MovementValue", "GeneratedValue")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column("name") = key
        column("caption") = ex_PEB_Export.fn_Config("text.ConflictColumn." & key)
        columns.Add column
    Next key
    Set result = New obj_LookupResult
    If Not result.Initialize(records, columns, False, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "MovementConflicts", result) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ConflictsVisibility", visibility) Then Exit Function
    private_PublishConflicts = True
    If flush Then
        private_PublishConflicts = m_pageBase.FlushUpdates()
    End If
End Function

Private Function private_PublishHistory( _
    ByVal events As Collection, _
    ByVal selectedId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim records As New Collection
    Dim columns As New Collection
    Dim rows As Object
    Dim row As Object
    Dim fields As Object
    Dim column As Object
    Dim key As Variant
    Dim index As Long
    Dim prefix As String
    Dim record As obj_LookupRecord
    Dim selected As obj_LookupRecord
    Dim result As obj_LookupResult

    Set rows = VBA.CreateObject("Scripting.Dictionary")
    For Each row In events
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        For Each key In row.Keys
            If Not VBA.IsObject(row(key)) Then fields(key) = row(key)
        Next key
        Set record = New obj_LookupRecord
        If Not record.Initialize(VBA.CStr(row("RowId")), fields) Then Exit Function
        records.Add record
        rows.Add record.Key, row
        If record.Key = selectedId Then Set selected = record
    Next row
    For index = 1 To ex_PEB_Export.fn_Count("history.list.columns.count")
        prefix = "history.list.columns." & VBA.CStr(index)
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column("name") = ex_PEB_Export.fn_Config(prefix & ".field")
        column("caption") = ex_PEB_Export.fn_Config(prefix & ".caption")
        columns.Add column
    Next index
    Set result = New obj_LookupResult
    If Not result.Initialize(records, columns, False, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "HistoryEvents", result) Then Exit Function
    Set m_historyResult = result
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedHistory", VBA.vbNullString) Then Exit Function
    Set m_historyRows = rows
    If VBA.Len(selectedId) > 0 Then
        If selected Is Nothing Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryListStale")
        If Not m_pageBase.BindingContext.SetObject("Data", "SelectedHistory", selected) Then Exit Function
        Set m_editHistoryRow = rows(selectedId)
        m_editHistoryStamp = ex_PEB_History.fn_Stamp(m_editHistoryRow)
    End If
    private_PublishHistory = True
End Function

Private Function private_Export(ByVal action As String) As Boolean
    Dim extensionRows As Collection
    Dim pebExportService As obj_PEB_ExportService
    Dim pebExptrMovement As obj_PEB_ExptrMovement
    Dim values As Object
    Dim diagnostic As String
    Dim exporter As String
    Dim success As Boolean
    Dim status As String
    Dim failureTemplate As String
    Dim wordContext As Object
    Dim scalar As Variant
    Dim isObject As Boolean
    Dim contextField As Variant

    If ex_PEB_Testing.fn_IsRunning() Then
        private_Export = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    failureTemplate = ex_PEB_Export.fn_Config("text.ExportFailed")
    status = ex_PEB_Export.fn_Config("text.Export" & action & "Success")
    exporter = action
    If action = "PreviewWord" Then
        exporter = "Word"
    End If
    If action = "PreviewWord" Then
        If Not m_pageBase.BindingContext.SetValue("Form", "WordPreviewText", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    If m_exportService Is Nothing Then
        Set m_exportService = New obj_PEB_ExportService
        If Not m_exportService.Initialize() Then
            Set m_exportService = Nothing
            Err.Raise 5, , "Cannot initialize export service."
        End If
    End If
    Set pebExportService = m_exportService
    If Not m_extensions.Collect(extensionRows, diagnostic) Then GoTo Failed
    If Not m_editHistoryRow Is Nothing Then
        If m_modeId <> VBA.CStr(m_editHistoryRow("EventId")) Then
            If m_pageBase.BindingContext.TryGetValue("Data", "WordContext", scalar, wordContext, isObject) Then
                If isObject And Not wordContext Is Nothing Then
                    For Each contextField In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.booleanInputs"), "|")
                        If wordContext.Exists(contextField) Then wordContext.Remove contextField
                    Next contextField
                End If
            End If
        End If
    End If
    If Not pebExportService.TryPrepare( _
        m_pageBase.BindingContext, _
        exporter, _
        values, _
        diagnostic, _
        action <> "PreviewWord", _
        Nothing, _
        extensionRows _
    ) Then GoTo Failed
    If exporter = "History" And Not m_editHistoryRow Is Nothing Then
        values("RowId") = m_editHistoryRow("RowId")
        values("ExpectedHistoryStamp") = m_editHistoryStamp
        If VBA.Len(values("DurationDays")) = 0 And VBA.Len(m_editHistoryRow("DateToShort")) > 0 Then
            If ex_PEB_History.fn_Date(m_editHistoryRow("DateToShort")) < ex_PEB_History.fn_Date(values("EventDate")) Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError028")
            End If
            values("DateToShort") = m_editHistoryRow("DateToShort")
        End If
    End If
    If exporter = "Word" Then
        If values("WordPolicy") = "Omit" Then
            status = ex_PEB_Export.fn_Config("text.ExportWordOmitted")
        End If
    End If
    If action = "PreviewWord" Then
        success = m_pageBase.BindingContext.SetValue( _
            "Form", _
            "WordPreviewText", _
            VBA.Replace$(VBA.CStr(values("WordText")), VBA.vbCrLf, VBA.vbLf) _
        )
        If Not success Then
            diagnostic = "Cannot publish the Word preview binding."
        End If
    ElseIf exporter = "History" Then
        success = ex_PEB_History.fn_Save(values, diagnostic, (m_editHistoryRow Is Nothing), m_exportService)
        If success Then
            If values("FormExtensionRows").Count <> extensionRows.Count Then
                If Not m_extensions.TryLoad(values("FormExtensionRows"), diagnostic, True) Then GoTo Failed
                If Not ex_PEB_ExtRefresh.fn_Request(m_extensions) Then GoTo Failed
            Else
                If Not m_extensions.TryAcceptSaved(values("FormExtensionRows"), diagnostic) Then GoTo Failed
            End If
        End If
        If success Then status = VBA.Replace$(status, "{action}", VBA.CStr(values("HistoryAction")))
        If success And Not m_editHistoryRow Is Nothing Then
            If Not private_RefreshHistory(VBA.CStr(values("RowId")), diagnostic) Then
                status = status & " " & diagnostic
                private_EndHistoryEdit
            End If
        End If
    Else
        Set pebExptrMovement = New obj_PEB_ExptrMovement
        If Not pebExptrMovement.Initialize() Then
            Err.Raise 5, , "Cannot initialize Movement exporter."
        End If
        success = pebExptrMovement.TryExport(values, diagnostic)
    End If
    If Not success Then
        GoTo Failed
    End If
    private_SetActionStatus status, False
    private_Export = True
Cleanup:
    If Not pebExptrMovement Is Nothing Then
        pebExptrMovement.Dispose
    End If
    Set pebExportService = Nothing
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    If VBA.Len(failureTemplate) > 0 Then
        diagnostic = VBA.Replace$(failureTemplate, "{details}", diagnostic)
    End If
    private_SetActionStatus diagnostic, True
    GoTo Cleanup
End Function

Private Function private_CaptureModeConfig(ByRef diagnostic As String) As Boolean
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim configTable As ListObject
    Dim values As Variant
    Dim keyColumn As Long
    Dim rowIndex As Long
    Dim key As String
    Dim prefix As String
    Dim field As Variant
    Dim fields As String

    On Error GoTo EH
    prefix = "PersonalEventBuilder::"
    For Each sheet In ThisWorkbook.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = "tbConfig" Then
                If Not configTable Is Nothing Then Err.Raise 5, , "Duplicate configuration table tbConfig."
                Set configTable = table
            End If
        Next table
    Next sheet
    If configTable Is Nothing Then Err.Raise 5, , "Configuration table tbConfig was not found."
    If configTable.DataBodyRange Is Nothing Then Err.Raise 5, , "Configuration table tbConfig is empty."
    keyColumn = configTable.ListColumns("Key").Index
    If keyColumn >= configTable.ListColumns.Count Then Err.Raise 5, , "Configuration values must follow Key."
    values = configTable.DataBodyRange.Value2
    Set m_modeConfig = VBA.CreateObject("Scripting.Dictionary")
    m_modeConfig.CompareMode = VBA.vbTextCompare
    ' Read the worksheet once; mode switches use the in-memory snapshot.
    For rowIndex = 1 To UBound(values, 1)
        key = VBA.Trim$(VBA.CStr(values(rowIndex, keyColumn)))
        If VBA.Left$(key, VBA.Len(prefix)) = prefix Then
            key = VBA.Mid$(key, VBA.Len(prefix) + 1)
            If m_modeConfig.Exists(key) Then Err.Raise 5, , "Duplicate configuration key: " & prefix & key
            m_modeConfig.Add key, VBA.Trim$(VBA.CStr(values(rowIndex, keyColumn + 1)))
        End If
    Next rowIndex
    ' Batch updates require existing scalar targets before loading the initial mode.
    If Not private_ReadModeConfig("forms.readOnly.fields", fields, diagnostic) Then Exit Function
    For Each field In Split(fields, "|")
        If Not m_pageBase.BindingContext.SetValue("Form", CStr(field) & "ReadOnly", False) Then
            diagnostic = "Cannot initialize read-only binding: " & field
            Exit Function
        End If
    Next field
    private_CaptureModeConfig = True
    Exit Function
EH:
    diagnostic = "Mode configuration: " & VBA.Err.Description
End Function

Private Function private_ReadModeConfig( _
    ByVal key As String, _
    ByRef value As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim parts As Variant

    value = VBA.vbNullString
    On Error GoTo EH
    If VBA.Left$(key, 6) = "modes." And key <> "modes.default" Then
        parts = VBA.Split(key, ".")
        If UBound(parts) <> 2 Then
            Err.Raise 5, , "Invalid event property: " & key
        End If
        value = ex_PEB_EventRoutes.fn_Value(VBA.CStr(parts(1)), VBA.CStr(parts(2)))
        private_ReadModeConfig = True
        Exit Function
    End If
    If Not m_modeConfig Is Nothing Then
        If m_modeConfig.Exists(key) Then
            value = m_modeConfig(key)
            private_ReadModeConfig = (VBA.Len(value) > 0)
        End If
    End If
    If Not private_ReadModeConfig Then diagnostic = "Required configuration key not found: PersonalEventBuilder::" & key
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Private Function private_PublishExtendedModes(ByRef diagnostic As String) As Boolean
    Dim items As Collection
    Dim captions As Object
    Dim eventId As Variant
    Dim caption As String

    On Error GoTo EH_EXTENDED
    Set items = New Collection
    Set captions = VBA.CreateObject("Scripting.Dictionary")
    captions.CompareMode = VBA.vbTextCompare
    For Each eventId In ex_PEB_EventRoutes.fn_Ids()
        If ex_PEB_EventRoutes.fn_Value(VBA.CStr(eventId), "visible") = "false" Then
            caption = ex_PEB_EventRoutes.fn_Value(VBA.CStr(eventId), "caption")
            If captions.Exists(caption) Then Err.Raise 5, , "Duplicate extended event caption: " & caption
            captions.Add caption, True
            items.Add caption
        End If
    Next eventId
    If Not m_pageBase.BindingContext.SetObject("Data", "ExtendedModes", items) Then Exit Function
    private_PublishExtendedModes = True
    Exit Function
EH_EXTENDED:
    diagnostic = VBA.Err.Description
End Function

Private Function private_LoadMode( _
    ByVal modeId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim direction As String
    Dim formKind As String
    Dim caption As String
    Dim countText As String
    Dim fieldName As String
    Dim captionKey As String
    Dim recordName As String
    Dim sourceAlias As String
    Dim fieldCaption As String
    Dim eventType As String
    Dim eventDateLabel As String
    Dim text As String
    Dim mode As Variant
    Dim field As Variant
    Dim index As Long
    Dim prefix As String
    Dim columns As New Collection
    Dim column As Object
    Dim records As Collection
    Dim projectedResult As obj_LookupResult
    Dim enrichedRecords As Collection
    Dim hasMore As Boolean
    Dim modes As Variant
    Dim currentForm As Variant
    Dim secondaryCandidates As Boolean

    On Error GoTo EH
    secondaryCandidates = m_reporterActive Or m_hospitalActive Or m_documentActive
    currentForm = vbNullString
    If Len(m_modeId) > 0 Then currentForm = ex_PEB_EventRoutes.fn_Value(m_modeId, "form")
    modes = ex_PEB_EventRoutes.fn_VisibleIds()
    caption = VBA.vbNullString
    text = "secondaryButton"
    If ex_PEB_EventRoutes.fn_Value(modeId, "visible") = "false" Then
        caption = ex_PEB_EventRoutes.fn_Value(modeId, "caption")
        text = "activeModeButton"
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "ExtendedMode", caption) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Settings", "ExtendedModeStyle", text) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".direction", direction, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".form", formKind, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".caption", caption, diagnostic) Then Exit Function
    If direction <> "Arrival" And direction <> "Departure" And direction <> "StatusChange" Then
        Err.Raise 5, , "Invalid mode direction: " & modeId
    End If
    If Not private_UpdateReadOnly(modeId, diagnostic) Then Exit Function
    If VBA.Len(m_modeId) > 0 And Not m_reporterActive And Not m_hospitalActive And Not m_documentActive Then
        If Not private_TryReadFormValue("ModeForm", currentForm) Then Exit Function
        If VBA.CStr(currentForm) = formKind And m_modeDirection = direction Then
            ' These modes share the same projection and derived event fields.
            If Not m_pageBase.BindingContext.SetValue("Settings", m_modeId & "Style", "secondaryButton") Then Exit Function
            If Not m_pageBase.BindingContext.SetValue("Settings", modeId & "Style", "activeModeButton") Then Exit Function
            If Not m_pageBase.BindingContext.SetValue("Form", "ModeId", modeId) Then Exit Function
            If Not m_pageBase.BindingContext.SetValue("Form", "WordPreviewText", VBA.vbNullString) Then Exit Function
            m_modeId = modeId
            If Not m_lookupResult Is Nothing Then
                Set projectedResult = m_lookupResult
                If Not private_TryEnrichResult(projectedResult, diagnostic) Then Exit Function
                If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", projectedResult) Then Exit Function
                Set m_lookupResult = projectedResult
            End If
            If Not m_extensions.Build(formKind, diagnostic) Then Exit Function
            private_LoadMode = True
            Exit Function
        End If
    End If
    If Not private_ReadModeConfig("forms." & formKind & ".fields.count", countText, diagnostic) Then Exit Function
    If Not VBA.IsNumeric(countText) Then Err.Raise 5, , "Invalid mode field count: " & formKind
    If VBA.CDbl(countText) < 1 Or VBA.CDbl(countText) > 16384 Or VBA.CDbl(countText) <> VBA.Fix(VBA.CDbl(countText)) Then
        Err.Raise 5, , "Mode field count must be an integer from 1 to 16384."
    End If
    If Not private_ReadModeConfig("text." & direction & "Type", eventType, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("text.Mode" & direction & "DateLabel", eventDateLabel, diagnostic) Then Exit Function
    For index = 1 To VBA.CLng(countText)
        prefix = "forms." & formKind & ".fields." & VBA.CStr(index)
        If Not private_ReadModeConfig(prefix & ".name", fieldName, diagnostic) Then Exit Function
        If Not private_ReadModeConfig(prefix & ".captionKey", captionKey, diagnostic) Then Exit Function
        If VBA.Left$(captionKey, VBA.Len("PersonalEventBuilder::")) <> "PersonalEventBuilder::" Then
            diagnostic = "Mode caption key must use PersonalEventBuilder:: prefix: " & captionKey
            Exit Function
        End If
        If Not private_ReadModeConfig(VBA.Mid$(captionKey, VBA.Len("PersonalEventBuilder::") + 1), fieldCaption, diagnostic) Then
            diagnostic = "Required caption key not found: " & captionKey
            Exit Function
        End If
        If fieldName = "EventDate" Then fieldCaption = eventDateLabel
        recordName = VBA.vbNullString
        sourceAlias = VBA.vbNullString
        If m_modeConfig.Exists(prefix & ".record") Then recordName = m_modeConfig(prefix & ".record")
        If m_modeConfig.Exists(prefix & ".sourceAlias") Then sourceAlias = m_modeConfig(prefix & ".sourceAlias")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column("formField") = fieldName
        column("apply") = (VBA.Len(recordName) > 0)
        If VBA.Len(recordName) = 0 Then recordName = "Manual." & fieldName
        column("name") = recordName
        column("caption") = fieldCaption
        column("sourceAlias") = sourceAlias
        columns.Add column
    Next index
    If Not private_PlaceFields(columns, diagnostic) Then Exit Function
    For Each field In VBA.Array("EventType", _
            "Destination", _
            "PersonRank", _
            "PersonName", _
            "PersonId", _
            "PositionCode", _
            "PersonPosition", _
            "ReportRank", _
            "ReportName", _
            "ReportPositionCode", _
            "IncomingNumber", _
            "IncomingDate", _
            "EventDate", _
            "DurationDays", _
            "TicketNumber", _
            "TicketDate", _
            "Hospital", _
            "HospitalShortName", _
            "Document", _
            "DocumentDate", _
            "DocumentNumber")
        If Not private_ReadModeConfig("text.Mode" & VBA.CStr(field) & "Label", text, diagnostic) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", "Mode" & VBA.CStr(field) & "Label", text) Then Exit Function
    Next field
    If Not m_pageBase.BindingContext.SetValue("Text", "ModeEventDateLabel", eventDateLabel) Then Exit Function
    For Each mode In modes
        If Not private_ReadModeConfig("modes." & VBA.CStr(mode) & ".caption", text, diagnostic) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", VBA.CStr(mode), text) Then Exit Function
        text = "secondaryButton"
        If VBA.CStr(mode) = modeId Then text = "activeModeButton"
        If Not m_pageBase.BindingContext.SetValue("Settings", VBA.CStr(mode) & "Style", text) Then Exit Function
    Next mode
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeId", modeId) Then Exit Function
    ' Secondary candidates must never be projected or reapplied as primary personnel.
    If secondaryCandidates Then
        If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    End If
    m_reporterActive = False
    m_hospitalActive = False
    m_documentActive = False
    If Not m_pageBase.BindingContext.SetValue("Form", "WordPreviewText", VBA.vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeDirection", direction) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeForm", formKind) Then Exit Function
    If Not private_PublishViewport(diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "EventType", eventType) Then Exit Function
    Set records = New Collection
    If Not secondaryCandidates Then
        If Not m_lookupResult Is Nothing Then
            Set records = m_lookupResult.Records
            hasMore = m_lookupResult.HasMore
        End If
    End If
    m_modeDirection = direction
    m_eventType = eventType
    If records.Count > 0 Then
        If Not m_candidateEnricher.TryEnrich( _
            m_lookupService, _
            records, _
            m_acceptedOrderNumber, _
            m_acceptedOrderDate, _
            modeId, _
            enrichedRecords, _
            diagnostic _
        ) Then Exit Function
        Set records = enrichedRecords
    End If
    If Not private_PrepareModeRecords(records, diagnostic) Then Exit Function
    Set projectedResult = New obj_LookupResult
    If Not private_SetCandidateSources(columns, modeId, diagnostic) Then Exit Function
    If Not projectedResult.Initialize(records, columns, hasMore, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", projectedResult) Then Exit Function
    Set m_lookupResult = projectedResult
    Set m_modeFields = columns
    m_modeId = modeId
    If Not m_extensions.Build(formKind, diagnostic) Then Exit Function
    private_LoadMode = True
    Exit Function
EH:
    diagnostic = "Mode configuration: " & VBA.Err.Description
End Function

Private Function private_UpdateReadOnly( _
    ByVal modeId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim fields As String
    Dim dependents As String
    Dim field As Variant
    Dim autoFill As Variant
    Dim updates As Object

    If Not private_ReadModeConfig("forms.readOnly.fields", fields, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("forms.readOnly.modes." & modeId, dependents, diagnostic) Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    Set updates = CreateObject("Scripting.Dictionary")
    For Each field In Split(dependents, "|")
        If field <> "-" Then
            If InStr(1, "|" & fields & "|", "|" & field & "|", vbBinaryCompare) = 0 Then
                diagnostic = "Unknown dependent field for mode " & modeId & ": " & field
                Exit Function
            End If
        End If
    Next field
    For Each field In Split(fields, "|")
        updates("Form." & field & "ReadOnly") = CBool(autoFill) And _
            (InStr(1, "|" & dependents & "|", "|" & field & "|", vbBinaryCompare) > 0)
    Next field
    private_UpdateReadOnly = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function

Private Function private_SwitchMode(ByVal modeId As String) As Boolean
    Dim diagnostic As String
    Dim caption As String
    Dim status As String
    Dim startedAt As Double
    Dim selectedValue As Variant
    Dim selectedObject As Object
    Dim isObject As Boolean
    Dim selectedRecord As obj_LookupRecord
    Dim record As obj_LookupRecord
    Dim autoFill As Variant
    Dim returnModes As String
    Dim oldReturnMode As Boolean
    Dim newReturnMode As Boolean
    Dim updates As Object
    Dim previousForm As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If modeId = m_modeId Then
        private_SwitchMode = True
        Exit Function
    End If
    previousForm = ex_PEB_EventRoutes.fn_Value(m_modeId, "form")
    startedAt = VBA.Timer
    If Not private_ReadModeConfig("lookup.ReturnTickets.modes", returnModes, diagnostic) Then GoTo Failed
    returnModes = "|" & returnModes & "|"
    oldReturnMode = VBA.InStr(1, returnModes, "|" & m_modeId & "|", VBA.vbTextCompare) > 0
    newReturnMode = VBA.InStr(1, returnModes, "|" & modeId & "|", VBA.vbTextCompare) > 0
    If oldReturnMode <> newReturnMode Then
        If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
        If VBA.CBool(autoFill) Then
            Set updates = VBA.CreateObject("Scripting.Dictionary")
            updates("Form.TicketNumber") = VBA.vbNullString
            updates("Form.TicketDate") = VBA.vbNullString
            If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
        End If
    End If
    If Not private_LoadMode(modeId, diagnostic) Then GoTo Failed
    If m_pageBase.BindingContext.TryGetValue("Data", "SelectedPersonnel", selectedValue, selectedObject, isObject) Then
        If isObject Then
            If TypeOf selectedObject Is obj_LookupRecord Then
                Set selectedRecord = selectedObject
                For Each record In m_lookupResult.Records
                    If record.Key = selectedRecord.Key Then
                        If Not m_pageBase.BindingContext.SetObject("Data", "SelectedPersonnel", record) Then Exit Function
                        If Not Me.SelectPersonnelHandler(record) Then Exit Function
                        Exit For
                    End If
                Next record
            End If
        End If
    End If
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.EventDate") = VBA.vbNullString
    updates("Form.DurationDays") = VBA.vbNullString
    If Not private_ReadModeConfig("forms.sourceStyles.default", status, diagnostic) Then GoTo Failed
    updates("Form.EventDateStyle") = status
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    ex_Core.fn_Diagnostic_WritePerf "Mode.Prepare", startedAt
    If Not private_ReadModeConfig("modes." & modeId & ".caption", caption, diagnostic) Then GoTo Failed
    If Not private_ReadModeConfig("text.ModeSelected", status, diagnostic) Then GoTo Failed
    If Not private_SetActionStatus(VBA.Replace$(status, "{mode}", caption), False) Then Exit Function
    startedAt = VBA.Timer
    If Not m_pageBase.RefreshContent("ExtensionMenus", diagnostic, True) Then GoTo Failed
    If Not m_pageBase.RefreshContent("ExtensionRows", diagnostic, True) Then GoTo Failed
    private_SwitchMode = m_pageBase.FlushUpdates()
    ex_Core.fn_Diagnostic_WritePerf "Mode.Flush", startedAt
    Exit Function
Failed:
    private_SetActionStatus diagnostic, True
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Event mode"
End Function

Private Function private_ClearPersonnelGroup( _
    ByVal includeKey As Boolean, _
    ByVal nextForm As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim currentForm As Variant
    Dim formKind As Variant
    Dim key As String
    Dim dependents As String
    Dim fieldName As Variant
    Dim updates As Object
    Dim records As New Collection
    Dim emptyResult As obj_LookupResult

    On Error GoTo EH
    If Not private_TryReadFormValue("ModeForm", currentForm) Then Exit Function
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    ' Очищаем зависимости при изменении ключа поиска; смена режима сохраняет значения формы.
    For Each formKind In VBA.Array(VBA.CStr(currentForm), nextForm)
        If VBA.Len(VBA.CStr(formKind)) > 0 Then
            If Not private_ReadModeConfig("forms." & formKind & ".lookup.key", key, diagnostic) Then Exit Function
            If Not private_ReadModeConfig("forms." & formKind & ".lookup.dependents", dependents, diagnostic) Then Exit Function
            If includeKey Then updates("Form." & key) = VBA.vbNullString
            For Each fieldName In VBA.Split(dependents, "|")
                fieldName = VBA.Trim$(VBA.CStr(fieldName))
                If VBA.Len(fieldName) = 0 Or fieldName = key Then
                    Err.Raise 5, , "Invalid lookup dependency in " & VBA.CStr(formKind)
                End If
                updates("Form." & fieldName) = VBA.vbNullString
            Next fieldName
        End If
    Next formKind
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    Set emptyResult = New obj_LookupResult
    If Not emptyResult.Initialize(records, m_modeFields, False, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", emptyResult) Then Exit Function
    Set m_lookupResult = emptyResult
    private_ClearPersonnelGroup = True
    Exit Function
EH:
    diagnostic = "Lookup dependencies: " & VBA.Err.Description
End Function

Private Function private_PrepareModeRecords( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim fields As Object

    For Each record In records
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields("Event.Type") = m_eventType
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then Exit Function
        If Not record.TryMerge(supplement, True) Then
            diagnostic = "Cannot apply event mode fields."
            Exit Function
        End If
    Next record
    private_PrepareModeRecords = True
End Function

Private Function private_ApplyModeRecord( _
    ByVal record As obj_LookupRecord, _
    ByRef diagnostic As String _
) As Boolean
    Dim column As Object
    Dim value As Variant
    Dim updates As Object
    Dim fields As Collection
    Dim sourceStyle As String
    Dim sourceAlias As String
    Dim lookupFields As String
    Dim defaultStyle As String
    Dim dynamicStyles As String

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    Set fields = m_modeFields
    If Not private_ReadModeConfig("forms.dynamicStylesEnabled", dynamicStyles, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("forms.lookupFields", lookupFields, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("forms.sourceStyles.default", defaultStyle, diagnostic) Then Exit Function
    If m_reporterActive Then
        Set fields = m_reporterFields
    End If
    If m_hospitalActive Then
        Set fields = m_hospitalFields
    End If
    If m_documentActive Then
        Set fields = m_documentFields
    End If
    For Each column In fields
        If column("apply") Then
            value = VBA.vbNullString
            If Not record.TryGetValue(column("name"), value) Then
                diagnostic = "Candidate field not found: " & column("name")
                Exit Function
            End If
            updates("Form." & column("formField")) = value
            If dynamicStyles = "false" Then GoTo NextAppliedField
            sourceAlias = VBA.CStr(column("sourceAlias"))
            If VBA.Len(VBA.Trim$(VBA.CStr(value))) = 0 Or VBA.Len(sourceAlias) = 0 Then sourceAlias = "default"
            If Not private_ReadModeConfig("forms.sourceStyles." & sourceAlias, sourceStyle, diagnostic) Then Exit Function
            If VBA.InStr(1, "|" & lookupFields & "|", "|" & column("formField") & "|", VBA.vbBinaryCompare) > 0 Then
                If sourceStyle = defaultStyle Then
                    If Not private_ReadModeConfig("forms.lookupDefaultStyle", sourceStyle, diagnostic) Then Exit Function
                Else
                    sourceStyle = sourceStyle & "Lookup"
                End If
            End If
            updates("Form." & column("formField") & "Style") = sourceStyle
        End If
NextAppliedField:
    Next column
    private_ApplyModeRecord = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function

Private Function private_TryEnrichResult( _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String _
) As Boolean
    Dim records As Collection
    Dim enrichedResult As obj_LookupResult

    If Not m_candidateEnricher.TryEnrich(m_lookupService, result.Records, _
            m_acceptedOrderNumber, m_acceptedOrderDate, m_modeId, records, diagnostic) Then Exit Function
    Set enrichedResult = New obj_LookupResult
    If Not private_PrepareModeRecords(records, diagnostic) Then Exit Function
    If Not private_SetCandidateSources(m_modeFields, m_modeId, diagnostic) Then Exit Function
    If Not enrichedResult.Initialize(records, m_modeFields, result.HasMore, diagnostic) Then Exit Function
    Set result = enrichedResult
    If Not Me.TryRefreshSourceLegend(diagnostic) Then Exit Function
    private_TryEnrichResult = True
End Function

Private Function private_PlaceFields( _
    ByVal columns As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim fixedColumns As Object
    Dim key As Variant
    Dim prefix As String
    Dim widthText As String
    Dim column As Object
    Dim groups As Object
    Dim fields As Variant
    Dim field As Variant
    Dim groupRanges As Object
    Dim rangeText As String
    Dim displayFields As Collection
    Dim activeFields As Object
    Dim caption As String
    Dim disabledStyle As String
    Dim sectionFields As String

    sectionFields = "|" & ex_PEB_Export.fn_Config("forms.sections.Ticket.fields") & "|"
    If Not private_ReadModeConfig("forms.columns.count", widthText, diagnostic) Then Exit Function
    If Not IsNumeric(widthText) Then
        diagnostic = "Invalid configuration: PersonalEventBuilder::forms.columns.count"
        Exit Function
    End If
    If CDbl(widthText) < 1 Or CDbl(widthText) > 16384 Or CDbl(widthText) <> Fix(CDbl(widthText)) Then
        diagnostic = "Invalid configuration: PersonalEventBuilder::forms.columns.count"
        Exit Function
    End If
    prefix = "forms.columns.fixed."
    Set fixedColumns = CreateObject("Scripting.Dictionary")
    For Each key In m_modeConfig.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then
            fixedColumns.Add Mid$(CStr(key), Len(prefix) + 1), m_modeConfig(key)
        End If
    Next key
    If fixedColumns.Count = 0 Then
        diagnostic = "Required configuration group not found: PersonalEventBuilder::forms.columns.fixed"
        Exit Function
    End If
    Set groups = CreateObject("Scripting.Dictionary")
    Set groupRanges = CreateObject("Scripting.Dictionary")
    prefix = "forms.columns.groups."
    For Each key In m_modeConfig.Keys
        If Left$(CStr(key), Len(prefix)) = prefix Then
            If Not private_ReadModeConfig("forms.columns.ranges." & Mid$(CStr(key), Len(prefix) + 1), rangeText, diagnostic) Then Exit Function
            fields = Split(rangeText, "|")
            If UBound(fields) <> 1 Then
                diagnostic = "Invalid form group range: " & CStr(key)
                Exit Function
            End If
            groupRanges.Add CStr(key), fields
            fields = Split(CStr(m_modeConfig(key)), "|")
            For Each field In fields
                If groups.Exists(CStr(field)) Then
                    diagnostic = "Duplicate form group field: " & CStr(field)
                    Exit Function
                End If
                groups.Add CStr(field), CStr(key)
            Next field
        End If
    Next key
    Set displayFields = New Collection
    If Not private_ReadModeConfig("forms.columns.disabledStyle", disabledStyle, diagnostic) Then Exit Function
    Set activeFields = CreateObject("Scripting.Dictionary")
    For Each column In columns
        If fixedColumns.Exists(CStr(column("formField"))) Then
            If Not private_ReadModeConfig("text.Mode" & CStr(column("formField")) & "Label", caption, diagnostic) Then Exit Function
            column("caption") = caption
        End If
        If InStr(1, sectionFields, "|" & CStr(column("formField")) & "|", vbBinaryCompare) > 0 Then
            column("column") = 0
        Else
            displayFields.Add column
        End If
        activeFields.Add CStr(column("formField")), True
    Next column
    For Each field In fixedColumns.Keys
        If Not activeFields.Exists(field) Then
            If Not private_ReadModeConfig("text.Mode" & CStr(field) & "Label", caption, diagnostic) Then Exit Function
            Set column = CreateObject("Scripting.Dictionary")
            column("formField") = CStr(field)
            column("caption") = caption
            column("disabled") = True
            column("disabledStyle") = disabledStyle
            displayFields.Add column
            If Not m_pageBase.BindingContext.SetValue("Form", "Disabled" & CStr(field), vbNullString) Then Exit Function
        End If
    Next field
    If Not ex_UiFormLayouts.fn_TryPlaceColumns(displayFields, fixedColumns, CLng(widthText), diagnostic, groups, groupRanges) Then Exit Function
    Set m_displayFields = displayFields
    For Each column In columns
        If CLng(column("column")) > 0 Then
            fixedColumns(CStr(column("formField"))) = CLng(column("column"))
        End If
    Next column
    private_PlaceFields = m_pageBase.BindingContext.SetObject("Data", "FormColumnMap", fixedColumns)
End Function

Private Function private_ScrollForm(ByVal direction As Long) As Boolean
    Dim diagnostic As String
    Dim stepText As String
    Dim previousStart As Long

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo FailedScroll
    ex_Core.fn_Diagnostic_BeginBatch
    If Not private_ReadModeConfig("forms.columns.scrollStep", stepText, diagnostic) Then GoTo FailedScroll
    If Not IsNumeric(stepText) Then
        diagnostic = "Invalid forms.columns.scrollStep."
        GoTo FailedScroll
    End If
    If CDbl(stepText) < 1 Or CDbl(stepText) > 16384 Or CDbl(stepText) <> Fix(CDbl(stepText)) Then
        diagnostic = "Invalid forms.columns.scrollStep."
        GoTo FailedScroll
    End If
    previousStart = m_viewportStart
    m_viewportStart = m_viewportStart + direction * CLng(stepText)
    If Not private_PublishViewport(diagnostic, previousStart) Then GoTo FailedScroll
    If previousStart = m_viewportStart Then
        private_ScrollForm = True
        GoTo CleanScroll
    End If
    If Not m_extensions.Build(ex_PEB_EventRoutes.fn_Value(m_modeId, "form"), diagnostic) Then GoTo FailedScroll
    If Not m_pageBase.RefreshContent("ExtensionMenus", diagnostic, True) Then GoTo FailedScroll
    If Not m_pageBase.RefreshContent("ExtensionRows", diagnostic, True) Then GoTo FailedScroll
    private_ScrollForm = m_pageBase.FlushUpdates()
CleanScroll:
    ex_Core.fn_Diagnostic_EndBatch
    Exit Function
FailedScroll:
    If Len(diagnostic) = 0 Then diagnostic = Err.Description
    ex_Core.fn_Diagnostic_EndBatch
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Form viewport"
End Function

Private Function private_PublishViewport( _
    ByRef diagnostic As String, _
    Optional ByVal unchangedStart As Variant _
) As Boolean
    Dim widthText As String
    Dim width As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    Dim position As Long
    Dim column As Object
    Dim visible As Collection
    Dim projected As Object
    Dim map As Object
    Dim key As Variant

    On Error GoTo FailedViewport
    If m_displayFields Is Nothing Then Err.Raise 5, , "Form fields are not initialized."
    If Not private_ReadModeConfig("forms.columns.count", widthText, diagnostic) Then Exit Function
    width = CLng(widthText)
    firstColumn = 1
    lastColumn = width
    For Each column In m_displayFields
        position = CLng(column("column"))
        If position < firstColumn Then firstColumn = position
        If position > lastColumn Then lastColumn = position
    Next column
    If lastColumn - firstColumn + 1 > width Then
        If Not m_pageBase.BindingContext.SetValue("Settings", "FormScrollVisibility", "Visible") Then Exit Function
    Else
        If Not m_pageBase.BindingContext.SetValue("Settings", "FormScrollVisibility", "Collapsed") Then Exit Function
    End If
    If Not m_hasViewport Then
        m_viewportStart = firstColumn
        m_hasViewport = True
    End If
    If m_viewportStart < firstColumn Then m_viewportStart = firstColumn
    If m_viewportStart > lastColumn - width + 1 Then m_viewportStart = lastColumn - width + 1
    If Not IsMissing(unchangedStart) Then
        If m_viewportStart = CLng(unchangedStart) Then
            private_PublishViewport = True
            Exit Function
        End If
    End If
    Set visible = New Collection
    Set map = CreateObject("Scripting.Dictionary")
    For Each column In m_displayFields
        position = CLng(column("column")) - m_viewportStart + 1
        map(CStr(column("formField"))) = -1
        If position >= 1 And position <= width Then
            Set projected = CreateObject("Scripting.Dictionary")
            For Each key In column.Keys
                projected.Add CStr(key), column(key)
            Next key
            projected("column") = position
            projected("row") = 1
            projected("viewport") = True
            projected("candidateColumn") = position
            If projected.Exists("disabled") Then projected("candidateColumn") = 0
            visible.Add projected
            map(CStr(column("formField"))) = position
        End If
    Next column
    If Not m_pageBase.BindingContext.SetValue("Settings", "FormViewportStart", m_viewportStart) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "FormColumnMap", map) Then Exit Function
    private_PublishViewport = m_pageBase.BindingContext.SetObject("Data", "EventFields", visible)
    Exit Function
FailedViewport:
    diagnostic = "Form viewport: " & Err.Description
End Function

Private Function private_SetCandidateSources( _
    ByVal columns As Collection, _
    ByVal modeId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim modes As String
    Dim fields As String
    Dim sourceAlias As String
    Dim orderDateField As String
    Dim orderDateAlias As String
    Dim column As Object
    Dim hospitalDateField As String
    Dim hospitalFields As String
    Dim hospitalSourceAlias As String
    Dim hospitalModes As String
    Dim isHospitalMode As Boolean

    If Not private_ReadModeConfig("lookup.ReturnTickets.modes", modes, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("lookup.ReturnTickets.styleFields", fields, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("legend.orderDate.field", orderDateField, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("legend.orderDate.sourceAlias", orderDateAlias, diagnostic) Then Exit Function
    If VBA.InStr(1, "|" & modes & "|", "|" & modeId & "|", VBA.vbTextCompare) > 0 Then
        If Not private_ReadModeConfig("lookup.ReturnTickets.sourceAlias", sourceAlias, diagnostic) Then Exit Function
    Else
        If Not private_ReadModeConfig("lookup.ReturnTickets.departureSourceAlias", sourceAlias, diagnostic) Then Exit Function
    End If
    hospitalDateField = ex_PEB_Export.fn_Config("lookup.ReturnHospital.dateField")
    hospitalModes = ex_PEB_Export.fn_Config("lookup.ReturnHospital.modes")
    isHospitalMode = InStr(1, "|" & hospitalModes & "|", "|" & modeId & "|", vbTextCompare) > 0
    If isHospitalMode Then
        hospitalFields = ex_PEB_Export.fn_Config("lookup.ReturnHospital.fields")
        hospitalSourceAlias = ex_PEB_Export.fn_Config("lookup.ReturnHospital.sourceAlias")
    End If
    For Each column In columns
        If Left$(column("name"), 9) = "Movement." Or _
            column("formField") = hospitalDateField Then
            column("name") = "Manual." & column("formField")
            column("apply") = False
            column("sourceAlias") = vbNullString
        End If
        If isHospitalMode Then
            If InStr(1, "|" & hospitalFields & "|", "|" & column("formField") & "|", vbBinaryCompare) > 0 Then
                column("name") = "Movement." & column("formField")
                column("apply") = True
                column("sourceAlias") = hospitalSourceAlias
            End If
            If column("formField") = hospitalDateField Then
                column("name") = orderDateField
                column("apply") = True
                column("sourceAlias") = orderDateAlias
            End If
        End If
        If VBA.InStr(1, "|" & fields & "|", "|" & column("name") & "|", VBA.vbTextCompare) > 0 Then
            column("sourceAlias") = sourceAlias
        End If
        If column("name") = orderDateField Then column("sourceAlias") = orderDateAlias
    Next column
    private_SetCandidateSources = True
End Function

Private Function private_SetActionStatus( _
    ByVal text As String, _
    ByVal isError As Boolean _
) As Boolean
    Dim updates As Object
    Dim diagnostic As String

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.ActionStatus") = text
    updates("Form.ActionStatusStyle") = "actionSuccess"
    If isError Then updates("Form.ActionStatusStyle") = "actionError"
    private_SetActionStatus = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function

Private Function private_OrderText( _
    ByVal record As obj_LookupRecord, _
    ByVal messageKey As String _
) As String
    Dim number As Variant
    Dim orderDate As Variant

    If Not record.TryGetValue("OrderNumber", number) Then Exit Function
    If Not record.TryGetValue("OrderDate", orderDate) Then Exit Function
    private_OrderText = VBA.Replace$(m_messages(messageKey), "{number}", VBA.CStr(number))
    private_OrderText = VBA.Replace$(private_OrderText, "{date}", VBA.Format$(VBA.CDate(orderDate), "dd.mm.yyyy"))
End Function

Private Function private_TryReadFormValue( _
    ByVal fieldName As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim outObject As Object
    Dim isObject As Boolean

    If m_pageBase Is Nothing Then Exit Function
    private_TryReadFormValue = m_pageBase.BindingContext.TryGetValue( _
        "Form", fieldName, outValue, outObject, isObject)
End Function