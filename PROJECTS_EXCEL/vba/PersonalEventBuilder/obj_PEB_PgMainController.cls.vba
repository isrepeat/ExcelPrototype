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
Private m_modeId As String
Private m_modeDirection As String
Private m_eventType As String
Private m_modeConfig As Object
Private m_reporterFields As Collection
Private m_reporterActive As Boolean
Private m_hospitalActive As Boolean
Private m_hospitalFields As Collection
Private m_hospitalProfile As obj_LookupProfile
Private m_hospitalService As obj_LookupService
Private m_exportService As obj_PEB_ExportService
Private m_environmentSwitchPending As Boolean
Private m_testBusy As Boolean
Private m_testScenario As String
Private m_testStepOrder As Long

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

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = pageBase
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
    If Not private_LoadMode(defaultMode, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
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
    Set m_modeConfig = Nothing
    Set m_reporterFields = Nothing
    If Not m_hospitalService Is Nothing Then m_hospitalService.Dispose
    If Not m_hospitalProfile Is Nothing Then m_hospitalProfile.Dispose
    Set m_hospitalService = Nothing
    Set m_hospitalProfile = Nothing
    Set m_hospitalFields = Nothing
    m_hospitalActive = False
    m_reporterActive = False
    Set m_pageBase = Nothing
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

Public Function SearchPersonnelHandler() As Boolean
    Dim text As Variant
    Dim formKind As Variant
    Dim keyField As String
    Dim diagnostic As String
    Dim result As obj_LookupResult
    Dim status As String
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_reporterActive Or m_hospitalActive Then
        If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    m_reporterActive = False
    m_hospitalActive = False
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
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(text))) < m_lookupProfile.MinChars Then
        status = VBA.Replace$(m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(m_lookupProfile.MinChars))
    ElseIf result.RowCount = 0 Then
        status = m_messages("PersonnelEmpty")
    Else
        status = VBA.Replace$(m_messages("PersonnelCount"), "{count}", VBA.CStr(result.RowCount))
        If result.HasMore Then status = VBA.Replace$(m_messages("PersonnelMore"), "{count}", VBA.CStr(result.RowCount))
    End If
    SearchPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", status)
    If SearchPersonnelHandler Then SearchPersonnelHandler = private_SetActionStatus(status, False)
    Exit Function
Failure:
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
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
    If m_reporterActive Or m_hospitalActive Then
        If Not private_SearchSecondary(False, m_hospitalActive) Then
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
    If Not Me.SearchPersonnelHandler() Then
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
    If m_reporterActive Or m_hospitalActive Then
        RefreshPersonnelHandler = private_SearchSecondary(False, m_hospitalActive)
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

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
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

Public Function FromHospitalHandler() As Boolean
    FromHospitalHandler = private_SwitchMode("FromHospital")
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

Public Function ExportWordHandler() As Boolean
    Dim enabled As String
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    enabled = ex_PEB_Export.fn_Config("export.Word.enabled")
    If enabled = "false" Then
        ExportWordHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.ExportWordUnavailable"), True)
        Exit Function
    End If
    If enabled <> "true" Then
        Err.Raise 5, , "Word enabled must be true or false."
    End If
    ExportWordHandler = private_Export("Word")
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    private_SetActionStatus diagnostic, True
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "PersonalEventBuilder"
End Function

Public Function ExportMovementHandler() As Boolean
    Dim enabled As String
    Dim diagnostic As String

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
        Err.Raise 5, , "Movement enabled must be true or false."
    End If
    ExportMovementHandler = private_Export("Movement")
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    private_SetActionStatus diagnostic, True
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "PersonalEventBuilder"
End Function

Public Function FromAnnualMainVacationHandler() As Boolean
    FromAnnualMainVacationHandler = private_SwitchMode("FromAnnualMainVacation")
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
    Dim oldFields As Collection
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
    oldMode = m_modeId
    oldDirection = m_modeDirection
    oldEventType = m_eventType
    oldOrder = m_acceptedOrderNumber
    oldOrderDate = m_acceptedOrderDate
    oldReporter = m_reporterActive
    oldHospital = m_hospitalActive
    formChanged = True

    ' Empty candidates without running personnel, reporter or hospital searches.
    Set emptyResult = New obj_LookupResult
    If Not emptyResult.Initialize(records, m_modeFields, False, diagnostic) Then GoTo Failed
    Set m_lookupResult = emptyResult
    m_modeId = VBA.vbNullString
    m_reporterActive = False
    m_hospitalActive = False
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
        Set m_lookupResult = oldResult
        m_reporterActive = oldReporter
        m_hospitalActive = oldHospital
        m_acceptedOrderNumber = oldOrder
        m_acceptedOrderDate = oldOrderDate
        If Not m_pageBase.BindingContext.TryApplyValues(snapshot, rollbackDiagnostic) Then
            diagnostic = diagnostic & " Restore form: " & rollbackDiagnostic
        End If
        If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", oldResult) Then
            diagnostic = diagnostic & " Cannot restore candidates."
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

Private Function private_SearchSecondary( _
    ByVal clearDependents As Boolean, _
    Optional ByVal hospital As Boolean = False _
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
    If hospital Then group = "hospital"
    If Not private_SecondaryFields(group, fields, diagnostic) Then
        GoTo Failed
    End If
    If hospital Then
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
    m_reporterActive = Not hospital
    m_hospitalActive = hospital

    ' Search the selected directory without enriching secondary records with absence data.
    If Not service.TrySearch(VBA.CStr(text), result, diagnostic) Then
        GoTo Failed
    End If
    Set projected = New obj_LookupResult
    If Not projected.Initialize(result.Records, fields, result.HasMore, diagnostic) Then
        GoTo Failed
    End If
    Set m_lookupResult = projected
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", projected) Then
        Err.Raise 5, , "Cannot publish secondary candidates."
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
            m_reporterActive = Not hospital
            m_hospitalActive = hospital
            m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", projected
        End If
    End If
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "PersonalEventBuilder"
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

    If group = "hospital" Then
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
    If group = "hospital" Then
        Set m_hospitalFields = fields
    Else
        Set m_reporterFields = fields
    End If
    Set result = fields
    private_SecondaryFields = True
End Function

Private Function private_Export(ByVal action As String) As Boolean
    Dim pebExportService As obj_PEB_ExportService
    Dim pebExptrWord As obj_PEB_ExptrWord
    Dim pebExptrMovement As obj_PEB_ExptrMovement
    Dim values As Object
    Dim diagnostic As String
    Dim exporter As String
    Dim success As Boolean
    Dim status As String
    Dim failureTemplate As String
    Dim messageTitle As String

    If ex_PEB_Testing.fn_IsRunning() Then
        private_Export = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    failureTemplate = ex_PEB_Export.fn_Config("text.ExportFailed")
    messageTitle = ex_PEB_Export.fn_Config("text.ExportTitle")
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
    If Not pebExportService.TryPrepare( _
        m_pageBase.BindingContext, _
        exporter, _
        values, _
        diagnostic, _
        action <> "PreviewWord" _
    ) Then GoTo Failed
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
    ElseIf exporter = "Word" Then
        Set pebExptrWord = New obj_PEB_ExptrWord
        If Not pebExptrWord.Initialize() Then
            Err.Raise 5, , "Cannot initialize Word exporter."
        End If
        success = pebExptrWord.TryExport(values, diagnostic)
        If values("WordPolicy") = "Omit" Then
            status = ex_PEB_Export.fn_Config("text.ExportWordOmitted")
        Else
            status = VBA.Replace$(status, "{path}", VBA.CStr(values("WordOutput")))
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
    If Not pebExptrWord Is Nothing Then
        pebExptrWord.Dispose
    End If
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
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, messageTitle
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
    value = VBA.vbNullString
    If Not m_modeConfig Is Nothing Then
        If m_modeConfig.Exists(key) Then
            value = m_modeConfig(key)
            private_ReadModeConfig = (VBA.Len(value) > 0)
        End If
    End If
    If Not private_ReadModeConfig Then diagnostic = "Required configuration key not found: PersonalEventBuilder::" & key
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
    Dim hasMore As Boolean
    Dim modes As Variant
    Dim currentForm As Variant

    On Error GoTo EH
    modes = VBA.Array("FromHospital", _
            "FromAnnualMainVacation", _
            "FromVacationForCuring", _
            "FromVacationForFamily", _
            "ToHospital", _
            "ToAnnualMainVacation", _
            "ToVacationForCuring", _
            "ToVacationForFamily")
    If Not private_ReadModeConfig("modes." & modeId & ".direction", direction, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".form", formKind, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".caption", caption, diagnostic) Then Exit Function
    If direction <> "Arrival" And direction <> "Departure" And direction <> "StatusChange" Then
        Err.Raise 5, , "Invalid mode direction: " & modeId
    End If
    If formKind <> "Leave" And formKind <> "Treatment" Then Err.Raise 5, , "Invalid mode form: " & modeId
    If VBA.Len(m_modeId) > 0 And Not m_reporterActive And Not m_hospitalActive Then
        If Not private_TryReadFormValue("ModeForm", currentForm) Then Exit Function
        If VBA.CStr(currentForm) = formKind And m_modeDirection = direction Then
            ' These modes share the same projection and derived event fields.
            If Not m_pageBase.BindingContext.SetValue("Settings", m_modeId & "Style", "secondaryButton") Then Exit Function
            If Not m_pageBase.BindingContext.SetValue("Settings", modeId & "Style", "activeModeButton") Then Exit Function
            If Not m_pageBase.BindingContext.SetValue("Form", "ModeId", modeId) Then Exit Function
            If Not m_pageBase.BindingContext.SetValue("Form", "WordPreviewText", VBA.vbNullString) Then Exit Function
            m_modeId = modeId
            private_LoadMode = True
            Exit Function
        End If
    End If
    If Not private_ReadModeConfig("forms." & formKind & ".fields.count", countText, diagnostic) Then Exit Function
    If Not VBA.IsNumeric(countText) Then Err.Raise 5, , "Invalid mode field count: " & formKind
    If VBA.CDbl(countText) < 1 Or VBA.CDbl(countText) > 16 Or VBA.CDbl(countText) <> VBA.Fix(VBA.CDbl(countText)) Then
        Err.Raise 5, , "Mode field count must be an integer from 1 to 16."
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
    m_reporterActive = False
    m_hospitalActive = False
    If Not m_pageBase.BindingContext.SetValue("Form", "WordPreviewText", VBA.vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeDirection", direction) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeForm", formKind) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "EventType", eventType) Then Exit Function
    Set records = New Collection
    If Not m_lookupResult Is Nothing Then
        Set records = m_lookupResult.Records
        hasMore = m_lookupResult.HasMore
    End If
    m_modeDirection = direction
    m_eventType = eventType
    If Not private_PrepareModeRecords(records, diagnostic) Then Exit Function
    Set projectedResult = New obj_LookupResult
    If Not projectedResult.Initialize(records, columns, hasMore, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", projectedResult) Then Exit Function
    Set m_lookupResult = projectedResult
    Set m_modeFields = columns
    m_modeId = modeId
    private_LoadMode = True
    Exit Function
EH:
    diagnostic = "Mode configuration: " & VBA.Err.Description
End Function

Private Function private_SwitchMode(ByVal modeId As String) As Boolean
    Dim diagnostic As String
    Dim caption As String
    Dim status As String
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If modeId = m_modeId Then
        private_SwitchMode = True
        Exit Function
    End If
    startedAt = VBA.Timer
    If Not private_LoadMode(modeId, diagnostic) Then GoTo Failed
    ex_Core.fn_Diagnostic_WritePerf "Mode.Prepare", startedAt
    If Not private_ReadModeConfig("modes." & modeId & ".caption", caption, diagnostic) Then GoTo Failed
    If Not private_ReadModeConfig("text.ModeSelected", status, diagnostic) Then GoTo Failed
    If Not private_SetActionStatus(VBA.Replace$(status, "{mode}", caption), False) Then Exit Function
    startedAt = VBA.Timer
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

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    Set fields = m_modeFields
    If m_reporterActive Then
        Set fields = m_reporterFields
    End If
    If m_hospitalActive Then
        Set fields = m_hospitalFields
    End If
    For Each column In fields
        If column("apply") Then
            value = VBA.vbNullString
            If Not record.TryGetValue(column("name"), value) Then
                diagnostic = "Candidate field not found: " & column("name")
                Exit Function
            End If
            updates("Form." & column("formField")) = value
        End If
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
            m_acceptedOrderNumber, m_acceptedOrderDate, records, diagnostic) Then Exit Function
    Set enrichedResult = New obj_LookupResult
    If Not private_PrepareModeRecords(records, diagnostic) Then Exit Function
    If Not enrichedResult.Initialize(records, m_modeFields, result.HasMore, diagnostic) Then Exit Function
    Set result = enrichedResult
    If Not Me.TryRefreshSourceLegend(diagnostic) Then Exit Function
    private_TryEnrichResult = True
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