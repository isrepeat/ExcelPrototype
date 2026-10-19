VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_ExportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_movementRenderer As obj_PEB_DslRenderer
Private m_pebDslRenderer As obj_PEB_DslRenderer
Private m_pebWordTemplates As obj_PEB_WordTemplates
Private m_lookupTables As Object

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
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    Set m_lookupTables = VBA.CreateObject("Scripting.Dictionary")
    m_lookupTables.CompareMode = VBA.vbTextCompare
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    Dim entry As Collection
    Dim key As Variant
    Dim table As obj_DataTable

    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_pebDslRenderer Is Nothing Then
        m_pebDslRenderer.Dispose
    End If
    If Not m_pebWordTemplates Is Nothing Then
        m_pebWordTemplates.Dispose
    End If
    If Not m_movementRenderer Is Nothing Then
        m_movementRenderer.Dispose
    End If
    Set m_movementRenderer = Nothing
    Set m_pebDslRenderer = Nothing
    Set m_pebWordTemplates = Nothing
    If Not m_lookupTables Is Nothing Then
        For Each key In m_lookupTables.Keys
            Set entry = m_lookupTables(key)
            Set table = entry(2)
            table.Dispose
        Next key
    End If
    Set m_lookupTables = Nothing
End Sub

Public Function TryPrepare( _
    ByVal uiBindingContext As obj_UiBindingContext, _
    ByVal exporter As String, _
    ByRef values As Object, _
    ByRef diagnostic As String, _
    Optional ByVal prepareOutput As Boolean = True, _
    Optional ByVal wordInputs As Object = Nothing _
) As Boolean
    Dim key As Variant
    Dim scalar As Variant
    Dim objectValue As Object
    Dim isObject As Boolean
    Dim fields As Variant
    Dim mode As String
    Dim required As Variant
    Dim index As Long
    Dim dateFields As Variant
    Dim template As String
    Dim rendered As String
    Dim requiredKey As String
    Dim templateName As String
    Dim startedAt As Double
    Dim stageStartedAt As Double
    Dim additionalInputs As Object
    Dim groupTemplate As Object
    Dim history As Boolean

    startedAt = VBA.Timer
    On Error GoTo EH
    ex_PEB_Export.fn_BeginConfigSnapshot
    ex_Core.fn_Diagnostic_WritePerf "Preview.ConfigSnapshot", startedAt
    ' Keep templates for an isolated test run; interactive operations still reload XML.
    If Not ex_PEB_Testing.fn_IsRunning() Then
        If Not m_pebDslRenderer Is Nothing Then m_pebDslRenderer.Dispose
        If Not m_pebWordTemplates Is Nothing Then m_pebWordTemplates.Dispose
        Set m_pebDslRenderer = Nothing
        Set m_pebWordTemplates = Nothing
    End If
    diagnostic = VBA.vbNullString
    Set values = Nothing
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Export service is not initialized."
    End If
    history = (exporter = "History")
    If history Then exporter = "Word"
    If exporter <> "Word" And exporter <> "Movement" Then
        Err.Raise 5, , "Unknown exporter: " & exporter
    End If
    Set values = VBA.CreateObject("Scripting.Dictionary")
    values.CompareMode = VBA.vbTextCompare

    ' Фиксируем значения формы для одной операции экспорта.
    fields = VBA.Split(ex_PEB_Export.fn_Config("export.fields"), "|")
    For Each key In fields
        key = VBA.Trim$(VBA.CStr(key))
        If Not uiBindingContext.TryGetValue("Form", VBA.CStr(key), scalar, objectValue, isObject) Then
            Err.Raise 5, , "Export form field not found: " & VBA.CStr(key)
        End If
        If isObject Or VBA.IsError(scalar) Or VBA.IsNull(scalar) Then
            Err.Raise 5, , "Invalid export form value: " & VBA.CStr(key)
        End If
        If values.Exists(key) Then
            Err.Raise 5, , "Duplicate export field: " & VBA.CStr(key)
        End If
        values.Add VBA.CStr(key), scalar
    Next key
    mode = ex_PEB_EventRoutes.fn_Apply(values, exporter)
    If exporter = "Word" And values("WordPolicy") = "Omit" And Not history Then
        values("WordOutcome") = "Omitted"
        values("WordText") = VBA.vbNullString
        TryPrepare = True
        GoTo CleanupPrepare
    End If
    requiredKey = "export." & exporter & "." & mode & ".required"
    If history Then requiredKey = "history.required"
    If exporter = "Word" And Not prepareOutput Then
        requiredKey = "export.Word." & mode & ".previewRequired"
    End If
    If exporter = "Movement" Then
        required = VBA.Split(ex_PEB_EventRoutes.fn_Value(mode, "movement.required"), "|")
    Else
        required = VBA.Split(ex_PEB_Export.fn_Config(requiredKey), "|")
    End If
    For Each key In required
        If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.Trim$(VBA.CStr(key))))) = 0 Then
            Err.Raise 5, , "Required export field is empty: " & VBA.CStr(key)
        End If
    Next key

    ' Требуем полный год, чтобы короткая дата не получила случайный контекст.
    dateFields = VBA.Split(ex_PEB_Export.fn_Config("export.dateFields"), "|")
    For Each key In dateFields
        scalar = ex_PEB_Export.fn_Required(values, VBA.CStr(key))
        If VBA.Len(ex_PEB_Export.fn_Text(scalar)) > 0 Then
            values(key) = private_Date(scalar, VBA.CStr(key))
        End If
    Next key

    private_PrepareTicket values
    private_PrepareFoodSupport values

    ' Both exporters consume the shared data providers; each owns its output templates.
    stageStartedAt = VBA.Timer
    For index = 1 To ex_PEB_Export.fn_Count("export.shared.lookups.count")
        private_Enrich values, "export.shared.lookups." & VBA.CStr(index)
    Next index
    Set additionalInputs = wordInputs
    If additionalInputs Is Nothing Then
        If uiBindingContext.TryGetValue("Data", "WordContext", scalar, objectValue, isObject) Then
            If Not isObject Then Err.Raise 5, , "Data.WordContext must be a dictionary."
            Set additionalInputs = objectValue
        End If
    End If
    If exporter = "Movement" Then
        private_PrepareMovementTexts values, additionalInputs
    End If
    If exporter = "Word" Then
        ' XML задаёт текст пунктов, конфигурация задаёт источники и соответствия полей.
        If values("WordPolicy") <> "Omit" Then
            For index = 1 To ex_PEB_Export.fn_Count("export.Word." & mode & ".lookups.count")
                private_Enrich values, "export.Word." & mode & ".lookups." & VBA.CStr(index)
            Next index
        End If
        ex_Core.fn_Diagnostic_WritePerf "Preview.Lookups", stageStartedAt
        stageStartedAt = VBA.Timer
        private_AddConstants values
        private_AddAliases values

        private_CompleteWordContext values, additionalInputs
        If history And values("WordPolicy") = "Omit" Then
            TryPrepare = True
            GoTo CleanupPrepare
        End If
        If m_pebWordTemplates Is Nothing Then
            Set m_pebWordTemplates = New obj_PEB_WordTemplates
            If Not m_pebWordTemplates.Initialize(diagnostic) Then
                GoTo Failed
            End If
            Set m_pebDslRenderer = New obj_PEB_DslRenderer
            If Not m_pebDslRenderer.Initialize(m_pebWordTemplates.SharedTemplates) Then
                Err.Raise 5, , "Cannot initialize DSL renderer."
            End If
        End If
        templateName = VBA.CStr(values("RoutedWordTemplate"))
        If Not additionalInputs Is Nothing Then
            If additionalInputs.Exists("TemplateName") Then
                If ex_PEB_Export.fn_Text(additionalInputs("TemplateName")) <> templateName Then
                    Err.Raise 5, , "Word template override conflicts with the business event route."
                End If
            End If
        End If
        values("TemplateName") = templateName
        Set groupTemplate = m_pebWordTemplates.GroupTemplate(templateName)
        values("WordGrouped") = Not groupTemplate Is Nothing
        For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.templates." & templateName & ".requiredKeys"), "|")
            If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(key)))) = 0 Then
                Err.Raise 5, , "Required Word context value is empty: " & key
            End If
        Next key
        If Not m_pebWordTemplates.TryRender(templateName, values, rendered, diagnostic) Then
            GoTo Failed
        End If
        ex_Core.fn_Diagnostic_WritePerf "Preview.TemplatesAndDsl", stageStartedAt
        If VBA.Len(rendered) = 0 Or VBA.Len(rendered) > 32767 Then
            Err.Raise 5, , "Word preview must contain between 1 and 32767 characters."
        End If
        values("WordText") = rendered
        values("WordOutcome") = "Success"
        If prepareOutput Then
            If ex_PEB_Export.fn_Config("export.Word.documentMode") = "template" Then
                private_PrepareOrderDocument values
                TryPrepare = True
                GoTo CleanupPrepare
            End If
            template = m_pebWordTemplates.DocumentFilepath
            If VBA.Len(template) > 0 Then
                If Not m_pebDslRenderer.TryRender(template, values, rendered, diagnostic) Then GoTo Failed
                values("XmlDocumentFilepath") = rendered
            End If
            If groupTemplate Is Nothing Then
                template = ex_PEB_Export.fn_Config("export.Word.output")
            Else
                template = ex_PEB_Export.fn_Config("export.Word.grouping.output")
            End If
            If Not m_pebDslRenderer.TryRender(template, values, rendered, diagnostic) Then
                GoTo Failed
            End If
            values("WordOutput") = ex_PEB_Export.fn_Path(rendered)
        End If
    End If
    TryPrepare = True
    GoTo CleanupPrepare
EH:
    diagnostic = VBA.Err.Description
Failed:
    Set values = Nothing
CleanupPrepare:
    ex_PEB_Export.fn_EndConfigSnapshot
    ex_Core.fn_Diagnostic_WritePerf "Preview.Prepare | Success=" & VBA.CStr(TryPrepare), startedAt
End Function

' Rebuild aliases from stored text without rereading personnel or institution sources.
Public Function TryPrepareHistoryMovement( _
    ByVal stored As Object, _
    ByRef values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim key As Variant
    Dim name As String
    Dim mode As String
    Dim index As Long
    Dim documents As Collection
    Dim document As Object
    Dim copy As Object

    On Error GoTo EH
    diagnostic = VBA.vbNullString
    Set values = Nothing
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError026")
    End If
    Set values = VBA.CreateObject("Scripting.Dictionary")
    values.CompareMode = VBA.vbTextCompare
    For Each key In stored.Keys
        If Not VBA.IsObject(stored(key)) Then
            values.Add VBA.CStr(key), stored(key)
        End If
    Next key
    If stored("HistoryReports").Count > 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.MovementHistoryReportsUnsupported")
    End If
    values("ModeId") = ex_PEB_Export.fn_Required(stored, "EventId")
    values("PersonAccusative") = ex_PEB_Export.fn_Required(stored, "FIOAccusative")
    For index = 1 To ex_PEB_Export.fn_Count("export.shared.aliases.count")
        name = ex_PEB_Export.fn_Config("export.shared.aliases." & index & ".name")
        If values.Exists(name) Then
            values.Remove name
        End If
    Next index
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.dateFields") & "|FoodSupportDate", "|")
        If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(key)))) > 0 Then
            values(key) = ex_PEB_History.fn_Date(values(key))
        End If
    Next key
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs.dateScalars"), "|")
        If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(key)))) > 0 Then
            values(key) = VBA.Format$(ex_PEB_History.fn_Date(values(key)), "dd.mm.yyyy")
        End If
    Next key
    For Each key In VBA.Array("RequiresFoodSupportEnroll", "RequiresFoodSupportRemove")
        Select Case VBA.LCase$(VBA.Trim$(VBA.CStr(ex_PEB_Export.fn_Required(values, VBA.CStr(key)))))
            Case "true"
                values(key) = True
            Case "false"
                values(key) = False
            Case Else
                Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError027"), "{value}", VBA.CStr(key))
        End Select
    Next key
    Set documents = New Collection
    For Each document In stored("MetaDocumentTables")
        Set copy = VBA.CreateObject("Scripting.Dictionary")
        For Each key In document.Keys
            copy.Add VBA.CStr(key), document(key)
        Next key
        If VBA.Len(copy("DocDateShort")) > 0 Then
            copy("DocDateShort") = VBA.Format$(ex_PEB_History.fn_Date(copy("DocDateShort")), "dd.mm.yyyy")
        End If
        documents.Add copy
    Next document
    Set values("MetaDocumentTables") = documents
    mode = ex_PEB_EventRoutes.fn_Apply(values, "Movement")
    For Each key In VBA.Split(ex_PEB_EventRoutes.fn_Value(mode, "movement.required"), "|")
        If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(key)))) = 0 Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError030"), "{value}", VBA.CStr(key))
        End If
    Next key
    private_PrepareTicket values
    private_PrepareMovementTexts values, Nothing, True
    TryPrepareHistoryMovement = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Set values = Nothing
End Function

' Rebuild Word context independently of the Movement preview.
Public Function TryPrepareHistory( _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim key As Variant
    Dim item As Object
    Dim reports As Collection
    Dim reportPattern As Object
    Dim report As Object
    Dim body As String
    Dim text As String
    Dim template As String
    Dim replacements As Object
    Dim index As Long
    Dim token As String
    Dim fileSystem As Object
    Dim filename As String
    Dim path As String
    Dim savedEnding As String
    Dim ending As Date
    Dim starting As Date

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError026")
    values("ModeId") = values("EventId")
    template = ex_PEB_EventRoutes.fn_Apply(values, "Word")
    If values("WordPolicy") = "Omit" Then
        TryPrepareHistory = True
        Exit Function
    End If
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.dateFields"), "|")
        If VBA.Len(VBA.CStr(values(key))) > 0 Then values(key) = ex_PEB_History.fn_Date(values(key))
    Next key
    values("PersonAccusative") = values("FIOAccusative")
    values.Remove "FIOAccusative"
    private_AddAliases values
    values("ManualOrderNo") = values("AcceptedOrderNumber")
    For Each key In VBA.Array("RequiresFoodSupportEnroll", "RequiresFoodSupportRemove")
        Select Case VBA.LCase$(VBA.Trim$(VBA.CStr(values(key))))
            Case "true"
                values(key) = True
            Case "false"
                values(key) = False
            Case Else
                Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError027"), "{value}", VBA.CStr(key))
        End Select
    Next key
    values("RequiresFoodSupportChange") = values("RequiresFoodSupportEnroll") Or values("RequiresFoodSupportRemove")
    values("EnrollToFoodSupportDateShort") = VBA.vbNullString
    values("RemoveFromFoodSupportDateShort") = VBA.vbNullString
    If values("RequiresFoodSupportChange") Then
        text = VBA.Format$(ex_PEB_History.fn_Date(values("FoodSupportDate")), "dd.mm.yyyy")
        values("EnrollToFoodSupportDateShort") = text
        values("RemoveFromFoodSupportDateShort") = text
    End If
    Set values("MetaTables") = New Collection
    For Each item In values("MetaDocumentTables")
        If VBA.Len(item("DocDateShort")) > 0 Then item("DocDateShort") = VBA.Format$(ex_PEB_History.fn_Date(item("DocDateShort")), "dd.mm.yyyy")
    Next item
    For Each key In VBA.Array("VlkDateShort", "PrevVacationTicketDateShort")
        If VBA.Len(values(key)) > 0 Then values(key) = VBA.Format$(ex_PEB_History.fn_Date(values(key)), "dd.mm.yyyy")
    Next key
    values("VacationDurationPhrase") = VBA.vbNullString
    values("AdditionalDurationPhrases") = VBA.vbNullString
    savedEnding = values("DateToShort")
    values("DateToShort") = VBA.vbNullString
    values("VacationDatesSameMonth") = False
    values("VacationDatesSameYear") = False
    private_PrepareVacationDuration values
    If VBA.Len(savedEnding) > 0 Then
        ending = ex_PEB_History.fn_Date(savedEnding)
        starting = ex_PEB_History.fn_Date(values("EventDate"))
        If ending < starting Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError028")
        If VBA.Len(values("DateToShort")) > 0 Then
            If savedEnding <> values("DateToShort") Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError029")
        Else
            values("DateToShort") = VBA.Format$(ending, "dd.mm.yyyy")
            values("VacationDatesSameMonth") = (VBA.Year(starting) = VBA.Year(ending) And VBA.Month(starting) = VBA.Month(ending))
            values("VacationDatesSameYear") = (VBA.Year(starting) = VBA.Year(ending))
        End If
    End If
    values("TemplateName") = template
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.documentTemplate"))
    values("WordTemplateFileName") = fileSystem.GetFileName(path)
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.requiredKeys"), "|")
        If Not values.Exists(key) Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError030"), "{value}", VBA.CStr(key))
    Next key
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.templates." & template & ".requiredKeys"), "|")
        If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(key)))) = 0 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError030"), "{value}", VBA.CStr(key))
    Next key
    If m_pebWordTemplates Is Nothing Then
        Set m_pebWordTemplates = New obj_PEB_WordTemplates
        If Not m_pebWordTemplates.Initialize(diagnostic) Then Exit Function
        Set m_pebDslRenderer = New obj_PEB_DslRenderer
        If Not m_pebDslRenderer.Initialize(m_pebWordTemplates.SharedTemplates) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError031")
    End If
    ' Keep each report structured; the basis block renders the complete ordered list.
    Set reports = values("HistoryReports")
    Set reportPattern = VBA.CreateObject("VBScript.RegExp")
    reportPattern.Pattern = ex_PEB_Export.fn_Config("history.selfReportPattern")
    For Each report In reports
        report("RankGenitive") = values("RankGenitive")
        report("FIOInitialsGenitive") = values("FIOInitialsGenitive")
        report("ReportPerson") = report("ReportName")
        If reportPattern.Test(VBA.CStr(report("ReportPerson"))) Then
            report("ReportPerson") = VBA.vbNullString
        End If
        report("IncomingNo") = report("IncomingNumber")
        If VBA.Len(report("IncomingNo")) = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError032")
        report("IncomingDateShort") = VBA.Format$(ex_PEB_History.fn_Date(report("IncomingDate")), "dd.mm.yyyy")
        report("ReporterGenitive") = VBA.Trim$(report("ReportPositionGenitive") & " " & report("ReportRankGenitive") & " " & report("ReportPersonInitialsGenitive"))
        If VBA.Len(report("ReportPerson")) > 0 And VBA.Len(report("ReporterGenitive")) = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError033")
    Next report
    If Not m_pebWordTemplates.TryRender(template, values, body, diagnostic) Then Exit Function
    values("WordText") = body
    Set replacements = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To ex_PEB_Export.fn_Count("export.Word.documentFields.count")
        token = ex_PEB_Export.fn_Config("export.Word.documentFields." & index & ".token")
        If Not m_pebDslRenderer.TryRender(ex_PEB_Export.fn_Config("export.Word.documentFields." & index & ".value"), values, text, diagnostic) Then Exit Function
        If VBA.Len(text) = 0 Or text Like "*[<>:""/\|?*]*" Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError034"), "{value}", VBA.CStr(token))
        replacements.Add token, text
    Next index
    Set values("WordDocumentReplacements") = replacements
    filename = fileSystem.GetFileName(path)
    For Each key In replacements.Keys
        filename = VBA.Replace$(filename, VBA.CStr(key), VBA.CStr(replacements(key)))
    Next key
    If VBA.InStr(filename, "{") > 0 Or VBA.InStr(filename, "}") > 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError035")
    values("WordTemplatePath") = path
    values("WordOutput") = fileSystem.BuildPath(fileSystem.GetParentFolderName(path), filename)
    TryPrepareHistory = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

' //
' // Private
' //
Private Sub private_PrepareOrderDocument(ByVal values As Object)
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim orderQueryFormat As obj_QueryValueFormat
    Dim table As obj_DataTable
    Dim replacements As Object
    Dim fileSystem As Object
    Dim diagnostic As String
    Dim dateHeader As String
    Dim numberHeader As String
    Dim signatoryHeader As String
    Dim signatory As String
    Dim candidate As String
    Dim row As Long
    Dim index As Long
    Dim prefix As String
    Dim token As Variant
    Dim rendered As String
    Dim filename As String
    Dim templatePath As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo EH
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    Set orderQueryFormat = New obj_QueryValueFormat
    If Not orderQueryFormat.InitializeFromConfig("PersonalEventBuilder::orders.queryFormat") Then
        Err.Raise 5, , "Cannot initialize order query value format."
    End If
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        Err.Raise 5, , "Cannot initialize order document lookup."
    End If
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("orders.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("orders.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("orders.source.range")
    dateHeader = ex_PEB_Export.fn_Config("orders.dateHeader")
    numberHeader = ex_PEB_Export.fn_Config("orders.numberHeader")
    signatoryHeader = ex_PEB_Export.fn_Config("orders.signatoryHeader")
    query.AddColumn signatoryHeader
    query.AddCondition _
        numberHeader, _
        QueryEquals, _
        values("AcceptedOrderNumber"), _
        QueryNumber, _
        True, _
        False, _
        orderQueryFormat, _
        orderQueryFormat
    query.AddCondition _
        dateHeader, _
        QueryEquals, _
        values("AcceptedOrderDate"), _
        QueryDate, _
        True, _
        False, _
        orderQueryFormat, _
        orderQueryFormat
    service.Backend = QueryAuto
    If Not service.TryExecute(source, query, table, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    For row = 1 To table.RowCount
        candidate = ex_PEB_Export.fn_Text(table.ValueAt(row, ex_TableQuery.fn_HeaderIndex(table.Headers, signatoryHeader)))
        If VBA.Len(candidate) = 0 Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderSignatoryRequired")
        End If
        If VBA.Len(signatory) > 0 And signatory <> candidate Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderSignatoryAmbiguous")
        End If
        signatory = candidate
    Next row
    If VBA.Len(signatory) = 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderSignatoryRequired")
    End If
    values("Signatory") = signatory
    Set replacements = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To ex_PEB_Export.fn_Count("export.Word.documentFields.count")
        prefix = "export.Word.documentFields." & VBA.CStr(index)
        token = ex_PEB_Export.fn_Config(prefix & ".token")
        If Not m_pebDslRenderer.TryRender(ex_PEB_Export.fn_Config(prefix & ".value"), values, rendered, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        replacements.Add VBA.CStr(token), rendered
    Next index
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    templatePath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.documentTemplate"))
    filename = fileSystem.GetFileName(templatePath)
    For Each token In replacements.Keys
        candidate = VBA.CStr(replacements(token))
        If VBA.Len(candidate) = 0 Or candidate Like "*[<>:""/\|?*]*" Or _
                VBA.InStr(candidate, VBA.vbCr) > 0 Or VBA.InStr(candidate, VBA.vbLf) > 0 Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderDocumentNameInvalid")
        End If
        filename = VBA.Replace$(filename, VBA.CStr(token), candidate)
    Next token
    If VBA.InStr(filename, "{") > 0 Or VBA.InStr(filename, "}") > 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderDocumentNameInvalid")
    End If
    values("WordOutput") = fileSystem.BuildPath(fileSystem.GetParentFolderName(templatePath), filename)
    If VBA.StrComp(values("WordOutput"), templatePath, VBA.vbTextCompare) = 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderDocumentNameInvalid")
    End If
    Set values("WordDocumentReplacements") = replacements
Cleanup:
    If Not table Is Nothing Then table.Dispose
    If Not service Is Nothing Then service.Dispose
    If Not orderQueryFormat Is Nothing Then orderQueryFormat.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, , errorText
    End If
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Sub

Private Sub private_CompleteWordContext( _
    ByVal values As Object, _
    ByVal inputs As Object _
)
    Dim key As Variant
    Dim scalarKeys As String
    Dim booleanKeys As String
    Dim collectionKeys As String
    Dim index As Long
    Dim prefix As String
    Dim output As String
    Dim items As Collection
    Dim normalized As Collection
    Dim item As Variant
    Dim row As Object
    Dim field As Variant
    Dim allowed As String
    Dim required As String

    scalarKeys = "|" & ex_PEB_Export.fn_Config("export.Word.context.scalarInputs") & "|"
    Set values("HistoryReports") = New Collection
    booleanKeys = "|" & ex_PEB_Export.fn_Config("export.Word.context.booleanInputs") & "|"
    collectionKeys = "|" & ex_PEB_Export.fn_Config("export.Word.context.collectionInputs") & "|"
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.optionalText"), "|")
        If Not values.Exists(key) Then values.Add VBA.CStr(key), VBA.vbNullString
    Next key
    values("RequiresFoodSupportRemove") = (values("ModeDirection") = ex_PEB_Export.fn_Config("export.Word.context.removeDirection"))
    values("RequiresFoodSupportEnroll") = (values("ModeDirection") = ex_PEB_Export.fn_Config("export.Word.context.enrollDirection"))
    If Not inputs Is Nothing Then
        If VBA.TypeName(inputs) <> "Dictionary" Then Err.Raise 5, , "Word context inputs must be a dictionary."
        For Each key In inputs.Keys
            If VBA.InStr(1, scalarKeys, "|" & key & "|", VBA.vbBinaryCompare) > 0 Then
                If VBA.IsObject(inputs(key)) Or VBA.IsArray(inputs(key)) Or VBA.IsNull(inputs(key)) Or VBA.IsError(inputs(key)) Then Err.Raise 5, , "Invalid Word context scalar: " & key
                values(key) = inputs(key)
            ElseIf VBA.InStr(1, booleanKeys, "|" & key & "|", VBA.vbBinaryCompare) > 0 Then
                If VBA.VarType(inputs(key)) <> VBA.vbBoolean Then Err.Raise 5, , "Word context Boolean required: " & key
                values(key) = inputs(key)
            ElseIf VBA.InStr(1, collectionKeys, "|" & key & "|", VBA.vbBinaryCompare) > 0 Then
                If Not VBA.IsObject(inputs(key)) Then Err.Raise 5, , "Word context collection required: " & key
                Set items = inputs(key)
                If VBA.TypeName(items) <> "Collection" Then Err.Raise 5, , "Word context requires Collection: " & key
                Set normalized = New Collection
                allowed = "|" & ex_PEB_Export.fn_Config("export.Word.context." & key & ".fields") & "|"
                required = ex_PEB_Export.fn_Config("export.Word.context." & key & ".required")
                For Each item In items
                    If Not VBA.IsObject(item) Then Err.Raise 5, , "Word metadata row must be a dictionary: " & key
                    If VBA.TypeName(item) <> "Dictionary" Then Err.Raise 5, , "Word metadata row must be a dictionary: " & key
                    Set row = VBA.CreateObject("Scripting.Dictionary")
                    row.CompareMode = VBA.vbBinaryCompare
                    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context." & key & ".fields"), "|")
                        row(field) = VBA.vbNullString
                    Next field
                    For Each field In item.Keys
                        If VBA.InStr(1, allowed, "|" & field & "|", VBA.vbBinaryCompare) = 0 Then Err.Raise 5, , "Unknown Word metadata field: " & field
                        If VBA.IsObject(item(field)) Or VBA.IsArray(item(field)) Or VBA.IsNull(item(field)) Or VBA.IsError(item(field)) Then Err.Raise 5, , "Invalid Word metadata value: " & field
                        row(field) = item(field)
                    Next field
                    For Each field In VBA.Split(required, "|")
                        If VBA.Len(ex_PEB_Export.fn_Text(row(field))) = 0 Then Err.Raise 5, , "Required Word metadata field is empty: " & field
                    Next field
                    If VBA.CStr(key) = "MetaTvoTables" Then
                        For index = 1 To ex_PEB_Export.fn_Count("export.Word.meta.lookups.count")
                            private_Enrich row, "export.Word.meta.lookups." & VBA.CStr(index)
                        Next index
                    End If
                    If row.Exists("DocDateShort") Then
                        If VBA.Len(ex_PEB_Export.fn_Text(row("DocDateShort"))) > 0 Then row("DocDateShort") = VBA.Format$(private_Date(row("DocDateShort"), "DocDateShort"), "dd.mm.yyyy")
                    End If
                    normalized.Add row
                Next item
                Set values(key) = normalized
            Else
                Err.Raise 5, , "Unknown Word context input: " & key
            End If
        Next key
    End If
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.dateInputs"), "|")
        If VBA.Len(ex_PEB_Export.fn_Text(values(key))) > 0 Then values(key) = VBA.Format$(private_Date(values(key), VBA.CStr(key)), "dd.mm.yyyy")
    Next key
    For index = 1 To ex_PEB_Export.fn_Count("export.Word.context.lookups.count")
        prefix = "export.Word.context.lookups." & VBA.CStr(index)
        output = ex_PEB_Export.fn_Config(prefix & ".output")
        If Not values.Exists(output) Then private_Enrich values, prefix
    Next index
    values("RequiresFoodSupportChange") = values("RequiresFoodSupportEnroll") Or values("RequiresFoodSupportRemove")
    values("EnrollToFoodSupportDateShort") = VBA.Format$(values("FoodSupportDate"), "dd.mm.yyyy")
    values("RemoveFromFoodSupportDateShort") = VBA.Format$(values("FoodSupportDate"), "dd.mm.yyyy")
    values("WordTemplateFileName") = VBA.CreateObject("Scripting.FileSystemObject").GetFileName(ex_PEB_Export.fn_Config("export.Word.documentTemplate"))
    values("ManualOrderNo") = values("AcceptedOrderNumber")
    private_PrepareVacationDuration values
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.requiredKeys"), "|")
        If Not values.Exists(key) Then Err.Raise 5, , "Word XML context provider is missing: " & key
    Next key
End Sub

Private Sub private_PrepareVacationDuration(ByVal values As Object)
    Dim text As String
    Dim rx As Object
    Dim matches As Object
    Dim operand As Object
    Dim days As Long
    Dim total As Long
    Dim additional As Long
    Dim unit As String
    Dim reason As String
    Dim reasonKey As String
    Dim mapping As Variant
    Dim parts As Variant
    Dim phrase As String
    Dim extras As String
    Dim ending As Date
    Dim starting As Date

    text = ex_PEB_Export.fn_Text(values("DurationDays"))
    If VBA.Len(text) = 0 Then Exit Sub
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = ex_PEB_Export.fn_Config("export.Word.duration.fullPattern")
    If Not rx.Test(text) Then Err.Raise 5, , "Invalid vacation duration: " & text
    Set matches = rx.Execute(text)
    days = VBA.CLng(matches(0).SubMatches(0))
    If days < 1 Then Err.Raise 5, , "Vacation duration must be positive."
    total = days
    If days Mod 100 >= 11 And days Mod 100 <= 14 Then
        unit = ex_PEB_Export.fn_Config("export.Word.duration.unit.many")
    ElseIf days Mod 10 = 1 Then
        unit = ex_PEB_Export.fn_Config("export.Word.duration.unit.one")
    ElseIf days Mod 10 >= 2 And days Mod 10 <= 4 Then
        unit = ex_PEB_Export.fn_Config("export.Word.duration.unit.few")
    Else
        unit = ex_PEB_Export.fn_Config("export.Word.duration.unit.many")
    End If
    phrase = ex_PEB_Export.fn_Config("export.Word.duration.phrase")
    phrase = VBA.Replace$(phrase, "{days}", VBA.CStr(days))
    values("VacationDurationPhrase") = VBA.Replace$(phrase, "{unit}", unit)
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = ex_PEB_Export.fn_Config("export.Word.duration.operandPattern")
    For Each operand In rx.Execute(text)
        additional = VBA.CLng(operand.SubMatches(0))
        If additional < 1 Then Err.Raise 5, , "Additional duration must be positive."
        reason = VBA.LCase$(VBA.Trim$(VBA.CStr(operand.SubMatches(1))))
        reasonKey = VBA.vbNullString
        If VBA.Len(reason) > 0 Then
            For Each mapping In VBA.Split(ex_PEB_Export.fn_Config("export.Word.duration.reasonKeys"), "|")
                parts = VBA.Split(mapping, "=", 2)
                If reason = parts(0) Then reasonKey = parts(1)
            Next mapping
            If VBA.Len(reasonKey) = 0 Then Err.Raise 5, , "Unknown additional vacation reason: " & reason
        End If
        If additional = 1 Then
            unit = ex_PEB_Export.fn_Config("export.Word.duration.additionalUnit.one")
        Else
            unit = ex_PEB_Export.fn_Config("export.Word.duration.additionalUnit.many")
        End If
        phrase = ex_PEB_Export.fn_Config("export.Word.duration.additionalPhrase")
        phrase = VBA.Replace$(phrase, "{days}", VBA.CStr(additional))
        phrase = VBA.Replace$(phrase, "{unit}", unit)
        If VBA.Len(reasonKey) > 0 Then
            If additional = 1 Then
                phrase = phrase & ex_PEB_Export.fn_Config("export.Word.duration." & reasonKey & ".one")
            Else
                phrase = phrase & ex_PEB_Export.fn_Config("export.Word.duration." & reasonKey & ".many")
            End If
        End If
        extras = extras & phrase
        total = total + additional
    Next operand
    If VBA.Len(extras) > 0 Then extras = extras & ","
    values("AdditionalDurationPhrases") = extras
    values("VacationDays") = days
    values("VacationTotalDays") = total
    starting = private_Date(values("DateFromShort"), "DateFromShort")
    ending = VBA.DateAdd("d", total - 1, starting)
    values("DateToShort") = VBA.Format$(ending, "dd.mm.yyyy")
    values("VacationDatesSameMonth") = (VBA.Year(starting) = VBA.Year(ending) And VBA.Month(starting) = VBA.Month(ending))
    values("VacationDatesSameYear") = (VBA.Year(starting) = VBA.Year(ending))
End Sub

Private Sub private_PrepareFoodSupport(ByVal values As Object)
    Dim effectiveDate As Date
    Dim eventDate As Date
    Dim offsetDays As Long

    offsetDays = ex_PEB_Export.fn_Count("export.Movement.foodSupport.minimumOrderOffsetDays")
    If offsetDays < 1 Then Err.Raise 5, , "Food support order offset must be at least one day."
    effectiveDate = VBA.DateAdd("d", offsetDays, private_Date(values("AcceptedOrderDate"), "AcceptedOrderDate"))
    eventDate = private_Date(values("EventDate"), "EventDate")
    If eventDate > effectiveDate Then effectiveDate = eventDate
    values("FoodSupportDate") = effectiveDate
End Sub

Private Sub private_PrepareTicket(ByVal values As Object)
    Dim ticket As String
    Dim pattern As Object
    Dim qualified As String

    values("TicketDocument") = VBA.vbNullString
    If ex_PEB_Export.fn_Text(values("ModeForm")) <> ex_PEB_Export.fn_Config("export.ticket.form") Then Exit Sub
    ticket = ex_PEB_Export.fn_Text(values("TicketNumber"))
    If VBA.Len(ticket) = 0 Then Exit Sub
    If VBA.InStr(1, "|" & ex_PEB_Export.fn_Config("lookup.ReturnTickets.modes") & "|", _
            "|" & VBA.CStr(values("ModeId")) & "|", VBA.vbTextCompare) > 0 Then
        values("TicketDocument") = ticket
        Exit Sub
    End If
    Set pattern = VBA.CreateObject("VBScript.RegExp")
    pattern.Pattern = ex_PEB_Export.fn_Config("export.ticket.qualifiedPattern")
    If pattern.Test(ticket) Then
        qualified = ticket
    Else
        qualified = ex_PEB_Export.fn_Config("export.ticket.numberTemplate")
        qualified = VBA.Replace(qualified, "{orderYear}", VBA.CStr(VBA.Year(private_Date(values("AcceptedOrderDate"), "AcceptedOrderDate"))))
        qualified = VBA.Replace(qualified, "{orderNum}", ex_PEB_Export.fn_Text(values("AcceptedOrderNumber")))
        qualified = VBA.Replace(qualified, "{ticketNum}", ticket)
    End If
    values("TicketNumber") = qualified
    values("TicketDocument") = qualified
End Sub

Private Sub private_PrepareMovementTexts( _
    ByVal values As Object, _
    ByVal inputs As Object, _
    Optional ByVal storedSnapshot As Boolean = False _
)
    Dim rendered As String
    Dim diagnostic As String
    Dim templateId As String
    Dim templates As Object
    Dim results As Object
    Dim column As Variant

    private_AddConstants values, storedSnapshot
    private_AddAliases values
    If Not storedSnapshot Then
        private_CompleteEventInputs values, inputs
    End If
    If m_movementRenderer Is Nothing Then
        Set m_movementRenderer = New obj_PEB_DslRenderer
        If Not m_movementRenderer.Initialize(ex_PEB_MovementTemplates.fn_SharedTemplates()) Then
            Err.Raise 5, , "Cannot initialize Movement DSL renderer."
        End If
    End If
    templateId = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "MovementTemplateId"))
    Set templates = ex_PEB_MovementTemplates.fn_Texts(templateId)
    Set results = VBA.CreateObject("Scripting.Dictionary")
    results.CompareMode = VBA.vbTextCompare
    For Each column In templates.Keys
        If Not m_movementRenderer.TryRender(VBA.CStr(templates(column)), values, rendered, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If VBA.Len(VBA.Trim$(rendered)) = 0 Then
            Err.Raise 5, , "Movement column text is empty: " & templateId & "/" & column
        End If
        results.Add VBA.CStr(column), rendered
    Next column
    Set values("MovementColumnTexts") = results
    ex_PEB_MovementRegistry.fn_ValidateTexts values
End Sub

Private Sub private_CompleteEventInputs( _
    ByVal values As Object, _
    ByVal inputs As Object _
)
    Dim name As Variant
    Dim field As Variant
    Dim fields As String
    Dim required As String
    Dim dates As String
    Dim scalar As Variant
    Dim items As Collection
    Dim normalized As Collection
    Dim item As Variant
    Dim row As Object

    If Not inputs Is Nothing Then
        If VBA.TypeName(inputs) <> "Dictionary" Then
            Err.Raise 5, , "Event inputs must be a dictionary."
        End If
    End If
    dates = "|" & ex_PEB_Export.fn_Config("export.shared.inputs.dateScalars") & "|"
    For Each name In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs.scalars"), "|")
        If Not values.Exists(name) Then
            values.Add VBA.CStr(name), VBA.vbNullString
        End If
        If Not inputs Is Nothing Then
            If inputs.Exists(name) Then
                If VBA.IsObject(inputs(name)) Or VBA.IsArray(inputs(name)) Or VBA.IsNull(inputs(name)) Or VBA.IsError(inputs(name)) Then
                    Err.Raise 5, , "Invalid event input scalar: " & name
                End If
                values(name) = inputs(name)
            End If
        End If
        If VBA.InStr(1, dates, "|" & name & "|", VBA.vbBinaryCompare) > 0 And VBA.Len(ex_PEB_Export.fn_Text(values(name))) > 0 Then
            values(name) = VBA.Format$(private_Date(values(name), VBA.CStr(name)), "dd.mm.yyyy")
        End If
    Next name
    For Each name In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs.collections"), "|")
        Set normalized = New Collection
        fields = "|" & ex_PEB_Export.fn_Config("export.shared.inputs." & name & ".fields") & "|"
        required = ex_PEB_Export.fn_Config("export.shared.inputs." & name & ".required")
        If Not inputs Is Nothing Then
            If inputs.Exists(name) Then
                If Not VBA.IsObject(inputs(name)) Then
                    Err.Raise 5, , "Event collection required: " & name
                End If
                If VBA.TypeName(inputs(name)) <> "Collection" Then
                    Err.Raise 5, , "Event Collection required: " & name
                End If
                Set items = inputs(name)
                For Each item In items
                    If Not VBA.IsObject(item) Then
                        Err.Raise 5, , "Event metadata dictionary required: " & name
                    End If
                    If VBA.TypeName(item) <> "Dictionary" Then
                        Err.Raise 5, , "Event metadata dictionary required: " & name
                    End If
                    Set row = VBA.CreateObject("Scripting.Dictionary")
                    row.CompareMode = VBA.vbBinaryCompare
                    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs." & name & ".fields"), "|")
                        row.Add VBA.CStr(field), VBA.vbNullString
                    Next field
                    For Each field In item.Keys
                        If VBA.InStr(1, fields, "|" & field & "|", VBA.vbBinaryCompare) = 0 Then
                            Err.Raise 5, , "Unknown event metadata field: " & field
                        End If
                        If VBA.IsObject(item(field)) Or VBA.IsArray(item(field)) Or VBA.IsNull(item(field)) Or VBA.IsError(item(field)) Then
                            Err.Raise 5, , "Invalid event metadata value: " & field
                        End If
                        scalar = item(field)
                        row(field) = scalar
                    Next field
                    For Each field In VBA.Split(required, "|")
                        If VBA.Len(ex_PEB_Export.fn_Text(row(field))) = 0 Then
                            Err.Raise 5, , "Required event metadata field is empty: " & field
                        End If
                    Next field
                    If row.Exists("DocDateShort") Then
                        If VBA.Len(ex_PEB_Export.fn_Text(row("DocDateShort"))) > 0 Then
                            row("DocDateShort") = VBA.Format$(private_Date(row("DocDateShort"), "DocDateShort"), "dd.mm.yyyy")
                        End If
                    End If
                    normalized.Add row
                Next item
            End If
        End If
        Set values(name) = normalized
    Next name
End Sub

Private Function private_Date( _
    ByVal value As Variant, _
    ByVal key As String _
) As Date
    Dim parts As Variant
    Dim parsed As Date
    Dim text As String
    Dim rx As Object

    If VBA.VarType(value) = VBA.vbDate Then
        private_Date = VBA.DateValue(value)
        Exit Function
    End If
    If VBA.IsNumeric(value) And VBA.VarType(value) <> VBA.vbString Then
        If VBA.CDbl(value) < 1 Or VBA.CDbl(value) > 2958465 Then
            Err.Raise 5, , "Invalid Excel date: " & key
        End If
        private_Date = VBA.DateValue(VBA.CDate(value))
        Exit Function
    End If
    text = ex_PEB_Export.fn_Text(value)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^\d{2}\.\d{2}\.\d{4}$"
    If Not rx.Test(text) Then
        Err.Raise 5, , "Full date DD.MM.YYYY required: " & key
    End If
    parts = VBA.Split(text, ".")
    If VBA.CLng(parts(2)) < 1900 Then
        Err.Raise 5, , "Invalid date year: " & key
    End If
    parsed = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(parsed, "dd.mm.yyyy") <> text Then
        Err.Raise 5, , "Invalid calendar date: " & key
    End If
    private_Date = parsed
End Function

Private Sub private_Enrich( _
    ByVal values As Object, _
    ByVal prefix As String _
)
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim tableQueryService As obj_TableQueryService
    Dim dataTable As obj_DataTable
    Dim header As String
    Dim inputField As String
    Dim outputField As String
    Dim valueHeader As String
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorText As String
    Dim whenField As String
    Dim whenPattern As String
    Dim rx As Object
    Dim lookupTable As obj_DataTable
    Dim lookupValues As Object
    Dim lookupKey As String

    On Error GoTo EH
    inputField = ex_PEB_Export.fn_Config(prefix & ".input")
    outputField = ex_PEB_Export.fn_Config(prefix & ".output")
    header = ex_PEB_Export.fn_Config(prefix & ".keyHeader")
    valueHeader = ex_PEB_Export.fn_Config(prefix & ".valueHeader")
    If values.Exists(outputField) Then
        Err.Raise 5, , "Lookup output would overwrite an export field: " & outputField
    End If
    whenField = ex_PEB_Export.fn_Config(prefix & ".when.field")
    whenPattern = ex_PEB_Export.fn_Config(prefix & ".when.pattern")
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = whenPattern
    rx.IgnoreCase = True
    If Not rx.Test(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, whenField))) Then
        values.Add outputField, VBA.vbNullString
        Exit Sub
    End If
    If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, inputField))) = 0 Then
        Err.Raise 5, , "Lookup input is empty: " & inputField
    End If
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    Set tableQueryService = New obj_TableQueryService
    If Not tableSource.Initialize() Then
        Err.Raise 5, , "Cannot initialize export lookup source."
    End If
    If Not tableQuery.Initialize() Then
        Err.Raise 5, , "Cannot initialize export lookup query."
    End If
    If Not tableQueryService.Initialize() Then
        Err.Raise 5, , "Cannot initialize export lookup service."
    End If
    tableSource.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(prefix & ".path"))
    tableSource.SheetName = ex_PEB_Export.fn_Config(prefix & ".sheet")
    tableSource.RangeAddress = ex_PEB_Export.fn_Config(prefix & ".range")
    tableQuery.AddColumn header
    If header <> valueHeader Then
        tableQuery.AddColumn valueHeader
    End If
    tableQuery.AddCondition header, QueryEquals, values(inputField), QueryText, True
    tableQuery.Limit = 2
    tableQueryService.Backend = QueryAuto
    Set lookupTable = private_LookupTable(tableSource, tableQueryService, lookupValues, diagnostic)
    If lookupTable Is Nothing Then
        Err.Raise 5, , diagnostic
    End If
    lookupKey = header & VBA.vbNullChar & valueHeader & VBA.vbNullChar & VBA.CStr(values(inputField))
    If lookupValues.Exists(lookupKey) Then
        values.Add outputField, lookupValues(lookupKey)
        GoTo Cleanup
    End If
    If Not ex_TableQuery.fn_TryApply(lookupTable, tableQuery, dataTable, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If dataTable.RowCount <> 1 Then
        Err.Raise 5, , "Export lookup must return exactly one row: " & prefix
    End If
    values.Add outputField, dataTable.ValueAt(1, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, valueHeader))
    If VBA.Len(ex_PEB_Export.fn_Text(values(outputField))) = 0 Then
        Err.Raise 5, , "Export lookup returned an empty value: " & outputField
    End If
    lookupValues.Add lookupKey, values(outputField)
Cleanup:
    If Not dataTable Is Nothing Then
        dataTable.Dispose
    End If
    If Not tableQueryService Is Nothing Then
        tableQueryService.Dispose
    End If
    If Not tableQuery Is Nothing Then
        tableQuery.Dispose
    End If
    If Not tableSource Is Nothing Then
        tableSource.Dispose
    End If
    If errorNumber <> 0 Then
        ' Propagate the saved error instead of entering this handler again.
        On Error GoTo 0
        Err.Raise errorNumber, , errorText
    End If
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Sub

Private Function private_LookupTable( _
    ByVal source As obj_TableSource, _
    ByVal service As obj_TableQueryService, _
    ByRef lookupValues As Object, _
    ByRef diagnostic As String _
) As obj_DataTable
    Dim key As String
    Dim stamp As String
    Dim fileSystem As Object
    Dim file As Object
    Dim book As Workbook
    Dim isOpen As Boolean
    Dim entry As Collection
    Dim table As obj_DataTable
    Dim previous As obj_DataTable
    Dim query As obj_TableQuery
    Dim startedAt As Double

    startedAt = VBA.Timer
    key = source.WorkbookPath & VBA.vbNullChar & source.SheetName & VBA.vbNullChar & _
        source.RangeAddress & VBA.vbNullChar & source.TableName
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set file = fileSystem.GetFile(source.WorkbookPath)
    stamp = VBA.CStr(file.DateLastModified) & "|" & VBA.CStr(file.Size)
    For Each book In Application.Workbooks
        If VBA.StrComp(book.FullName, source.WorkbookPath, VBA.vbTextCompare) = 0 Then
            isOpen = True
            Exit For
        End If
    Next book
    If m_lookupTables.Exists(key) Then
        Set entry = m_lookupTables(key)
        If Not isOpen And entry(1) = stamp Then
            Set private_LookupTable = entry(2)
            Set lookupValues = entry(3)
            ex_Core.fn_Diagnostic_WritePerf "Preview.LookupCacheHit | Sheet=" & source.SheetName, startedAt
            Exit Function
        End If
    End If
    Set query = New obj_TableQuery
    If Not query.Initialize() Then Err.Raise 5, , "Cannot initialize export directory query."
    If Not service.TryExecute(source, query, table, diagnostic) Then
        query.Dispose
        Exit Function
    End If
    query.Dispose
    If m_lookupTables.Exists(key) Then
        Set entry = m_lookupTables(key)
        Set previous = entry(2)
        previous.Dispose
    End If
    Set entry = New Collection
    If isOpen Then stamp = VBA.vbNullString
    entry.Add stamp
    entry.Add table
    Set lookupValues = VBA.CreateObject("Scripting.Dictionary")
    entry.Add lookupValues
    Set m_lookupTables(key) = entry
    Set private_LookupTable = table
    ex_Core.fn_Diagnostic_WritePerf "Preview.LookupCacheLoad | Sheet=" & source.SheetName, startedAt
End Function

Private Sub private_AddConstants( _
    ByVal values As Object, _
    Optional ByVal preserveExisting As Boolean = False _
)
    Dim index As Long
    Dim prefix As String
    Dim name As String
    Dim kind As String
    Dim value As String
    Dim items As Collection

    For index = 1 To ex_PEB_Export.fn_Count("export.shared.constants.count")
        prefix = "export.shared.constants." & VBA.CStr(index)
        name = ex_PEB_Export.fn_Config(prefix & ".name")
        kind = ex_PEB_Export.fn_Config(prefix & ".kind")
        If values.Exists(name) Then
            If Not preserveExisting Then
                Err.Raise 5, , "Constant would overwrite an export value: " & name
            End If
        Else
            Select Case kind
                Case "text"
                    value = ex_PEB_Export.fn_Config(prefix & ".value")
                    values.Add name, value
                Case "emptyText"
                    values.Add name, VBA.vbNullString
                Case "boolean"
                    value = ex_PEB_Export.fn_Config(prefix & ".value")
                    If value <> "true" And value <> "false" Then
                        Err.Raise 5, , "Invalid Boolean constant: " & name
                    End If
                    values.Add name, (value = "true")
                Case "emptyCollection"
                    Set items = New Collection
                    values.Add name, items
                Case Else
                    Err.Raise 5, , "Unknown export constant kind: " & kind
            End Select
        End If
    Next index
End Sub

Private Sub private_AddAliases(ByVal values As Object)
    Dim index As Long
    Dim prefix As String
    Dim name As String
    Dim kind As String
    Dim source As String
    Dim field As Variant
    Dim scalar As Variant
    Dim text As String
    Dim part As String

    For index = 1 To ex_PEB_Export.fn_Count("export.shared.aliases.count")
        prefix = "export.shared.aliases." & VBA.CStr(index)
        name = ex_PEB_Export.fn_Config(prefix & ".name")
        kind = ex_PEB_Export.fn_Config(prefix & ".kind")
        source = ex_PEB_Export.fn_Config(prefix & ".source")
        If values.Exists(name) Then
            Err.Raise 5, , "Alias would overwrite an export value: " & name
        End If
        Select Case kind
            Case "value"
                scalar = ex_PEB_Export.fn_Required(values, source)
            Case "shortDate"
                scalar = ex_PEB_Export.fn_Required(values, source)
                If VBA.Len(ex_PEB_Export.fn_Text(scalar)) > 0 Then
                    scalar = VBA.Format$(private_Date(scalar, source), "dd.mm.yyyy")
                End If
            Case "joinLowerFirst"
                text = VBA.vbNullString
                For Each field In VBA.Split(source, "|")
                    part = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(field)))
                    If VBA.Len(part) > 0 Then
                        If VBA.Len(text) > 0 Then
                            text = text & " "
                        End If
                        text = text & part
                    End If
                Next field
                scalar = VBA.LCase$(VBA.Left$(text, 1)) & VBA.Mid$(text, 2)
            Case Else
                Err.Raise 5, , "Unknown export alias kind: " & kind
        End Select
        values.Add name, scalar
    Next index
End Sub