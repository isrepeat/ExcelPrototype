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
    Optional ByVal prepareOutput As Boolean = True _
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

    startedAt = VBA.Timer
    On Error GoTo EH
    ex_PEB_Export.fn_BeginConfigSnapshot
    ex_Core.fn_Diagnostic_WritePerf "Preview.ConfigSnapshot", startedAt
    ' Reload templates per operation so edits to the XML are visible immediately.
    If Not m_pebDslRenderer Is Nothing Then m_pebDslRenderer.Dispose
    If Not m_pebWordTemplates Is Nothing Then m_pebWordTemplates.Dispose
    Set m_pebDslRenderer = Nothing
    Set m_pebWordTemplates = Nothing
    diagnostic = VBA.vbNullString
    Set values = Nothing
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Export service is not initialized."
    End If
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
    mode = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "ModeId"))
    values.Add "EventName", ex_PEB_Export.fn_Config("export.events." & mode)
    ' Select the destination input for each Movement mode through configuration.
    If exporter = "Movement" Then
        key = ex_PEB_Export.fn_Config("export.Movement." & mode & ".destinationField")
        values("Destination") = ex_PEB_Export.fn_Required(values, VBA.CStr(key))
    End If
    requiredKey = "export." & exporter & "." & mode & ".required"
    If exporter = "Word" And Not prepareOutput Then
        requiredKey = "export.Word." & mode & ".previewRequired"
    End If
    required = VBA.Split(ex_PEB_Export.fn_Config(requiredKey), "|")
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

    ' Enrich the snapshot using only the selected exporter's directories.
    stageStartedAt = VBA.Timer
    For index = 1 To ex_PEB_Export.fn_Count("export." & exporter & ".lookups.count")
        private_Enrich values, "export." & exporter & ".lookups." & VBA.CStr(index)
    Next index
    If exporter = "Movement" Then
        private_PrepareMovementBasis values
    End If
    If exporter = "Word" Then
        ' XML задаёт текст пунктов, конфигурация задаёт источники и соответствия полей.
        For index = 1 To ex_PEB_Export.fn_Count("export.Word." & mode & ".lookups.count")
            private_Enrich values, "export.Word." & mode & ".lookups." & VBA.CStr(index)
        Next index
        ex_Core.fn_Diagnostic_WritePerf "Preview.Lookups", stageStartedAt
        stageStartedAt = VBA.Timer
        private_AddConstants values
        private_AddAliases values
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
        templateName = ex_PEB_Export.fn_Config("export.Word." & mode & ".templateName")
        If Not m_pebWordTemplates.TryGetText(templateName, template, diagnostic) Then
            GoTo Failed
        End If
        If Not m_pebDslRenderer.TryRender(template, values, rendered, diagnostic) Then
            GoTo Failed
        End If
        ex_Core.fn_Diagnostic_WritePerf "Preview.TemplatesAndDsl", stageStartedAt
        If VBA.Len(rendered) = 0 Or VBA.Len(rendered) > 32767 Then
            Err.Raise 5, , "Word preview must contain between 1 and 32767 characters."
        End If
        values("WordText") = rendered
        If prepareOutput Then
            template = ex_PEB_Export.fn_Config("export.Word.output")
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

' //
' // Private
' //
Private Sub private_PrepareTicket(ByVal values As Object)
    Dim ticket As String
    Dim pattern As Object
    Dim qualified As String

    values("TicketDocument") = VBA.vbNullString
    If ex_PEB_Export.fn_Text(values("ModeForm")) <> ex_PEB_Export.fn_Config("export.ticket.form") Then Exit Sub
    ticket = ex_PEB_Export.fn_Text(values("TicketNumber"))
    If VBA.Len(ticket) = 0 Then Exit Sub
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

Private Sub private_PrepareMovementBasis(ByVal values As Object)
    Dim index As Long
    Dim source As String
    Dim template As String
    Dim rendered As String
    Dim diagnostic As String
    Dim prefix As String

    source = ex_PEB_Export.fn_Config("export.Movement.basis.lookupSource")
    For index = 1 To ex_PEB_Export.fn_Count("export." & source & ".lookups.count")
        private_Enrich values, "export." & source & ".lookups." & VBA.CStr(index)
    Next index
    private_AddConstants values
    private_AddAliases values
    Set m_pebWordTemplates = New obj_PEB_WordTemplates
    If Not m_pebWordTemplates.Initialize(diagnostic) Then Err.Raise 5, , diagnostic
    Set m_pebDslRenderer = New obj_PEB_DslRenderer
    If Not m_pebDslRenderer.Initialize(m_pebWordTemplates.SharedTemplates) Then Err.Raise 5, , "Cannot initialize DSL renderer."
    template = ex_PEB_Export.fn_Config("export.Movement.basis.template")
    If Not m_pebDslRenderer.TryRender(template, values, rendered, diagnostic) Then Err.Raise 5, , diagnostic
    prefix = ex_PEB_Export.fn_Config("export.Movement.basis.prefix")
    If VBA.Left$(rendered, VBA.Len(prefix)) <> prefix Then Err.Raise 5, , "Movement basis prefix is missing."
    values("MovementBasis") = VBA.Trim$(VBA.Mid$(rendered, VBA.Len(prefix) + 1))
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

Private Sub private_AddConstants(ByVal values As Object)
    Dim index As Long
    Dim prefix As String
    Dim name As String
    Dim kind As String
    Dim value As String
    Dim items As Collection

    For index = 1 To ex_PEB_Export.fn_Count("export.Word.constants.count")
        prefix = "export.Word.constants." & VBA.CStr(index)
        name = ex_PEB_Export.fn_Config(prefix & ".name")
        kind = ex_PEB_Export.fn_Config(prefix & ".kind")
        If values.Exists(name) Then
            Err.Raise 5, , "Constant would overwrite an export value: " & name
        End If
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

    For index = 1 To ex_PEB_Export.fn_Count("export.Word.aliases.count")
        prefix = "export.Word.aliases." & VBA.CStr(index)
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