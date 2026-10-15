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
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
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

    On Error GoTo EH
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

    ' Обогащаем снимок только справочниками выбранного экспортера.
    For index = 1 To ex_PEB_Export.fn_Count("export." & exporter & ".lookups.count")
        private_Enrich values, "export." & exporter & ".lookups." & VBA.CStr(index)
    Next index
    If exporter = "Word" Then
        ' XML задаёт текст пунктов, конфигурация задаёт источники и соответствия полей.
        For index = 1 To ex_PEB_Export.fn_Count("export.Word." & mode & ".lookups.count")
            private_Enrich values, "export.Word." & mode & ".lookups." & VBA.CStr(index)
        Next index
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
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    Set values = Nothing
End Function

' //
' // Private
' //
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
    If Not tableQueryService.TryExecute(tableSource, tableQuery, dataTable, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If dataTable.RowCount <> 1 Then
        Err.Raise 5, , "Export lookup must return exactly one row: " & prefix
    End If
    values.Add outputField, dataTable.ValueAt(1, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, valueHeader))
    If VBA.Len(ex_PEB_Export.fn_Text(values(outputField))) = 0 Then
        Err.Raise 5, , "Export lookup returned an empty value: " & outputField
    End If
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