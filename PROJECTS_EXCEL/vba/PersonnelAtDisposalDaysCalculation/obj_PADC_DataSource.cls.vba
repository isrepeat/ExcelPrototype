VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PADC_Configuration
Private m_runContext As obj_PADC_RunContext
Private Const ERROR_SOURCE As String = "PersonnelAtDisposalDays"

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
Public Function Initialize( _
    ByVal configuration As obj_PADC_Configuration, _
    ByVal runContext As obj_PADC_RunContext _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    If runContext Is Nothing Then
        Exit Function
    End If
    Set m_runContext = runContext
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_runContext = Nothing
    Set m_configuration = Nothing
End Sub

Public Function ReadParameters(ByVal parameters As obj_PADC_Parameters) As obj_DataTable
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim data As obj_DataTable
    Dim diagnostic As String
    Dim columnKey As Variant
    Dim errorNumber As Long
    Dim errorText As String

    private_EnsureReady
    On Error GoTo Failed
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not tableSource.Initialize() Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    If Not tableQuery.Initialize() Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    If Not service.Initialize() Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    tableSource.WorkbookPath = parameters.InputPath
    tableSource.SheetName = parameters.SheetName
    tableSource.RangeAddress = parameters.RangeAddress
    service.Backend = QueryAdo
    For Each columnKey In VBA.Array("TaxId", "FullName", "StartDate", "IncludeBusinessTripDays")
        tableQuery.AddColumn m_configuration.GetText("input.columns." & VBA.CStr(columnKey))
    Next columnKey
    m_runContext.CheckCancel 0
    If Not service.TryExecute(tableSource, tableQuery, data, diagnostic, False) Then
        m_configuration.Fail diagnostic
    End If
    Set ReadParameters = data
Cleanup:
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not tableQuery Is Nothing Then
        tableQuery.Dispose
    End If
    If Not tableSource Is Nothing Then
        tableSource.Dispose
    End If
    If errorNumber <> 0 Then
        If Not data Is Nothing Then
            data.Dispose
        End If
        VBA.Err.Raise errorNumber, ERROR_SOURCE, errorText
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Function

Public Function FindSource() As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim workbookIndex As Long
    Dim worksheetIndex As Long
    Dim sourceSheetName As String
    Dim sourceTableName As String

    private_EnsureReady
    private_Trace "SourceSearch.Configuration.Started"
    sourceSheetName = m_configuration.GetText("source.events.sheet")
    sourceTableName = m_configuration.GetText("source.events.table")
    private_Trace "SourceSearch.Started", "Sheet=" & sourceSheetName & " | Table=" & sourceTableName
    For Each wb In Application.Workbooks
        workbookIndex = workbookIndex + 1
        private_Trace "SourceSearch.Workbook.Started", "Index=" & workbookIndex
        private_Trace "SourceSearch.Workbook.Ready", "Workbook=" & wb.Name
        worksheetIndex = 0
        For Each ws In wb.Worksheets
            worksheetIndex = worksheetIndex + 1
            private_Trace "SourceSearch.Worksheet.Started", "Index=" & worksheetIndex
            private_Trace "SourceSearch.Worksheet.Ready", "Sheet=" & ws.Name
            If ws.Name = sourceSheetName Then
                For Each lo In ws.ListObjects
                    private_Trace "SourceSearch.Table.Started"
                    private_Trace "SourceSearch.Table.Ready", "Table=" & lo.Name
                    If lo.Name = sourceTableName Then
                        If Not FindSource Is Nothing Then
                            m_configuration.Fail m_configuration.GetText("text.MultipleSources")
                        End If
                        Set FindSource = lo
                        private_Trace "SourceSearch.Match", "Workbook=" & wb.Name & " | Table=" & lo.Name
                    End If
                Next lo
            End If
        Next ws
    Next wb
    private_Trace "SourceSearch.Completed", "Found=" & VBA.CStr(Not FindSource Is Nothing)
    If FindSource Is Nothing Then
        m_configuration.Fail m_configuration.GetText("text.OpenSource") & m_configuration.GetText("source.events.sheet") & m_configuration.GetText("text.Table") & m_configuration.GetText("source.events.table") & "'."
    End If
End Function

Public Function RequireTable( _
    ByVal sheetName As String, _
    ByVal tableName As String, _
    Optional ByVal workbook As Workbook = Nothing _
) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    private_EnsureReady
    If workbook Is Nothing Then
        Set workbook = ThisWorkbook
    End If
    private_Trace "RequiredTable.Started", "Sheet=" & sheetName & " | Table=" & tableName
    For Each ws In workbook.Worksheets
        private_Trace "RequiredTable.Worksheet", "Sheet=" & ws.Name
        If ws.Name = sheetName Then
            For Each lo In ws.ListObjects
                If lo.Name = tableName Then
                    Set RequireTable = lo
                    private_Trace "RequiredTable.Completed", "Table=" & lo.Name
                    Exit Function
                End If
            Next lo
        End If
    Next ws
    m_configuration.Fail m_configuration.GetText("text.TableNotFound") & tableName & m_configuration.GetText("text.OnWorksheet") & sheetName & m_configuration.GetText("text.CheckTableConfiguration")
End Function

Public Function ColumnIndex( _
    ByVal lo As ListObject, _
    ByVal header As String _
) As Long
    Dim col As ListColumn

    private_EnsureReady
    private_Trace "ColumnLookup.Started", "Header=" & header
    For Each col In lo.ListColumns
        If col.Name = header Then
            ColumnIndex = col.Index
            private_Trace "ColumnLookup.Completed", "Index=" & ColumnIndex
            Exit Function
        End If
    Next col
    m_configuration.Fail m_configuration.GetText("text.TableText") & lo.Name & m_configuration.GetText("text.IsMissingColumn") & header & "'."
End Function

Public Function ReadQueryTable(ByVal table As ListObject) As Variant
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim tableQueryService As obj_TableQueryService
    Dim dataTable As obj_DataTable
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim stageName As String
    Dim startedAt As Double

    private_EnsureReady
    On Error GoTo Failed
    startedAt = VBA.Timer
    stageName = "Query.CreateObjects"
    private_Trace stageName & ".Started"
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    Set tableQueryService = New obj_TableQueryService
    stageName = "Query.InitializeSource"
    private_Trace stageName & ".Started"
    If Not tableSource.Initialize() Then
        m_configuration.Fail table.Name
    End If
    stageName = "Query.InitializeDefinition"
    private_Trace stageName & ".Started"
    If Not tableQuery.Initialize() Then
        m_configuration.Fail table.Name
    End If
    stageName = "Query.InitializeService"
    private_Trace stageName & ".Started"
    If Not tableQueryService.Initialize() Then
        m_configuration.Fail table.Name
    End If

    stageName = "Query.DescribeSource"
    private_Trace stageName & ".Started"
    tableSource.WorkbookPath = table.Parent.Parent.FullName
    tableSource.SheetName = table.Parent.Name
    tableSource.TableName = table.Name
    tableQueryService.Backend = QueryExcel
    stageName = "Query.Execute"
    private_Trace stageName & ".Started", "Path=" & tableSource.WorkbookPath & _
        " | Sheet=" & tableSource.SheetName & " | Table=" & tableSource.TableName
    If Not tableQueryService.TryExecute(tableSource, tableQuery, dataTable, diagnostic) Then
        m_configuration.Fail diagnostic
    End If
    private_Trace stageName & ".Completed", "Rows=" & dataTable.RowCount
    stageName = "Query.CopyValues"
    private_Trace stageName & ".Started"
    ReadQueryTable = dataTable.Values
    private_Trace stageName & ".Completed"
Cleanup:
    On Error GoTo 0
    private_Trace "Query.Cleanup.Started"
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
    private_Trace "Query.Cleanup.Completed"
    ex_Core.fn_Diagnostic_WritePerf "PADC.ReadQueryTable", startedAt
    If errorNumber <> 0 Then
        VBA.Err.Raise errorNumber, "ReadQueryTable", errorDescription
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    private_Trace "Query.Failed", "Stage=" & stageName & " | Number=" & errorNumber & _
        " | Description=" & errorDescription
    Resume Cleanup
End Function

Public Function CompactParameterRows( _
    ByRef data As Variant, _
    ByRef sourceRows() As Long _
) As Long
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim hasValue As Boolean

    private_EnsureReady
    ReDim sourceRows(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        m_runContext.CheckCancel i
        hasValue = False
        For j = 1 To UBound(data, 2)
            If VBA.IsError(data(i, j)) Then
                hasValue = True
            ElseIf Not VBA.IsNull(data(i, j)) Then
                If VBA.Len(VBA.Trim$(VBA.CStr(data(i, j)))) > 0 Then
                    hasValue = True
                End If
            End If
        Next j
        If hasValue Then
            count = count + 1
            sourceRows(count) = i
            If count <> i Then
                For j = 1 To UBound(data, 2)
                    data(count, j) = data(i, j)
                Next j
            End If
        End If
    Next i
    CompactParameterRows = count
End Function

' //
' // Private
' //
Private Sub private_Trace( _
    ByVal stageName As String, _
    Optional ByVal details As String = vbNullString _
)
    ex_Core.fn_Diagnostic_WriteLog "PADC_STAGE | Name=" & stageName & " | " & details
    ex_Core.fn_Diagnostic_Flush
End Sub

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_DataSource", "Service is not initialized."
    End If
End Sub