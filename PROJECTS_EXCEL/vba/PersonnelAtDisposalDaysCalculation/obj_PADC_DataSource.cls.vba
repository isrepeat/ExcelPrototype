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
Private m_parameterWorkbook As Workbook
Private m_openedHere As Boolean
Private Const INPUT_SHEET_SEPARATOR As String = "!"
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
    If m_openedHere And Not m_parameterWorkbook Is Nothing Then
        private_CloseParameterWorkbook m_parameterWorkbook
    End If
    Set m_parameterWorkbook = Nothing
    Set m_runContext = Nothing
    Set m_configuration = Nothing
End Sub

Public Function OpenParameterTable(ByVal parameters As obj_PADC_Parameters) As ListObject
    private_EnsureReady
    private_Trace "ParameterTable.Started"
    Set m_parameterWorkbook = private_OpenParameterWorkbook(parameters.InputPath, m_openedHere)
    private_Trace "ParameterWorkbook.Ready", "OpenedHere=" & VBA.CStr(m_openedHere)
    m_runContext.CheckCancel 0
    private_Trace "ParameterTable.Lookup.Started", "Sheet=" & parameters.SheetName & _
        " | Header=" & parameters.HeaderAddress
    Set OpenParameterTable = private_ReadParameterTable(m_parameterWorkbook, parameters)
    private_Trace "ParameterTable.Lookup.Completed", "Table=" & OpenParameterTable.Name
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
    sourceSheetName = m_configuration.GetText("legacy.SOURCE_SHEET_NAME")
    sourceTableName = m_configuration.GetText("legacy.SOURCE_TABLE_NAME")
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
                            m_configuration.Fail m_configuration.GetText("legacy.MSG_MULTIPLE_SOURCES")
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
        m_configuration.Fail m_configuration.GetText("legacy.MSG_OPEN_SOURCE") & m_configuration.GetText("legacy.SOURCE_SHEET_NAME") & m_configuration.GetText("legacy.MSG_TABLE") & m_configuration.GetText("legacy.SOURCE_TABLE_NAME") & "'."
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
    m_configuration.Fail m_configuration.GetText("legacy.MSG_TABLE_NOT_FOUND") & tableName & m_configuration.GetText("legacy.MSG_ON_WORKSHEET") & sheetName & m_configuration.GetText("legacy.MSG_CHECK_TABLE_CONFIGURATION")
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
    m_configuration.Fail m_configuration.GetText("legacy.MSG_TABLE_TEXT") & lo.Name & m_configuration.GetText("legacy.MSG_IS_MISSING_COLUMN") & header & "'."
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
            ElseIf VBA.Len(VBA.Trim$(VBA.CStr(data(i, j)))) > 0 Then
                hasValue = True
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
Private Function private_OpenParameterWorkbook( _
    ByVal filePath As String, _
    ByRef openedHere As Boolean _
) As Workbook
    Dim workbook As Workbook
    Dim oldSecurity As Long
    Dim errorNumber As Long
    Dim errorText As String

    private_Trace "ParameterWorkbook.SearchOpen.Started", "Path=" & filePath
    For Each workbook In Application.Workbooks
        private_Trace "ParameterWorkbook.SearchOpen.Candidate.Started"
        private_Trace "ParameterWorkbook.SearchOpen.Candidate.Ready", "Path=" & workbook.FullName
        If VBA.StrComp(workbook.FullName, filePath, vbTextCompare) = 0 Then
            Set private_OpenParameterWorkbook = workbook
            private_Trace "ParameterWorkbook.Reused", "Path=" & filePath
            Exit Function
        End If
    Next workbook
    oldSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    private_Trace "ParameterWorkbook.Open.Started", "Path=" & filePath
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set private_OpenParameterWorkbook = Application.Workbooks.Open(Filename:=filePath, _
        UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    openedHere = True
    Application.AutomationSecurity = oldSecurity
    private_Trace "ParameterWorkbook.Open.Completed", "Path=" & filePath
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Application.AutomationSecurity = oldSecurity
    private_Trace "ParameterWorkbook.Open.Failed", "Number=" & errorNumber & " | Description=" & errorText
    VBA.Err.Raise errorNumber, ERROR_SOURCE, errorText
End Function

Private Sub private_CloseParameterWorkbook(ByVal workbook As Workbook)
    On Error GoTo Failed
    private_Trace "ParameterWorkbook.Close.Started"
    workbook.Close SaveChanges:=False
    private_Trace "ParameterWorkbook.Close.Completed"
    Exit Sub
Failed:
    ex_WindowsUi.fn_ShowMessage m_configuration.GetText("legacy.MSG_PARAMETER_CLOSE_FAILED") & VBA.Err.Description, vbExclamation
End Sub

Private Function private_ReadParameterTable( _
    ByVal workbook As Workbook, _
    ByVal parameters As obj_PADC_Parameters _
) As ListObject
    Dim ws As Worksheet
    Dim parameterSheet As Worksheet
    Dim headerCell As Range
    Dim table As ListObject

    private_Trace "ParameterTable.EnumerateWorksheets.Started"
    For Each ws In workbook.Worksheets
        private_Trace "ParameterTable.Worksheet", "Sheet=" & ws.Name
        If VBA.StrComp(ws.Name, parameters.SheetName, vbTextCompare) = 0 Then
            Set parameterSheet = ws
            Exit For
        End If
    Next ws
    If parameterSheet Is Nothing Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_PARAMETER_SHEET_MISSING") & parameters.SheetName
    End If
    private_Trace "ParameterTable.ResolveHeader.Started", "Address=" & parameters.HeaderAddress
    Set headerCell = parameterSheet.Range(parameters.HeaderAddress)
    private_Trace "ParameterTable.ResolveHeader.Completed"
    For Each table In parameterSheet.ListObjects
        private_Trace "ParameterTable.Candidate", "Table=" & table.Name
        If Not table.HeaderRowRange Is Nothing Then
            If table.HeaderRowRange.Cells(1, 1).Address = headerCell.Address Then
                Set private_ReadParameterTable = table
                Exit Function
            End If
        End If
    Next table
    m_configuration.Fail m_configuration.GetText("legacy.MSG_PARAMETER_TABLE_MISSING") & parameters.SheetName & INPUT_SHEET_SEPARATOR & parameters.HeaderAddress
End Function

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