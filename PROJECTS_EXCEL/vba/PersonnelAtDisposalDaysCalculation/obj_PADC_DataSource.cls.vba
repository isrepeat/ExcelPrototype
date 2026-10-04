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
    Set m_parameterWorkbook = private_OpenParameterWorkbook(parameters.InputPath, m_openedHere)
    m_runContext.CheckCancel 0
    Set OpenParameterTable = private_ReadParameterTable(m_parameterWorkbook, parameters)
End Function

Public Function FindSource() As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    private_EnsureReady
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If ws.Name = m_configuration.GetText("legacy.SOURCE_SHEET_NAME") Then
                For Each lo In ws.ListObjects
                    If lo.Name = m_configuration.GetText("legacy.SOURCE_TABLE_NAME") Then
                        If Not FindSource Is Nothing Then
                            m_configuration.Fail m_configuration.GetText("legacy.MSG_MULTIPLE_SOURCES")
                        End If
                        Set FindSource = lo
                    End If
                Next lo
            End If
        Next ws
    Next wb
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
    For Each ws In workbook.Worksheets
        If ws.Name = sheetName Then
            For Each lo In ws.ListObjects
                If lo.Name = tableName Then
                    Set RequireTable = lo
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
    For Each col In lo.ListColumns
        If col.Name = header Then
            ColumnIndex = col.Index
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

    private_EnsureReady
    On Error GoTo Failed
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    Set tableQueryService = New obj_TableQueryService
    If Not tableSource.Initialize() Then
        m_configuration.Fail table.Name
    End If
    If Not tableQuery.Initialize() Then
        m_configuration.Fail table.Name
    End If
    If Not tableQueryService.Initialize() Then
        m_configuration.Fail table.Name
    End If

    tableSource.WorkbookPath = table.Parent.Parent.FullName
    tableSource.SheetName = table.Parent.Name
    tableSource.TableName = table.Name
    tableQueryService.Backend = QueryExcel
    If Not tableQueryService.TryExecute(tableSource, tableQuery, dataTable, diagnostic) Then
        m_configuration.Fail diagnostic
    End If
    ReadQueryTable = dataTable.Values
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
        VBA.Err.Raise errorNumber, "ReadQueryTable", errorDescription
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
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

    For Each workbook In Application.Workbooks
        If VBA.StrComp(workbook.FullName, filePath, vbTextCompare) = 0 Then
            Set private_OpenParameterWorkbook = workbook
            Exit Function
        End If
    Next workbook
    oldSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set private_OpenParameterWorkbook = Application.Workbooks.Open(Filename:=filePath, _
        UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    openedHere = True
    Application.AutomationSecurity = oldSecurity
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Application.AutomationSecurity = oldSecurity
    VBA.Err.Raise errorNumber, ERROR_SOURCE, errorText
End Function

Private Sub private_CloseParameterWorkbook(ByVal workbook As Workbook)
    On Error GoTo Failed
    workbook.Close SaveChanges:=False
    Exit Sub
Failed:
    VBA.MsgBox m_configuration.GetText("legacy.MSG_PARAMETER_CLOSE_FAILED") & VBA.Err.Description, vbExclamation
End Sub

Private Function private_ReadParameterTable( _
    ByVal workbook As Workbook, _
    ByVal parameters As obj_PADC_Parameters _
) As ListObject
    Dim ws As Worksheet
    Dim parameterSheet As Worksheet
    Dim headerCell As Range
    Dim table As ListObject

    For Each ws In workbook.Worksheets
        If VBA.StrComp(ws.Name, parameters.SheetName, vbTextCompare) = 0 Then
            Set parameterSheet = ws
            Exit For
        End If
    Next ws
    If parameterSheet Is Nothing Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_PARAMETER_SHEET_MISSING") & parameters.SheetName
    End If
    Set headerCell = parameterSheet.Range(parameters.HeaderAddress)
    For Each table In parameterSheet.ListObjects
        If Not table.HeaderRowRange Is Nothing Then
            If table.HeaderRowRange.Cells(1, 1).Address = headerCell.Address Then
                Set private_ReadParameterTable = table
                Exit Function
            End If
        End If
    Next table
    m_configuration.Fail m_configuration.GetText("legacy.MSG_PARAMETER_TABLE_MISSING") & parameters.SheetName & INPUT_SHEET_SEPARATOR & parameters.HeaderAddress
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_DataSource", "Service is not initialized."
    End If
End Sub