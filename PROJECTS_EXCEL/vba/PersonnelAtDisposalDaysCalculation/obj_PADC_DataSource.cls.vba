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

Public Function ReadConfiguredSource( _
    ByVal prefix As String, _
    ByVal workbookPath As String, _
    ByVal columnKeys As Variant, _
    Optional ByVal sheetName As String, _
    Optional ByVal rangeAddress As String _
) As obj_DataTable
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
    If Not tableSource.Initialize() Or Not tableQuery.Initialize() Or Not service.Initialize() Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    If VBA.Len(workbookPath) = 0 Then
        workbookPath = VBA.CreateObject("WScript.Shell").ExpandEnvironmentStrings(m_configuration.GetText(prefix & ".path"))
    End If
    If VBA.Len(sheetName) = 0 Then
        sheetName = m_configuration.GetText(prefix & ".sheet")
    End If
    If VBA.Len(rangeAddress) = 0 Then
        rangeAddress = m_configuration.GetText(prefix & ".range")
    End If
    tableSource.WorkbookPath = workbookPath
    tableSource.SheetName = sheetName
    tableSource.RangeAddress = rangeAddress
    service.Backend = QueryAdo
    For Each columnKey In columnKeys
        tableQuery.AddColumn m_configuration.GetText(prefix & ".columns." & VBA.CStr(columnKey))
    Next columnKey
    m_runContext.CheckCancel 0
    If Not service.TryExecute(tableSource, tableQuery, data, diagnostic, False) Then
        m_configuration.Fail workbookPath & " | " & sheetName & "!" & rangeAddress & " | " & diagnostic
    End If
    m_runContext.CheckCancel 0
    Set ReadConfiguredSource = data
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

' //
' // Private
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_DataSource", "Service is not initialized."
    End If
End Sub