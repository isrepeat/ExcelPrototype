VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PADC_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_calculationRun As obj_PADC_CalculationRun
Private m_pageBase As obj_PageBase
Private m_resultTable As obj_UiRawTable
Private m_errorTable As obj_UiRawTable
Private m_calculateCommand As obj_UiCommand
Private m_cancelCommand As obj_UiCommand
Private m_refreshCommand As obj_UiCommand
Private m_exportCommand As obj_UiCommand
Private m_title As String

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
    Dim bindings As obj_UiBindingContext
    Dim textKey As Variant
    Dim configuredText As String
    Dim dateFormat As String
    Dim currentDate As Date
    Dim endDate As String
    Dim headers(0 To 5) As Variant
    Dim headerIndex As Long
    Dim errorCaption As String

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If

    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        ex_WindowsUi.fn_ShowMessage "The PersonnelAtDisposalDaysCalculation page controller requires a page base.", _
            VBA.vbExclamation, "PersonnelAtDisposalDaysCalculation"
        Exit Function
    End If

    Set bindings = m_pageBase.BindingContext
    If bindings Is Nothing Then
        GoTo Failed
    End If
    For Each textKey In VBA.Array("Calculate", "Cancel", "RefreshPage", "Export", "EndDateLabel")
        If Not private_TryReadConfiguration("text." & VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
        If Not bindings.SetValue("Text", VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
    Next textKey
    If Not private_TryReadConfiguration("text.Title", m_title) Then
        GoTo Failed
    End If
    If Not private_TryReadConfiguration("format.date", dateFormat) Then
        GoTo Failed
    End If
    currentDate = VBA.Date
    endDate = VBA.Format$(VBA.DateSerial(VBA.Year(currentDate), VBA.Month(currentDate), 1), dateFormat)
    If Not bindings.SetValue("Form", "EndDate", endDate) Then
        GoTo Failed
    End If

    For Each textKey In VBA.Array("FullName", "TaxId", "StartDate", _
            "Periods", "CountedDays", "ThresholdDate")
        If Not private_TryReadConfiguration("header." & VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
        headers(headerIndex) = configuredText
        headerIndex = headerIndex + 1
    Next textKey
    Set m_resultTable = New obj_UiRawTable
    If Not m_resultTable.InitializeEmpty(headers) Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Data", "CalculationRows", m_resultTable) Then
        GoTo Failed
    End If
    Set m_errorTable = New obj_UiRawTable
    If Not private_TryReadConfiguration("header.errors.Description", errorCaption) Then
        GoTo Failed
    End If
    If Not m_errorTable.InitializeEmpty(VBA.Array(headers(0), headers(1), errorCaption)) Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Data", "Errors", m_errorTable) Then
        GoTo Failed
    End If

    Set m_calculationRun = New obj_PADC_CalculationRun
    If Not m_calculationRun.Initialize(m_pageBase) Then
        GoTo Failed
    End If

    Set m_calculateCommand = New obj_UiCommand
    If Not m_calculateCommand.Initialize(Me, "CalculateHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "Calculate", m_calculateCommand) Then
        GoTo Failed
    End If
    Set m_cancelCommand = New obj_UiCommand
    If Not m_cancelCommand.Initialize(Me, "CancelHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "Cancel", m_cancelCommand) Then
        GoTo Failed
    End If
    Set m_refreshCommand = New obj_UiCommand
    If Not m_refreshCommand.Initialize(Me, "RefreshPageHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "RefreshPage", m_refreshCommand) Then
        GoTo Failed
    End If

    Set m_exportCommand = New obj_UiCommand
    If Not m_exportCommand.Initialize(Me, "ExportHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "Export", m_exportCommand) Then
        GoTo Failed
    End If

    ex_Core.fn_Diagnostic_WriteLog "PADC_BINDINGS_READY"
    Initialize = True
    m_isInitialized = Initialize
    Exit Function
Failed:
    ex_Core.fn_Diagnostic_WriteLog "PADC_BINDINGS_INITIALIZATION_FAILED"
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If

    m_isDisposed = True
    m_isInitialized = False
    If Not m_exportCommand Is Nothing Then
        m_exportCommand.Dispose
    End If
    Set m_exportCommand = Nothing
    If Not m_calculateCommand Is Nothing Then
        m_calculateCommand.Dispose
    End If
    If Not m_cancelCommand Is Nothing Then
        m_cancelCommand.Dispose
    End If
    If Not m_refreshCommand Is Nothing Then
        m_refreshCommand.Dispose
    End If
    If Not m_resultTable Is Nothing Then
        m_resultTable.Dispose
    End If
    If Not m_errorTable Is Nothing Then
        m_errorTable.Dispose
    End If
    If Not m_calculationRun Is Nothing Then
        m_calculationRun.Dispose
    End If
    Set m_calculationRun = Nothing
    Set m_errorTable = Nothing
    Set m_calculateCommand = Nothing
    Set m_cancelCommand = Nothing
    Set m_refreshCommand = Nothing
    Set m_resultTable = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function CalculateHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_calculationRun.Calculate
    CalculateHandler = True
End Function

Public Function RefreshPageHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    ex_Core.fn_Diagnostic_WriteLog "PADC_PAGE_REFRESH_REQUESTED"
    RefreshPageHandler = m_pageBase.UpdatePage()
End Function

Public Function ExportHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_calculationRun.ExportResult
    ExportHandler = True
End Function

Public Function CancelHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_calculationRun.RequestCancel
    CancelHandler = True
End Function

' //
' // Private
' //
Private Function private_TryReadConfiguration( _
    ByVal suffix As String, _
    ByRef value As String _
) As Boolean
    Dim key As String

    key = "PersonnelAtDisposalDaysCalculation::" & suffix
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
        ex_WindowsUi.fn_ShowMessage "Required configuration key not found: " & key, _
            vbExclamation, "Configuration"
        Exit Function
    End If
    private_TryReadConfiguration = True
End Function