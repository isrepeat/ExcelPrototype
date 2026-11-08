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
Private m_tvoRun As obj_PADC_TvoRun
Private m_disposalRows As Object
Private m_tvoRows As Object
Private m_pageBase As obj_PageBase
Private m_resultTable As obj_UiRawTable
Private m_calculateCommand As obj_UiCommand
Private m_cancelCommand As obj_UiCommand
Private m_exportCommand As obj_UiCommand
Private m_disposalModeCommand As obj_UiCommand
Private m_tvoModeCommand As obj_UiCommand
Private m_tvoMode As Boolean
Private m_running As Boolean
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
    Dim headers(0 To 12) As Variant
    Dim headerIndex As Long

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If

    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        ex_Core.fn_StatusBar_Show "The PersonnelAtDisposalDaysCalculation page controller requires a page base."
        Exit Function
    End If

    Set bindings = m_pageBase.BindingContext
    If bindings Is Nothing Then
        GoTo Failed
    End If
    For Each textKey In VBA.Array("Calculate", "Cancel", "Export", "EndDateLabel", "LastStatus", "DisposalMode", "TvoMode")
        If Not private_TryReadConfiguration("text." & VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
        If Not bindings.SetValue("Text", VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
    Next textKey
    If Not private_TryReadConfiguration("status.Ready", configuredText) Then
        GoTo Failed
    End If
    If Not bindings.SetValue("Form", "ActionStatus", configuredText) Then
        GoTo Failed
    End If
    If Not bindings.SetValue("Form", "ActionStatusStyle", "statusNormal") Then
        GoTo Failed
    End If
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

    For Each textKey In VBA.Array("header.FullName", "header.TaxId", "export.header.Status", _
            "export.header.OpenStatus", "header.StartDate", _
            "header.Periods", "header.CountedDays", "header.ThresholdDate", _
            "export.header.From", "export.header.Through", "export.header.Period", "export.header.Include", "header.Notes")
        If Not private_TryReadConfiguration(VBA.CStr(textKey), configuredText) Then
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
    Set m_calculationRun = New obj_PADC_CalculationRun
    If Not m_calculationRun.Initialize(m_pageBase) Then
        GoTo Failed
    End If
    Set m_tvoRun = New obj_PADC_TvoRun
    If Not m_tvoRun.Initialize(m_pageBase) Then
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
    Set m_exportCommand = New obj_UiCommand
    If Not m_exportCommand.Initialize(Me, "ExportHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "Export", m_exportCommand) Then
        GoTo Failed
    End If
    Set m_disposalModeCommand = New obj_UiCommand
    If Not m_disposalModeCommand.Initialize(Me, "DisposalModeHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "DisposalMode", m_disposalModeCommand) Then
        GoTo Failed
    End If
    Set m_tvoModeCommand = New obj_UiCommand
    If Not m_tvoModeCommand.Initialize(Me, "TvoModeHandler") Then
        GoTo Failed
    End If
    If Not bindings.SetObject("Commands", "TvoMode", m_tvoModeCommand) Then
        GoTo Failed
    End If
    If Not bindings.SetValue("Form", "DisposalModeStyle", "activeModeButton") Then
        GoTo Failed
    End If
    If Not bindings.SetValue("Form", "TvoModeStyle", "secondaryButton") Then
        GoTo Failed
    End If

    m_isInitialized = True
    If Not private_SelectMode(True, False) Then
        GoTo Failed
    End If
    ex_Core.fn_Diagnostic_WriteLog "PADC_BINDINGS_READY"
    Initialize = True
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
    If Not m_tvoRun Is Nothing Then
        m_tvoRun.Dispose
    End If
    Set m_tvoRun = Nothing
    Set m_disposalRows = Nothing
    Set m_tvoRows = Nothing
    If Not m_disposalModeCommand Is Nothing Then
        m_disposalModeCommand.Dispose
    End If
    If Not m_tvoModeCommand Is Nothing Then
        m_tvoModeCommand.Dispose
    End If
    Set m_disposalModeCommand = Nothing
    Set m_tvoModeCommand = Nothing
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
    If Not m_resultTable Is Nothing Then
        m_resultTable.Dispose
    End If
    If Not m_calculationRun Is Nothing Then
        m_calculationRun.Dispose
    End If
    Set m_calculationRun = Nothing
    Set m_calculateCommand = Nothing
    Set m_cancelCommand = Nothing
    Set m_resultTable = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function CalculateHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If m_running Then
        Exit Function
    End If
    On Error GoTo Failed
    m_running = True
    If m_tvoMode Then
        m_tvoRun.Calculate
    Else
        m_calculationRun.Calculate
    End If
    m_running = False
    CalculateHandler = True
    Exit Function
Failed:
    m_running = False
    ex_Core.fn_StatusBar_Show VBA.Err.Description
End Function

Public Function ExportHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If m_tvoMode Then
        m_tvoRun.ExportResult
        ExportHandler = True
        Exit Function
    End If
    m_calculationRun.ExportResult
    ExportHandler = True
End Function

Public Function CancelHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If m_tvoMode Then
        m_tvoRun.RequestCancel
    Else
        m_calculationRun.RequestCancel
    End If
    CancelHandler = True
End Function

Public Function DisposalModeHandler() As Boolean
    DisposalModeHandler = private_SelectMode(False)
End Function

Public Function TvoModeHandler() As Boolean
    TvoModeHandler = private_SelectMode(True)
End Function

' //
' // Private
' //
Private Function private_SelectMode( _
    ByVal tvoMode As Boolean, _
    Optional ByVal updatePage As Boolean = True _
) As Boolean
    Dim statusText As String
    Dim disposalStyle As String
    Dim tvoStyle As String
    Dim value As Variant
    Dim currentRows As Object
    Dim isObject As Boolean
    Dim headers(0 To 15) As Variant
    Dim headerKeys As Variant
    Dim headerIndex As Long
    Dim emptyTable As obj_UiRawTable

    If Not m_isInitialized Or m_isDisposed Or m_running Then
        Exit Function
    End If
    disposalStyle = "activeModeButton"
    tvoStyle = "secondaryButton"
    If tvoMode Then
        disposalStyle = "secondaryButton"
        tvoStyle = "activeModeButton"
        If Not private_TryReadConfiguration("tvo.text.Ready", statusText) Then
            Exit Function
        End If
    Else
        If Not private_TryReadConfiguration("status.Ready", statusText) Then
            Exit Function
        End If
    End If
    If tvoMode <> m_tvoMode Then
        If Not m_pageBase.BindingContext.TryGetValue("Data", "CalculationRows", value, currentRows, isObject) Then
            Exit Function
        End If
        If Not isObject Then
            Exit Function
        End If
        If tvoMode Then
            Set m_disposalRows = currentRows
            If m_tvoRows Is Nothing Then
                headerKeys = VBA.Array("tvo.header.Rank", "header.FullName", "header.TaxId", "tvo.header.OwnPosition", "tvo.header.Position", _
                    "tvo.header.Duration", "tvo.header.AccessOrder", "tvo.header.StartOrder", _
                    "tvo.header.StartDate", "tvo.header.ThresholdDate", "tvo.header.PlannedEndDate", _
                    "tvo.header.EndDate", "tvo.header.EndOrder", "tvo.header.StreakStart", "tvo.header.Days", "header.Periods")
                For headerIndex = 0 To 15
                    If Not private_TryReadConfiguration(VBA.CStr(headerKeys(headerIndex)), statusText) Then
                        Exit Function
                    End If
                    headers(headerIndex) = statusText
                Next headerIndex
                Set emptyTable = New obj_UiRawTable
                If Not emptyTable.InitializeEmpty(headers) Then
                    Exit Function
                End If
                Set m_tvoRows = emptyTable
                If Not private_TryReadConfiguration("tvo.text.Ready", statusText) Then
                    Exit Function
                End If
            End If
            If Not m_pageBase.BindingContext.SetObject("Data", "CalculationRows", m_tvoRows) Then
                Exit Function
            End If
        Else
            Set m_tvoRows = currentRows
            If Not m_pageBase.BindingContext.SetObject("Data", "CalculationRows", m_disposalRows) Then
                Exit Function
            End If
        End If
    End If
    m_tvoMode = tvoMode
    If Not m_pageBase.BindingContext.SetValue("Form", "DisposalModeStyle", disposalStyle) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "TvoModeStyle", tvoStyle) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatusStyle", "statusNormal") Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatus", statusText) Then
        Exit Function
    End If
    If updatePage Then
        private_SelectMode = m_pageBase.UpdatePage()
    Else
        private_SelectMode = True
    End If
End Function

Private Function private_TryReadConfiguration( _
    ByVal suffix As String, _
    ByRef value As String _
) As Boolean
    Dim key As String

    key = "PersonnelAtDisposalDaysCalculation::" & suffix
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
        ex_Core.fn_StatusBar_Show "Required configuration key not found: " & key
        Exit Function
    End If
    private_TryReadConfiguration = True
End Function