VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PG_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_calculationRun As obj_PG_CalculationRun
Private m_pageBase As obj_PageBase
Private m_resultTable As obj_UiRawTable
Private m_calculateCommand As obj_UiCommand
Private m_cancelCommand As obj_UiCommand
Private m_refreshCommand As obj_UiCommand
Private m_configuration As obj_PG_Configuration

' //
' // Жизненный цикл
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
    Dim headers(0 To 9) As Variant
    Dim headerIndex As Long

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If

    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        Exit Function
    End If

    Set m_configuration = New obj_PG_Configuration
    If Not m_configuration.Initialize() Then
        GoTo Failed
    End If
    Set bindings = m_pageBase.BindingContext
    If bindings Is Nothing Then
        GoTo Failed
    End If
    For Each textKey In VBA.Array("Calculate", "Cancel", "RefreshPage", "StartDateLabel", "EndDateLabel")
        If Not private_TryReadConfiguration("text." & VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
        If Not bindings.SetValue("Text", VBA.CStr(textKey), configuredText) Then
            GoTo Failed
        End If
    Next textKey
    If Not bindings.SetValue("Form", "StartDate", vbNullString) Then
        GoTo Failed
    End If
    If Not bindings.SetValue("Form", "EndDate", vbNullString) Then
        GoTo Failed
    End If

    For Each textKey In VBA.Array( _
        "TARGET_COL_RANK", _
        "TARGET_COL_NAME", _
        "TARGET_COL_TAX_ID", _
        "TARGET_COL_POSITION", _
        "TARGET_COL_EVENT", _
        "TARGET_COL_DEPARTURE_ORDER", _
        "TARGET_COL_PERIOD_FROM", _
        "TARGET_COL_PERIOD_TO", _
        "TARGET_COL_ARRIVAL_ORDER", _
        "TARGET_COL_PERIOD_COUNT" _
    )
        If Not private_TryReadConfiguration("legacy." & VBA.CStr(textKey), configuredText) Then
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
    Set m_calculationRun = New obj_PG_CalculationRun
    If Not m_calculationRun.Initialize(m_pageBase, m_configuration) Then
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

    ex_Core.fn_Diagnostic_WriteLog "PG_BINDINGS_READY"
    Initialize = True
    m_isInitialized = Initialize
    Exit Function
Failed:
    ex_Core.fn_Diagnostic_WriteLog "PG_BINDINGS_INITIALIZATION_FAILED"
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If

    m_isDisposed = True
    m_isInitialized = False
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
    If Not m_calculationRun Is Nothing Then
        m_calculationRun.Dispose
    End If
    Set m_calculationRun = Nothing
    Set m_calculateCommand = Nothing
    Set m_cancelCommand = Nothing
    Set m_refreshCommand = Nothing
    Set m_resultTable = Nothing
    If Not m_configuration Is Nothing Then
        m_configuration.Dispose
    End If
    Set m_configuration = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function CalculateHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    CalculateHandler = m_calculationRun.Calculate()
End Function

Public Function RefreshPageHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If m_calculationRun.IsRunning Then
        ex_WindowsUi.fn_ShowMessage m_configuration.GetText("message.AlreadyRunning"), vbExclamation
        Exit Function
    End If
    ex_Core.fn_Diagnostic_WriteLog "PG_PAGE_REFRESH_REQUESTED"
    RefreshPageHandler = m_pageBase.UpdatePage()
End Function

Public Function CancelHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_calculationRun.RequestCancel
    CancelHandler = True
End Function

' //
' // Вспомогательные методы
' //
Private Function private_TryReadConfiguration( _
    ByVal suffix As String, _
    ByRef value As String _
) As Boolean
    value = m_configuration.GetText(suffix)
    private_TryReadConfiguration = True
End Function