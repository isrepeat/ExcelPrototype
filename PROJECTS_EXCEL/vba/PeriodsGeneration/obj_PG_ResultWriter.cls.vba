VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_ResultWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
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
Public Function Initialize( _
    ByVal pageBase As obj_PageBase, _
    ByVal configuration As obj_PG_Configuration _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If pageBase Is Nothing Or configuration Is Nothing Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    Set m_configuration = configuration
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_pageBase = Nothing
    Set m_configuration = Nothing
End Sub

Public Sub Publish( _
    ByVal values As Variant, _
    ByVal count As Long _
)
    Dim resultTable As obj_UiRawTable
    Dim headers As Variant
    Dim oldScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    private_EnsureReady
    headers = VBA.Array( _
        m_configuration.GetText("legacy.TARGET_COL_RANK"), _
        m_configuration.GetText("legacy.TARGET_COL_NAME"), _
        m_configuration.GetText("legacy.TARGET_COL_TAX_ID"), _
        m_configuration.GetText("legacy.TARGET_COL_POSITION"), _
        m_configuration.GetText("legacy.TARGET_COL_EVENT"), _
        m_configuration.GetText("legacy.TARGET_COL_DEPARTURE_ORDER"), _
        m_configuration.GetText("legacy.TARGET_COL_PERIOD_FROM"), _
        m_configuration.GetText("legacy.TARGET_COL_PERIOD_TO"), _
        m_configuration.GetText("legacy.TARGET_COL_ARRIVAL_ORDER"), _
        m_configuration.GetText("legacy.TARGET_COL_PERIOD_COUNT") _
    )
    Set resultTable = New obj_UiRawTable
    If count > 0 Then
        If Not resultTable.Initialize(values, headers) Then
            m_configuration.Fail m_configuration.GetText("message.PublishFailed")
        End If
    Else
        If Not resultTable.InitializeEmpty(headers) Then
            m_configuration.Fail m_configuration.GetText("message.PublishFailed")
        End If
    End If
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    If Not m_pageBase.BindingContext.SetObject("Data", "CalculationRows", resultTable) Then
        m_configuration.Fail m_configuration.GetText("message.PublishFailed")
    End If
    If Not m_pageBase.Render() Then
        m_configuration.Fail m_configuration.GetText("message.PublishFailed")
    End If
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Application.ScreenUpdating = oldScreenUpdating
    VBA.Err.Raise errorNumber, "obj_PG_ResultWriter", errorText
End Sub

' //
' // Вспомогательные методы
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_ResultWriter", "Service is not initialized."
    End If
End Sub