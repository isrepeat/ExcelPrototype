Attribute VB_Name = "ex_PEB_Testing"
Option Explicit

Private m_testUi As obj_PEB_TestUi
Private m_tickAt As Date
Private m_tickPending As Boolean
Private m_tickProcedure As String

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    fn_CancelTick
    If Not m_testUi Is Nothing Then m_testUi.Dispose
    Set m_testUi = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_IsRunning() As Boolean
    If Not m_testUi Is Nothing Then fn_IsRunning = m_testUi.IsRunning
End Function

Public Function fn_Open() As Boolean
    If m_testUi Is Nothing Then
        Set m_testUi = New obj_PEB_TestUi
        If Not m_testUi.Initialize() Then Exit Function
    End If
    fn_Open = m_testUi.OpenPage()
End Function

Public Sub fn_RequestStop()
    If m_testUi Is Nothing Then Exit Sub
    If Not m_testUi.IsRunning Then Exit Sub
    m_testUi.RequestStop
    fn_CancelTick
    fn_ScheduleTick True
End Sub

Public Sub fn_ScheduleTick(Optional ByVal immediate As Boolean = False)
    Dim delay As Long

    If m_testUi Is Nothing Or m_tickPending Then Exit Sub
    If Not m_testUi.IsRunning Then Exit Sub
    delay = ex_PEB_Export.fn_Count("test.runner.tickSeconds")
    If delay < 1 Then Err.Raise 5, , "test.runner.tickSeconds must be positive."
    If immediate Then delay = 0
    m_tickAt = VBA.Now + VBA.TimeSerial(0, 0, delay)
    m_tickProcedure = private_Callback()
    Application.OnTime m_tickAt, m_tickProcedure
    m_tickPending = True
End Sub

Public Sub fn_CancelTick()
    If Not m_tickPending Then Exit Sub
    On Error Resume Next
    Application.OnTime m_tickAt, m_tickProcedure, , False
    On Error GoTo 0
    m_tickPending = False
End Sub

Public Sub fn_Tick()
    Dim runtimeContext As Object
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean
    Dim entered As Boolean
    Dim more As Boolean
    Dim diagnostic As String

    If Not m_tickPending Then Exit Sub
    fn_CancelTick
    If m_testUi Is Nothing Then Exit Sub
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        fn_Module_Dispose
        Exit Sub
    End If
    entered = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    more = m_testUi.Advance()
    private_FlushMain
CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    If entered Then
        ex_RuntimeLifecycle.fn_Leave runtimeContext
        entered = False
    End If
    If more Then fn_ScheduleTick m_testUi.StopRequested
    If VBA.Len(diagnostic) > 0 Then ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    Exit Sub
EH:
    diagnostic = VBA.Err.Description
    more = False
    fn_Module_Dispose
    Resume CleanExit
End Sub

Public Sub fn_Click()
    Dim runtimeContext As Object
    Dim caller As String
    Dim entered As Boolean
    Dim previousEvents As Boolean

    previousEvents = Application.EnableEvents
    On Error GoTo EH
    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    Application.EnableEvents = False
    caller = VBA.CStr(Application.Caller)
    If caller = "pebt_Stop" Then
        fn_RequestStop
        private_FlushMain
        GoTo CleanExit
    End If
    If m_testUi Is Nothing Then
        fn_Open
        GoTo CleanExit
    End If
    If m_testUi.IsRunning Then GoTo CleanExit
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then GoTo CleanExit
    entered = True
    Select Case caller
        Case "pebt_All"
            m_testUi.SelectAll True
        Case "pebt_None"
            m_testUi.SelectAll False
        Case "pebt_Run"
            m_testUi.RunSelected
        Case "pebt_Refresh"
            m_testUi.OpenPage
        Case Else
            If VBA.Left$(caller, 6) <> "pebt_C" Then Err.Raise 5, , "Unknown test UI command: " & caller
            m_testUi.ToggleCase VBA.CLng(VBA.Mid$(caller, 7))
    End Select
    private_FlushMain
CleanExit:
    Application.EnableEvents = previousEvents
    If entered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Callback() As String
    private_Callback = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_PEB_Testing.fn_Tick"
End Function

Private Sub private_FlushMain()
    Dim sheet As Worksheet
    Dim context As obj_UiRenderContext
    Dim diagnostic As String

    Set sheet = ThisWorkbook.Worksheets(ex_PEB_Export.fn_Config("test.ui.mainSheet"))
    If ex_UiRuntime.fn_TryGetContext(sheet, context) Then
        If Not context.FlushLayout(diagnostic) Then Err.Raise 5, , diagnostic
    End If
End Sub