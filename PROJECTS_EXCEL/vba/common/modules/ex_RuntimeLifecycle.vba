Option Explicit

Private m_context As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Function fn_PrepareReload() As Boolean
    Dim context As Object

    Set context = fn_Context()
    On Error GoTo EH_PREPARE
    If Not context("StopRequested") Or VBA.CLng(context("ActiveCalls")) <> 0 Then _
        VBA.Err.Raise VBA.vbObjectError + 2204, "fn_PrepareReload", "Runtime is not quiescent."

    ' Сначала отключаем внешние точки входа, затем освобождаем корни объектов.
    ex_AppHotkeys.fn_PrepareReload
    If Not ex_Core.fn_Diagnostic_Flush() Then _
        VBA.Err.Raise VBA.vbObjectError + 2205, "fn_PrepareReload", "Diagnostic log flush failed."
    ex_UiPageManager.fn_Module_Dispose
    ex_UiBindings.fn_Module_Dispose
    ex_UiRuntime.fn_Module_Dispose
    ex_StylePipeline.fn_Module_Dispose
    ex_RuntimePaths.fn_Module_Dispose
    ex_AppHotkeys.fn_Module_Dispose
    ex_Core.fn_Module_Dispose
    context("Phase") = "Prepared"
    fn_PrepareReload = True
    Exit Function
EH_PREPARE:
    context("Error") = VBA.Err.Description
    context("Phase") = "Faulted"
End Function

Public Function fn_InitializeReloaded(ByVal context As Object, ByVal uiFolder As String) As Boolean
    On Error GoTo EH_INITIALIZE
    fn_AttachContext context
    If context("Phase") <> "Initializing" Then _
        VBA.Err.Raise VBA.vbObjectError + 2206, "fn_InitializeReloaded", "Unexpected initialization phase."
    ex_RuntimePaths.fn_SetUiFolder uiFolder
    If Not ex_PersonalEventBuilder.fn_Initialize() Then _
        VBA.Err.Raise VBA.vbObjectError + 2207, "fn_InitializeReloaded", "Runtime initialization failed."
    ex_AppHotkeys.fn_Activate
    fn_InitializeReloaded = True
    Exit Function
EH_INITIALIZE:
    context("Error") = VBA.Err.Description
    context("Phase") = "Faulted"
End Function
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Context() As Object
    Dim blockedMarker As Name
    Dim updater As Workbook

    If m_context Is Nothing Then
        On Error Resume Next
        Set updater = Application.Workbooks("WorkbookUpdater.xlam")
        On Error GoTo 0
        If Not updater Is Nothing Then
            Set m_context = Application.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_FindContext", ThisWorkbook)
        End If
    End If
    If m_context Is Nothing Then
        Set m_context = VBA.CreateObject("Scripting.Dictionary")
        m_context.Add "StopRequested", False
        m_context.Add "ActiveCalls", 0&
        m_context.Add "Phase", "Running"
        m_context.Add "Generation", 0&
        ' Маркер книги сохраняет запрет запуска даже при сбросе VBA-переменных.
        On Error Resume Next
        Set blockedMarker = ThisWorkbook.Names("_RuntimeReloadBlocked")
        On Error GoTo 0
        If Not blockedMarker Is Nothing Then
            m_context("StopRequested") = True
            m_context("Phase") = "Faulted"
        End If
    End If
    Set fn_Context = m_context
End Function

Public Function fn_TryEnter(ByRef context As Object) As Boolean
    Set context = fn_Context()
    If context("StopRequested") Or context("Phase") <> "Running" Then Exit Function
    context("ActiveCalls") = VBA.CLng(context("ActiveCalls")) + 1
    fn_TryEnter = True
End Function

Public Sub fn_Leave(ByVal context As Object)
    If context Is Nothing Then _
        VBA.Err.Raise VBA.vbObjectError + 2200, "fn_Leave", "Missing runtime context."
    If VBA.CLng(context("ActiveCalls")) <= 0 Then _
        VBA.Err.Raise VBA.vbObjectError + 2201, "fn_Leave", "Unbalanced runtime entry."
    context("ActiveCalls") = VBA.CLng(context("ActiveCalls")) - 1
End Sub

Public Function fn_StopRequested(ByVal context As Object) As Boolean
    If context Is Nothing Then _
        VBA.Err.Raise VBA.vbObjectError + 2202, "fn_StopRequested", "Missing runtime context."
    fn_StopRequested = VBA.CBool(context("StopRequested"))
End Function

Public Sub fn_AttachContext(ByVal context As Object)
    If context Is Nothing Then _
        VBA.Err.Raise VBA.vbObjectError + 2203, "fn_AttachContext", "Missing runtime context."
    Set m_context = context
End Sub

' --------------------------------------
' } // namespace API
' --------------------------------------