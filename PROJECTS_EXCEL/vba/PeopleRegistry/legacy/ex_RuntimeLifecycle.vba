Attribute VB_Name = "ex_RuntimeLifecycle"
Option Explicit

Private m_context As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Context() As Object
    Dim marker As Name
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
        Set m_context = CreateObject("Scripting.Dictionary")
        m_context.Add "StopRequested", False
        m_context.Add "ActiveCalls", 0&
        m_context.Add "Phase", "Running"
        m_context.Add "Generation", 0&
        On Error Resume Next
        Set marker = ThisWorkbook.Names("_RuntimeReloadBlocked")
        On Error GoTo 0
        If Not marker Is Nothing Then
            m_context("StopRequested") = True
            m_context("Phase") = "Faulted"
        End If
    End If
    Set fn_Context = m_context
End Function

Public Function fn_TryEnter(ByRef context As Object) As Boolean
    Set context = fn_Context()
    If context("StopRequested") Or context("Phase") <> "Running" Then
        Exit Function
    End If
    context("ActiveCalls") = CLng(context("ActiveCalls")) + 1
    fn_TryEnter = True
End Function

Public Sub fn_Leave(ByVal context As Object)
    If context Is Nothing Then
        Err.Raise vbObjectError + 2200, , private_Message("MissingContext")
    End If
    If CLng(context("ActiveCalls")) <= 0 Then
        Err.Raise vbObjectError + 2201, , private_Message("UnbalancedCalls")
    End If
    context("ActiveCalls") = CLng(context("ActiveCalls")) - 1
End Sub

Public Sub fn_AttachContext(ByVal context As Object)
    If context Is Nothing Then
        Err.Raise vbObjectError + 2203, , private_Message("MissingContext")
    End If
    Set m_context = context
End Sub

Public Function fn_PrepareReload() As Boolean
    Dim context As Object

    Set context = fn_Context()
    On Error GoTo Failed
    If Not context("StopRequested") Or CLng(context("ActiveCalls")) <> 0 Then
        Err.Raise vbObjectError + 2204, , private_Message("RuntimeBusy")
    End If
    ' Workbook events use the entry guard; no hotkeys or persistent services need disposal.
    context("Phase") = "Prepared"
    fn_PrepareReload = True
    Exit Function
Failed:
    context("Error") = Err.Description
    context("Phase") = "Faulted"
End Function

Public Function fn_InitializeReloaded( _
    ByVal context As Object, _
    ByVal uiFolder As String _
) As Boolean
    On Error GoTo Failed
    fn_AttachContext context
    If context("Phase") <> "Initializing" Then
        Err.Raise vbObjectError + 2206, , private_Message("InvalidPhase")
    End If
    ' The updater passes a UI folder for every profile; this profile has no UI resources.
    If Not ex_PeopleRegistry.fn_Initialize() Then
        Err.Raise vbObjectError + 2207, , private_Message("InitializeFailed")
    End If
    fn_InitializeReloaded = True
    Exit Function
Failed:
    context("Error") = Err.Description
    context("Phase") = "Faulted"
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Message(ByVal key As String) As String
    Dim item As ListRow

    For Each item In ThisWorkbook.Worksheets("wsConfig").ListObjects("tbConfig").ListRows
        If CStr(item.Range.Cells(1, 2).Value2) = "PeopleRegistry::Message." & key Then
            private_Message = CStr(item.Range.Cells(1, 3).Value2)
            If Len(private_Message) > 0 Then
                Exit Function
            End If
            Exit For
        End If
    Next item
    Err.Raise vbObjectError + 2050, , "Missing configuration key: PeopleRegistry::Message." & key
End Function