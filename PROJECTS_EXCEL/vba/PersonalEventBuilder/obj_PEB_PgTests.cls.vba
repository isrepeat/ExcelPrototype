VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_PgTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IPage

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_controller As obj_PEB_PgTestsController

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    obj_IPage_Initialize = Me.Initialize(profileId)
End Function

Private Function obj_IPage_Render() As Boolean
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If Not m_controller.TryLoadCatalog(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
        Exit Function
    End If
    obj_IPage_Render = m_pageBase.Render()
    If obj_IPage_Render Then
        ex_PEB_Testing.fn_Attach m_controller
    End If
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If target Is Nothing Then
        Exit Function
    End If
    obj_IPage_HandleCellChange = True
End Function

Private Sub obj_IPage_Dispose()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal profileId As String) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then
        ex_WindowsUi.fn_ShowMessage ex_PEB_Export.fn_Config("text.TestOnly"), VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
        Exit Function
    End If
    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, ex_PEB_Export.fn_Config("ui.folder")) Then
        Exit Function
    End If
    If Not private_TryRegisterBindings() Then
        Exit Function
    End If
    Set m_controller = New obj_PEB_PgTestsController
    If Not m_controller.Initialize(m_pageBase) Then
        Exit Function
    End If
    If Not private_TryRegisterCommands() Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_controller Is Nothing Then
        m_controller.Dispose
    End If
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then
        m_pageBase.Dispose
    End If
    Set m_pageBase = Nothing
End Sub

' //
' // Private
' //
Private Function private_TryRegisterBindings() As Boolean
    Dim key As Variant

    For Each key In VBA.Array("TestsRun", "TestsStop", "TestsResults", "EnvironmentSwitch", "CheckboxChecked", "CheckboxUnchecked", "CheckboxPartial")
        If Not m_pageBase.BindingContext.SetValue("Text", VBA.CStr(key), ex_PEB_Export.fn_Config("text." & key)) Then
            Exit Function
        End If
    Next key
    If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatus", VBA.vbNullString) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatusStyle", "actionSuccess") Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Settings", "TestsRunning", False) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "Caption", VBA.vbNullString) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "AllC", False) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "AllP", False) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "PartialC", False) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "PartialP", False) Then
        Exit Function
    End If
    private_TryRegisterBindings = True
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim name As Variant
    Dim command As obj_UiCommand

    For Each name In VBA.Array("SelectionChanged", "TestRun", "CancelTests", "TestPreview", "EnvironmentSwitch")
        Set command = New obj_UiCommand
        If Not command.Initialize(m_controller, VBA.CStr(name) & "Handler") Then
            Exit Function
        End If
        If Not m_pageBase.BindingContext.SetObject("Commands", VBA.CStr(name), command) Then
            Exit Function
        End If
    Next name
    private_TryRegisterCommands = True
End Function