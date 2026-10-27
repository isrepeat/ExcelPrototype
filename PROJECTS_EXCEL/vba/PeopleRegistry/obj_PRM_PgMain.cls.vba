VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PRM_PgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IPage

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_commands As Collection

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_commands = New Collection
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal profileId As String) As Boolean
    Dim bindings As obj_UiBindingContext
    Dim key As Variant
    Dim text As String
    Dim uiCommand As obj_UiCommand

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo Failed
    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, ex_PRM.fn_Config("ui.folder")) Then GoTo Failed
    Set bindings = m_pageBase.BindingContext

    ' Load captions independently of the layout and field values.
    For Each key In Array("NewId", "Origin", "Register", "InputHelp", "PersonName", _
            "ReplaceId", "InputTitle", "PreviousId", "ResultHelp", "ReplaceHelp", _
            "ResultTitle", "RegistryPath", "ReplaceTitle", "LayoutReady")
        text = ex_PRM.fn_Config("text." & CStr(key))
        If CStr(key) = "RegistryPath" Then
            text = Replace(text, "{path}", ex_PRM.fn_Config("registry.path"))
        End If
        If Not bindings.SetValue("Text", CStr(key), text) Then GoTo Failed
    Next key
    For Each key In Array("PersonName", "Origin", "PreviousId", "NewId")
        If Not bindings.SetValue("Form", CStr(key), vbNullString) Then GoTo Failed
    Next key
    If Not bindings.SetValue("Form", "Result", ex_PRM.fn_Config("text.LayoutReady")) Then GoTo Failed

    ' Expose explicit commands without calling the legacy in-workbook registry.
    For Each key In Array("Register", "ReplaceId")
        Set uiCommand = New obj_UiCommand
        If Not uiCommand.Initialize(Me, "ShowLayoutStatus") Then GoTo Failed
        m_commands.Add uiCommand
        If Not bindings.SetObject("Commands", CStr(key), uiCommand) Then GoTo Failed
    Next key
    m_isInitialized = True
    Initialize = True
    Exit Function
Failed:
    If Err.Number <> 0 Then ex_WindowsUi.fn_ShowMessage Err.Description, vbExclamation
    Me.Dispose
End Function

Public Sub Dispose()
    Dim uiCommand As obj_UiCommand

    If m_isDisposed Then Exit Sub
    m_isInitialized = False
    m_isDisposed = True
    For Each uiCommand In m_commands
        uiCommand.Dispose
    Next uiCommand
    Set m_commands = Nothing
    If Not m_pageBase Is Nothing Then m_pageBase.Dispose
    Set m_pageBase = Nothing
End Sub

Public Function ShowLayoutStatus() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ex_WindowsUi.fn_ShowMessage ex_PRM.fn_Config("text.LayoutReady"), vbInformation
    ShowLayoutStatus = True
End Function

' //
' // Interface
' //
Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    obj_IPage_Initialize = Me.Initialize(profileId)
End Function

Private Sub obj_IPage_Dispose()
    Me.Dispose
End Sub

Private Function obj_IPage_Render() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    obj_IPage_Render = m_pageBase.Render()
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If target Is Nothing Then Exit Function
    obj_IPage_HandleCellChange = True
End Function