VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_profileId As String
Private m_uiFolderRelativePath As String
Private m_uiBindingContext As obj_UiBindingContext

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get ProfileId() As String
    ProfileId = m_profileId
End Property

Public Property Get BindingContext() As obj_UiBindingContext
    Set BindingContext = m_uiBindingContext
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal profileId As String, _
    ByVal uiFolderRelativePath As String _
) As Boolean
    m_profileId = VBA.Trim$(profileId)
    m_uiFolderRelativePath = VBA.Trim$(uiFolderRelativePath)
    If VBA.Len(m_profileId) = 0 Or VBA.Len(m_uiFolderRelativePath) = 0 Then
        VBA.MsgBox "A page requires a profile and UI folder.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If

    Set m_uiBindingContext = New obj_UiBindingContext
    If Not m_uiBindingContext.Initialize() Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    If Not m_uiBindingContext Is Nothing Then m_uiBindingContext.Dispose
    Set m_uiBindingContext = Nothing
    m_profileId = VBA.vbNullString
    m_uiFolderRelativePath = VBA.vbNullString
End Sub

Public Function Render() As Boolean
    If m_uiBindingContext Is Nothing Then Exit Function
    ex_UiRenderer.fn_RenderPages m_uiFolderRelativePath, m_uiBindingContext
    Render = True
End Function

Public Function RenderActivePage() As Boolean
    If m_uiBindingContext Is Nothing Then Exit Function
    ex_UiRenderer.fn_RenderActivePage m_uiFolderRelativePath, m_uiBindingContext
    RenderActivePage = True
End Function

Public Function UpdatePage() As Boolean
    Dim previousScreenUpdating As Boolean

    If m_uiBindingContext Is Nothing Then Exit Function
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    Application.ScreenUpdating = False
    UpdatePage = Me.RenderActivePage()
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Exit Function
EH:
    VBA.MsgBox "The page cannot be updated: " & VBA.Err.Description, _
        VBA.vbExclamation, "Page update"
    Resume CleanExit
End Function