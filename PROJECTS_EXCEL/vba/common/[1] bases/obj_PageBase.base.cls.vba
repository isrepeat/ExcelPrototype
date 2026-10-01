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

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize( _
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
    If Not m_uiBindingContext.fn_Initialize() Then Exit Function
    fn_Initialize = True
End Function

Public Function fn_Render() As Boolean
    If m_uiBindingContext Is Nothing Then Exit Function
    ex_UiRenderer.fn_RenderPages m_uiFolderRelativePath, m_uiBindingContext
    fn_Render = True
End Function

Public Function fn_RenderActivePage() As Boolean
    If m_uiBindingContext Is Nothing Then Exit Function
    ex_UiRenderer.fn_RenderActivePage m_uiFolderRelativePath, m_uiBindingContext
    fn_RenderActivePage = True
End Function

Public Function fn_UpdatePage() As Boolean
    Dim previousScreenUpdating As Boolean

    If m_uiBindingContext Is Nothing Then Exit Function
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    Application.ScreenUpdating = False
    fn_UpdatePage = fn_RenderActivePage()
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Exit Function
EH:
    VBA.MsgBox "The page cannot be updated: " & VBA.Err.Description, _
        VBA.vbExclamation, "Page update"
    Resume CleanExit
End Function

Public Property Get fn_ProfileId() As String
    fn_ProfileId = m_profileId
End Property

Public Property Get fn_BindingContext() As obj_UiBindingContext
    Set fn_BindingContext = m_uiBindingContext
End Property

Public Sub fn_Dispose()
    If Not m_uiBindingContext Is Nothing Then m_uiBindingContext.fn_Dispose
    Set m_uiBindingContext = Nothing
    m_profileId = VBA.vbNullString
    m_uiFolderRelativePath = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------