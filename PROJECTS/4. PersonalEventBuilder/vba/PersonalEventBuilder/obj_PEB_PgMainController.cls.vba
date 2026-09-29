VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pageBase As obj_PageBase

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize(ByVal pageBase As obj_PageBase) As Boolean
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        VBA.MsgBox "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_Initialize = True
End Function

Public Function fn_HelloWorld() As Boolean
    ex_Core.fn_Diagnostic_WriteLog "HELLO_WORLD_CLICKED"
    VBA.MsgBox "Hello World from PersonalEventBuilder.", VBA.vbInformation, "PersonalEventBuilder"
    fn_HelloWorld = True
End Function

Public Function fn_UpdatePage() As Boolean
    Dim previousScreenUpdating As Boolean

    If m_pageBase Is Nothing Then Exit Function
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    Application.ScreenUpdating = False
    fn_UpdatePage = m_pageBase.fn_RenderActivePage()
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Exit Function
EH:
    VBA.MsgBox "The page cannot be updated: " & VBA.Err.Description, _
        VBA.vbExclamation, "PersonalEventBuilder"
    Resume CleanExit
End Function

Public Sub fn_Dispose()
    Set m_pageBase = Nothing
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------