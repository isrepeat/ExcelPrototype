VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExportEditingScen"
Option Explicit

Private Const ERROR_TITLE As String = "PrsnlEventBuilder / Редагування РУХ / WORD"
Private Const MOVEMENT_EVENTS_CONTROL_NAME As String = "MovementEventsMenu"

Private m_IsDisposed As Boolean
Private m_Page As obj_IPage
Private m_Controller As obj_PagePrsnlEvntBuilderCtrl

Public Function Initialize( _
    ByVal page As obj_IPage, _
    ByVal controller As obj_PagePrsnlEvntBuilderCtrl _
) As Boolean
    If page Is Nothing Or controller Is Nothing Then Exit Function
    m_IsDisposed = False
    Set m_Page = page
    Set m_Controller = controller
    Initialize = True
End Function

Public Function PrepareOpen() As Boolean
    If m_IsDisposed Or m_Controller Is Nothing Then Exit Function
    If Not m_Controller.HasResolvedOrderPair Then
        VBA.MsgBox "Спочатку прийміть номер і дату наказу.", _
            VBA.vbExclamation, "PrsnlEventBuilder / Наказ"
        Exit Function
    End If
    PrepareOpen = m_Controller.TryLoadExportEditingEvents(True, False)
End Function

Public Function Refresh() As Boolean
    Dim pageBase As obj_PageBase

    If m_IsDisposed Or m_Page Is Nothing Or m_Controller Is Nothing Then _
        Exit Function
    If Not m_Controller.TryLoadExportEditingEvents(True, False) Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.TryReflowControl(MOVEMENT_EVENTS_CONTROL_NAME) Then
        VBA.MsgBox "Не вдалося частково оновити таблицю редагування.", _
            VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If
    Refresh = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    Set m_Page = Nothing
    Set m_Controller = Nothing
End Sub

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub
