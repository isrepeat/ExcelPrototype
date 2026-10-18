Option Explicit

Private m_depth As Long
Private m_previousCursor As XlMousePointer
Private m_previousCursorCaptured As Boolean
Private m_statusBarShown As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    fn_Reset
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Begin(Optional ByVal statusText As String = VBA.vbNullString) As Boolean
    On Error Resume Next
    If m_depth = 0 Then
        m_previousCursorCaptured = False
        m_statusBarShown = False
        Err.Clear
        m_previousCursor = Application.Cursor
        m_previousCursorCaptured = (Err.Number = 0)
        Err.Clear
        Application.Cursor = xlWait
        Err.Clear
        If VBA.Len(VBA.Trim$(statusText)) > 0 Then
            ex_Core.fn_StatusBar_Show statusText
            m_statusBarShown = (Err.Number = 0)
            Err.Clear
        End If
    End If
    m_depth = m_depth + 1
    fn_Refresh
    fn_Begin = True
    Err.Clear
    On Error GoTo 0
End Function

Public Sub fn_End()
    If m_depth <= 0 Then Exit Sub

    m_depth = m_depth - 1
    If m_depth <> 0 Then
        fn_Refresh
        Exit Sub
    End If
    private_Restore
End Sub

Public Sub fn_Refresh()
    If m_depth <= 0 Then Exit Sub
    ' Message pumping and external automation can replace the displayed cursor.
    On Error Resume Next
    Application.Cursor = xlWait
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub fn_Reset()
    m_depth = 0
    private_Restore
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_Restore()
    On Error Resume Next
    If m_previousCursorCaptured Then Application.Cursor = m_previousCursor
    m_previousCursorCaptured = False
    If m_statusBarShown Then ex_Core.fn_StatusBar_Clear
    m_statusBarShown = False
    Err.Clear
    On Error GoTo 0
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------