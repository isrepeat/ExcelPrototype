Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function private_RedrawWindow Lib "user32" Alias "RedrawWindow" ( _
    ByVal windowHandle As LongPtr, _
    ByVal updateRect As LongPtr, _
    ByVal updateRegion As LongPtr, _
    ByVal flags As Long _
) As Long
Private Declare PtrSafe Function private_GetCursor Lib "user32" Alias "GetCursor" () As LongPtr
Private Declare PtrSafe Function private_LoadCursor Lib "user32" Alias "LoadCursorW" ( _
    ByVal instance As LongPtr, _
    ByVal cursorName As LongPtr _
) As LongPtr
Private Declare PtrSafe Function private_SetCursor Lib "user32" Alias "SetCursor" (ByVal cursor As LongPtr) As LongPtr
Private m_previousNativeCursor As LongPtr
#Else
Private Declare Function private_RedrawWindow Lib "user32" Alias "RedrawWindow" ( _
    ByVal windowHandle As Long, _
    ByVal updateRect As Long, _
    ByVal updateRegion As Long, _
    ByVal flags As Long _
) As Long
Private Declare Function private_GetCursor Lib "user32" Alias "GetCursor" () As Long
Private Declare Function private_LoadCursor Lib "user32" Alias "LoadCursorW" ( _
    ByVal instance As Long, _
    ByVal cursorName As Long _
) As Long
Private Declare Function private_SetCursor Lib "user32" Alias "SetCursor" (ByVal cursor As Long) As Long
Private m_previousNativeCursor As Long
#End If

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
        m_previousNativeCursor = private_GetCursor()
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

Public Sub fn_End(Optional ByVal redraw As Boolean = False)
    Const RDW_INVALIDATE As Long = &H1
    Const RDW_ALLCHILDREN As Long = &H80
    Const RDW_UPDATENOW As Long = &H100

    If m_depth <= 0 Then Exit Sub

    If redraw And m_depth = 1 And Application.ScreenUpdating Then
        fn_Refresh
        ' Complete the Excel window paint while the wait cursor is still active.
        private_RedrawWindow Application.Hwnd, 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
    End If
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
    ' Update the visible cursor without pumping messages or allowing command reentry.
    private_SetCursor private_LoadCursor(0, 32514)
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
    ' Do not restore a transient Excel wait cursor after completing our own busy scope.
    If m_previousCursorCaptured And m_previousCursor = xlDefault Then
        If m_previousNativeCursor = private_LoadCursor(0, 32514) Or _
                m_previousNativeCursor = private_LoadCursor(0, 32650) Then
            m_previousNativeCursor = private_LoadCursor(0, 32512)
        End If
    End If
    If m_previousCursorCaptured Then Application.Cursor = m_previousCursor
    If m_previousNativeCursor <> 0 Then
        private_SetCursor m_previousNativeCursor
    End If
    m_previousNativeCursor = 0
    m_previousCursorCaptured = False
    If m_statusBarShown Then ex_Core.fn_StatusBar_Clear
    m_statusBarShown = False
    Err.Clear
    On Error GoTo 0
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------