Attribute VB_Name = "ex_SheetViewZoom"
Option Explicit

Private Const DEFAULT_PROFILE_ATTR As String = "resultZoom"
Private Const ZOOM_MIN As Long = 10
Private Const ZOOM_MAX As Long = 400

Private g_ZoomCacheBySheet As Object

Public Sub m_ApplyProfileZoomForResultSheet( _
    ByVal ws As Worksheet, _
    ByVal sheetExistedBeforeRender As Boolean, _
    Optional ByVal profileAttrName As String = DEFAULT_PROFILE_ATTR, _
    Optional ByVal defaultZoom As Long = 115 _
)
    Dim zoomValue As Long

    If ws Is Nothing Then Exit Sub

    ws.Activate

    If sheetExistedBeforeRender Then
        ' Existing sheet keeps its own view zoom; cache is only a fallback.
        zoomValue = mp_ReadActiveWindowZoom()
        If zoomValue <= 0 Then
            If mp_TryGetCachedZoomBySheetName(ws.Name, zoomValue) Then
                On Error Resume Next
                ActiveWindow.Zoom = zoomValue
                On Error GoTo 0
                zoomValue = mp_ReadActiveWindowZoom()
            End If
        End If
    Else
        zoomValue = mp_GetProfileZoom(profileAttrName, defaultZoom)
        On Error Resume Next
        ActiveWindow.Zoom = zoomValue
        On Error GoTo 0
        zoomValue = mp_ReadActiveWindowZoom()
    End If

    If zoomValue > 0 Then
        mp_SetCachedZoomBySheetName ws.Name, zoomValue
    End If
End Sub

Public Sub m_ResetZoomCache(Optional ByVal sheetName As String = vbNullString)
    Dim normalizedSheetName As String

    normalizedSheetName = LCase$(Trim$(sheetName))
    If Len(normalizedSheetName) = 0 Then
        Set g_ZoomCacheBySheet = Nothing
        Exit Sub
    End If

    If g_ZoomCacheBySheet Is Nothing Then Exit Sub
    If g_ZoomCacheBySheet.Exists(normalizedSheetName) Then
        g_ZoomCacheBySheet.Remove normalizedSheetName
    End If
End Sub

Private Function mp_GetProfileZoom(ByVal profileAttrName As String, ByVal defaultZoom As Long) As Long
    Dim zoomText As String
    Dim zoomValue As Long

    profileAttrName = Trim$(profileAttrName)
    If Len(profileAttrName) = 0 Then profileAttrName = DEFAULT_PROFILE_ATTR

    zoomText = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileAttribute(profileAttrName, CStr(defaultZoom), ws_Dev))
    If Not mp_TryParseZoomValue(zoomText, zoomValue) Then
        zoomValue = defaultZoom
    End If

    mp_GetProfileZoom = mp_NormalizeZoomValue(zoomValue)
End Function

Private Function mp_ReadActiveWindowZoom() As Long
    Dim zoomValue As Long

    On Error Resume Next
    zoomValue = CLng(ActiveWindow.Zoom)
    If Err.Number <> 0 Then
        Err.Clear
        zoomValue = 0
    End If
    On Error GoTo 0

    If zoomValue <= 0 Then Exit Function
    mp_ReadActiveWindowZoom = mp_NormalizeZoomValue(zoomValue)
End Function

Private Function mp_TryParseZoomValue(ByVal rawText As String, ByRef outZoom As Long) As Boolean
    On Error GoTo EH

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then Exit Function
    If Not IsNumeric(rawText) Then Exit Function

    outZoom = CLng(CDbl(rawText))
    If outZoom <= 0 Then Exit Function
    outZoom = mp_NormalizeZoomValue(outZoom)
    mp_TryParseZoomValue = True
    Exit Function

EH:
    outZoom = 0
End Function

Private Function mp_NormalizeZoomValue(ByVal zoomValue As Long) As Long
    If zoomValue < ZOOM_MIN Then zoomValue = ZOOM_MIN
    If zoomValue > ZOOM_MAX Then zoomValue = ZOOM_MAX
    mp_NormalizeZoomValue = zoomValue
End Function

Private Sub mp_EnsureZoomCacheBySheet()
    If Not g_ZoomCacheBySheet Is Nothing Then Exit Sub

    Set g_ZoomCacheBySheet = CreateObject("Scripting.Dictionary")
    g_ZoomCacheBySheet.CompareMode = 1
End Sub

Private Sub mp_SetCachedZoomBySheetName(ByVal sheetName As String, ByVal zoomValue As Long)
    sheetName = LCase$(Trim$(sheetName))
    If Len(sheetName) = 0 Then Exit Sub

    zoomValue = mp_NormalizeZoomValue(zoomValue)
    If zoomValue <= 0 Then Exit Sub

    mp_EnsureZoomCacheBySheet
    g_ZoomCacheBySheet(sheetName) = zoomValue
End Sub

Private Function mp_TryGetCachedZoomBySheetName(ByVal sheetName As String, ByRef outZoom As Long) As Boolean
    On Error GoTo EH

    sheetName = LCase$(Trim$(sheetName))
    If Len(sheetName) = 0 Then Exit Function
    If g_ZoomCacheBySheet Is Nothing Then Exit Function
    If Not g_ZoomCacheBySheet.Exists(sheetName) Then Exit Function

    outZoom = CLng(g_ZoomCacheBySheet(sheetName))
    outZoom = mp_NormalizeZoomValue(outZoom)
    mp_TryGetCachedZoomBySheetName = (outZoom > 0)
    Exit Function

EH:
    outZoom = 0
End Function
