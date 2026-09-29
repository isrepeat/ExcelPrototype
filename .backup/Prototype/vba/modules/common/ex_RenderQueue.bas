Attribute VB_Name = "ex_RenderQueue"
Option Explicit

Private g_QueueBySheet As Object
Private g_ActiveBySheet As Object

Public Sub m_BeginForSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    mp_EnsureStores
    m_ClearQueueForSheet ws
    m_SetActiveForSheet ws, True
End Sub

Public Sub m_EndForSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    m_ClearQueueForSheet ws
    m_SetActiveForSheet ws, False
End Sub

Public Sub m_ClearAll()
    Set g_QueueBySheet = Nothing
    Set g_ActiveBySheet = Nothing
End Sub

Public Function m_GetOrCreateQueueForSheet(ByVal ws As Worksheet) As Collection
    Dim sheetKey As String
    Dim newQueue As Collection

    If ws Is Nothing Then Exit Function
    mp_EnsureStores
    sheetKey = mp_BuildSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function

    If Not g_QueueBySheet.Exists(sheetKey) Then
        Set newQueue = New Collection
        g_QueueBySheet.Add sheetKey, newQueue
    End If
    Set m_GetOrCreateQueueForSheet = g_QueueBySheet(sheetKey)
End Function

Public Sub m_ClearQueueForSheet(ByVal ws As Worksheet)
    Dim sheetKey As String

    If ws Is Nothing Then Exit Sub
    If g_QueueBySheet Is Nothing Then Exit Sub
    sheetKey = mp_BuildSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If g_QueueBySheet.Exists(sheetKey) Then g_QueueBySheet.Remove sheetKey
End Sub

Public Sub m_SetActiveForSheet(ByVal ws As Worksheet, ByVal isActive As Boolean)
    Dim sheetKey As String

    If ws Is Nothing Then Exit Sub
    mp_EnsureStores
    sheetKey = mp_BuildSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If isActive Then
        g_ActiveBySheet(sheetKey) = "1"
    ElseIf g_ActiveBySheet.Exists(sheetKey) Then
        g_ActiveBySheet.Remove sheetKey
    End If
End Sub

Public Function m_IsActiveForSheet(ByVal ws As Worksheet) As Boolean
    Dim sheetKey As String

    If ws Is Nothing Then Exit Function
    If g_ActiveBySheet Is Nothing Then Exit Function
    sheetKey = mp_BuildSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function
    If Not g_ActiveBySheet.Exists(sheetKey) Then Exit Function

    m_IsActiveForSheet = (CStr(g_ActiveBySheet(sheetKey)) = "1")
End Function

Private Sub mp_EnsureStores()
    If g_QueueBySheet Is Nothing Then
        Set g_QueueBySheet = CreateObject("Scripting.Dictionary")
        g_QueueBySheet.CompareMode = 1 ' vbTextCompare
    End If
    If g_ActiveBySheet Is Nothing Then
        Set g_ActiveBySheet = CreateObject("Scripting.Dictionary")
        g_ActiveBySheet.CompareMode = 1 ' vbTextCompare
    End If
End Sub

Private Function mp_BuildSheetKey(ByVal ws As Worksheet) As String
    Dim wb As Workbook

    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set wb = ws.Parent
    On Error GoTo 0
    If wb Is Nothing Then Exit Function

    mp_BuildSheetKey = wb.Name & "|" & ws.Name
End Function

