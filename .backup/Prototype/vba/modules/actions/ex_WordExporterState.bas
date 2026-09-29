Attribute VB_Name = "ex_WordExporterState"
Option Explicit

Private Const SOURCE_EXPORT_INSERT_MODE As String = "export.insertmode"
Private Const DEFAULT_INSERT_MODE As String = "AppendToBottom"
Private Const VALUE_APPEND_TO_TOP As String = "appendtotop"
Private Const VALUE_APPEND_TO_BOTTOM As String = "appendtobottom"
Private Const VALUE_REPLACE_ALL As String = "replaceall"

Private g_InsertModeBySheet As Object

Public Function m_IsInsertModeSource(ByVal toggleSource As String) As Boolean
    m_IsInsertModeSource = (StrComp(mp_NormalizeSource(toggleSource), SOURCE_EXPORT_INSERT_MODE, vbTextCompare) = 0)
End Function

Public Function m_GetInsertMode( _
    Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim normalizedDefault As String
    Dim candidateValue As String
    Dim cacheKey As String
    Dim cache As Object

    normalizedDefault = mp_NormalizeInsertMode(defaultValue, False)
    If Len(normalizedDefault) = 0 Then normalizedDefault = DEFAULT_INSERT_MODE

    cacheKey = mp_BuildSheetKey(ws)
    If Len(cacheKey) = 0 Then
        m_GetInsertMode = normalizedDefault
        Exit Function
    End If

    Set cache = g_InsertModeBySheet
    If Not cache Is Nothing Then
        If cache.Exists(cacheKey) Then
            candidateValue = CStr(cache(cacheKey))
            m_GetInsertMode = mp_NormalizeInsertMode(candidateValue, True)
            Exit Function
        End If
    End If

    m_GetInsertMode = normalizedDefault
End Function

Public Sub m_SetInsertMode( _
    ByVal valueText As String, _
    Optional ByVal ws As Worksheet = Nothing _
)
    Dim cacheKey As String
    Dim normalizedValue As String
    Dim cache As Object

    cacheKey = mp_BuildSheetKey(ws)
    If Len(cacheKey) = 0 Then Exit Sub

    normalizedValue = mp_NormalizeInsertMode(valueText, True)
    Set cache = mp_EnsureCache()
    cache(cacheKey) = normalizedValue
End Sub

Private Function mp_EnsureCache() As Object
    If g_InsertModeBySheet Is Nothing Then
        Set g_InsertModeBySheet = CreateObject("Scripting.Dictionary")
        g_InsertModeBySheet.CompareMode = 1
    End If
    Set mp_EnsureCache = g_InsertModeBySheet
End Function

Private Function mp_BuildSheetKey(ByVal ws As Worksheet) As String
    Dim sheetKey As String

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveSheet
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Function

    sheetKey = Trim$(CStr(ws.CodeName))
    If Len(sheetKey) = 0 Then sheetKey = Trim$(CStr(ws.Name))
    If Len(sheetKey) = 0 Then Exit Function

    mp_BuildSheetKey = LCase$(sheetKey)
End Function

Private Function mp_NormalizeInsertMode(ByVal valueText As String, ByVal useDefaultIfUnknown As Boolean) As String
    Dim normalized As String

    normalized = LCase$(Trim$(CStr(valueText)))
    Select Case normalized
        Case VALUE_APPEND_TO_TOP
            mp_NormalizeInsertMode = "AppendToTop"
        Case VALUE_APPEND_TO_BOTTOM
            mp_NormalizeInsertMode = "AppendToBottom"
        Case VALUE_REPLACE_ALL
            mp_NormalizeInsertMode = "ReplaceAll"
        Case Else
            If useDefaultIfUnknown Then
                mp_NormalizeInsertMode = DEFAULT_INSERT_MODE
            Else
                mp_NormalizeInsertMode = vbNullString
            End If
    End Select
End Function

Private Function mp_NormalizeSource(ByVal sourceText As String) As String
    mp_NormalizeSource = LCase$(Trim$(CStr(sourceText)))
End Function
