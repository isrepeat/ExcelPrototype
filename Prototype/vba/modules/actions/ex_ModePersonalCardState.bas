Attribute VB_Name = "ex_ModePersonalCardState"
Option Explicit

Private Const SOURCE_VALIDATION_MODE As String = "postprocess.validationmode"
Private Const DEFAULT_VALIDATION_MODE As String = "Enabled"
Private Const VALUE_ENABLED As String = "enabled"
Private Const VALUE_DISABLED As String = "disabled"

Private g_ValidationModeBySheet As Object

Public Function m_IsValidationModeSource(ByVal toggleSource As String) As Boolean
    m_IsValidationModeSource = (StrComp(mp_NormalizeSource(toggleSource), SOURCE_VALIDATION_MODE, vbTextCompare) = 0)
End Function

Public Function m_GetValidationMode( _
    Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim normalizedDefault As String
    Dim candidateValue As String
    Dim cacheKey As String
    Dim cache As Object

    normalizedDefault = mp_NormalizeValidationMode(defaultValue, False)
    If Len(normalizedDefault) = 0 Then normalizedDefault = DEFAULT_VALIDATION_MODE

    cacheKey = mp_BuildSheetKey(ws)
    If Len(cacheKey) = 0 Then
        m_GetValidationMode = normalizedDefault
        Exit Function
    End If

    Set cache = g_ValidationModeBySheet
    If Not cache Is Nothing Then
        If cache.Exists(cacheKey) Then
            candidateValue = CStr(cache(cacheKey))
            m_GetValidationMode = mp_NormalizeValidationMode(candidateValue, True)
            Exit Function
        End If
    End If

    m_GetValidationMode = normalizedDefault
End Function

Public Sub m_SetValidationMode( _
    ByVal valueText As String, _
    Optional ByVal ws As Worksheet = Nothing _
)
    Dim cacheKey As String
    Dim normalizedValue As String
    Dim cache As Object

    cacheKey = mp_BuildSheetKey(ws)
    If Len(cacheKey) = 0 Then Exit Sub

    normalizedValue = mp_NormalizeValidationMode(valueText, True)
    Set cache = mp_EnsureCache()
    cache(cacheKey) = normalizedValue
End Sub

Public Function m_IsValidationDisabled( _
    Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal defaultDisabled As Boolean = False _
) As Boolean
    Dim defaultMode As String

    If defaultDisabled Then
        defaultMode = "Disabled"
    Else
        defaultMode = "Enabled"
    End If

    m_IsValidationDisabled = (StrComp(m_GetValidationMode(ws, defaultMode), "Disabled", vbTextCompare) = 0)
End Function

Private Function mp_EnsureCache() As Object
    If g_ValidationModeBySheet Is Nothing Then
        Set g_ValidationModeBySheet = CreateObject("Scripting.Dictionary")
        g_ValidationModeBySheet.CompareMode = 1
    End If
    Set mp_EnsureCache = g_ValidationModeBySheet
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

Private Function mp_NormalizeValidationMode(ByVal valueText As String, ByVal useDefaultIfUnknown As Boolean) As String
    Dim normalized As String

    normalized = LCase$(Trim$(CStr(valueText)))
    Select Case normalized
        Case VALUE_ENABLED
            mp_NormalizeValidationMode = "Enabled"
        Case VALUE_DISABLED
            mp_NormalizeValidationMode = "Disabled"
        Case Else
            If useDefaultIfUnknown Then
                mp_NormalizeValidationMode = DEFAULT_VALIDATION_MODE
            Else
                mp_NormalizeValidationMode = vbNullString
            End If
    End Select
End Function

Private Function mp_NormalizeSource(ByVal sourceText As String) As String
    mp_NormalizeSource = LCase$(Trim$(CStr(sourceText)))
End Function
