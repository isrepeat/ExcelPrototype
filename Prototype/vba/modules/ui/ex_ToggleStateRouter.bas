Attribute VB_Name = "ex_ToggleStateRouter"
Option Explicit

' Thin routing wrapper for toggle state sources.
' Source ownership:
' - Export.InsertMode -> ex_WordExporterState
' - PostProcess.ValidationMode -> ex_ModePersonalCardState

Private Const SOURCE_EXPORT_INSERT_MODE As String = "export.insertmode"
Private Const SOURCE_POSTPROCESS_VALIDATION_MODE As String = "postprocess.validationmode"

Public Function m_IsRuntimeSource(ByVal toggleSource As String) As Boolean
    toggleSource = mp_NormalizeToggleSource(toggleSource)
    Select Case toggleSource
        Case SOURCE_EXPORT_INSERT_MODE
            m_IsRuntimeSource = ex_WordExporterState.m_IsInsertModeSource(toggleSource)
        Case SOURCE_POSTPROCESS_VALIDATION_MODE
            m_IsRuntimeSource = ex_ModePersonalCardState.m_IsValidationModeSource(toggleSource)
        Case Else
            m_IsRuntimeSource = False
    End Select
End Function

Public Function m_GetToggleValue( _
    ByVal toggleSource As String, _
    Optional ByVal defaultValue As String = vbNullString, _
    Optional ByVal ws As Worksheet = Nothing _
) As String
    Dim normalizedSource As String

    normalizedSource = mp_NormalizeToggleSource(toggleSource)
    If Len(normalizedSource) = 0 Then
        m_GetToggleValue = Trim$(CStr(defaultValue))
        Exit Function
    End If

    Select Case normalizedSource
        Case SOURCE_EXPORT_INSERT_MODE
            m_GetToggleValue = ex_WordExporterState.m_GetInsertMode(ws, CStr(defaultValue))
        Case SOURCE_POSTPROCESS_VALIDATION_MODE
            m_GetToggleValue = ex_ModePersonalCardState.m_GetValidationMode(ws, CStr(defaultValue))
        Case Else
            m_GetToggleValue = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue(toggleSource, defaultValue)))
    End Select
End Function

Public Sub m_SetToggleValue( _
    ByVal toggleSource As String, _
    ByVal valueText As String, _
    Optional ByVal ws As Worksheet = Nothing _
)
    Dim normalizedSource As String

    normalizedSource = mp_NormalizeToggleSource(toggleSource)
    If Len(normalizedSource) = 0 Then Exit Sub

    Select Case normalizedSource
        Case SOURCE_EXPORT_INSERT_MODE
            ex_WordExporterState.m_SetInsertMode valueText, ws
        Case SOURCE_POSTPROCESS_VALIDATION_MODE
            ex_ModePersonalCardState.m_SetValidationMode valueText, ws
        Case Else
            ex_ConfigProvider.m_SetConfigValue toggleSource, valueText, True
    End Select
End Sub

Private Function mp_NormalizeToggleSource(ByVal toggleSource As String) As String
    mp_NormalizeToggleSource = LCase$(Trim$(CStr(toggleSource)))
End Function
