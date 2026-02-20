Attribute VB_Name = "ex_UIActions"
Option Explicit

' UI entrypoints layer: keeps user-triggered callbacks in actions/*
' and delegates work to domain/config modules.

Public Sub m_DeleteResultSheets_OnClick()
    ex_SheetStylesXmlProvider.m_DeleteResultSheets
End Sub

Public Sub m_SwitchMode_OnClick()
    ex_Settings.m_SwitchMode_OnClick
End Sub

Public Sub m_OnProfileChanged_OnClick()
    ex_ConfigProfilesManager.m_OnProfileChanged
End Sub

Public Sub m_OnModeChanged_OnClick()
    ex_ConfigProfilesManager.m_OnModeChanged
End Sub

Public Sub m_HelloWorld_OnClick()
    ex_Startup.m_HelloWorld
End Sub

Public Sub m_ShowPersonalCard_OnClick()
    ex_PersonTimeline.m_ShowPersonTimeline_UI
End Sub

Public Sub m_RunComparingTables_OnClick()
    ex_TableComparing.m_RunComparing
End Sub

Public Sub m_OutputPanelStartSearch_OnClick()
    Dim ws As Worksheet
    Dim searchKey As String
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle
    Dim configKey As String
    Dim callerName As String
    Dim fieldIndex As Long
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2401, "ex_UIActions.m_OutputPanelStartSearch_OnClick", "Active sheet is not available for output panel search."
    End If

    configKey = "Context.PersonValue"
    If ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook) Then
        callerName = vbNullString
        On Error Resume Next
        callerName = CStr(Application.Caller)
        On Error GoTo EH

        If ex_OutputPanel.m_TryGetClickedFieldIndex(ws, callerName, fieldIndex) Then
            If fieldIndex >= 1 And fieldIndex <= outputStyle.PanelFieldCount Then
                searchKey = ex_OutputPanel.m_ReadFieldValue(ws, outputStyle.PanelFields(fieldIndex).InputName)
                configKey = Trim$(outputStyle.PanelFields(fieldIndex).InputConfigKey)
            End If
        End If

        If Len(searchKey) = 0 Then
            searchKey = ex_OutputPanel.m_ReadSearchValue(ws)
        End If
        If Len(Trim$(configKey)) = 0 Then
            If outputStyle.PanelFieldCount >= 1 Then
                configKey = Trim$(outputStyle.PanelFields(1).InputConfigKey)
            End If
        End If
    End If
    If Len(searchKey) = 0 Then
        searchKey = ex_OutputPanel.m_ReadSearchValue(ws)
    End If
    If Len(searchKey) = 0 Then
        Err.Raise vbObjectError + 2402, "ex_UIActions.m_OutputPanelStartSearch_OnClick", "Введите значение ключа в панели поиска."
    End If

    ex_ConfigProvider.m_SetConfigValue configKey, searchKey, True
    ex_PersonTimeline.m_ShowPersonTimeline searchKey
    Exit Sub

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveSheet
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    If ex_SheetStylesXmlProvider.m_InitializeStyles(ThisWorkbook) Then
        If ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook) Then
            ex_OutputPanel.m_RenderForSheet ws, outputStyle
        End If
    End If
    On Error GoTo 0
    ex_Messaging.m_RenderErrorBanner ws, errDescription, errSource, errNumber, "ERROR: Timeline generation failed", ex_SheetStylesXmlProvider.m_GetOutputErrorBannerRangeAddress(ThisWorkbook)
End Sub
