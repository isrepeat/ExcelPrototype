Attribute VB_Name = "ex_UIActions"
Option Explicit

' UI entrypoints layer: keeps user-triggered callbacks in actions/*
' and delegates work to domain/config modules.

Public Sub m_DeleteResultSheets_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_SheetStylesXmlProvider.m_DeleteResultSheets
End Sub

Public Sub m_SwitchMode_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_Settings.m_SwitchMode_OnClick
End Sub

Public Sub m_OnProfileChanged_OnClick()
    ex_ConfigProfilesManager.m_OnProfileChanged
End Sub

Public Sub m_OnModeChanged_OnClick()
    ex_ConfigProfilesManager.m_OnModeChanged
End Sub

Public Sub m_ToggleDropdownButton_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_CustomDropdown.m_ToggleDropdownButton ThisWorkbook
End Sub

Public Sub m_UpdateUi_OnClick()
    On Error GoTo EH

    ex_CustomDropdown.m_OnManagedButtonClick
    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws_Dev
    ex_UILoader.m_LoadUiFromConfig ThisWorkbook
    Application.Run "ex_ConfigProfilesManager.m_RestoreSelectionState"
    ex_CustomDropdown.m_InitDevTestDropdown ThisWorkbook
    Exit Sub
EH:
    MsgBox "Update UI failed: " & Err.Description, vbExclamation
End Sub

Public Sub m_SelectDropdownOption_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_CustomDropdown.m_SelectDropdownOption ThisWorkbook
End Sub

Public Sub m_OnModeOptionSelected(Optional ByVal selectedKey As String = vbNullString, Optional ByVal selectedCaption As String = vbNullString, Optional ByVal sourceControlName As String = vbNullString)
    Dim modeKey As String

    modeKey = Trim$(selectedKey)
    If Len(modeKey) = 0 Then
        modeKey = Trim$(selectedCaption)
    End If
    If Len(modeKey) = 0 Then
        MsgBox "Mode option selection is empty for control '" & sourceControlName & "'.", vbExclamation
        Exit Sub
    End If

    ex_ConfigProfilesManager.m_SetActiveModeKey modeKey
    If StrComp(Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey()), modeKey, vbTextCompare) <> 0 Then
        Exit Sub
    End If

    ex_ConfigProfilesManager.m_OnModeChanged
End Sub

Public Sub m_OnProfileOptionSelected(Optional ByVal selectedKey As String = vbNullString, Optional ByVal selectedCaption As String = vbNullString, Optional ByVal sourceControlName As String = vbNullString)
    Dim profileName As String
    Dim activeModeKey As String

    profileName = Trim$(selectedKey)
    If Len(profileName) = 0 Then
        profileName = Trim$(selectedCaption)
    End If
    If Len(profileName) = 0 Then
        MsgBox "Profile option selection is empty for control '" & sourceControlName & "'.", vbExclamation
        Exit Sub
    End If

    activeModeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey())
    If Len(activeModeKey) = 0 Then
        MsgBox "Active mode key is empty while applying profile selection.", vbExclamation
        Exit Sub
    End If

    ex_ConfigProfilesManager.m_SetActiveProfileName profileName, activeModeKey
    If StrComp(Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName()), profileName, vbTextCompare) <> 0 Then
        Exit Sub
    End If

    ex_ConfigProfilesManager.m_OnProfileChanged
End Sub

Public Sub m_HelloWorld_OnClick()
    ex_Startup.m_HelloWorld
End Sub

Public Sub m_ShowPersonalCard_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_PersonTimeline.m_ShowPersonTimeline_UI
End Sub

Public Sub m_RunComparingTables_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_TableComparing.m_RunComparing
End Sub

Public Sub m_OutputPanelRunPostProcess_OnClick()
    ex_PersonTimeline.m_RunPostProcessForActiveSheet
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

    configKey = "CommonKey"
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
    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws
    If ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook) Then
        ex_OutputPanel.m_RenderForSheet ws, outputStyle
    End If
    On Error GoTo 0
    ex_Messaging.m_RenderErrorBanner ws, errDescription, errSource, errNumber, "ERROR: Timeline generation failed", ex_SheetStylesXmlProvider.m_GetOutputErrorBannerRangeAddress(ThisWorkbook)
End Sub
