Attribute VB_Name = "ex_UIActions"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const ASCII_UPPER As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const ASCII_LOWER As String = "abcdefghijklmnopqrstuvwxyz"

' UI entrypoints layer: keeps user-triggered callbacks in actions/*
' and delegates work to domain/config modules.

Public Sub m_DeleteResultSheets_OnClick()
    On Error GoTo EH
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_SheetStylesXmlProvider.m_DeleteResultSheets
    Exit Sub
EH:
    MsgBox "Clear failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
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
    ex_ModePersonalCard.m_RunPersonalCard
End Sub

Public Sub m_ShowReportCreation_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_ModeReportCreation.m_RunKeysCollectionReport
End Sub

Public Sub m_ShowSimpleTest_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_ModeSimpleTest.m_RunSimpleTest
End Sub

Public Sub m_ShowMultiSources_OnClick()
    On Error GoTo EH
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_ModeMultiSources.m_RunMultiSources
    Exit Sub
EH:
    MsgBox "MultiSources action failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
End Sub

Public Sub m_RunComparingTables_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_ModeTablesComparing.m_RunComparing
End Sub

Public Sub m_OutputPanelRunPostProcess_OnClick()
    ex_ModePersonalCard.m_RunPostProcessForActiveSheet
End Sub

Public Sub m_OpenPreProcessScript_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    mp_OpenActiveProfileScriptSource True
End Sub

Public Sub m_OpenPostProcessScript_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    mp_OpenActiveProfileScriptSource False
End Sub

Public Sub m_ReportCreationRunPostProcess_OnClick()
    ' ReportCreation already applies implicit postprocess during generation,
    ' so re-running ReportCreation refreshes output + postprocess in one action.
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_ModeReportCreation.m_RunKeysCollectionReport
End Sub

Public Sub m_ReportCreationExport_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    MsgBox "ReportCreation export is not implemented yet.", vbInformation
End Sub

Public Sub m_ExportFooterReportToWord_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_WordPlaceholderReports.m_API_ExportActiveSheetFooterPlaceholderReport
End Sub

Public Sub m_ExportFooterReportDone_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_WordPlaceholderReports.m_API_CleanupExportAnchorMarkers
End Sub

Public Sub m_OutputPanelToggleButton_OnClick()
    ex_CustomDropdown.m_OnManagedButtonClick
    ex_OutputPanel.m_HandleToggleButtonOnClick
End Sub

Public Sub m_OutputPanelStartSearch_OnClick()
    Dim ws As Worksheet
    Dim searchKey As String
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle
    Dim configKey As String
    Dim callerName As String
    Dim fieldIndex As Long
    Dim activeModeKey As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2401, "ex_UIActions.m_OutputPanelStartSearch_OnClick", "Active sheet is not available for output panel search."
    End If

    ' Self-heal: auto-search relies on Workbook_SheetChange, which won't fire when events are disabled.
    If Application.EnableEvents = False Then
        Application.EnableEvents = True
    End If

    ' Keep UI (Dev table) as the authoritative runtime config during searches.
    ' Profile XML reapply is intentionally not triggered from output-panel search,
    ' to avoid overwriting current UI config/state mid-session.
    ' Profile refresh is handled explicitly by mode/profile change and Update UI.

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
    activeModeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey())
    If Len(activeModeKey) = 0 Then
        Err.Raise vbObjectError + 2403, "ex_UIActions.m_OutputPanelStartSearch_OnClick", "Active mode key is empty."
    End If

    Select Case LCase$(activeModeKey)
        Case "personalcard"
            ex_ModePersonalCard.m_RunPersonalCard
        Case "multisources"
            ex_ModeMultiSources.m_RunMultiSources
        Case "simpletest"
            ex_ModeSimpleTest.m_RunSimpleTest
        Case "reportcreation"
            ex_ModeReportCreation.m_RunKeysCollectionReport
        Case Else
            Err.Raise vbObjectError + 2404, "ex_UIActions.m_OutputPanelStartSearch_OnClick", _
                "Output panel Search is not configured for mode '" & activeModeKey & "'."
    End Select
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
    ex_Messaging.m_RenderErrorBanner ws, errDescription, errSource, errNumber, "ERROR: Search failed", ex_SheetStylesXmlProvider.m_GetOutputErrorBannerRangeAddress(ThisWorkbook)
End Sub

Private Sub mp_OpenActiveProfileScriptSource(ByVal isPreProcess As Boolean)
    Dim modeKey As String
    Dim profileName As String
    Dim profilesFilePath As String
    Dim targetFilePath As String
    Dim sourceLabel As String
    Dim openLabel As String
    Dim errText As String

    modeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev))
    profileName = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName(ws_Dev))
    If Len(modeKey) = 0 Or Len(profileName) = 0 Then
        MsgBox "Cannot open script source: active mode/profile is not selected.", vbExclamation
        Exit Sub
    End If

    profilesFilePath = Trim$(ex_ProfilesStore.m_GetProfilesFilePath(modeKey, ThisWorkbook))
    If Len(profilesFilePath) = 0 Then
        MsgBox "Cannot open script source: profiles file path is empty for mode '" & modeKey & "'.", vbExclamation
        Exit Sub
    End If

    If Not mp_TryResolveActiveProfileScriptSourcePath(profilesFilePath, profileName, isPreProcess, targetFilePath, sourceLabel, errText) Then
        MsgBox "Script open failed: " & errText, vbExclamation
        Exit Sub
    End If

    If Len(Dir$(targetFilePath)) = 0 Then
        MsgBox "Script source file was not found: " & targetFilePath, vbExclamation
        Exit Sub
    End If

    If isPreProcess Then
        openLabel = "Pre Process Script"
    Else
        openLabel = "Post Process Script"
    End If

    If Not mp_OpenFileInNotepad(targetFilePath, errText) Then
        MsgBox "Failed to open " & openLabel & " (" & sourceLabel & "): " & errText, vbExclamation
    End If
End Sub

Private Function mp_TryResolveActiveProfileScriptSourcePath( _
    ByVal profilesFilePath As String, _
    ByVal profileName As String, _
    ByVal isPreProcess As Boolean, _
    ByRef outSourcePath As String, _
    ByRef outSourceLabel As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim doc As Object
    Dim profileNode As Object
    Dim scriptNode As Object
    Dim includePath As String

    outSourcePath = vbNullString
    outSourceLabel = vbNullString
    outErrorText = vbNullString

    If Len(Dir$(profilesFilePath)) = 0 Then
        outErrorText = "Profiles file was not found: " & profilesFilePath
        Exit Function
    End If

    Set doc = ex_XmlCore.m_CreateDom(PROFILES_NS)
    If Not doc.Load(profilesFilePath) Then
        outErrorText = "Failed to parse profiles file: " & profilesFilePath
        Exit Function
    End If

    Set profileNode = doc.selectSingleNode("/p:profiles/p:profile[@name=" & ex_XmlCore.m_XPathLiteral(profileName) & "]")
    If profileNode Is Nothing Then
        outErrorText = "Active profile '" & profileName & "' was not found in " & profilesFilePath
        Exit Function
    End If

    If isPreProcess Then
        Set scriptNode = mp_GetPreProcessScriptNode(profileNode)
        outSourceLabel = "preProcessScript"
    Else
        Set scriptNode = mp_GetPostProcessScriptNode(profileNode)
        outSourceLabel = "postProcessScript"
    End If

    If scriptNode Is Nothing Then
        outErrorText = "Script definition is not configured for active profile '" & profileName & "'."
        Exit Function
    End If

    includePath = Trim$(ex_XmlCore.m_NodeAttrText(scriptNode, "include"))
    If Len(includePath) > 0 Then
        outSourcePath = mp_ResolveIncludeFilePath(profilesFilePath, includePath)
        If Len(outSourcePath) = 0 Then
            outErrorText = "Unable to resolve include path: '" & includePath & "'."
            Exit Function
        End If
        outSourceLabel = outSourceLabel & " include"
    Else
        outSourcePath = profilesFilePath
        outSourceLabel = outSourceLabel & " inline"
    End If

    mp_TryResolveActiveProfileScriptSourcePath = True
End Function

Private Function mp_GetPreProcessScriptNode(ByVal profileNode As Object) As Object
    Set mp_GetPreProcessScriptNode = profileNode.selectSingleNode("p:preProcessScript")
End Function

Private Function mp_GetPostProcessScriptNode(ByVal profileNode As Object) As Object
    Dim nodes As Object

    Set nodes = profileNode.selectNodes("p:postProcessScript[translate(normalize-space(@execution), '" & ASCII_UPPER & "', '" & ASCII_LOWER & "')='explicit']")
    If Not nodes Is Nothing Then
        If nodes.Length > 0 Then
            Set mp_GetPostProcessScriptNode = nodes.Item(0)
            Exit Function
        End If
    End If

    Set nodes = profileNode.selectNodes("p:postProcessScript[translate(normalize-space(@execution), '" & ASCII_UPPER & "', '" & ASCII_LOWER & "')='implicit']")
    If Not nodes Is Nothing Then
        If nodes.Length > 0 Then
            Set mp_GetPostProcessScriptNode = nodes.Item(0)
            Exit Function
        End If
    End If

    Set nodes = profileNode.selectNodes("p:postProcessScript")
    If Not nodes Is Nothing Then
        If nodes.Length > 0 Then
            Set mp_GetPostProcessScriptNode = nodes.Item(0)
        End If
    End If
End Function

Private Function mp_OpenFileInNotepad(ByVal filePath As String, ByRef outErrorText As String) As Boolean
    Dim commandText As String
    Dim shellRunner As Object

    outErrorText = vbNullString
    commandText = "notepad.exe """ & filePath & """"

    On Error GoTo EH
    Set shellRunner = CreateObject("WScript.Shell")
    shellRunner.Run commandText, vbNormalFocus, False
    mp_OpenFileInNotepad = True
    Exit Function
EH:
    outErrorText = Err.Description
End Function

Private Function mp_ResolveIncludeFilePath(ByVal ownerFilePath As String, ByVal includePath As String) As String
    Dim normalizedIncludePath As String
    Dim ownerDir As String
    Dim combinedPath As String

    normalizedIncludePath = mp_NormalizeFilePath(includePath)
    If Len(normalizedIncludePath) = 0 Then Exit Function

    If mp_IsAbsolutePath(normalizedIncludePath) Then
        mp_ResolveIncludeFilePath = normalizedIncludePath
        Exit Function
    End If

    ownerDir = mp_GetParentDirectory(ownerFilePath)
    If Len(ownerDir) = 0 Then Exit Function

    combinedPath = ownerDir & "\" & normalizedIncludePath
    mp_ResolveIncludeFilePath = mp_NormalizeFilePath(combinedPath)
End Function

Private Function mp_GetParentDirectory(ByVal filePath As String) As String
    Dim slashPos As Long
    Dim normalized As String

    normalized = mp_NormalizeFilePath(filePath)
    If Len(normalized) = 0 Then Exit Function

    slashPos = InStrRev(normalized, "\", -1, vbBinaryCompare)
    If slashPos <= 0 Then Exit Function
    If slashPos = 1 Then
        mp_GetParentDirectory = "\"
    Else
        mp_GetParentDirectory = Left$(normalized, slashPos - 1)
    End If
End Function

Private Function mp_IsAbsolutePath(ByVal filePath As String) As Boolean
    Dim normalized As String

    normalized = mp_NormalizeFilePath(filePath)
    If Len(normalized) = 0 Then Exit Function

    If Left$(normalized, 2) = "\\" Then
        mp_IsAbsolutePath = True
        Exit Function
    End If

    If Len(normalized) >= 3 Then
        If Mid$(normalized, 2, 1) = ":" And Mid$(normalized, 3, 1) = "\" Then
            mp_IsAbsolutePath = True
            Exit Function
        End If
    End If
End Function

Private Function mp_NormalizeFilePath(ByVal filePath As String) As String
    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then Exit Function
    filePath = Replace$(filePath, "/", "\\")
    mp_NormalizeFilePath = filePath
End Function
