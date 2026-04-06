Attribute VB_Name = "ex_ResultLayoutXmlProvider"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const OUTPUT_UI_FILE_SUFFIX As String = "UI.xml"
Private Const DEBUG_LOG_PATH As String = "Logs\layout_engine.log"
Private Const DEBUG_LOG_ENABLED As Boolean = False

Public Function m_TryLoadActiveModeUiDom( _
    ByVal wb As Workbook, _
    ByRef outDoc As Object, _
    ByRef outUiFilePath As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim activeModeKey As String
    Dim uiFilePath As String
    Dim doc As Object
    Dim parseReason As String
    Dim parseLine As Long
    Dim parsePos As Long

    outErrorText = vbNullString
    outUiFilePath = vbNullString
    Set outDoc = Nothing
    mp_DebugLog "m_TryLoadActiveModeUiDom: start."

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        outErrorText = "Workbook is not available for active mode UI resolution."
        mp_DebugLog "m_TryLoadActiveModeUiDom: " & outErrorText
        Exit Function
    End If

    activeModeKey = mp_GetCurrentActiveModeKey()
    If Len(activeModeKey) = 0 Then
        outErrorText = "Active mode key is empty for mode UI mapping."
        mp_DebugLog "m_TryLoadActiveModeUiDom: " & outErrorText
        Exit Function
    End If
    mp_DebugLog "m_TryLoadActiveModeUiDom: activeMode='" & activeModeKey & "'."

    uiFilePath = mp_GetOutputUiFilePathByMode(activeModeKey, wb, outErrorText)
    If Len(uiFilePath) = 0 Then Exit Function

    If Len(Dir$(uiFilePath)) = 0 Then
        outErrorText = "Mode UI config file was not found: " & uiFilePath
        mp_DebugLog "m_TryLoadActiveModeUiDom: " & outErrorText
        Exit Function
    End If
    mp_DebugLog "m_TryLoadActiveModeUiDom: uiFilePath='" & uiFilePath & "'."

    Set doc = ex_XmlCore.m_CreateDom(PROFILES_NS)
    If doc Is Nothing Then
        outErrorText = "Failed to create XML DOM for mode UI config file '" & uiFilePath & "'."
        mp_DebugLog "m_TryLoadActiveModeUiDom: " & outErrorText
        Exit Function
    End If

    If Not doc.Load(uiFilePath) Then
        parseReason = Trim$(CStr(doc.parseError.reason))
        parseLine = CLng(doc.parseError.Line)
        parsePos = CLng(doc.parseError.linepos)
        If Len(parseReason) = 0 Then parseReason = "Unknown XML parse error."

        outErrorText = "Failed to parse mode UI config file '" & uiFilePath & "': " & parseReason
        If parseLine > 0 Then
            outErrorText = outErrorText & " (line " & CStr(parseLine)
            If parsePos > 0 Then outErrorText = outErrorText & ", pos " & CStr(parsePos)
            outErrorText = outErrorText & ")."
        End If
        mp_DebugLog "m_TryLoadActiveModeUiDom: " & outErrorText
        Exit Function
    End If

    outUiFilePath = uiFilePath
    Set outDoc = doc
    m_TryLoadActiveModeUiDom = True
    mp_DebugLog "m_TryLoadActiveModeUiDom: loaded."
End Function

Public Function m_HasResultLayoutGrid(ByVal doc As Object) As Boolean
    Dim gridNodes As Object

    If doc Is Nothing Then Exit Function
    Set gridNodes = doc.selectNodes("/p:uiDefinition/p:layout/p:grid")
    If gridNodes Is Nothing Then Exit Function
    m_HasResultLayoutGrid = (gridNodes.Length > 0)
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ResultLayoutXmlProvider] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub

Private Function mp_GetCurrentActiveModeKey() As String
    Dim defaultModeKey As String

    On Error Resume Next
    mp_GetCurrentActiveModeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey())
    On Error GoTo 0

    If Len(mp_GetCurrentActiveModeKey) = 0 Then
        defaultModeKey = Trim$(ex_UiXmlProvider.m_GetDefaultModeKey(ThisWorkbook))
        If Len(defaultModeKey) > 0 Then
            mp_GetCurrentActiveModeKey = defaultModeKey
        Else
            mp_GetCurrentActiveModeKey = Trim$(ex_UiXmlProvider.m_GetModeKeyByIndex(1, ThisWorkbook))
        End If
    End If
End Function

Private Function mp_GetOutputUiFilePathByMode( _
    ByVal modeKey As String, _
    ByVal wb As Workbook, _
    ByRef outErrorText As String _
) As String
    Dim profilesFilePath As String
    Dim slashPos As Long
    Dim modeDirPath As String
    Dim modeDirName As String

    modeKey = Trim$(modeKey)
    If Len(modeKey) = 0 Then
        outErrorText = "Active mode key is empty for mode UI mapping."
        Exit Function
    End If

    profilesFilePath = Trim$(ex_UiXmlProvider.m_GetProfilesFilePathByMode(modeKey, wb, "profilesFileByMode"))
    If Len(profilesFilePath) = 0 Then
        outErrorText = "Profiles file path is not resolved for active mode key '" & modeKey & "'."
        Exit Function
    End If

    slashPos = InStrRev(profilesFilePath, "\")
    If slashPos <= 1 Then
        outErrorText = "Invalid profiles file path for active mode key '" & modeKey & "': " & profilesFilePath
        Exit Function
    End If
    modeDirPath = Left$(profilesFilePath, slashPos - 1)

    slashPos = InStrRev(modeDirPath, "\")
    If slashPos <= 0 Then
        modeDirName = modeDirPath
    Else
        modeDirName = Mid$(modeDirPath, slashPos + 1)
    End If
    modeDirName = Trim$(modeDirName)
    If Len(modeDirName) = 0 Then
        outErrorText = "Invalid mode directory in profiles file path for active mode key '" & modeKey & "': " & profilesFilePath
        Exit Function
    End If

    mp_GetOutputUiFilePathByMode = modeDirPath & "\" & modeDirName & OUTPUT_UI_FILE_SUFFIX
End Function
