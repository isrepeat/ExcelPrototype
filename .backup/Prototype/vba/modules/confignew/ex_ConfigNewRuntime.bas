Attribute VB_Name = "ex_ConfigNewRuntime"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const UI_CONFIG_REL_PATH As String = "config\DevUI.xml"
Private Const DEFAULT_SHEET_NAME As String = "Dev"
Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_CONFIG_MARKER_COL As Long = 1
Private Const DEV_CONFIG_KEY_COL As Long = 2
Private Const DEV_CONFIG_VALUE_COL As Long = 3
Private Const DEV_CONFIG_STYLES_COL As Long = 4
Private Const DEV_CONFIG_COL_COUNT As Long = 4
Private Const STATE_ACTIVE_PROFILE_PROP As String = "Settings.ActiveProfile"

Public Sub m_ApplyConfigControls(Optional ByVal wb As Workbook)
    Dim uiDoc As Object
    Dim configNodes As Object
    Dim controlNode As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "ConfigNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    ex_UiXmlProvider.m_RefreshConfigBindings wb

    Set uiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        UI_CONFIG_REL_PATH, _
        UI_NS, _
        "ConfigNew: UI config file was not found: ", _
        "ConfigNew: failed to parse UI config file: ")
    If uiDoc Is Nothing Then Exit Sub

    Set configNodes = uiDoc.selectNodes( _
        "/p:uiDefinition/p:layout//p:control[" & _
        "translate(normalize-space(@type), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='config' and " & _
        "string-length(normalize-space(@profileData)) > 0]")
    If configNodes Is Nothing Then Exit Sub

    For Each controlNode In configNodes
        If Not mp_ApplyConfigControlNode(controlNode, wb) Then Exit Sub
    Next controlNode
End Sub

Public Sub m_ApplyConfigControlByName(ByVal controlName As String, Optional ByVal wb As Workbook)
    Dim uiDoc As Object
    Dim controlNode As Object

    controlName = Trim$(controlName)
    If Len(controlName) = 0 Then Exit Sub

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "ConfigNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    Set uiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        UI_CONFIG_REL_PATH, _
        UI_NS, _
        "ConfigNew: UI config file was not found: ", _
        "ConfigNew: failed to parse UI config file: ")
    If uiDoc Is Nothing Then Exit Sub

    Set controlNode = uiDoc.selectSingleNode( _
        "/p:uiDefinition/p:layout//p:control[" & _
        "translate(normalize-space(@type), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='config' and " & _
        "@name=" & ex_XmlCore.m_XPathLiteral(controlName) & "]")
    If controlNode Is Nothing Then Exit Sub

    If Len(Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "profileData"))) = 0 Then Exit Sub

    mp_ApplyConfigControlNode controlNode, wb
End Sub

Private Function mp_ApplyConfigControlNode(ByVal controlNode As Object, ByVal wb As Workbook) As Boolean
    Dim controlName As String
    Dim sheetName As String
    Dim ws As Worksheet
    Dim profileDataValue As String
    Dim relPath As String
    Dim selectorText As String
    Dim filePath As String
    Dim profileDoc As Object
    Dim profileNode As Object
    Dim profileName As String
    Dim entries As Variant

    If controlNode Is Nothing Then
        mp_ApplyConfigControlNode = True
        Exit Function
    End If

    controlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(controlName) = 0 Then
        MsgBox "ConfigNew: <control type='config'> is missing required attribute 'name'.", vbExclamation
        Exit Function
    End If

    profileDataValue = Trim$(ex_UiXmlProvider.m_GetConfigBindingValue( _
        controlName, _
        "profileData", _
        Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "profileData"))))
    If Len(profileDataValue) = 0 Then
        MsgBox "ConfigNew: control '" & controlName & "' has empty resolved profileData.", vbExclamation
        Exit Function
    End If

    If Not mp_TrySplitProfileData(profileDataValue, relPath, selectorText) Then
        MsgBox "ConfigNew: invalid profileData format for control '" & controlName & "': " & profileDataValue, vbExclamation
        Exit Function
    End If

    filePath = ex_XmlCore.m_CombineBasePath(wb, relPath)
    If Len(Trim$(filePath)) = 0 Then
        MsgBox "ConfigNew: failed to build profile file path for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    Set profileDoc = ex_XmlCore.m_LoadDomByFilePath( _
        filePath, _
        UI_NS, _
        "ConfigNew: profile file was not found: ", _
        "ConfigNew: failed to parse profile file: ")
    If profileDoc Is Nothing Then Exit Function

    Set profileNode = mp_ResolveProfileNode(profileDoc, selectorText)
    If profileNode Is Nothing Then
        MsgBox "ConfigNew: profileData selector did not resolve profile item for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    sheetName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "sheet"))
    If Len(sheetName) = 0 Then sheetName = DEFAULT_SHEET_NAME
    Set ws = mp_GetWorksheetByName(wb, sheetName)
    If ws Is Nothing Then
        MsgBox "ConfigNew: sheet '" & sheetName & "' for control '" & controlName & "' was not found.", vbExclamation
        Exit Function
    End If

    entries = ex_ProfilesEntriesMapper.m_ReadProfileEntries(ws, profileNode)
    mp_WriteEntriesToConfigTable ws, entries

    profileName = Trim$(ex_XmlCore.m_NodeAttrText(profileNode, "key"))
    If Len(profileName) > 0 Then
        On Error Resume Next
        ex_ConfigProvider.m_RefreshConfigTitle ws, profileName
        On Error GoTo 0
    End If

    mp_ApplyConfigControlNode = True
End Function

Private Function mp_TrySplitProfileData( _
    ByVal profileDataValue As String, _
    ByRef outRelativePath As String, _
    ByRef outSelector As String) As Boolean

    Dim hashPos As Long

    profileDataValue = Trim$(profileDataValue)
    If Len(profileDataValue) = 0 Then Exit Function

    hashPos = InStr(1, profileDataValue, "#", vbBinaryCompare)
    If hashPos > 0 Then
        outRelativePath = Trim$(Left$(profileDataValue, hashPos - 1))
        outSelector = Trim$(Mid$(profileDataValue, hashPos + 1))
    Else
        outRelativePath = profileDataValue
        outSelector = vbNullString
    End If

    If Len(outRelativePath) = 0 Then Exit Function
    mp_TrySplitProfileData = True
End Function

Private Function mp_ResolveProfileNode(ByVal doc As Object, ByVal selectorText As String) As Object
    Dim xpathExpr As String
    Dim activeProfile As String
    Dim resolvedNode As Object

    If doc Is Nothing Then Exit Function

    selectorText = Trim$(selectorText)
    If Len(selectorText) > 0 Then
        If StrComp(Left$(selectorText, 6), "xpath=", vbTextCompare) = 0 Then
            xpathExpr = Trim$(Mid$(selectorText, 7))
        Else
            xpathExpr = selectorText
        End If

        If Len(xpathExpr) = 0 Then Exit Function
        Set resolvedNode = doc.selectSingleNode(xpathExpr)
    End If

    If Not resolvedNode Is Nothing Then
        If StrComp(LCase$(resolvedNode.baseName), "item", vbTextCompare) = 0 Then
            Set mp_ResolveProfileNode = resolvedNode
            Exit Function
        End If
    End If

    activeProfile = Trim$(mp_GetStatePropertyText(STATE_ACTIVE_PROFILE_PROP))
    If Len(activeProfile) > 0 Then
        Set resolvedNode = doc.selectSingleNode("/*/*[local-name()='list'][1]/*[local-name()='item'][@key=" & ex_XmlCore.m_XPathLiteral(activeProfile) & "]")
        If Not resolvedNode Is Nothing Then
            Set mp_ResolveProfileNode = resolvedNode
            Exit Function
        End If
    End If

    Set resolvedNode = doc.selectSingleNode("/*/*[local-name()='list'][1]/*[local-name()='item'][1]")
    Set mp_ResolveProfileNode = resolvedNode
End Function

Private Sub mp_WriteEntriesToConfigTable(ByVal ws As Worksheet, ByVal entries As Variant)
    Dim tbl As ListObject
    Dim rowCount As Long
    Dim values() As Variant
    Dim i As Long

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "ConfigNew: config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ex_ConfigTableStore.m_ClearConfigDataArea ws, tbl

    rowCount = mp_ArrayRowCount(entries)
    ex_ConfigTableStore.m_ResizeConfigTableRows ws, tbl, rowCount

    If rowCount > 0 Then
        ReDim values(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
        For i = 1 To rowCount
            values(i, DEV_CONFIG_MARKER_COL) = CStr(entries(i, DEV_CONFIG_MARKER_COL))
            values(i, DEV_CONFIG_KEY_COL) = CStr(entries(i, DEV_CONFIG_KEY_COL))
            values(i, DEV_CONFIG_VALUE_COL) = CStr(entries(i, DEV_CONFIG_VALUE_COL))
            values(i, DEV_CONFIG_STYLES_COL) = CStr(entries(i, DEV_CONFIG_STYLES_COL))
        Next i

        ex_ConfigTableStore.m_EnsureConfigTableTextFormat tbl
        tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = values
    End If

    ex_ConfigTableStore.m_ApplyConfigMarkerStyles tbl
End Sub

Private Function mp_ArrayRowCount(ByVal values As Variant) As Long
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayRowCount = UBound(values, 1) - LBound(values, 1) + 1
    End If
    Exit Function
EH:
    mp_ArrayRowCount = 0
End Function

Private Function mp_GetStatePropertyText(ByVal propName As String, Optional ByVal defaultValue As String = vbNullString) As String
    On Error GoTo EH
    mp_GetStatePropertyText = CStr(ThisWorkbook.CustomDocumentProperties(propName).Value)
    Exit Function
EH:
    mp_GetStatePropertyText = defaultValue
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
