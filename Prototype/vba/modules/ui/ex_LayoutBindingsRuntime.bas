Attribute VB_Name = "ex_LayoutBindingsRuntime"
Option Explicit

Private Const BINDINGS_LOG_PATH As String = "Logs\layout_engine.log"
Private Const BIND_SCOPE_PREFIX As String = "sheet::"
Private Const KEY_BY_ADDRESS As String = "byAddress"
Private Const KEY_BY_NAME As String = "byName"
Private Const KEY_PRIMARY As String = "primaryAddress"
Private Const PERSIST_PRIMARY_NAME As String = "__layoutPrimaryInputCell"
Private Const PERSIST_INPUT_PREFIX As String = "__layoutInput_"

Private g_SheetBindings As Object

Public Sub m_ClearSheetBindings(ByVal ws As Worksheet)
    Dim sheetKey As String

    If ws Is Nothing Then Exit Sub
    sheetKey = mp_GetSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If g_SheetBindings Is Nothing Then Exit Sub
    If g_SheetBindings.Exists(sheetKey) Then g_SheetBindings.Remove sheetKey
    mp_ClearPersistentPrimaryCell ws
    mp_ClearPersistentNamedInputCells ws
End Sub

Public Sub m_RegisterInputBinding( _
    ByVal ws As Worksheet, _
    ByVal inputCell As Range, _
    Optional ByVal inputName As String = vbNullString, _
    Optional ByVal bindSpec As String = vbNullString, _
    Optional ByVal onChangeMacro As String = vbNullString, _
    Optional ByVal isPrimaryInput As Boolean = False _
)
    Dim stageName As String
    Dim state As Object
    Dim byAddress As Object
    Dim byName As Object
    Dim cellKey As String
    Dim nameKey As String
    Dim configKey As String
    Dim bindingMeta As Object

    On Error GoTo EH

    stageName = "validate-args"
    If ws Is Nothing Then Exit Sub
    If inputCell Is Nothing Then Exit Sub
    If inputCell.Cells.Count <> 1 Then Exit Sub

    stageName = "resolve-state"
    Set state = mp_GetSheetState(ws, True)
    If state Is Nothing Then Exit Sub
    Set byAddress = state(KEY_BY_ADDRESS)
    Set byName = state(KEY_BY_NAME)

    stageName = "resolve-cell-key"
    cellKey = mp_GetCellAddressKey(inputCell)
    If Len(cellKey) = 0 Then Exit Sub

    stageName = "create-meta"
    Set bindingMeta = CreateObject("Scripting.Dictionary")
    bindingMeta.CompareMode = 1

    bindingMeta("address") = cellKey
    bindingMeta("macro") = Trim$(onChangeMacro)
    bindingMeta("bind") = Trim$(bindSpec)
    bindingMeta("configKey") = vbNullString
    If m_TryResolveConfigKeyFromBindSpec(bindSpec, configKey) Then
        bindingMeta("configKey") = configKey
    End If

    stageName = "save-by-address"
    Set byAddress(cellKey) = bindingMeta

    stageName = "save-by-name"
    nameKey = mp_NormalizeInputNameKey(inputName)
    If Len(nameKey) > 0 Then
        byName(nameKey) = cellKey
        mp_SetPersistentNamedInputCell ws, nameKey, inputCell
    End If

    stageName = "set-primary"
    If isPrimaryInput Or Len(Trim$(CStr(state(KEY_PRIMARY)))) = 0 Then
        state(KEY_PRIMARY) = cellKey
        mp_SetPersistentPrimaryCell ws, inputCell
    End If
    Exit Sub

EH:
    On Error Resume Next
    ex_Messaging.m_LogToFile _
        "[ex_LayoutBindingsRuntime] m_RegisterInputBinding failed stage='" & stageName & _
        "' ws='" & ws.Name & "' cell='" & cellKey & "' inputName='" & inputName & _
        "' bindSpec='" & bindSpec & "' error='[" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description & "'.", _
        BINDINGS_LOG_PATH
    On Error GoTo 0
    Err.Raise Err.Number, "ex_LayoutBindingsRuntime", "m_RegisterInputBinding failed at stage '" & stageName & "': " & Err.Description
End Sub

Public Sub m_HandleSheetInputChange(ByVal ws As Worksheet, ByVal target As Range)
    Dim state As Object
    Dim byAddress As Object
    Dim cellKey As String
    Dim bindingMeta As Object
    Dim macroName As String
    Dim configKey As String
    Dim valueText As String
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean

    If ws Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If target.Cells.Count <> 1 Then Exit Sub

    Set state = mp_GetSheetState(ws, False)
    If state Is Nothing Then Exit Sub
    Set byAddress = state(KEY_BY_ADDRESS)
    If byAddress Is Nothing Then Exit Sub

    cellKey = mp_GetCellAddressKey(target)
    If Len(cellKey) = 0 Then Exit Sub
    If Not byAddress.Exists(cellKey) Then Exit Sub

    If Not IsObject(byAddress(cellKey)) Then Exit Sub
    Set bindingMeta = byAddress(cellKey)
    If bindingMeta Is Nothing Then Exit Sub

    On Error Resume Next
    macroName = Trim$(CStr(bindingMeta("macro")))
    configKey = Trim$(CStr(bindingMeta("configKey")))
    On Error GoTo 0

    On Error GoTo EH
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If Len(configKey) > 0 Then
        valueText = Trim$(CStr(target.Value))
        ex_ConfigProvider.m_SetConfigValue configKey, valueText, True
    End If

    If Len(macroName) > 0 Then
        Application.Run macroName
    End If

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Exit Sub

EH:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    ex_Messaging.m_LogToFile "[ex_LayoutBindingsRuntime] m_HandleSheetInputChange failed ws='" & ws.Name & "' cell='" & cellKey & "' macro='" & macroName & "' configKey='" & configKey & "' error='" & Err.Description & "'.", BINDINGS_LOG_PATH
    On Error GoTo 0
    MsgBox "Input change handling failed: " & Err.Description, vbExclamation
End Sub

Public Function m_ReadPrimaryInputValue(ByVal ws As Worksheet) As String
    Dim targetCell As Range

    Set targetCell = mp_GetPersistentPrimaryCell(ws)
    If targetCell Is Nothing Then
        Set targetCell = mp_GetPrimaryCell(ws)
    End If
    If targetCell Is Nothing Then
        Set targetCell = mp_GetSinglePersistentNamedInputCell(ws)
    End If
    If targetCell Is Nothing Then
        On Error Resume Next
        ex_Messaging.m_LogToFile _
            "[ex_LayoutBindingsRuntime] m_ReadPrimaryInputValue: primary input is not resolved for ws='" & IIf(ws Is Nothing, "<none>", ws.Name) & "'.", _
            BINDINGS_LOG_PATH
        On Error GoTo 0
        Exit Function
    End If

    On Error Resume Next
    ex_Messaging.m_LogToFile _
        "[ex_LayoutBindingsRuntime] m_ReadPrimaryInputValue: ws='" & ws.Name & "' cell='" & targetCell.Address(False, False) & "' value='" & CStr(targetCell.Value) & "'.", _
        BINDINGS_LOG_PATH
    On Error GoTo 0
    m_ReadPrimaryInputValue = Trim$(CStr(targetCell.Value))
End Function

Public Function m_ReadInputValueByName(ByVal ws As Worksheet, ByVal inputName As String) As String
    Dim targetCell As Range
    Dim nameKey As String
    Dim resolvedAddress As String

    nameKey = mp_NormalizeInputNameKey(inputName)
    If Len(nameKey) > 0 Then
        Set targetCell = mp_GetPersistentNamedInputCell(ws, nameKey)
    End If
    If targetCell Is Nothing Then
        Set targetCell = mp_GetNamedInputCell(ws, inputName)
    End If
    If targetCell Is Nothing Then Exit Function
    resolvedAddress = targetCell.Address(False, False)
    m_ReadInputValueByName = Trim$(CStr(targetCell.Value))
    On Error Resume Next
    ex_Messaging.m_LogToFile _
        "[ex_LayoutBindingsRuntime] m_ReadInputValueByName: ws='" & ws.Name & "' inputName='" & inputName & "' cell='" & resolvedAddress & "' valueLen=" & CStr(Len(m_ReadInputValueByName)) & ".", _
        BINDINGS_LOG_PATH
    On Error GoTo 0
End Function

Public Function m_TryGetPrimaryConfigKey(ByVal ws As Worksheet, ByRef outConfigKey As String) As Boolean
    Dim state As Object
    Dim byAddress As Object
    Dim primaryAddress As String
    Dim bindingMeta As Object

    outConfigKey = vbNullString
    If ws Is Nothing Then Exit Function

    Set state = mp_GetSheetState(ws, False)
    If state Is Nothing Then Exit Function
    Set byAddress = state(KEY_BY_ADDRESS)
    If byAddress Is Nothing Then Exit Function

    primaryAddress = Trim$(CStr(state(KEY_PRIMARY)))
    If Len(primaryAddress) = 0 Then Exit Function
    If Not byAddress.Exists(primaryAddress) Then Exit Function
    If Not IsObject(byAddress(primaryAddress)) Then Exit Function

    Set bindingMeta = byAddress(primaryAddress)
    If bindingMeta Is Nothing Then Exit Function

    On Error Resume Next
    outConfigKey = Trim$(CStr(bindingMeta("configKey")))
    On Error GoTo 0
    If Len(outConfigKey) = 0 Then Exit Function

    m_TryGetPrimaryConfigKey = True
End Function

Public Function m_TryResolveConfigKeyFromBindSpec(ByVal bindSpec As String, ByRef outConfigKey As String) As Boolean
    Dim normalized As String
    Dim bindingPath As String
    Dim cfgPrefix As String
    Dim isBindingExpression As Boolean

    outConfigKey = vbNullString
    normalized = Trim$(bindSpec)
    If Len(normalized) = 0 Then Exit Function

    If mp_TryExtractBindingPath(normalized, bindingPath) Then
        isBindingExpression = True
        normalized = bindingPath
    End If

    cfgPrefix = "config."
    If isBindingExpression Then
        If StrComp(LCase$(Left$(normalized, Len(cfgPrefix))), cfgPrefix, vbBinaryCompare) <> 0 Then
            Exit Function
        End If
    End If

    If StrComp(LCase$(Left$(normalized, Len(cfgPrefix))), cfgPrefix, vbBinaryCompare) = 0 Then
        normalized = Mid$(normalized, Len(cfgPrefix) + 1)
    End If

    normalized = Trim$(normalized)
    If Len(normalized) = 0 Then Exit Function

    outConfigKey = normalized
    m_TryResolveConfigKeyFromBindSpec = True
End Function

Private Function mp_GetPrimaryCell(ByVal ws As Worksheet) As Range
    Dim state As Object
    Dim primaryAddress As String

    If ws Is Nothing Then Exit Function
    Set state = mp_GetSheetState(ws, False)
    If state Is Nothing Then Exit Function

    primaryAddress = Trim$(CStr(state(KEY_PRIMARY)))
    If Len(primaryAddress) = 0 Then Exit Function

    On Error Resume Next
    Set mp_GetPrimaryCell = ws.Range(primaryAddress)
    On Error GoTo 0
End Function

Private Sub mp_SetPersistentPrimaryCell(ByVal ws As Worksheet, ByVal targetCell As Range)
    Dim refersToText As String

    If ws Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub
    If targetCell.Cells.Count <> 1 Then Exit Sub

    refersToText = "=" & targetCell.Address(True, True, xlA1)

    On Error Resume Next
    ws.Names(PERSIST_PRIMARY_NAME).Delete
    On Error GoTo 0

    On Error GoTo EH
    ws.Names.Add Name:=PERSIST_PRIMARY_NAME, RefersTo:=refersToText
    On Error Resume Next
    ex_Messaging.m_LogToFile _
        "[ex_LayoutBindingsRuntime] mp_SetPersistentPrimaryCell: ws='" & ws.Name & "' name='" & PERSIST_PRIMARY_NAME & "' refersTo='" & refersToText & "'.", _
        BINDINGS_LOG_PATH
    On Error GoTo 0
    Exit Sub

EH:
    On Error Resume Next
    ex_Messaging.m_LogToFile _
        "[ex_LayoutBindingsRuntime] mp_SetPersistentPrimaryCell failed ws='" & ws.Name & "' refersTo='" & refersToText & "' err='[" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description & "'.", _
        BINDINGS_LOG_PATH
    On Error GoTo 0
End Sub

Private Sub mp_ClearPersistentPrimaryCell(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Names(PERSIST_PRIMARY_NAME).Delete
    On Error GoTo 0
End Sub

Private Function mp_GetPersistentPrimaryCell(ByVal ws As Worksheet) As Range
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetPersistentPrimaryCell = ws.Names(PERSIST_PRIMARY_NAME).RefersToRange
    On Error GoTo 0
End Function

Private Sub mp_SetPersistentNamedInputCell(ByVal ws As Worksheet, ByVal nameKey As String, ByVal targetCell As Range)
    Dim persistentName As String
    Dim refersToText As String

    If ws Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub
    If targetCell.Cells.Count <> 1 Then Exit Sub
    nameKey = mp_NormalizeInputNameKey(nameKey)
    If Len(nameKey) = 0 Then Exit Sub

    persistentName = mp_BuildPersistentInputName(nameKey)
    refersToText = "=" & targetCell.Address(True, True, xlA1)

    On Error Resume Next
    ws.Names(persistentName).Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Names.Add Name:=persistentName, RefersTo:=refersToText
    On Error GoTo 0
End Sub

Private Function mp_GetPersistentNamedInputCell(ByVal ws As Worksheet, ByVal nameKey As String) As Range
    Dim persistentName As String

    If ws Is Nothing Then Exit Function
    nameKey = mp_NormalizeInputNameKey(nameKey)
    If Len(nameKey) = 0 Then Exit Function
    persistentName = mp_BuildPersistentInputName(nameKey)

    On Error Resume Next
    Set mp_GetPersistentNamedInputCell = ws.Names(persistentName).RefersToRange
    On Error GoTo 0
End Function

Private Function mp_GetSinglePersistentNamedInputCell(ByVal ws As Worksheet) As Range
    Dim nm As Name
    Dim matchedCount As Long
    Dim localName As String

    If ws Is Nothing Then Exit Function

    For Each nm In ws.Names
        localName = mp_GetLocalNamePart(nm.Name)
        If LCase$(Left$(localName, Len(PERSIST_INPUT_PREFIX))) = LCase$(PERSIST_INPUT_PREFIX) Then
            matchedCount = matchedCount + 1
            If matchedCount > 1 Then Exit Function
            On Error Resume Next
            Set mp_GetSinglePersistentNamedInputCell = nm.RefersToRange
            On Error GoTo 0
        End If
    Next nm
End Function

Private Sub mp_ClearPersistentNamedInputCells(ByVal ws As Worksheet)
    Dim nm As Name
    Dim namesToDelete As Collection
    Dim item As Variant
    Dim localName As String

    If ws Is Nothing Then Exit Sub
    Set namesToDelete = New Collection

    On Error Resume Next
    For Each nm In ws.Names
        localName = mp_GetLocalNamePart(nm.Name)
        If LCase$(Left$(localName, Len(PERSIST_INPUT_PREFIX))) = LCase$(PERSIST_INPUT_PREFIX) Then
            namesToDelete.Add nm.Name
        End If
    Next nm
    On Error GoTo 0

    For Each item In namesToDelete
        On Error Resume Next
        ws.Names(CStr(item)).Delete
        On Error GoTo 0
    Next item
End Sub

Private Function mp_BuildPersistentInputName(ByVal nameKey As String) As String
    Dim resultText As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long

    nameKey = LCase$(Trim$(nameKey))
    If Len(nameKey) = 0 Then
        mp_BuildPersistentInputName = PERSIST_INPUT_PREFIX & "unnamed"
        Exit Function
    End If

    For i = 1 To Len(nameKey)
        ch = Mid$(nameKey, i, 1)
        codePoint = AscW(ch)
        If (codePoint >= 48 And codePoint <= 57) Or _
           (codePoint >= 97 And codePoint <= 122) Or _
           codePoint = 95 Then
            resultText = resultText & ch
        Else
            resultText = resultText & "_"
        End If
    Next i

    If Len(resultText) = 0 Then resultText = "unnamed"
    mp_BuildPersistentInputName = PERSIST_INPUT_PREFIX & resultText
End Function

Private Function mp_GetLocalNamePart(ByVal fullName As String) As String
    Dim bangPos As Long

    fullName = Trim$(fullName)
    bangPos = InStrRev(fullName, "!")
    If bangPos > 0 Then
        mp_GetLocalNamePart = Mid$(fullName, bangPos + 1)
    Else
        mp_GetLocalNamePart = fullName
    End If
End Function

Private Function mp_GetNamedInputCell(ByVal ws As Worksheet, ByVal inputName As String) As Range
    Dim state As Object
    Dim byName As Object
    Dim cellAddress As String
    Dim nameKey As String

    If ws Is Nothing Then Exit Function
    nameKey = mp_NormalizeInputNameKey(inputName)
    If Len(nameKey) = 0 Then Exit Function

    Set state = mp_GetSheetState(ws, False)
    If state Is Nothing Then Exit Function
    Set byName = state(KEY_BY_NAME)
    If byName Is Nothing Then Exit Function
    If Not byName.Exists(nameKey) Then Exit Function

    cellAddress = Trim$(CStr(byName(nameKey)))
    If Len(cellAddress) = 0 Then Exit Function

    On Error Resume Next
    Set mp_GetNamedInputCell = ws.Range(cellAddress)
    On Error GoTo 0
End Function

Private Function mp_GetSheetState(ByVal ws As Worksheet, ByVal createIfMissing As Boolean) As Object
    Dim sheetKey As String
    Dim state As Object
    Dim byAddress As Object
    Dim byName As Object

    If ws Is Nothing Then Exit Function
    sheetKey = mp_GetSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function

    If g_SheetBindings Is Nothing Then
        If Not createIfMissing Then Exit Function
        Set g_SheetBindings = CreateObject("Scripting.Dictionary")
        g_SheetBindings.CompareMode = 1
    End If

    If g_SheetBindings.Exists(sheetKey) Then
        Set mp_GetSheetState = g_SheetBindings(sheetKey)
        Exit Function
    End If

    If Not createIfMissing Then Exit Function

    Set state = CreateObject("Scripting.Dictionary")
    state.CompareMode = 1

    Set byAddress = CreateObject("Scripting.Dictionary")
    byAddress.CompareMode = 1
    Set state(KEY_BY_ADDRESS) = byAddress

    Set byName = CreateObject("Scripting.Dictionary")
    byName.CompareMode = 1
    Set state(KEY_BY_NAME) = byName

    state(KEY_PRIMARY) = vbNullString

    Set g_SheetBindings(sheetKey) = state
    Set mp_GetSheetState = state
End Function

Private Function mp_GetSheetKey(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Exit Function
    mp_GetSheetKey = BIND_SCOPE_PREFIX & LCase$(Trim$(ws.CodeName))
End Function

Private Function mp_NormalizeInputNameKey(ByVal inputName As String) As String
    mp_NormalizeInputNameKey = LCase$(Trim$(inputName))
End Function

Private Function mp_GetCellAddressKey(ByVal target As Range) As String
    If target Is Nothing Then Exit Function
    mp_GetCellAddressKey = LCase$(target.Address(False, False))
End Function

Private Function mp_TryExtractBindingPath(ByVal rawText As String, ByRef outPath As String) As Boolean
    Dim normalized As String

    normalized = Trim$(rawText)
    If Len(normalized) < 10 Then Exit Function
    If Left$(normalized, 9) <> "{Binding " Then Exit Function
    If Right$(normalized, 1) <> "}" Then Exit Function

    outPath = Trim$(Mid$(normalized, 10, Len(normalized) - 10))
    If Len(outPath) = 0 Then outPath = "."
    mp_TryExtractBindingPath = True
End Function
