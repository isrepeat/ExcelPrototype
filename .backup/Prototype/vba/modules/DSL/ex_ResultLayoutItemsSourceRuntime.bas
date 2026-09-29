Attribute VB_Name = "ex_ResultLayoutItemsRt"
Option Explicit

Private Const INPUT_KEY_LAYOUT_ITEMSOURCES As String = "__ResultLayoutItemsSources"
Private Const INPUT_KEY_LAYOUT_FIELDRANGES As String = "__ResultLayoutFieldRanges"
Private Const INPUT_KEY_LAYOUT_KINDRANGES As String = "__ResultLayoutKindRanges"
Private Const LAYOUT_ITEMSOURCE_BANNER_CONTROL_PREFIX As String = "__LayoutBannerControl."
Private Const DEBUG_LOG_PATH As String = "Logs\layout_engine.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Private Const SESSION_KEY_PREFIX As String = "sheet::"
Private Const SESSION_FIELD_DOC As String = "doc"
Private Const SESSION_FIELD_RESULT_TABLES As String = "resultTables"
Private Const SESSION_FIELD_INPUT_OBJECT As String = "inputObject"
Private Const SESSION_FIELD_ITEMS_MAP As String = "itemsMap"
Private Const SESSION_FIELD_BATCH_DEPTH As String = "batchDepth"
Private Const SESSION_FIELD_DIRTY_KEYS As String = "dirtyKeys"

Private g_Sessions As Object

Public Sub m_ClearSession(ByVal ws As Worksheet)
    Dim sheetKey As String

    If ws Is Nothing Then Exit Sub
    sheetKey = mp_GetSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If g_Sessions Is Nothing Then Exit Sub
    If g_Sessions.Exists(sheetKey) Then g_Sessions.Remove sheetKey

    mp_DebugLog "m_ClearSession: ws='" & ws.Name & "'."
End Sub

Public Sub m_RegisterSession( _
    ByVal ws As Worksheet, _
    ByVal layoutDoc As Object, _
    ByVal resultTables As Collection, _
    ByVal inputObject As Object _
)
    Dim session As Object
    Dim itemsMap As Object
    Dim sheetKey As String
    Dim layoutInput As Object

    If ws Is Nothing Then Exit Sub
    If layoutDoc Is Nothing Then Exit Sub
    If resultTables Is Nothing Then Exit Sub

    sheetKey = mp_GetSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If inputObject Is Nothing Then
        Set layoutInput = New obj_ScriptIOPayload
    Else
        Set layoutInput = inputObject
    End If

    Set itemsMap = mp_EnsureItemsSourceMap(layoutInput)

    Set session = CreateObject("Scripting.Dictionary")
    session.CompareMode = 1
    Set session(SESSION_FIELD_DOC) = layoutDoc
    Set session(SESSION_FIELD_RESULT_TABLES) = resultTables
    Set session(SESSION_FIELD_INPUT_OBJECT) = layoutInput
    Set session(SESSION_FIELD_ITEMS_MAP) = itemsMap
    session(SESSION_FIELD_BATCH_DEPTH) = CLng(0)
    Set session(SESSION_FIELD_DIRTY_KEYS) = mp_CreateDirtyKeysMap()

    mp_EnsureSessionsStore
    If g_Sessions.Exists(sheetKey) Then g_Sessions.Remove sheetKey
    Set g_Sessions(sheetKey) = session

    mp_DebugLog "m_RegisterSession: ws='" & ws.Name & "' itemsSourceMapCount=" & mp_TryGetDictionaryCountText(itemsMap) & "."
End Sub

Public Sub m_BeginBatchUpdate(ByVal ws As Worksheet)
    Dim session As Object
    Dim batchDepth As Long

    If ws Is Nothing Then Exit Sub
    If Not mp_TryGetSession(ws, session) Then Exit Sub

    batchDepth = mp_GetSessionBatchDepth(session)
    batchDepth = batchDepth + 1
    session(SESSION_FIELD_BATCH_DEPTH) = batchDepth
    mp_DebugLog "m_BeginBatchUpdate: ws='" & ws.Name & "' depth=" & CStr(batchDepth) & "."
End Sub

Public Function m_IsBatchUpdateActive(ByVal ws As Worksheet) As Boolean
    Dim session As Object

    If ws Is Nothing Then Exit Function
    If Not mp_TryGetSession(ws, session) Then Exit Function
    m_IsBatchUpdateActive = mp_IsSessionBatchActive(session)
End Function

Public Function m_EndBatchUpdate( _
    ByVal ws As Worksheet, _
    Optional ByVal applyRefresh As Boolean = True, _
    Optional ByRef outErrorText As String _
) As Boolean
    Dim session As Object
    Dim batchDepth As Long
    Dim dirtyKeys As Object
    Dim dirtyKey As Variant

    outErrorText = vbNullString
    If ws Is Nothing Then
        outErrorText = "Worksheet is required for m_EndBatchUpdate."
        Exit Function
    End If
    If Not mp_TryGetSession(ws, session) Then
        m_EndBatchUpdate = True
        Exit Function
    End If

    batchDepth = mp_GetSessionBatchDepth(session)
    If batchDepth > 0 Then batchDepth = batchDepth - 1
    session(SESSION_FIELD_BATCH_DEPTH) = batchDepth

    If batchDepth > 0 Then
        m_EndBatchUpdate = True
        Exit Function
    End If

    Set dirtyKeys = mp_GetOrCreateSessionDirtyKeys(session)
    If applyRefresh And Not dirtyKeys Is Nothing Then
        If dirtyKeys.Count > 0 Then
            For Each dirtyKey In dirtyKeys.Keys
                If Len(Trim$(CStr(dirtyKey))) = 0 Then
                    If Not m_Refresh(ws, vbNullString, outErrorText) Then Exit Function
                    Exit For
                End If
                If Not m_Refresh(ws, CStr(dirtyKey), outErrorText) Then Exit Function
            Next dirtyKey
        End If
    End If

    If applyRefresh Then
        ex_PostProcessActions.m_FlushPostLayoutDeferredBanners ws
    Else
        mp_DebugLog "m_EndBatchUpdate: deferred banner flush postponed ws='" & ws.Name & "' applyRefresh=false."
    End If
    mp_ClearSessionDirtyKeys session

    mp_DebugLog "m_EndBatchUpdate: ws='" & ws.Name & "' refreshed=" & LCase$(CStr(applyRefresh)) & "."
    m_EndBatchUpdate = True
End Function

Public Function m_SetItemsSource( _
    ByVal ws As Worksheet, _
    ByVal itemsSourceKey As String, _
    ByVal itemsSourceCollection As Collection, _
    Optional ByVal autoRefresh As Boolean = True, _
    Optional ByRef outErrorText As String _
) As Boolean
    Dim session As Object
    Dim itemsMap As Object
    Dim normalizedKey As String

    outErrorText = vbNullString
    normalizedKey = Trim$(itemsSourceKey)

    If ws Is Nothing Then
        outErrorText = "Worksheet is required for m_SetItemsSource."
        Exit Function
    End If
    If Len(normalizedKey) = 0 Then
        outErrorText = "itemsSource key is empty in m_SetItemsSource."
        Exit Function
    End If
    If itemsSourceCollection Is Nothing Then
        outErrorText = "itemsSource collection is Nothing for key '" & normalizedKey & "'."
        Exit Function
    End If

    If Not mp_TryGetSession(ws, session) Then
        outErrorText = "Result layout session is not registered for sheet '" & ws.Name & "'."
        Exit Function
    End If

    Set itemsMap = session(SESSION_FIELD_ITEMS_MAP)
    If itemsMap Is Nothing Then
        outErrorText = "itemsSource map is not initialized for sheet '" & ws.Name & "'."
        Exit Function
    End If

    Set itemsMap(normalizedKey) = itemsSourceCollection
    mp_DebugLog "m_SetItemsSource: ws='" & ws.Name & "' key='" & normalizedKey & "' count=" & CStr(itemsSourceCollection.Count) & "."

    If autoRefresh Then
        If mp_IsSessionBatchActive(session) Then
            mp_MarkSessionDirtyKey session, normalizedKey
        Else
            If Not m_Refresh(ws, normalizedKey, outErrorText) Then Exit Function
        End If
    End If

    m_SetItemsSource = True
End Function

Public Function m_NotifyItemsSourceChanged( _
    ByVal ws As Worksheet, _
    ByVal itemsSourceKey As String, _
    Optional ByVal autoRefresh As Boolean = True, _
    Optional ByRef outErrorText As String _
) As Boolean
    Dim normalizedKey As String
    Dim session As Object

    outErrorText = vbNullString
    normalizedKey = Trim$(itemsSourceKey)

    If ws Is Nothing Then
        outErrorText = "Worksheet is required for m_NotifyItemsSourceChanged."
        Exit Function
    End If

    If autoRefresh Then
        If mp_TryGetSession(ws, session) Then
            If mp_IsSessionBatchActive(session) Then
                mp_MarkSessionDirtyKey session, normalizedKey
            Else
                If Not m_Refresh(ws, normalizedKey, outErrorText) Then Exit Function
            End If
        Else
            If Not m_Refresh(ws, normalizedKey, outErrorText) Then Exit Function
        End If
    End If

    m_NotifyItemsSourceChanged = True
End Function

Public Function m_Refresh( _
    ByVal ws As Worksheet, _
    Optional ByVal changedItemsSourceKey As String = vbNullString, _
    Optional ByRef outErrorText As String _
) As Boolean
    Dim session As Object
    Dim layoutDoc As Object
    Dim resultTables As Collection
    Dim inputObject As Object
    Dim changedKey As String

    outErrorText = vbNullString
    If ws Is Nothing Then
        outErrorText = "Worksheet is required for m_Refresh."
        Exit Function
    End If

    If Not mp_TryGetSession(ws, session) Then
        outErrorText = "Result layout session is not registered for sheet '" & ws.Name & "'."
        Exit Function
    End If

    changedKey = Trim$(changedItemsSourceKey)

    Set layoutDoc = session(SESSION_FIELD_DOC)
    Set resultTables = session(SESSION_FIELD_RESULT_TABLES)
    Set inputObject = session(SESSION_FIELD_INPUT_OBJECT)

    If layoutDoc Is Nothing Then
        outErrorText = "Layout DOM is missing in layout session for sheet '" & ws.Name & "'."
        Exit Function
    End If
    If resultTables Is Nothing Then
        outErrorText = "ResultTables are missing in layout session for sheet '" & ws.Name & "'."
        Exit Function
    End If

    If Not ex_ResultLayoutXmlEngine.m_ApplyResultLayoutFromDom(layoutDoc, ws, resultTables, inputObject, outErrorText) Then
        If Len(outErrorText) = 0 Then outErrorText = "Unknown XML layout refresh error."
        Exit Function
    End If
    If Not mp_ApplyRefreshSheetStyling(ws, inputObject, outErrorText) Then Exit Function

    mp_DebugLog "m_Refresh: applied ws='" & ws.Name & "' key='" & changedKey & "'."
    m_Refresh = True
End Function

Public Function m_TryGetItemsSource( _
    ByVal ws As Worksheet, _
    ByVal itemsSourceKey As String, _
    ByRef outItemsSource As Object _
) As Boolean
    Dim session As Object
    Dim itemsMap As Object
    Dim normalizedKey As String

    Set outItemsSource = Nothing
    normalizedKey = Trim$(itemsSourceKey)

    If ws Is Nothing Then Exit Function
    If Len(normalizedKey) = 0 Then Exit Function
    If Not mp_TryGetSession(ws, session) Then Exit Function

    Set itemsMap = session(SESSION_FIELD_ITEMS_MAP)
    If itemsMap Is Nothing Then Exit Function
    If Not itemsMap.Exists(normalizedKey) Then Exit Function

    Set outItemsSource = itemsMap(normalizedKey)
    m_TryGetItemsSource = Not (outItemsSource Is Nothing)
End Function

Public Function m_TryGetItemsSourcesMap( _
    ByVal ws As Worksheet, _
    ByRef outItemsMap As Object _
) As Boolean
    Dim session As Object

    Set outItemsMap = Nothing
    If ws Is Nothing Then Exit Function
    If Not mp_TryGetSession(ws, session) Then Exit Function

    Set outItemsMap = session(SESSION_FIELD_ITEMS_MAP)
    If outItemsMap Is Nothing Then Exit Function
    m_TryGetItemsSourcesMap = True
End Function

Public Function m_IsItemsSourceBound( _
    ByVal ws As Worksheet, _
    ByVal itemsSourceKey As String _
) As Boolean
    Dim session As Object
    Dim normalizedKey As String
    Dim layoutDoc As Object

    normalizedKey = Trim$(itemsSourceKey)
    If ws Is Nothing Then Exit Function
    If Len(normalizedKey) = 0 Then Exit Function
    If Not mp_TryGetSession(ws, session) Then Exit Function

    Set layoutDoc = session(SESSION_FIELD_DOC)
    If layoutDoc Is Nothing Then Exit Function

    m_IsItemsSourceBound = mp_LayoutUsesItemsSourceKey(layoutDoc, normalizedKey)
End Function

Public Function m_TryResolveItemsPanelItemsSourceKey( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByRef outItemsSourceKey As String, _
    Optional ByRef outErrorText As String _
) As Boolean
    Dim session As Object
    Dim layoutDoc As Object
    Dim controlNode As Object
    Dim controlType As String
    Dim sourceText As String
    Dim normalizedName As String

    outItemsSourceKey = vbNullString
    outErrorText = vbNullString
    normalizedName = Trim$(controlName)

    If ws Is Nothing Then
        outErrorText = "Worksheet is required for control lookup."
        Exit Function
    End If
    If Len(normalizedName) = 0 Then
        outErrorText = "Control name is required for itemsPanel lookup."
        Exit Function
    End If
    If Not mp_TryGetSession(ws, session) Then
        outErrorText = "Result layout session is not registered for sheet '" & ws.Name & "'."
        Exit Function
    End If

    Set layoutDoc = session(SESSION_FIELD_DOC)
    If layoutDoc Is Nothing Then
        outErrorText = "Layout DOM is not available for control lookup."
        Exit Function
    End If
    If Not mp_TryGetControlNodeByName(layoutDoc, normalizedName, controlNode) Then
        outErrorText = "Control '" & normalizedName & "' is not present in current layout."
        Exit Function
    End If

    controlType = LCase$(Trim$(mp_NodeAttrText(controlNode, "type")))
    If StrComp(controlType, "itemspanel", vbTextCompare) <> 0 Then
        outErrorText = "Control '" & normalizedName & "' must be type='itemsPanel', got type='" & controlType & "'."
        Exit Function
    End If

    sourceText = Trim$(mp_NodeAttrText(controlNode, "itemsSource"))
    If Len(sourceText) = 0 Then
        outErrorText = "itemsPanel control '" & normalizedName & "' requires non-empty itemsSource."
        Exit Function
    End If
    If mp_IsBindingExpression(sourceText) Then
        outErrorText = "itemsPanel control '" & normalizedName & "' uses binding itemsSource; direct key is required for script updates."
        Exit Function
    End If

    outItemsSourceKey = sourceText
    m_TryResolveItemsPanelItemsSourceKey = True
End Function

Public Function m_TryResolveBannerControlItemsSourceKey( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByRef outItemsSourceKey As String, _
    Optional ByRef outErrorText As String _
) As Boolean
    Dim session As Object
    Dim layoutDoc As Object
    Dim controlNode As Object
    Dim controlType As String
    Dim normalizedName As String

    outItemsSourceKey = vbNullString
    outErrorText = vbNullString
    normalizedName = Trim$(controlName)

    If ws Is Nothing Then
        outErrorText = "Worksheet is required for control lookup."
        Exit Function
    End If
    If Len(normalizedName) = 0 Then
        outErrorText = "Control name is required for banner layout target."
        Exit Function
    End If
    If Not mp_TryGetSession(ws, session) Then
        outErrorText = "Result layout session is not registered for sheet '" & ws.Name & "'."
        Exit Function
    End If

    Set layoutDoc = session(SESSION_FIELD_DOC)
    If layoutDoc Is Nothing Then
        outErrorText = "Layout DOM is not available for control lookup."
        Exit Function
    End If
    If Not mp_TryGetControlNodeByName(layoutDoc, normalizedName, controlNode) Then
        outErrorText = "Control '" & normalizedName & "' is not present in current layout."
        Exit Function
    End If

    controlType = LCase$(Trim$(mp_NodeAttrText(controlNode, "type")))
    If StrComp(controlType, "banner", vbTextCompare) <> 0 Then
        outErrorText = "Control '" & normalizedName & "' must be type='banner', got type='" & controlType & "'."
        Exit Function
    End If

    outItemsSourceKey = LAYOUT_ITEMSOURCE_BANNER_CONTROL_PREFIX & normalizedName
    m_TryResolveBannerControlItemsSourceKey = True
End Function

Private Function mp_EnsureItemsSourceMap(ByVal inputObject As Object) As Object
    Dim itemsMap As Object

    Set itemsMap = Nothing
    If Not inputObject Is Nothing Then
        ex_ScriptIO.m_TryGetObject inputObject, INPUT_KEY_LAYOUT_ITEMSOURCES, itemsMap
    End If

    If itemsMap Is Nothing Then
        Set itemsMap = CreateObject("Scripting.Dictionary")
        itemsMap.CompareMode = 1
        If Not inputObject Is Nothing Then
            ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_ITEMSOURCES, itemsMap
        End If
        Set mp_EnsureItemsSourceMap = itemsMap
        Exit Function
    End If

    If Not mp_IsDictionary(itemsMap) Then
        Set itemsMap = CreateObject("Scripting.Dictionary")
        itemsMap.CompareMode = 1
        If Not inputObject Is Nothing Then
            ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_ITEMSOURCES, itemsMap
        End If
    End If

    Set mp_EnsureItemsSourceMap = itemsMap
End Function

Private Function mp_TryGetSession(ByVal ws As Worksheet, ByRef outSession As Object) As Boolean
    Dim sheetKey As String

    Set outSession = Nothing
    If ws Is Nothing Then Exit Function

    sheetKey = mp_GetSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function

    If g_Sessions Is Nothing Then Exit Function
    If Not g_Sessions.Exists(sheetKey) Then Exit Function

    Set outSession = g_Sessions(sheetKey)
    mp_TryGetSession = Not (outSession Is Nothing)
End Function

Private Sub mp_EnsureSessionsStore()
    If Not g_Sessions Is Nothing Then Exit Sub

    Set g_Sessions = CreateObject("Scripting.Dictionary")
    g_Sessions.CompareMode = 1
End Sub

Private Function mp_GetSheetKey(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Exit Function
    mp_GetSheetKey = SESSION_KEY_PREFIX & LCase$(Trim$(ws.CodeName))
End Function

Private Function mp_CreateDirtyKeysMap() As Object
    Set mp_CreateDirtyKeysMap = CreateObject("Scripting.Dictionary")
    mp_CreateDirtyKeysMap.CompareMode = 1
End Function

Private Function mp_GetOrCreateSessionDirtyKeys(ByVal session As Object) As Object
    If session Is Nothing Then Exit Function

    If session.Exists(SESSION_FIELD_DIRTY_KEYS) Then
        Set mp_GetOrCreateSessionDirtyKeys = session(SESSION_FIELD_DIRTY_KEYS)
        If Not mp_GetOrCreateSessionDirtyKeys Is Nothing Then Exit Function
    End If

    Set mp_GetOrCreateSessionDirtyKeys = mp_CreateDirtyKeysMap()
    Set session(SESSION_FIELD_DIRTY_KEYS) = mp_GetOrCreateSessionDirtyKeys
End Function

Private Sub mp_ClearSessionDirtyKeys(ByVal session As Object)
    Dim dirtyKeys As Object

    If session Is Nothing Then Exit Sub
    Set dirtyKeys = mp_GetOrCreateSessionDirtyKeys(session)
    If dirtyKeys Is Nothing Then Exit Sub
    If dirtyKeys.Count > 0 Then dirtyKeys.RemoveAll
End Sub

Private Function mp_GetSessionBatchDepth(ByVal session As Object) As Long
    On Error Resume Next
    If session Is Nothing Then Exit Function
    If session.Exists(SESSION_FIELD_BATCH_DEPTH) Then
        mp_GetSessionBatchDepth = CLng(session(SESSION_FIELD_BATCH_DEPTH))
    End If
    If Err.Number <> 0 Then
        Err.Clear
        mp_GetSessionBatchDepth = 0
    End If
    On Error GoTo 0
End Function

Private Function mp_IsSessionBatchActive(ByVal session As Object) As Boolean
    mp_IsSessionBatchActive = (mp_GetSessionBatchDepth(session) > 0)
End Function

Private Sub mp_MarkSessionDirtyKey(ByVal session As Object, ByVal sourceKey As String)
    Dim dirtyKeys As Object
    sourceKey = Trim$(sourceKey)
    If Len(sourceKey) = 0 Then Exit Sub
    Set dirtyKeys = mp_GetOrCreateSessionDirtyKeys(session)
    If dirtyKeys Is Nothing Then Exit Sub
    If Not dirtyKeys.Exists(sourceKey) Then dirtyKeys.Add sourceKey, True
End Sub

Private Function mp_ApplyRefreshSheetStyling( _
    ByVal ws As Worksheet, _
    ByVal layoutInput As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim objectValue As Object
    Dim resultFieldRanges As Collection
    Dim kindRanges As Object

    On Error GoTo EH

    If ws Is Nothing Then
        outErrorText = "Worksheet is required for layout refresh styling."
        Exit Function
    End If

    Set resultFieldRanges = Nothing

    If Not layoutInput Is Nothing Then
        Set objectValue = Nothing
        If ex_ScriptIO.m_TryGetObject(layoutInput, INPUT_KEY_LAYOUT_FIELDRANGES, objectValue) Then
            If Not objectValue Is Nothing Then
                If TypeName(objectValue) = "Collection" Then
                    Set resultFieldRanges = objectValue
                End If
            End If
        End If

        Set objectValue = Nothing
        If ex_ScriptIO.m_TryGetObject(layoutInput, INPUT_KEY_LAYOUT_KINDRANGES, objectValue) Then
            If Not objectValue Is Nothing Then
                Set kindRanges = objectValue
            End If
        End If
    End If

    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, resultFieldRanges, Nothing, kindRanges
    mp_ApplyRefreshSheetStyling = True
    Exit Function
EH:
    outErrorText = "Failed to apply refresh sheet styling: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_LayoutUsesItemsSourceKey(ByVal layoutDoc As Object, ByVal itemsSourceKey As String) As Boolean
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim sourceText As String

    itemsSourceKey = Trim$(itemsSourceKey)
    If Len(itemsSourceKey) = 0 Then
        mp_LayoutUsesItemsSourceKey = True
        Exit Function
    End If
    If layoutDoc Is Nothing Then Exit Function

    On Error GoTo EH
    Set controlNodes = layoutDoc.selectNodes("//*[local-name()='control'][@itemsSource]")
    If controlNodes Is Nothing Then Exit Function

    For Each controlNode In controlNodes
        sourceText = Trim$(mp_NodeAttrText(controlNode, "itemsSource"))
        If Len(sourceText) = 0 Then GoTo ContinueNode
        If mp_IsBindingExpression(sourceText) Then GoTo ContinueNode
        If StrComp(sourceText, itemsSourceKey, vbTextCompare) = 0 Then
            mp_LayoutUsesItemsSourceKey = True
            Exit Function
        End If
ContinueNode:
    Next controlNode
    Exit Function
EH:
    On Error Resume Next
    mp_DebugLog "mp_LayoutUsesItemsSourceKey failed key='" & itemsSourceKey & "' err='[" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description & "'."
    On Error GoTo 0
End Function

Private Function mp_TryGetControlNodeByName( _
    ByVal layoutDoc As Object, _
    ByVal controlName As String, _
    ByRef outControlNode As Object _
) As Boolean
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim nodeName As String

    Set outControlNode = Nothing
    controlName = Trim$(controlName)
    If layoutDoc Is Nothing Then Exit Function
    If Len(controlName) = 0 Then Exit Function

    On Error GoTo EH
    Set controlNodes = layoutDoc.selectNodes("//*[local-name()='control'][@name]")
    If controlNodes Is Nothing Then Exit Function

    For Each controlNode In controlNodes
        nodeName = Trim$(mp_NodeAttrText(controlNode, "name"))
        If StrComp(nodeName, controlName, vbTextCompare) = 0 Then
            Set outControlNode = controlNode
            mp_TryGetControlNodeByName = True
            Exit Function
        End If
    Next controlNode
    Exit Function
EH:
    On Error Resume Next
    mp_DebugLog "mp_TryGetControlNodeByName failed control='" & controlName & "' err='[" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description & "'."
    On Error GoTo 0
End Function

Private Function mp_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    Dim attrNode As Object

    If node Is Nothing Then Exit Function
    On Error Resume Next
    Set attrNode = node.Attributes.getNamedItem(attrName)
    If Not attrNode Is Nothing Then mp_NodeAttrText = CStr(attrNode.Text)
    On Error GoTo 0
End Function

Private Function mp_IsBindingExpression(ByVal rawText As String) As Boolean
    rawText = Trim$(rawText)
    If Len(rawText) < 10 Then Exit Function
    If Left$(rawText, 9) <> "{Binding " Then Exit Function
    If Right$(rawText, 1) <> "}" Then Exit Function
    mp_IsBindingExpression = True
End Function

Private Function mp_IsDictionary(ByVal sourceObject As Object) As Boolean
    If sourceObject Is Nothing Then Exit Function
    mp_IsDictionary = (TypeName(sourceObject) = "Dictionary" Or TypeName(sourceObject) = "Scripting.Dictionary")
End Function

Private Function mp_TryGetDictionaryCountText(ByVal dictObj As Object) As String
    On Error Resume Next
    If dictObj Is Nothing Then
        mp_TryGetDictionaryCountText = "0"
    Else
        mp_TryGetDictionaryCountText = CStr(dictObj.Count)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        mp_TryGetDictionaryCountText = "n/a"
    End If
    On Error GoTo 0
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ResultLayoutItemsSourceRuntime] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
