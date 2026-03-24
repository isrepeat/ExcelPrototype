Attribute VB_Name = "ex_ProfilesEntriesMapper"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_CONFIG_HEADER_ROW As Long = 2
Private Const DEV_CONFIG_MARKER_COL As Long = 1
Private Const DEV_CONFIG_KEY_COL As Long = 2
Private Const DEV_CONFIG_VALUE_COL As Long = 3
Private Const DEV_CONFIG_STYLES_COL As Long = 4
Private Const DEV_CONFIG_COL_COUNT As Long = 4
Private Const XML_ATTR_HIDDEN As String = "hidden"
Private Const XML_ATTR_MUTABLE As String = "mutable"

Public Sub m_WriteSheetValuesToProfile(ByVal ws As Worksheet, ByVal doc As Object, ByVal profileNode As Object)
    Dim entries As Variant
    Dim i As Long
    Dim existingNodes As Collection
    Dim existingNode As Object
    Dim entryQueueByKey As Object
    Dim handledEntryIndexes As Object
    Dim visibleKeySet As Object
    Dim keyIdentity As String
    Dim entryIndex As Long
    Dim isHiddenNode As Boolean
    Dim keyName As String
    Dim vNode As Object
    Dim preservedByKey As Object

    Set preservedByKey = mp_ReadPreservedByKey(profileNode)
    Set existingNodes = mp_ReadProfileVNodeList(profileNode)
    Set handledEntryIndexes = CreateObject("Scripting.Dictionary")
    handledEntryIndexes.CompareMode = 1

    entries = m_ReadConfigTableEntries(ws)
    Set entryQueueByKey = mp_BuildEntryIndexQueueByKey(entries)
    Set visibleKeySet = mp_BuildVisibleKeySet(entries)

    For Each existingNode In existingNodes
        keyName = mp_NodeAttrText(existingNode, "key")
        keyIdentity = mp_NormalizeEntryKey(keyName)
        isHiddenNode = mp_NodeHasAttr(existingNode, XML_ATTR_HIDDEN)

        If isHiddenNode Then
            If Len(keyIdentity) > 0 Then
                If visibleKeySet.Exists(keyIdentity) Then
                    profileNode.removeChild existingNode
                    GoTo ContinueExistingNode
                End If
            End If
            GoTo ContinueExistingNode
        End If

        entryIndex = mp_DequeueEntryIndexForKey(entryQueueByKey, keyIdentity)
        If entryIndex <= 0 Then
            profileNode.removeChild existingNode
            GoTo ContinueExistingNode
        End If

        mp_UpdateVisibleProfileVNode existingNode, entries, entryIndex, preservedByKey
        handledEntryIndexes(CStr(entryIndex)) = True
ContinueExistingNode:
    Next existingNode

    If mp_ArrayHasItems(entries) Then
        For i = LBound(entries, 1) To UBound(entries, 1)
            If handledEntryIndexes.Exists(CStr(i)) Then GoTo ContinueAppendRemaining
            Set vNode = mp_CreateVisibleProfileVNode(doc, entries, i, preservedByKey)
            profileNode.appendChild vNode
ContinueAppendRemaining:
        Next i
    End If
End Sub

Private Sub mp_UpdateVisibleProfileVNode( _
    ByVal vNode As Object, _
    ByVal entries As Variant, _
    ByVal entryIndex As Long, _
    ByVal preservedByKey As Object _
)
    Dim keyName As String
    Dim markerText As String
    Dim entryValue As String
    Dim preservedItem As Object

    keyName = CStr(entries(entryIndex, DEV_CONFIG_KEY_COL))
    markerText = Trim$(CStr(entries(entryIndex, DEV_CONFIG_MARKER_COL)))
    entryValue = CStr(entries(entryIndex, DEV_CONFIG_VALUE_COL))

    vNode.setAttribute "key", keyName
    mp_SetOrClearNodeAttr vNode, "type", markerText
    mp_RemoveNodeAttr vNode, XML_ATTR_HIDDEN

    If Not preservedByKey Is Nothing Then
        If preservedByKey.Exists(keyName) Then
            Set preservedItem = preservedByKey(keyName)
            If CBool(preservedItem("HasMutable")) Then
                mp_SetOrClearNodeAttr vNode, XML_ATTR_MUTABLE, CStr(preservedItem("MutableAttrValue"))
            Else
                mp_RemoveNodeAttr vNode, XML_ATTR_MUTABLE
            End If
            vNode.Text = entryValue
            Exit Sub
        End If
    End If

    mp_RemoveNodeAttr vNode, XML_ATTR_MUTABLE
    vNode.Text = entryValue
End Sub

Private Function mp_CreateVisibleProfileVNode( _
    ByVal doc As Object, _
    ByVal entries As Variant, _
    ByVal entryIndex As Long, _
    ByVal preservedByKey As Object _
) As Object
    Dim vNode As Object
    Dim keyName As String
    Dim preservedItem As Object

    Set vNode = doc.createNode(1, "v", PROFILES_NS)
    If Len(Trim$(CStr(entries(entryIndex, DEV_CONFIG_MARKER_COL)))) > 0 Then
        vNode.setAttribute "type", CStr(entries(entryIndex, DEV_CONFIG_MARKER_COL))
    End If

    keyName = CStr(entries(entryIndex, DEV_CONFIG_KEY_COL))
    vNode.setAttribute "key", keyName

    If Not preservedByKey Is Nothing Then
        If preservedByKey.Exists(keyName) Then
            Set preservedItem = preservedByKey(keyName)
            If CBool(preservedItem("HasMutable")) Then
                vNode.setAttribute XML_ATTR_MUTABLE, CStr(preservedItem("MutableAttrValue"))
            End If
            vNode.Text = CStr(entries(entryIndex, DEV_CONFIG_VALUE_COL))
        Else
            vNode.Text = CStr(entries(entryIndex, DEV_CONFIG_VALUE_COL))
        End If
    Else
        vNode.Text = CStr(entries(entryIndex, DEV_CONFIG_VALUE_COL))
    End If

    Set mp_CreateVisibleProfileVNode = vNode
End Function

Private Sub mp_SetOrClearNodeAttr(ByVal node As Object, ByVal attrName As String, ByVal attrValue As String)
    attrValue = Trim$(CStr(attrValue))
    If Len(attrValue) = 0 Then
        mp_RemoveNodeAttr node, attrName
        Exit Sub
    End If
    node.setAttribute attrName, attrValue
End Sub

Private Sub mp_RemoveNodeAttr(ByVal node As Object, ByVal attrName As String)
    Dim attrNode As Object

    On Error Resume Next
    Set attrNode = node.selectSingleNode("@*[local-name()='" & attrName & "']")
    If Not attrNode Is Nothing Then
        node.Attributes.removeNamedItem attrNode.nodeName
    End If
    On Error GoTo 0
End Sub

Private Function mp_ReadProfileVNodeList(ByVal profileNode As Object) As Collection
    Dim result As New Collection
    Dim nodes As Object
    Dim i As Long

    If profileNode Is Nothing Then
        Set mp_ReadProfileVNodeList = result
        Exit Function
    End If

    Set nodes = profileNode.selectNodes("p:v")
    If nodes Is Nothing Then
        Set mp_ReadProfileVNodeList = result
        Exit Function
    End If

    For i = 0 To nodes.Length - 1
        result.Add nodes.Item(i)
    Next i

    Set mp_ReadProfileVNodeList = result
End Function

Private Function mp_BuildEntryIndexQueueByKey(ByVal entries As Variant) As Object
    Dim result As Object
    Dim keyName As String
    Dim queue As Collection
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not mp_ArrayHasItems(entries) Then
        Set mp_BuildEntryIndexQueueByKey = result
        Exit Function
    End If

    For i = LBound(entries, 1) To UBound(entries, 1)
        keyName = mp_NormalizeEntryKey(CStr(entries(i, DEV_CONFIG_KEY_COL)))
        If Not result.Exists(keyName) Then
            Set queue = New Collection
            result.Add keyName, queue
        End If
        Set queue = result(keyName)
        queue.Add CLng(i)
    Next i

    Set mp_BuildEntryIndexQueueByKey = result
End Function

Private Function mp_BuildVisibleKeySet(ByVal entries As Variant) As Object
    Dim result As Object
    Dim keyName As String
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not mp_ArrayHasItems(entries) Then
        Set mp_BuildVisibleKeySet = result
        Exit Function
    End If

    For i = LBound(entries, 1) To UBound(entries, 1)
        keyName = mp_NormalizeEntryKey(CStr(entries(i, DEV_CONFIG_KEY_COL)))
        If Len(keyName) = 0 Then GoTo ContinueEntry
        result(keyName) = True
ContinueEntry:
    Next i

    Set mp_BuildVisibleKeySet = result
End Function

Private Function mp_HasQueuedEntryForKey(ByVal queueByKey As Object, ByVal keyName As String) As Boolean
    Dim queue As Collection

    If queueByKey Is Nothing Then Exit Function
    keyName = mp_NormalizeEntryKey(keyName)
    If Not queueByKey.Exists(keyName) Then Exit Function

    Set queue = queueByKey(keyName)
    mp_HasQueuedEntryForKey = (queue.Count > 0)
End Function

Private Function mp_DequeueEntryIndexForKey(ByVal queueByKey As Object, ByVal keyName As String) As Long
    Dim queue As Collection

    If Not mp_HasQueuedEntryForKey(queueByKey, keyName) Then Exit Function

    keyName = mp_NormalizeEntryKey(keyName)
    Set queue = queueByKey(keyName)
    mp_DequeueEntryIndexForKey = CLng(queue(1))
    queue.Remove 1
End Function

Private Function mp_NormalizeEntryKey(ByVal keyName As String) As String
    mp_NormalizeEntryKey = Trim$(CStr(keyName))
End Function

Private Function mp_ReadPreservedByKey(ByVal profileNode As Object) As Object
    Dim result As Object
    Dim nodes As Object
    Dim node As Object
    Dim keyName As String
    Dim item As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If profileNode Is Nothing Then
        Set mp_ReadPreservedByKey = result
        Exit Function
    End If

    Set nodes = profileNode.selectNodes("p:v")
    If nodes Is Nothing Then
        Set mp_ReadPreservedByKey = result
        Exit Function
    End If

    For Each node In nodes
        keyName = mp_NodeAttrText(node, "key")
        If Len(keyName) = 0 Then GoTo ContinueNode

        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = 1
        item("HasMutable") = mp_NodeHasAttr(node, XML_ATTR_MUTABLE)
        item("MutableAttrValue") = mp_NodeAttrText(node, XML_ATTR_MUTABLE)
        Set result(keyName) = item

ContinueNode:
    Next node

    Set mp_ReadPreservedByKey = result
End Function

Private Function mp_ReadHiddenNodes(ByVal profileNode As Object) As Collection
    Dim result As New Collection
    Dim nodes As Object
    Dim node As Object
    Dim item As Object

    If profileNode Is Nothing Then
        Set mp_ReadHiddenNodes = result
        Exit Function
    End If

    Set nodes = profileNode.selectNodes("p:v")
    If nodes Is Nothing Then
        Set mp_ReadHiddenNodes = result
        Exit Function
    End If

    For Each node In nodes
        If Not mp_NodeHasAttr(node, XML_ATTR_HIDDEN) Then GoTo ContinueNode

        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = 1
        item("Type") = mp_NodeAttrText(node, "type")
        item("Key") = mp_NodeAttrText(node, "key")
        item("Value") = CStr(node.Text)
        item("HiddenAttrValue") = mp_NodeAttrText(node, XML_ATTR_HIDDEN)
        item("HasMutable") = mp_NodeHasAttr(node, XML_ATTR_MUTABLE)
        item("MutableAttrValue") = mp_NodeAttrText(node, XML_ATTR_MUTABLE)
        result.Add item
ContinueNode:
    Next node

    Set mp_ReadHiddenNodes = result
End Function

Public Function m_ReadProfileEntries(ByVal ws As Worksheet, ByVal profileNode As Object) As Variant
    Dim nodes As Object
    Dim hasKeyFormat As Boolean
    Dim i As Long
    Dim node As Object
    Dim entries() As Variant
    Dim visibleCount As Long
    Dim writeIndex As Long

    Set nodes = profileNode.selectNodes("p:v")
    If nodes Is Nothing Then
        m_ReadProfileEntries = Array()
        Exit Function
    End If
    If nodes.Length = 0 Then
        m_ReadProfileEntries = Array()
        Exit Function
    End If

    hasKeyFormat = False
    For i = 0 To nodes.Length - 1
        If Len(mp_NodeAttrText(nodes.Item(i), "key")) > 0 _
           Or Len(mp_NodeAttrText(nodes.Item(i), "type")) > 0 Then
            hasKeyFormat = True
            Exit For
        End If
    Next i

    If hasKeyFormat Then
        For i = 0 To nodes.Length - 1
            If Not mp_NodeHasAttr(nodes.Item(i), XML_ATTR_HIDDEN) Then
                visibleCount = visibleCount + 1
            End If
        Next i

        If visibleCount <= 0 Then
            m_ReadProfileEntries = Array()
            Exit Function
        End If

        ReDim entries(1 To visibleCount, 1 To DEV_CONFIG_COL_COUNT)
        For i = 0 To nodes.Length - 1
            Set node = nodes.Item(i)
            If mp_NodeHasAttr(node, XML_ATTR_HIDDEN) Then GoTo ContinueVisible

            writeIndex = writeIndex + 1
            entries(writeIndex, DEV_CONFIG_MARKER_COL) = mp_NodeAttrText(node, "type")
            entries(writeIndex, DEV_CONFIG_KEY_COL) = mp_NodeAttrText(node, "key")
            entries(writeIndex, DEV_CONFIG_VALUE_COL) = CStr(node.Text)
            entries(writeIndex, DEV_CONFIG_STYLES_COL) = vbNullString
            ex_ConfigTableStore.m_NormalizeLegacyMarkerEntry entries, writeIndex
ContinueVisible:
        Next i
        m_ReadProfileEntries = entries
        Exit Function
    End If

    m_ReadProfileEntries = mp_ReadLegacyProfileEntries(ws, nodes)
End Function

Public Function m_ReadConfigTableEntries(ByVal ws As Worksheet) As Variant
    Dim tbl As ListObject
    Dim values As Variant
    Dim entries() As Variant
    Dim rowCount As Long
    Dim i As Long

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        m_ReadConfigTableEntries = Array()
        Exit Function
    End If

    rowCount = ex_ConfigTableStore.m_GetTableDataRowCount(tbl)
    If rowCount = 0 Then
        m_ReadConfigTableEntries = Array()
        Exit Function
    End If

    values = mp_ReadConfigTableValues(tbl)
    ReDim entries(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
    For i = 1 To rowCount
        entries(i, DEV_CONFIG_MARKER_COL) = CStr(values(i, DEV_CONFIG_MARKER_COL))
        entries(i, DEV_CONFIG_KEY_COL) = CStr(values(i, DEV_CONFIG_KEY_COL))
        entries(i, DEV_CONFIG_VALUE_COL) = CStr(values(i, DEV_CONFIG_VALUE_COL))
        entries(i, DEV_CONFIG_STYLES_COL) = CStr(values(i, DEV_CONFIG_STYLES_COL))
        ex_ConfigTableStore.m_NormalizeLegacyMarkerEntry entries, i
    Next i

    m_ReadConfigTableEntries = entries
End Function

Private Function mp_ReadLegacyProfileEntries(ByVal ws As Worksheet, ByVal nodes As Object) As Variant
    Dim tbl As ListObject
    Dim tableValues As Variant
    Dim rowCount As Long
    Dim entries() As Variant
    Dim i As Long
    Dim rowAttr As String
    Dim entryIndex As Long
    Dim node As Object
    Dim maxIndex As Long

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        mp_ReadLegacyProfileEntries = Array()
        Exit Function
    End If

    rowCount = ex_ConfigTableStore.m_GetTableDataRowCount(tbl)
    maxIndex = mp_GetMaxLegacyIndex(nodes)
    If maxIndex > rowCount Then
        rowCount = maxIndex
        ex_ConfigTableStore.m_ResizeConfigTableRows ws, tbl, rowCount
    End If
    If rowCount = 0 Then
        mp_ReadLegacyProfileEntries = Array()
        Exit Function
    End If

    tableValues = mp_ReadConfigTableValues(tbl)
    ReDim entries(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)

    For i = 1 To rowCount
        entries(i, DEV_CONFIG_MARKER_COL) = CStr(tableValues(i, DEV_CONFIG_MARKER_COL))
        entries(i, DEV_CONFIG_KEY_COL) = CStr(tableValues(i, DEV_CONFIG_KEY_COL))
        entries(i, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(i, DEV_CONFIG_STYLES_COL) = vbNullString
    Next i

    For i = 0 To nodes.Length - 1
        Set node = nodes.Item(i)
        rowAttr = mp_NodeAttrText(node, "row")
        If Len(rowAttr) > 0 And IsNumeric(rowAttr) Then
            entryIndex = CLng(rowAttr) - DEV_CONFIG_HEADER_ROW
            If entryIndex >= 1 And entryIndex <= rowCount Then
                entries(entryIndex, DEV_CONFIG_VALUE_COL) = CStr(node.Text)
            End If
        End If
    Next i

    mp_ReadLegacyProfileEntries = entries
End Function

Private Function mp_GetMaxLegacyIndex(ByVal nodes As Object) As Long
    Dim i As Long
    Dim rowAttr As String
    Dim idx As Long

    For i = 0 To nodes.Length - 1
        rowAttr = mp_NodeAttrText(nodes.Item(i), "row")
        If Len(rowAttr) > 0 And IsNumeric(rowAttr) Then
            idx = CLng(rowAttr) - DEV_CONFIG_HEADER_ROW
            If idx > mp_GetMaxLegacyIndex Then mp_GetMaxLegacyIndex = idx
        End If
    Next i
End Function

Private Function mp_ReadConfigTableValues(ByVal tbl As ListObject) As Variant
    Dim rowCount As Long
    Dim rawValues As Variant
    Dim values() As Variant

    rowCount = ex_ConfigTableStore.m_GetTableDataRowCount(tbl)
    If rowCount = 0 Then
        mp_ReadConfigTableValues = Array()
        Exit Function
    End If

    rawValues = tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value
    If rowCount = 1 Then
        ReDim values(1 To 1, 1 To DEV_CONFIG_COL_COUNT)
        values(1, 1) = rawValues(1, 1)
        values(1, 2) = rawValues(1, 2)
        values(1, 3) = rawValues(1, 3)
        values(1, 4) = rawValues(1, 4)
        mp_ReadConfigTableValues = values
        Exit Function
    End If

    mp_ReadConfigTableValues = rawValues
End Function

Private Function mp_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    mp_NodeAttrText = CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        mp_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

Private Function mp_NodeHasAttr(ByVal node As Object, ByVal attrName As String) As Boolean
    On Error Resume Next
    mp_NodeHasAttr = Not node.selectSingleNode("@*[local-name()='" & attrName & "']") Is Nothing
    If Err.Number <> 0 Then
        Err.Clear
        mp_NodeHasAttr = False
    End If
    On Error GoTo 0
End Function

Private Function mp_ArrayHasItems(ByVal values As Variant) As Boolean
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayHasItems = (UBound(values) >= LBound(values))
    End If
    Exit Function
EH:
    mp_ArrayHasItems = False
End Function
