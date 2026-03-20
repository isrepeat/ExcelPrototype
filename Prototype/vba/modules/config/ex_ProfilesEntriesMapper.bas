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
Private Const XML_ATTR_LOCKED_WITH_PLACEHOLDER As String = "lockedWithPlaceholder"
Private Const XML_ATTR_MUTABLE As String = "mutable"

Public Sub m_WriteSheetValuesToProfile(ByVal ws As Worksheet, ByVal doc As Object, ByVal profileNode As Object)
    Dim entries As Variant
    Dim i As Long
    Dim vNode As Object
    Dim child As Object
    Dim keyName As String
    Dim preservedByKey As Object
    Dim preservedItem As Object
    Dim hiddenNodes As Collection
    Dim hiddenItem As Object
    Dim visibleKeys As Object
    Dim hiddenKey As String

    Set preservedByKey = mp_ReadPreservedByKey(profileNode)
    Set hiddenNodes = mp_ReadHiddenNodes(profileNode)
    Set visibleKeys = CreateObject("Scripting.Dictionary")
    visibleKeys.CompareMode = 1

    For Each child In profileNode.selectNodes("p:v")
        profileNode.removeChild child
    Next child

    entries = m_ReadConfigTableEntries(ws)
    If mp_ArrayHasItems(entries) Then
        For i = LBound(entries, 1) To UBound(entries, 1)
            Set vNode = doc.createNode(1, "v", PROFILES_NS)
            If Len(Trim$(CStr(entries(i, DEV_CONFIG_MARKER_COL)))) > 0 Then
                vNode.setAttribute "type", CStr(entries(i, DEV_CONFIG_MARKER_COL))
            End If
            keyName = CStr(entries(i, DEV_CONFIG_KEY_COL))
            vNode.setAttribute "key", keyName

            If Len(Trim$(keyName)) > 0 Then visibleKeys(Trim$(keyName)) = True

            If Not preservedByKey Is Nothing Then
                If preservedByKey.Exists(keyName) Then
                    Set preservedItem = preservedByKey(keyName)
                    If CBool(preservedItem("HasMutable")) Then
                        vNode.setAttribute XML_ATTR_MUTABLE, CStr(preservedItem("MutableAttrValue"))
                    End If
                    If CBool(preservedItem("HasLockedWithPlaceholder")) Then
                        vNode.setAttribute XML_ATTR_LOCKED_WITH_PLACEHOLDER, CStr(preservedItem("LockedWithPlaceholder"))
                        vNode.Text = CStr(preservedItem("PreservedValue"))
                    Else
                        vNode.Text = CStr(entries(i, DEV_CONFIG_VALUE_COL))
                    End If
                Else
                    vNode.Text = CStr(entries(i, DEV_CONFIG_VALUE_COL))
                End If
            Else
                vNode.Text = CStr(entries(i, DEV_CONFIG_VALUE_COL))
            End If
            profileNode.appendChild vNode
        Next i
    End If

    If Not hiddenNodes Is Nothing Then
        For Each hiddenItem In hiddenNodes
            hiddenKey = Trim$(CStr(hiddenItem("Key")))
            If Len(hiddenKey) > 0 Then
                If visibleKeys.Exists(hiddenKey) Then GoTo ContinueHiddenNode
            End If

            Set vNode = doc.createNode(1, "v", PROFILES_NS)
            If Len(Trim$(CStr(hiddenItem("Type")))) > 0 Then
                vNode.setAttribute "type", CStr(hiddenItem("Type"))
            End If
            vNode.setAttribute "key", CStr(hiddenItem("Key"))

            vNode.setAttribute XML_ATTR_HIDDEN, CStr(hiddenItem("HiddenAttrValue"))
            If CBool(hiddenItem("HasLockedWithPlaceholder")) Then
                vNode.setAttribute XML_ATTR_LOCKED_WITH_PLACEHOLDER, CStr(hiddenItem("LockedWithPlaceholder"))
            End If
            If CBool(hiddenItem("HasMutable")) Then
                vNode.setAttribute XML_ATTR_MUTABLE, CStr(hiddenItem("MutableAttrValue"))
            End If
            vNode.Text = CStr(hiddenItem("Value"))
            profileNode.appendChild vNode
ContinueHiddenNode:
        Next hiddenItem
    End If
End Sub

Private Function mp_ReadPreservedByKey(ByVal profileNode As Object) As Object
    Dim result As Object
    Dim nodes As Object
    Dim node As Object
    Dim keyName As String
    Dim placeholderText As String
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
        placeholderText = mp_NodeAttrText(node, XML_ATTR_LOCKED_WITH_PLACEHOLDER)
        If Len(keyName) = 0 Then GoTo ContinueNode

        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = 1
        item("HasLockedWithPlaceholder") = mp_NodeHasAttr(node, XML_ATTR_LOCKED_WITH_PLACEHOLDER)
        item("LockedWithPlaceholder") = placeholderText
        item("HasMutable") = mp_NodeHasAttr(node, XML_ATTR_MUTABLE)
        item("MutableAttrValue") = mp_NodeAttrText(node, XML_ATTR_MUTABLE)
        item("PreservedValue") = CStr(node.Text)
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
        item("HasLockedWithPlaceholder") = mp_NodeHasAttr(node, XML_ATTR_LOCKED_WITH_PLACEHOLDER)
        item("LockedWithPlaceholder") = mp_NodeAttrText(node, XML_ATTR_LOCKED_WITH_PLACEHOLDER)
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
