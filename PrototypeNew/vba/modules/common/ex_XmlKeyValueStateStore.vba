Attribute VB_Name = "ex_XmlKeyValueStateStore"
Option Explicit

' Универсальное key/value хранилище поверх CustomXMLPart.
' Позволяет читать/писать valueAttr у entry-узла, найденного по keyAttr=keyValue.

Public Function m_TryGetValue( _
    ByVal namespaceUri As String, _
    ByVal rootNodeName As String, _
    ByVal entryNodeName As String, _
    ByVal keyAttrName As String, _
    ByVal keyValue As String, _
    ByVal valueAttrName As String, _
    ByRef outValue As String _
) As Boolean
    Dim partObj As Object
    Dim dom As Object
    Dim rootNode As Object
    Dim entryNode As Object

    namespaceUri = Trim$(namespaceUri)
    rootNodeName = LCase$(Trim$(rootNodeName))
    entryNodeName = LCase$(Trim$(entryNodeName))
    keyAttrName = Trim$(keyAttrName)
    valueAttrName = Trim$(valueAttrName)
    keyValue = Trim$(keyValue)

    If Len(namespaceUri) = 0 Then
        MsgBox "XmlKeyValueStateStore: namespace is empty.", vbExclamation
        Exit Function
    End If
    If Len(rootNodeName) = 0 Then
        MsgBox "XmlKeyValueStateStore: root node name is empty.", vbExclamation
        Exit Function
    End If
    If Len(entryNodeName) = 0 Then
        MsgBox "XmlKeyValueStateStore: entry node name is empty.", vbExclamation
        Exit Function
    End If
    If Len(keyAttrName) = 0 Then
        MsgBox "XmlKeyValueStateStore: key attr name is empty.", vbExclamation
        Exit Function
    End If
    If Len(valueAttrName) = 0 Then
        MsgBox "XmlKeyValueStateStore: value attr name is empty.", vbExclamation
        Exit Function
    End If
    If Len(keyValue) = 0 Then
        MsgBox "XmlKeyValueStateStore: key value is empty.", vbExclamation
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(namespaceUri, partObj) Then Exit Function
    If partObj Is Nothing Then
        outValue = vbNullString
        m_TryGetValue = True
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryLoadPartDom(partObj, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        outValue = vbNullString
        m_TryGetValue = True
        Exit Function
    End If
    If LCase$(CStr(rootNode.baseName)) <> rootNodeName Then
        MsgBox "XmlKeyValueStateStore: unexpected root node '" & CStr(rootNode.baseName) & "'. Expected '" & rootNodeName & "'.", vbExclamation
        Exit Function
    End If

    Set entryNode = mp_FindEntryByKey(rootNode, entryNodeName, keyAttrName, keyValue)
    If entryNode Is Nothing Then
        outValue = vbNullString
        m_TryGetValue = True
        Exit Function
    End If

    outValue = Trim$(CStr(entryNode.getAttribute(valueAttrName)))
    m_TryGetValue = True
End Function

Public Function m_SetValue( _
    ByVal namespaceUri As String, _
    ByVal rootNodeName As String, _
    ByVal entryNodeName As String, _
    ByVal keyAttrName As String, _
    ByVal keyValue As String, _
    ByVal valueAttrName As String, _
    ByVal valueText As String _
) As Boolean
    Dim partObj As Object
    Dim dom As Object
    Dim rootNode As Object
    Dim entryNode As Object

    namespaceUri = Trim$(namespaceUri)
    rootNodeName = LCase$(Trim$(rootNodeName))
    entryNodeName = LCase$(Trim$(entryNodeName))
    keyAttrName = Trim$(keyAttrName)
    valueAttrName = Trim$(valueAttrName)
    keyValue = Trim$(keyValue)
    valueText = Trim$(valueText)

    If Len(namespaceUri) = 0 Then
        MsgBox "XmlKeyValueStateStore: namespace is empty.", vbExclamation
        Exit Function
    End If
    If Len(rootNodeName) = 0 Then
        MsgBox "XmlKeyValueStateStore: root node name is empty.", vbExclamation
        Exit Function
    End If
    If Len(entryNodeName) = 0 Then
        MsgBox "XmlKeyValueStateStore: entry node name is empty.", vbExclamation
        Exit Function
    End If
    If Len(keyAttrName) = 0 Then
        MsgBox "XmlKeyValueStateStore: key attr name is empty.", vbExclamation
        Exit Function
    End If
    If Len(valueAttrName) = 0 Then
        MsgBox "XmlKeyValueStateStore: value attr name is empty.", vbExclamation
        Exit Function
    End If
    If Len(keyValue) = 0 Then
        MsgBox "XmlKeyValueStateStore: key value is empty.", vbExclamation
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(namespaceUri, partObj) Then Exit Function
    If partObj Is Nothing Then
        If Not ex_CustomXmlPartStore.m_TryCreateEmptyDom(rootNodeName, namespaceUri, dom) Then Exit Function
    Else
        If Not ex_CustomXmlPartStore.m_TryLoadPartDom(partObj, dom) Then Exit Function
    End If

    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        MsgBox "XmlKeyValueStateStore: root node is missing.", vbExclamation
        Exit Function
    End If
    If LCase$(CStr(rootNode.baseName)) <> rootNodeName Then
        MsgBox "XmlKeyValueStateStore: unexpected root node '" & CStr(rootNode.baseName) & "'. Expected '" & rootNodeName & "'.", vbExclamation
        Exit Function
    End If

    Set entryNode = mp_FindEntryByKey(rootNode, entryNodeName, keyAttrName, keyValue)

    If Len(valueText) = 0 Then
        If Not entryNode Is Nothing Then rootNode.removeChild entryNode
    Else
        If entryNode Is Nothing Then
            Set entryNode = dom.createNode(1, entryNodeName, namespaceUri)
            entryNode.setAttribute keyAttrName, keyValue
            rootNode.appendChild entryNode
        End If
        entryNode.setAttribute valueAttrName, valueText
    End If

    If Not ex_CustomXmlPartStore.m_TrySaveDom(dom, partObj) Then Exit Function
    m_SetValue = True
End Function

Private Function mp_FindEntryByKey( _
    ByVal rootNode As Object, _
    ByVal entryNodeName As String, _
    ByVal keyAttrName As String, _
    ByVal keyValue As String _
) As Object
    Dim entries As Object
    Dim entryNode As Object
    Dim entryKey As String

    If rootNode Is Nothing Then Exit Function

    Set entries = rootNode.ChildNodes
    If entries Is Nothing Then Exit Function

    For Each entryNode In entries
        If entryNode.NodeType <> 1 Then GoTo ContinueEntry
        If LCase$(CStr(entryNode.baseName)) <> entryNodeName Then GoTo ContinueEntry

        entryKey = Trim$(CStr(entryNode.getAttribute(keyAttrName)))
        If StrComp(entryKey, keyValue, vbTextCompare) = 0 Then
            Set mp_FindEntryByKey = entryNode
            Exit Function
        End If

ContinueEntry:
    Next entryNode
End Function

