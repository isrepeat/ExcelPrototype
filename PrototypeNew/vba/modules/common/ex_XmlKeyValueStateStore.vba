Attribute VB_Name = "ex_XmlKeyValueStateStore"
Option Explicit

' Универсальное key/value хранилище поверх CustomXMLPart.
' Позволяет читать/писать valueAttr у entry-узла, найденного по keyAttr=keyValue.

' //
' // API
' //
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

    namespaceUri = VBA.Trim$(namespaceUri)
    rootNodeName = VBA.LCase$(VBA.Trim$(rootNodeName))
    entryNodeName = VBA.LCase$(VBA.Trim$(entryNodeName))
    keyAttrName = VBA.Trim$(keyAttrName)
    valueAttrName = VBA.Trim$(valueAttrName)
    keyValue = VBA.Trim$(keyValue)

    If VBA.Len(namespaceUri) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: namespace is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(rootNodeName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: root node name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(entryNodeName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: entry node name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(keyAttrName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: key attr name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(valueAttrName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: value attr name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(keyValue) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: key value is empty.", VBA.vbExclamation
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(namespaceUri, partObj) Then Exit Function
    If partObj Is Nothing Then
        outValue = VBA.vbNullString
        m_TryGetValue = True
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryLoadPartDom(partObj, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        outValue = VBA.vbNullString
        m_TryGetValue = True
        Exit Function
    End If
    If VBA.LCase$(VBA.CStr(rootNode.baseName)) <> rootNodeName Then
        VBA.MsgBox "XmlKeyValueStateStore: unexpected root node '" & VBA.CStr(rootNode.baseName) & "'. Expected '" & rootNodeName & "'.", VBA.vbExclamation
        Exit Function
    End If

    Set entryNode = private_FindEntryByKey(rootNode, entryNodeName, keyAttrName, keyValue)
    If entryNode Is Nothing Then
        outValue = VBA.vbNullString
        m_TryGetValue = True
        Exit Function
    End If

    outValue = VBA.Trim$(VBA.CStr(entryNode.getAttribute(valueAttrName)))
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

    namespaceUri = VBA.Trim$(namespaceUri)
    rootNodeName = VBA.LCase$(VBA.Trim$(rootNodeName))
    entryNodeName = VBA.LCase$(VBA.Trim$(entryNodeName))
    keyAttrName = VBA.Trim$(keyAttrName)
    valueAttrName = VBA.Trim$(valueAttrName)
    keyValue = VBA.Trim$(keyValue)
    valueText = VBA.Trim$(valueText)

    If VBA.Len(namespaceUri) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: namespace is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(rootNodeName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: root node name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(entryNodeName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: entry node name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(keyAttrName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: key attr name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(valueAttrName) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: value attr name is empty.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(keyValue) = 0 Then
        VBA.MsgBox "XmlKeyValueStateStore: key value is empty.", VBA.vbExclamation
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
        VBA.MsgBox "XmlKeyValueStateStore: root node is missing.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.LCase$(VBA.CStr(rootNode.baseName)) <> rootNodeName Then
        VBA.MsgBox "XmlKeyValueStateStore: unexpected root node '" & VBA.CStr(rootNode.baseName) & "'. Expected '" & rootNodeName & "'.", VBA.vbExclamation
        Exit Function
    End If

    Set entryNode = private_FindEntryByKey(rootNode, entryNodeName, keyAttrName, keyValue)

    If VBA.Len(valueText) = 0 Then
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

' //
' // Internal
' //
Private Function private_FindEntryByKey( _
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
        If VBA.LCase$(VBA.CStr(entryNode.baseName)) <> entryNodeName Then GoTo ContinueEntry

        entryKey = VBA.Trim$(VBA.CStr(entryNode.getAttribute(keyAttrName)))
        If VBA.StrComp(entryKey, keyValue, VBA.vbTextCompare) = 0 Then
            Set private_FindEntryByKey = entryNode
            Exit Function
        End If

ContinueEntry:
    Next entryNode
End Function
