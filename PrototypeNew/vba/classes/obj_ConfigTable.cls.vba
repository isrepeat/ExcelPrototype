VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigTable"
Option Explicit

Private m_Items As Collection

Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub

' //
' // API
' //
Public Property Get Items() As Collection
    private_EnsureItems
    Set Items = m_Items
End Property

Public Property Get Count() As Long
    private_EnsureItems
    Count = m_Items.Count
End Property

Public Sub Clear()
    Set m_Items = New Collection
End Sub

Public Function AddItem(ByVal cfgItem As obj_ConfigEntry) As Boolean
    private_EnsureItems

    If cfgItem Is Nothing Then
        VBA.MsgBox "ConfigTable: item is not specified.", VBA.vbExclamation
        Exit Function
    End If

    m_Items.Add cfgItem
    AddItem = True
End Function

Public Function AddRow( _
    ByVal attrText As String, _
    ByVal keyText As String, _
    ByVal valueText As String _
) As Boolean
    Dim cfgItem As obj_ConfigEntry

    Set cfgItem = New obj_ConfigEntry
    cfgItem.Attr = VBA.CStr(attrText)
    cfgItem.Key = VBA.CStr(keyText)
    cfgItem.Value = VBA.CStr(valueText)

    AddRow = Me.AddItem(cfgItem)
End Function

Public Function TryLoadFromXmlNode( _
    ByVal profileNode As Object, _
    Optional ByVal clearBefore As Boolean = True _
) As Boolean
    TryLoadFromXmlNode = private_TryLoadFromXmlNodeInternal(profileNode, clearBefore)
End Function

Public Function TryAppendFromXmlNode(ByVal profileNode As Object) As Boolean
    TryAppendFromXmlNode = private_TryLoadFromXmlNodeInternal(profileNode, False)
End Function

' //
' // Internal
' //
Private Sub private_EnsureItems()
    If Not m_Items Is Nothing Then Exit Sub
    Set m_Items = New Collection
End Sub

Private Function private_TryLoadFromXmlNodeInternal( _
    ByVal profileNode As Object, _
    ByVal clearBefore As Boolean _
) As Boolean
    Dim rowNodes As Object
    Dim rowNode As Object
    Dim cfgItem As obj_ConfigEntry

    private_EnsureItems

    If profileNode Is Nothing Then
        VBA.MsgBox "ConfigTable: profile XML node is not specified.", VBA.vbExclamation
        Exit Function
    End If

    If clearBefore Then Me.Clear

    If Not private_TryCollectRowNodes(profileNode, rowNodes) Then Exit Function

    If rowNodes Is Nothing Then
        If Not private_TryResolveSingleNodeAsRow(profileNode, cfgItem) Then Exit Function
        If Not cfgItem Is Nothing Then
            If Not Me.AddItem(cfgItem) Then Exit Function
            private_TryLoadFromXmlNodeInternal = True
            Exit Function
        End If

        VBA.MsgBox "ConfigTable: profile node does not contain config rows.", VBA.vbExclamation
        Exit Function
    End If

    If rowNodes.Length = 0 Then
        If Not private_TryResolveSingleNodeAsRow(profileNode, cfgItem) Then Exit Function
        If Not cfgItem Is Nothing Then
            If Not Me.AddItem(cfgItem) Then Exit Function
            private_TryLoadFromXmlNodeInternal = True
            Exit Function
        End If

        VBA.MsgBox "ConfigTable: profile node does not contain config rows.", VBA.vbExclamation
        Exit Function
    End If

    For Each rowNode In rowNodes
        Set cfgItem = Nothing
        If Not private_TryCreateItemFromNode(rowNode, cfgItem) Then Exit Function
        If cfgItem Is Nothing Then GoTo ContinueRow
        If Not Me.AddItem(cfgItem) Then Exit Function
ContinueRow:
    Next rowNode

    private_TryLoadFromXmlNodeInternal = True
End Function

Private Function private_TryCollectRowNodes(ByVal profileNode As Object, ByRef outRowNodes As Object) As Boolean
    On Error GoTo EH_XML

    Set outRowNodes = profileNode.selectNodes("./*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']")
    If Not outRowNodes Is Nothing Then
        If outRowNodes.Length > 0 Then
            private_TryCollectRowNodes = True
            Exit Function
        End If
    End If

    Set outRowNodes = profileNode.selectNodes(".//*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']")
    private_TryCollectRowNodes = True
    Exit Function

EH_XML:
    VBA.MsgBox "ConfigTable: failed to collect row nodes from profile XML: " & Err.Description, VBA.vbExclamation
End Function

Private Function private_TryResolveSingleNodeAsRow(ByVal profileNode As Object, ByRef outItem As obj_ConfigEntry) As Boolean
    Dim keyAttr As String
    Dim keyChildText As String

    Set outItem = Nothing

    keyAttr = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "key")))
    If VBA.Len(keyAttr) > 0 Then
        private_TryResolveSingleNodeAsRow = private_TryCreateItemFromNode(profileNode, outItem)
        Exit Function
    End If

    keyChildText = VBA.vbNullString
    If Not private_TryReadChildNodeText(profileNode, "key", keyChildText) Then Exit Function
    If VBA.Len(VBA.Trim$(keyChildText)) > 0 Then
        private_TryResolveSingleNodeAsRow = private_TryCreateItemFromNode(profileNode, outItem)
        Exit Function
    End If

    private_TryResolveSingleNodeAsRow = True
End Function

Private Function private_TryCreateItemFromNode(ByVal rowNode As Object, ByRef outItem As obj_ConfigEntry) As Boolean
    Dim attrText As String
    Dim keyText As String
    Dim valueText As String
    Dim localName As String
    Dim hasElementChildren As Boolean
    Dim childNode As Object

    Set outItem = Nothing
    If rowNode Is Nothing Then Exit Function

    attrText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(rowNode, "attr")))
    If VBA.Len(attrText) = 0 Then attrText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(rowNode, "marker")))

    keyText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(rowNode, "key")))
    If VBA.Len(keyText) = 0 Then keyText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(rowNode, "name")))
    If VBA.Len(keyText) = 0 Then
        If Not private_TryReadChildNodeText(rowNode, "key", keyText) Then Exit Function
    End If
    If VBA.Len(keyText) = 0 Then
        If Not private_TryReadChildNodeText(rowNode, "name", keyText) Then Exit Function
    End If

    valueText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(rowNode, "value")))
    If VBA.Len(valueText) = 0 Then
        If Not private_TryReadChildNodeText(rowNode, "value", valueText) Then Exit Function
    End If

    If VBA.Len(VBA.Trim$(keyText)) = 0 Then
        localName = private_ReadNodeLocalName(rowNode)
        VBA.MsgBox "ConfigTable: row node '" & localName & "' must contain non-empty key.", VBA.vbExclamation
        Exit Function
    End If

    If VBA.Len(valueText) = 0 Then
        hasElementChildren = False
        For Each childNode In rowNode.ChildNodes
            If Not childNode Is Nothing Then
                If VBA.CLng(childNode.nodeType) = 1 Then
                    hasElementChildren = True
                    Exit For
                End If
            End If
        Next childNode

        If Not hasElementChildren Then
            valueText = VBA.Trim$(VBA.CStr(rowNode.Text))
        End If
    End If

    Set outItem = New obj_ConfigEntry
    outItem.Attr = attrText
    outItem.Key = keyText
    outItem.Value = valueText

    private_TryCreateItemFromNode = True
End Function

Private Function private_TryReadChildNodeText( _
    ByVal parentNode As Object, _
    ByVal childLocalName As String, _
    ByRef outText As String _
) As Boolean
    Dim childNode As Object

    outText = VBA.vbNullString
    If parentNode Is Nothing Then
        private_TryReadChildNodeText = True
        Exit Function
    End If

    On Error GoTo EH_XML
    Set childNode = parentNode.selectSingleNode("./*[local-name()='" & childLocalName & "']")
    On Error GoTo 0

    If Not childNode Is Nothing Then outText = VBA.Trim$(VBA.CStr(childNode.Text))
    private_TryReadChildNodeText = True
    Exit Function

EH_XML:
    VBA.MsgBox "ConfigTable: failed to read child node '" & childLocalName & "': " & Err.Description, VBA.vbExclamation
End Function

Private Function private_ReadNodeLocalName(ByVal nodeObj As Object) As String
    On Error Resume Next
    private_ReadNodeLocalName = VBA.CStr(nodeObj.baseName)
    If Err.Number <> 0 Then
        Err.Clear
        private_ReadNodeLocalName = VBA.TypeName(nodeObj)
    End If
    On Error GoTo 0

    private_ReadNodeLocalName = VBA.Trim$(private_ReadNodeLocalName)
    If VBA.Len(private_ReadNodeLocalName) = 0 Then private_ReadNodeLocalName = VBA.TypeName(nodeObj)
End Function
