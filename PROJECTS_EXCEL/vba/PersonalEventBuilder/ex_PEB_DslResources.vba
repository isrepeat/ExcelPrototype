Attribute VB_Name = "ex_PEB_DslResources"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Load( _
    ByVal root As Object, _
    ByVal knownKeys As String _
) As Object
    Dim resources As Object
    Dim section As Object
    Dim node As Object
    Dim entryNode As Object
    Dim variantNode As Object
    Dim dictionary As Object
    Dim entries As Object
    Dim entry As Object
    Dim variants As Object
    Dim choices As Object
    Dim name As String
    Dim key As String
    Dim field As String
    Dim expected As String
    Dim value As String
    Dim compare As Long
    Dim aliasKey As Variant

    Set resources = private_Dictionary(vbBinaryCompare)
    If root.selectNodes("p:resources").Length > 1 Then
        Err.Raise 5, , "Duplicate DSL resources section."
    End If
    For Each section In root.selectNodes("p:resources")
        private_Attributes section, ""
        private_Children section, "dictionary|itemOrder"
        For Each node In section.selectNodes("p:dictionary")
            private_Attributes node, "name|ignoreCase|normalize|onMissing"
            private_Children node, "entry"
            name = private_Required(node, "name")
            If resources.Exists(name) Then
                Err.Raise 5, , "Duplicate DSL dictionary: " & name
            End If
            value = private_Required(node, "ignoreCase")
            If value <> "true" And value <> "false" Then
                Err.Raise 5, , "Invalid dictionary ignoreCase: " & name
            End If
            compare = vbBinaryCompare
            If value = "true" Then
                compare = vbTextCompare
            End If
            Set dictionary = private_Dictionary(vbBinaryCompare)
            dictionary.Add "Compare", compare
            value = private_Required(node, "normalize")
            If value <> "trim" And value <> "trim,slashSpaces" Then
                Err.Raise 5, , "Unsupported dictionary normalization: " & name
            End If
            dictionary.Add "Normalize", value
            value = private_Required(node, "onMissing")
            If value <> "keep" And value <> "error" Then
                Err.Raise 5, , "Unsupported dictionary onMissing: " & name
            End If
            dictionary.Add "OnMissing", value
            Set entries = private_Dictionary(compare)
            dictionary.Add "Entries", entries
            For Each entryNode In node.selectNodes("p:entry")
                private_Attributes entryNode, "key"
                private_Children entryNode, "variant"
                key = private_Normalize(private_Required(entryNode, "key"), dictionary("Normalize"))
                If Len(key) = 0 Then
                    Err.Raise 5, , "Empty normalized dictionary key: " & name
                End If
                If entries.Exists(key) Then
                    Err.Raise 5, , "Duplicate normalized dictionary key: " & name & "/" & key
                End If
                Set entry = private_Dictionary(vbBinaryCompare)
                Set variants = private_Dictionary(vbBinaryCompare)
                entry.Add "Variants", variants
                For Each variantNode In entryNode.selectNodes("p:variant")
                    private_Attributes variantNode, "field|equals|value"
                    private_Children variantNode, ""
                    value = private_Required(variantNode, "value")
                    field = VBA.vbNullString
                    expected = VBA.vbNullString
                    If Not variantNode.getAttributeNode("field") Is Nothing Then
                        field = VBA.CStr(variantNode.getAttribute("field"))
                    End If
                    If Not variantNode.getAttributeNode("equals") Is Nothing Then
                        expected = VBA.CStr(variantNode.getAttribute("equals"))
                    End If
                    If Len(field) = 0 Then
                        If Not variantNode.getAttributeNode("field") Is Nothing Or _
                                Not variantNode.getAttributeNode("equals") Is Nothing Or entry.Exists("Default") Then
                            Err.Raise 5, , "Invalid or duplicate default variant: " & name & "/" & key
                        End If
                        entry.Add "Default", value
                    Else
                        If InStr(1, knownKeys, "|" & field & "|", vbBinaryCompare) = 0 Then
                            Err.Raise 5, , "Unknown dictionary context field: " & field
                        End If
                        If variantNode.getAttributeNode("equals") Is Nothing Then
                            Err.Raise 5, , "Dictionary variant requires equals: " & field
                        End If
                        If Not variants.Exists(field) Then
                            variants.Add field, private_Dictionary(compare)
                        End If
                        Set choices = variants(field)
                        If choices.Exists(expected) Then
                            Err.Raise 5, , "Duplicate conditional dictionary variant: " & field & "/" & expected
                        End If
                        choices.Add expected, value
                    End If
                Next variantNode
                If Not entry.Exists("Default") Then
                    Err.Raise 5, , "Dictionary entry requires one default variant: " & name & "/" & key
                End If
                entries.Add key, entry
            Next entryNode
            If entries.Count = 0 Then
                Err.Raise 5, , "Empty DSL dictionary: " & name
            End If
            resources.Add name, dictionary
        Next node
        For Each node In section.selectNodes("p:itemOrder")
            private_Attributes node, "name|field|ignoreCase|normalize"
            private_Children node, "item|others"
            name = "@order:" & private_Required(node, "name")
            If resources.Exists(name) Then
                Err.Raise 5, , "Duplicate item order: " & name
            End If
            Set dictionary = private_Dictionary(vbBinaryCompare)
            dictionary.Add "Field", private_Required(node, "field")
            value = private_Required(node, "ignoreCase")
            If value <> "true" And value <> "false" Then
                Err.Raise 5, , "Invalid item order ignoreCase: " & name
            End If
            compare = vbBinaryCompare
            If value = "true" Then
                compare = vbTextCompare
            End If
            Set entries = private_Dictionary(compare)
            value = private_Required(node, "normalize")
            If value <> "none" And value <> "trim" And value <> "trim,slashSpaces" Then
                Err.Raise 5, , "Unsupported item order normalization: " & name
            End If
            dictionary.Add "Normalize", value
            dictionary.Add "Ranks", entries
            compare = 0
            For Each entryNode In node.childNodes
                If entryNode.nodeType = 1 Then
                    compare = compare + 1
                    private_Children entryNode, ""
                    If entryNode.baseName = "others" Then
                        private_Attributes entryNode, ""
                        If entries.Exists("@others") Then
                            Err.Raise 5, , "Duplicate others placeholder: " & name
                        End If
                        entries.Add "@others", compare
                    Else
                        private_Attributes entryNode, "key|aliases"
                        key = private_Required(entryNode, "key")
                        If Not entryNode.getAttributeNode("aliases") Is Nothing Then
                            key = key & "|" & VBA.CStr(entryNode.getAttribute("aliases"))
                        End If
                        For Each aliasKey In VBA.Split(key, "|")
                            expected = private_Normalize(VBA.CStr(aliasKey), dictionary("Normalize"))
                            If Len(expected) = 0 Or VBA.StrComp(expected, "@others", vbTextCompare) = 0 Or entries.Exists(expected) Then
                                Err.Raise 5, , "Duplicate/empty item order key: " & expected
                            End If
                            entries.Add expected, compare
                        Next aliasKey
                    End If
                End If
            Next entryNode
            If Not entries.Exists("@others") Then
                Err.Raise 5, , "Item order requires one others placeholder: " & name
            End If
            resources.Add name, dictionary
        Next node
    Next section
    Set fn_Load = resources
End Function

Public Function fn_OrderItems( _
    ByVal resources As Object, _
    ByVal name As String, _
    ByVal items As Collection _
) As Collection
    Dim ranks As Object
    Dim itemOrder As Object
    Dim result As New Collection
    Dim item As Object
    Dim other As Object
    Dim rank As Long
    Dim otherRank As Long
    Dim index As Long
    Dim inserted As Boolean
    Dim key As String

    If resources Is Nothing Then
        Err.Raise 5, , "DSL resources are not loaded: " & name
    End If
    If Not resources.Exists("@order:" & name) Then
        Err.Raise 5, , "Unknown item order: " & name
    End If
    Set itemOrder = resources("@order:" & name)
    Set ranks = itemOrder("Ranks")
    For Each item In items
        key = private_Normalize(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(item, itemOrder("Field"))), itemOrder("Normalize"))
        rank = ranks("@others")
        If ranks.Exists(key) Then
            rank = ranks(key)
        End If
        index = 0
        inserted = False
        For Each other In result
            index = index + 1
            key = private_Normalize(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(other, itemOrder("Field"))), itemOrder("Normalize"))
            otherRank = ranks("@others")
            If ranks.Exists(key) Then
                otherRank = ranks(key)
            End If
            If rank < otherRank Then
                result.Add item, Before:=index
                inserted = True
                Exit For
            End If
        Next other
        If Not inserted Then
            result.Add item
        End If
    Next item
    Set fn_OrderItems = result
End Function

Public Function fn_Lookup( _
    ByVal resources As Object, _
    ByVal name As String, _
    ByVal value As String, _
    ByVal context As Object _
) As String
    Dim dictionary As Object
    Dim entries As Object
    Dim entry As Object
    Dim variants As Object
    Dim choices As Object
    Dim field As Variant
    Dim key As String
    Dim expected As String
    Dim matched As Boolean

    If resources Is Nothing Then
        Err.Raise 5, , "DSL resources are not loaded: " & name
    End If
    If Not resources.Exists(name) Then
        Err.Raise 5, , "Unknown DSL dictionary: " & name
    End If
    Set dictionary = resources(name)
    Set entries = dictionary("Entries")
    key = private_Normalize(value, dictionary("Normalize"))
    If Not entries.Exists(key) Then
        If dictionary("OnMissing") = "error" Then
            Err.Raise 5, , "Unknown dictionary key: " & name & "/" & key
        End If
        fn_Lookup = value
        Exit Function
    End If
    Set entry = entries(key)
    fn_Lookup = entry("Default")
    Set variants = entry("Variants")
    For Each field In variants.Keys
        expected = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(context, VBA.CStr(field)))
        Set choices = variants(field)
        If choices.Exists(expected) Then
            If matched Then
                Err.Raise 5, , "Ambiguous dictionary variants: " & name & "/" & key
            End If
            matched = True
            fn_Lookup = choices(expected)
        End If
    Next field
End Function

Public Sub fn_ValidateText( _
    ByVal text As String, _
    ByVal resources As Object _
)
    Dim expression As Object
    Dim reference As Object
    Dim name As String

    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "\|lookup\s*:\s*""([^""]+)"""
    For Each reference In expression.Execute(text)
        name = reference.SubMatches(0)
        If Not resources.Exists(name) Then
            Err.Raise 5, , "Unknown DSL dictionary: " & name
        End If
    Next reference
    expression.Pattern = "orderBy\s*:\s*""([^""]+)"""
    For Each reference In expression.Execute(text)
        name = reference.SubMatches(0)
        If Not resources.Exists("@order:" & name) Then
            Err.Raise 5, , "Unknown item order: " & name
        End If
    Next reference
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Dictionary(ByVal compare As Long) As Object
    Set private_Dictionary = VBA.CreateObject("Scripting.Dictionary")
    private_Dictionary.CompareMode = compare
End Function

Private Function private_Normalize( _
    ByVal value As String, _
    ByVal normalization As String _
) As String
    If normalization = "none" Then
        private_Normalize = value
        Exit Function
    End If
    value = VBA.Replace$(value, VBA.ChrW$(160), " ")
    value = VBA.Replace$(value, VBA.ChrW$(8239), " ")
    value = VBA.Replace$(value, VBA.vbTab, " ")
    value = VBA.Trim$(value)
    If normalization = "trim,slashSpaces" Then
        Do While InStr(value, " /") > 0 Or InStr(value, "/ ") > 0
            value = VBA.Replace$(value, " /", "/")
            value = VBA.Replace$(value, "/ ", "/")
        Loop
    End If
    private_Normalize = value
End Function

Private Function private_Required( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    If node.getAttributeNode(name) Is Nothing Then
        Err.Raise 5, , "Missing DSL resource attribute: " & node.nodeName & "/" & name
    End If
    private_Required = VBA.CStr(node.getAttribute(name))
    If Len(VBA.Trim$(private_Required)) = 0 Then
        Err.Raise 5, , "Empty DSL resource attribute: " & node.nodeName & "/" & name
    End If
End Function

Private Sub private_Attributes( _
    ByVal node As Object, _
    ByVal allowed As String _
)
    Dim attributeNode As Object

    For Each attributeNode In node.Attributes
        If InStr(1, "|" & allowed & "|", "|" & attributeNode.nodeName & "|", vbBinaryCompare) = 0 Then
            Err.Raise 5, , "Unknown DSL resource attribute: " & attributeNode.nodeName
        End If
    Next attributeNode
End Sub

Private Sub private_Children( _
    ByVal node As Object, _
    ByVal allowed As String _
)
    Dim child As Object

    For Each child In node.childNodes
        If child.nodeType = 1 Then
            If child.namespaceURI <> "urn:excelprototype:profiles" Or _
                    InStr(1, "|" & allowed & "|", "|" & child.baseName & "|", vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unknown DSL resource element: " & child.nodeName
            End If
        ElseIf Len(VBA.Trim$(child.Text)) > 0 And child.nodeType <> 8 Then
            Err.Raise 5, , "DSL resources cannot contain text: " & node.nodeName
        End If
    Next child
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------