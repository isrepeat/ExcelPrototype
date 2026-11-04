VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_WordTableTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_document As Object
Private m_styles As Object
Private m_templates As Object
Private m_text As obj_WordTableText
Private m_maxDepth As Long
Private m_maxRows As Long
Private m_maxColumns As Long
Private m_images As Object
Private m_drawingId As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal path As String, _
    ByVal maxDepth As Long, _
    ByVal maxRows As Long, _
    ByVal maxColumns As Long, _
    ByVal months As String, _
    ByRef diagnostic As String, _
    Optional ByVal imagePaths As Object = Nothing _
) As Boolean
    Dim root As Object
    Dim node As Object
    Dim entry As Object
    Dim sharedTemplates As Object
    Dim resources As Object
    Dim key As Variant
    Dim properties As Object
    Dim stream As Object
    Dim image As Object
    Dim binaryNode As Object
    Dim errorDescription As String

    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    diagnostic = VBA.vbNullString
    If maxDepth < 1 Or maxDepth > 64 Or maxRows < 1 Or maxColumns < 1 Or maxColumns > 63 Then
        Err.Raise 5, , "Invalid Word table DSL limits."
    End If
    m_maxDepth = maxDepth
    m_maxRows = maxRows
    m_maxColumns = maxColumns
    Set m_images = private_Dictionary()
    If Not imagePaths Is Nothing Then
        For Each key In imagePaths.Keys
            Set stream = VBA.CreateObject("ADODB.Stream")
            stream.Type = 1
            stream.Open
            stream.LoadFromFile VBA.CStr(imagePaths(key))
            Set binaryNode = VBA.CreateObject("MSXML2.DOMDocument.6.0").createElement("binary")
            binaryNode.DataType = "bin.base64"
            binaryNode.nodeTypedValue = stream.Read
            stream.Close
            Set image = private_Dictionary()
            image.Add "name", "image" & VBA.CStr(m_images.Count + 1) & ".png"
            image.Add "id", "rIdImage" & VBA.CStr(m_images.Count + 1)
            image.Add "base64", VBA.Replace$(VBA.Replace$(binaryNode.Text, VBA.vbCr, ""), VBA.vbLf, "")
            If VBA.Left$(image("base64"), 11) <> "iVBORw0KGgo" Then
                Err.Raise 5, , "Table image must be a PNG: " & key
            End If
            m_images.Add VBA.CStr(key), image
        Next key
    End If
    Set m_document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    m_document.async = False
    m_document.validateOnParse = False
    m_document.resolveExternals = False
    m_document.preserveWhiteSpace = True
    m_document.setProperty "ProhibitDTD", True
    If Not m_document.Load(path) Then
        Err.Raise 5, , path & ": " & m_document.parseError.reason
    End If
    m_document.setProperty "SelectionNamespaces", "xmlns:p='urn:excelprototype:profiles'"
    Set root = m_document.documentElement
    If root.baseName <> "wordTableTemplates" Or root.namespaceURI <> "urn:excelprototype:profiles" Then
        Err.Raise 5, , "Expected wordTableTemplates in urn:excelprototype:profiles."
    End If
    private_Attributes root, "version"
    If private_Required(root, "version") <> "1" Then
        Err.Raise 5, , "Unsupported Word table DSL version."
    End If
    private_Children root, "styles|sharedTemplates|resources|template"
    For Each key In VBA.Array("styles", "sharedTemplates", "resources")
        If root.selectNodes("p:" & key).Length > 1 Then
            Err.Raise 5, , "Duplicate table DSL section: " & key
        End If
    Next key
    Set m_styles = private_Dictionary()
    For Each node In root.selectNodes("p:styles")
        private_Attributes node, ""
        private_Children node, "tableStyle|cellStyle"
        For Each entry In node.ChildNodes
            If entry.NodeType = 1 Then
                private_Attributes entry, "name|basedOn|" & private_StyleKeys(entry.baseName)
                private_Children entry, ""
                m_styles.Add private_Required(entry, "name"), entry
            End If
        Next entry
    Next node
    For Each key In m_styles.Keys
        Set node = m_styles(key)
        Set properties = private_Style(VBA.CStr(key), node.baseName, 0)
        private_ValidateStyle properties, node.baseName
    Next key
    Set sharedTemplates = private_Dictionary()
    For Each node In root.selectNodes("p:sharedTemplates")
        private_Attributes node, ""
        private_Children node, "sharedTemplate"
        For Each entry In node.selectNodes("p:sharedTemplate")
            private_Attributes entry, "id"
            private_Children entry, "text"
            If entry.selectNodes("p:text").Length <> 1 Then
                Err.Raise 5, , "Shared table DSL block requires one text."
            End If
            private_Attributes entry.selectSingleNode("p:text"), ""
            private_Children entry.selectSingleNode("p:text"), ""
            sharedTemplates.Add private_Required(entry, "id"), entry.selectSingleNode("p:text").Text
        Next entry
    Next node
    Set resources = ex_WordDslResources.fn_Load(root, "")
    Set m_text = New obj_WordTableText
    If Not m_text.Initialize(sharedTemplates, maxDepth, months, resources) Then
        Err.Raise 5, , "Cannot initialize table text DSL."
    End If
    Set m_templates = private_Dictionary()
    For Each node In root.selectNodes("p:template")
        private_Attributes node, "name"
        private_Children node, "table"
        If node.selectNodes("p:table").Length <> 1 Then
            Err.Raise 5, , "A table template requires exactly one table."
        End If
        m_templates.Add private_Required(node, "name"), node.selectSingleNode("p:table")
        private_ValidateStructure node.selectSingleNode("p:table"), 0
    Next node
    If m_templates.Count = 0 Then
        Err.Raise 5, , "No Word table templates were found."
    End If
    m_isInitialized = True
    Initialize = True
    Exit Function
Failed:
    errorDescription = VBA.Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
    End If
    Me.Dispose
    On Error GoTo 0
    diagnostic = "Word table DSL: " & errorDescription
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    If Not m_text Is Nothing Then
        m_text.Dispose
    End If
    Set m_text = Nothing
    Set m_document = Nothing
    Set m_styles = Nothing
    Set m_templates = Nothing
    Set m_images = Nothing
End Sub

Public Function TryPrepare( _
    ByVal name As String, _
    ByVal values As Object, _
    ByRef plan As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim tableNode As Object
    Dim node As Object
    Dim columns As Collection
    Dim rows As Collection
    Dim cells As Collection
    Dim occupied As Object
    Dim row As Object
    Dim cell As Object
    Dim properties As Object
    Dim width As Double
    Dim totalWidth As Double
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim endRow As Long
    Dim endColumn As Long
    Dim r As Long
    Dim c As Long
    Dim headerEnded As Boolean

    On Error GoTo Failed
    Set plan = Nothing
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word table DSL is not initialized."
    End If
    If values Is Nothing Then
        Err.Raise 5, , "Word table context is required."
    End If
    If Not m_templates.Exists(name) Then
        Err.Raise 5, , "Unknown Word table template: " & name
    End If
    Set tableNode = m_templates(name)
    Set properties = private_NodeStyle(tableNode, "tableStyle")
    private_ValidateStyle properties, "tableStyle"
    Set columns = New Collection
    Set occupied = private_Dictionary()
    For Each node In tableNode.selectNodes("p:columns/p:column")
        occupied.Add private_Required(node, "name"), True
        width = private_Length(private_Required(node, "width"), False)
        columns.Add width
        totalWidth = totalWidth + width
    Next node
    If columns.Count = 0 Or columns.Count > m_maxColumns Then
        Err.Raise 5, , "Invalid Word table column count."
    End If
    width = totalWidth
    If properties.Exists("width") Then
        width = private_Length(properties("width"), False)
    End If
    If VBA.Abs(width - totalWidth) > 0.1 Then
        Err.Raise 5, , "Table width must equal the sum of column widths."
    End If
    Set rows = New Collection
    m_drawingId = 0
    private_Expand tableNode, values, rows, 0
    If rows.Count = 0 Then
        Err.Raise 5, , "Word table expanded to zero rows."
    End If
    Set occupied = private_Dictionary()
    Set cells = New Collection
    For rowIndex = 1 To rows.Count
        Set row = rows(rowIndex)
        If row("header") And headerEnded Then
            Err.Raise 5, , "Repeated table headers must be consecutive first rows."
        End If
        If Not row("header") Then
            headerEnded = True
        End If
        columnIndex = 1
        For Each cell In row("cells")
            Do While occupied.Exists(VBA.CStr(rowIndex) & ":" & VBA.CStr(columnIndex))
                columnIndex = columnIndex + 1
            Loop
            endColumn = columnIndex + cell("columnSpan") - 1
            endRow = rowIndex + cell("rowSpan") - 1
            If endColumn > columns.Count Or endRow > rows.Count Then
                Err.Raise 5, , "Cell span exceeds table grid."
            End If
            For r = rowIndex To endRow
                If rows(r)("header") <> row("header") Then
                    Err.Raise 5, , "A merged cell cannot cross the header boundary."
                End If
                For c = columnIndex To endColumn
                    occupied.Add VBA.CStr(r) & ":" & VBA.CStr(c), True
                Next c
            Next r
            cell.Add "row", rowIndex
            cell.Add "column", columnIndex
            cells.Add cell
            columnIndex = endColumn + 1
        Next cell
        For c = 1 To columns.Count
            If Not occupied.Exists(VBA.CStr(rowIndex) & ":" & VBA.CStr(c)) Then
                Err.Raise 5, , "Incomplete Word table row: " & rowIndex
            End If
        Next c
    Next rowIndex
    Set plan = private_Dictionary()
    plan.Add "name", name
    plan.Add "columns", columns
    plan.Add "rows", rows
    plan.Add "cells", cells
    plan.Add "style", properties
    plan.Add "width", width
    plan.Add "xml", ex_WordTableXml.fn_Build(plan, m_images)
    TryPrepare = True
    Exit Function
Failed:
    diagnostic = "Word table DSL: " & VBA.Err.Description
    Set plan = Nothing
End Function

Public Function TryInsert( _
    ByVal document As Object, _
    ByVal plan As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim first As Object
    Dim last As Object
    Dim marker As String

    On Error GoTo Failed
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word table DSL is not initialized."
    End If
    If plan Is Nothing Then
        Err.Raise 5, , "A prepared Word table is required."
    End If
    marker = "{\exportTable:" & plan("name")
    Set first = private_FindMarker(document, marker & "_Begin}")
    Set last = private_FindMarker(document, marker & "_End}")
    If first.End > last.Start Then
        Err.Raise 5, , "Word table markers are reversed or overlap."
    End If
    first.SetRange first.Paragraphs(1).Range.Start, last.Paragraphs(1).Range.End
    ' Insert a validated Flat OPC package in one Word call, preserving content outside the markers.
    first.InsertXML plan("xml")
    TryInsert = True
    Exit Function
Failed:
    diagnostic = "Word table DSL: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_Dictionary() As Object
    Set private_Dictionary = VBA.CreateObject("Scripting.Dictionary")
End Function

Private Function private_Required( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    If node.getAttributeNode(name) Is Nothing Then
        Err.Raise 5, , "Required table DSL attribute: " & node.baseName & "." & name
    End If
    private_Required = node.getAttribute(name)
    If VBA.Len(VBA.Trim$(private_Required)) = 0 Then
        Err.Raise 5, , "Required table DSL attribute: " & node.baseName & "." & name
    End If
End Function

Private Sub private_Attributes( _
    ByVal node As Object, _
    ByVal allowed As String _
)
    Dim attributeNode As Object

    For Each attributeNode In node.Attributes
        If attributeNode.namespaceURI <> "http://www.w3.org/2000/xmlns/" And attributeNode.nodeName <> "xmlns" Then
            If VBA.InStr(1, "|" & allowed & "|", "|" & attributeNode.nodeName & "|", VBA.vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unknown table DSL attribute: " & node.baseName & "." & attributeNode.nodeName
            End If
        End If
    Next attributeNode
End Sub

Private Sub private_Children( _
    ByVal node As Object, _
    ByVal allowed As String _
)
    Dim child As Object

    For Each child In node.ChildNodes
        If child.NodeType = 1 Then
            If child.namespaceURI <> "urn:excelprototype:profiles" Or _
                    VBA.InStr(1, "|" & allowed & "|", "|" & child.baseName & "|", VBA.vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unexpected table DSL element: " & child.nodeName
            End If
        ElseIf child.NodeType = 3 Or child.NodeType = 4 Then
            If node.baseName <> "text" And VBA.Len(VBA.Trim$(VBA.Replace$(VBA.Replace$(child.Text, VBA.vbCr, ""), VBA.vbLf, ""))) > 0 Then
                Err.Raise 5, , "Table DSL text must be inside text elements."
            End If
        End If
    Next child
End Sub

Private Function private_StyleKeys(ByVal kind As String) As String
    If kind = "tableStyle" Then
        private_StyleKeys = "width|alignment|layout|borders|cellPadding|cellPaddingTop|cellPaddingLeft|cellPaddingBottom|cellPaddingRight"
    Else
        private_StyleKeys = "fontName|fontSize|fontBold|fontItalic|horizontal|vertical|spaceBefore|spaceAfter|lineSpacing|firstLineIndent|borderBottom"
    End If
End Function

Private Function private_Style( _
    ByVal name As String, _
    ByVal kind As String, _
    ByVal depth As Long _
) As Object
    Dim node As Object
    Dim result As Object
    Dim attributeNode As Object

    If depth > m_maxDepth Then
        Err.Raise 5, , "Table style inheritance cycle or depth exceeded: " & name
    End If
    If Not m_styles.Exists(name) Then
        Err.Raise 5, , "Unknown table DSL style: " & name
    End If
    Set node = m_styles(name)
    If node.baseName <> kind Then
        Err.Raise 5, , "Incorrect table DSL style kind: " & name
    End If
    Set result = private_Dictionary()
    If Not node.getAttributeNode("basedOn") Is Nothing Then
        Set result = private_Style(private_Required(node, "basedOn"), kind, depth + 1)
    End If
    For Each attributeNode In node.Attributes
        If attributeNode.nodeName <> "name" And attributeNode.nodeName <> "basedOn" Then
            result(attributeNode.nodeName) = attributeNode.Text
        End If
    Next attributeNode
    Set private_Style = result
End Function

Private Function private_NodeStyle( _
    ByVal node As Object, _
    ByVal kind As String _
) As Object
    Set private_NodeStyle = private_Dictionary()
    If Not node.getAttributeNode("style") Is Nothing Then
        Set private_NodeStyle = private_Style(private_Required(node, "style"), kind, 0)
    End If
End Function

Private Function private_Length( _
    ByVal value As String, _
    ByVal allowZero As Boolean _
) As Double
    Dim number As Double
    Dim factor As Double
    Dim rx As Object

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^[0-9]+([.][0-9]+)?(mm|cm|pt)$"
    If Not rx.Test(value) Then
        Err.Raise 5, , "Expected a length in mm, cm or pt: " & value
    End If
    number = VBA.Val(value)
    Select Case VBA.Right$(value, 2)
        Case "mm"
            factor = 72 / 25.4
        Case "cm"
            factor = 72 / 2.54
        Case "pt"
            factor = 1
    End Select
    If number = 0 And Not allowZero Then
        Err.Raise 5, , "A positive length is required: " & value
    End If
    If number * factor > 1584 Then
        Err.Raise 5, , "Word table length exceeds 22 inches: " & value
    End If
    private_Length = number * factor
End Function

Private Function private_Boolean(ByVal value As String) As Boolean
    If value <> "true" And value <> "false" Then
        Err.Raise 5, , "Expected true or false: " & value
    End If
    private_Boolean = (value = "true")
End Function

Private Function private_Alignment(ByVal value As String) As Long
    Select Case value
        Case "left"
            private_Alignment = 0
        Case "center"
            private_Alignment = 1
        Case "right"
            private_Alignment = 2
        Case "justify"
            private_Alignment = 3
        Case Else
            Err.Raise 5, , "Invalid Word table alignment: " & value
    End Select
End Function

Private Sub private_ValidateStyle( _
    ByVal properties As Object, _
    ByVal kind As String _
)
    Dim key As Variant
    Dim number As Double
    Dim flag As Boolean

    For Each key In properties.Keys
        Select Case VBA.CStr(key)
            Case "width"
                number = private_Length(properties(key), False)
            Case "cellPadding", "cellPaddingTop", "cellPaddingLeft", "cellPaddingBottom", "cellPaddingRight", "spaceBefore", "spaceAfter", "firstLineIndent"
                number = private_Length(properties(key), True)
            Case "fontSize"
                number = private_FontSize(properties(key))
            Case "lineSpacing"
                number = private_Length(properties(key), False)
            Case "fontBold", "fontItalic"
                flag = private_Boolean(properties(key))
            Case "alignment", "horizontal"
                number = private_Alignment(properties(key))
                If kind = "tableStyle" And number = 3 Then
                    Err.Raise 5, , "Tables cannot use justify alignment."
                End If
            Case "vertical"
                If VBA.InStr(1, "|top|center|bottom|", "|" & properties(key) & "|", VBA.vbBinaryCompare) = 0 Then
                    Err.Raise 5, , "Invalid vertical alignment."
                End If
            Case "layout"
                If properties(key) <> "fixed" Then
                    Err.Raise 5, , "Word table DSL v1 requires fixed layout."
                End If
            Case "borders", "borderBottom"
                If properties(key) <> "none" And properties(key) <> "single" Then
                    Err.Raise 5, , "Invalid table borders."
                End If
            Case "fontName"
                If VBA.Len(VBA.Trim$(properties(key))) = 0 Then
                    Err.Raise 5, , "Empty Word font name."
                End If
        End Select
    Next key
End Sub

Private Sub private_ValidateStructure( _
    ByVal node As Object, _
    ByVal depth As Long _
)
    Dim child As Object
    Dim properties As Object
    Dim ignored As String

    If depth > m_maxDepth Then
        Err.Raise 5, , "Table structure exceeds maximum depth."
    End If
    Select Case node.baseName
        Case "table"
            private_Attributes node, "style"
            private_Children node, "columns|row|if|forEach"
            If node.selectNodes("p:columns").Length <> 1 Then
                Err.Raise 5, , "Table requires one columns section."
            End If
            Set properties = private_NodeStyle(node, "tableStyle")
        Case "columns"
            private_Attributes node, ""
            private_Children node, "column"
        Case "column"
            private_Attributes node, "name|width"
            private_Children node, ""
        Case "row"
            private_Attributes node, "header|allowBreakAcrossPages|height|heightRule"
            private_Children node, "cell"
        Case "cell"
            private_Attributes node, "style|columnSpan|rowSpan"
            private_Children node, "text|paragraph"
            If node.selectNodes("p:text").Length > 1 Or _
                    (node.selectNodes("p:text").Length > 0 And node.selectNodes("p:paragraph").Length > 0) Then
                Err.Raise 5, , "Use one text or a sequence of paragraphs in a cell."
            End If
            Set properties = private_NodeStyle(node, "cellStyle")
        Case "paragraph"
            private_Attributes node, "style"
            private_Children node, "text|image"
            If node.selectNodes("p:text|p:image").Length <> 1 Then
                Err.Raise 5, , "Paragraph requires one text or image."
            End If
            Set properties = private_NodeStyle(node, "cellStyle")
        Case "image"
            private_Attributes node, "source|width|height"
            private_Children node, ""
        Case "text"
            private_Attributes node, ""
            private_Children node, ""
        Case "if"
            private_Attributes node, "test"
            private_Children node, "row|if|forEach"
            ignored = private_Required(node, "test")
        Case "forEach"
            private_Attributes node, "itemsSource|item"
            private_Children node, "row|if|forEach"
            ignored = private_Required(node, "itemsSource")
            ignored = private_Required(node, "item")
    End Select
    For Each child In node.ChildNodes
        If child.NodeType = 1 Then
            private_ValidateStructure child, depth + 1
        End If
    Next child
End Sub

Private Sub private_Expand( _
    ByVal parent As Object, _
    ByVal values As Object, _
    ByVal rows As Collection, _
    ByVal depth As Long _
)
    Dim node As Object
    Dim items As Collection
    Dim item As Variant
    Dim child As Object
    Dim key As Variant
    Dim aliasName As String
    Dim source As String
    Dim rendered As String
    Dim diagnostic As String
    Dim index As Long

    If depth > m_maxDepth Then
        Err.Raise 5, , "Table expansion exceeds maximum depth."
    End If
    For Each node In parent.ChildNodes
        If node.NodeType = 1 Then
            Select Case node.baseName
                Case "row"
                    If rows.Count >= m_maxRows Then
                        Err.Raise 5, , "Word table row limit exceeded."
                    End If
                    rows.Add private_PrepareRow(node, values)
                Case "if"
                    If Not m_text.TryRender("{#if " & private_Required(node, "test") & "}1{#else}0{#endif}", values, rendered, diagnostic) Then
                        Err.Raise 5, , diagnostic
                    End If
                    If rendered = "1" Then
                        private_Expand node, values, rows, depth + 1
                    End If
                Case "forEach"
                    source = private_Required(node, "itemsSource")
                    aliasName = private_Required(node, "item")
                    If Not values.Exists(source) Then
                        Err.Raise 5, , "Unknown table collection: " & source
                    End If
                    If Not VBA.IsObject(values(source)) Then
                        Err.Raise 5, , "Table collection required: " & source
                    End If
                    If VBA.TypeName(values(source)) <> "Collection" Then
                        Err.Raise 5, , "Table forEach requires a Collection: " & source
                    End If
                    Set items = values(source)
                    index = 0
                    For Each item In items
                        index = index + 1
                        If Not VBA.IsObject(item) Then
                            Err.Raise 5, , "Table collection items must be dictionaries."
                        End If
                        Set child = private_Dictionary()
                        For Each key In values.Keys
                            If VBA.Left$(VBA.CStr(key), VBA.Len(aliasName) + 1) <> aliasName & "." Then
                                If VBA.IsObject(values(key)) Then
                                    Set child(key) = values(key)
                                Else
                                    child(key) = values(key)
                                End If
                            End If
                        Next key
                        For Each key In item.Keys
                            If VBA.IsObject(item(key)) Then
                                Set child(aliasName & "." & key) = item(key)
                            Else
                                child(aliasName & "." & key) = item(key)
                            End If
                        Next key
                        child(aliasName & ".__index") = index
                        child(aliasName & ".__first") = (index = 1)
                        child(aliasName & ".__last") = (index = items.Count)
                        private_Expand node, child, rows, depth + 1
                    Next item
            End Select
        End If
    Next node
End Sub

Private Function private_Span( _
    ByVal node As Object, _
    ByVal name As String _
) As Long
    Dim value As String
    Dim rx As Object

    private_Span = 1
    If node.getAttributeNode(name) Is Nothing Then
        Exit Function
    End If
    value = private_Required(node, name)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^[1-9][0-9]*$"
    If Not rx.Test(value) Then
        Err.Raise 5, , "Invalid table cell span: " & value
    End If
    private_Span = VBA.CLng(value)
End Function

Private Function private_PrepareRow( _
    ByVal node As Object, _
    ByVal values As Object _
) As Object
    Dim row As Object
    Dim cells As Collection
    Dim cell As Object
    Dim child As Object
    Dim paragraph As Object
    Dim paragraphs As Collection
    Dim properties As Object
    Dim heightRule As String
    Dim textNode As Object
    Dim imageNode As Object
    Dim imageParagraph As Object

    Set row = private_Dictionary()
    Set cells = New Collection
    row.Add "header", False
    row.Add "allowBreak", False
    If Not node.getAttributeNode("header") Is Nothing Then
        row("header") = private_Boolean(node.getAttribute("header"))
    End If
    If Not node.getAttributeNode("allowBreakAcrossPages") Is Nothing Then
        row("allowBreak") = private_Boolean(node.getAttribute("allowBreakAcrossPages"))
    End If
    If Not node.getAttributeNode("height") Is Nothing Then
        row.Add "height", private_Length(node.getAttribute("height"), False)
        heightRule = private_Required(node, "heightRule")
        If heightRule <> "atLeast" And heightRule <> "exact" Then
            Err.Raise 5, , "Row heightRule must be atLeast or exact."
        End If
        row.Add "heightRule", VBA.IIf(heightRule = "exact", 2, 1)
    ElseIf Not node.getAttributeNode("heightRule") Is Nothing Then
        If node.getAttribute("heightRule") <> "auto" Then
            Err.Raise 5, , "Row height is required."
        End If
    End If
    For Each child In node.selectNodes("p:cell")
        Set cell = private_Dictionary()
        Set properties = private_NodeStyle(child, "cellStyle")
        private_ValidateStyle properties, "cellStyle"
        cell.Add "style", properties
        cell.Add "columnSpan", private_Span(child, "columnSpan")
        cell.Add "rowSpan", private_Span(child, "rowSpan")
        Set paragraphs = New Collection
        Set textNode = child.selectSingleNode("p:text")
        If Not textNode Is Nothing Then
            paragraphs.Add private_Paragraph(textNode.Text, properties, values)
        Else
            For Each paragraph In child.selectNodes("p:paragraph")
                Set properties = private_NodeStyle(paragraph, "cellStyle")
                private_ValidateStyle properties, "cellStyle"
                Set imageNode = paragraph.selectSingleNode("p:image")
                If imageNode Is Nothing Then
                    paragraphs.Add private_Paragraph(paragraph.selectSingleNode("p:text").Text, properties, values)
                Else
                    Set imageParagraph = private_Paragraph("", properties, values)
                    imageParagraph.Add "image", private_Required(imageNode, "source")
                    If Not m_images.Exists(imageParagraph("image")) Then
                        Err.Raise 5, , "Unknown Word table image: " & imageParagraph("image")
                    End If
                    imageParagraph.Add "width", private_Length(private_Required(imageNode, "width"), False)
                    imageParagraph.Add "height", private_Length(private_Required(imageNode, "height"), False)
                    m_drawingId = m_drawingId + 1
                    imageParagraph.Add "drawingId", m_drawingId
                    paragraphs.Add imageParagraph
                End If
            Next paragraph
        End If
        If paragraphs.Count = 0 Then
            paragraphs.Add private_Paragraph("", cell("style"), values)
        End If
        cell.Add "paragraphs", paragraphs
        cells.Add cell
    Next child
    row.Add "cells", cells
    Set private_PrepareRow = row
End Function

Private Function private_FindMarker( _
    ByVal document As Object, _
    ByVal marker As String _
) As Object
    Dim target As Object
    Dim second As Object

    Set target = document.Content.Duplicate
    target.Find.ClearFormatting
    If Not target.Find.Execute(marker, True, False, False, False, False, True, 0, False) Then
        Err.Raise 5, , "Word table marker not found: " & marker
    End If
    Set second = document.Content.Duplicate
    second.Start = target.End
    second.Find.ClearFormatting
    If second.Find.Execute(marker, True, False, False, False, False, True, 0, False) Then
        Err.Raise 5, , "Duplicate Word table marker: " & marker
    End If
    If target.Information(12) Then
        Err.Raise 5, , "Word table markers must be outside existing tables."
    End If
    If VBA.Replace$(target.Paragraphs(1).Range.Text, VBA.vbCr, "") <> marker Then
        Err.Raise 5, , "A Word table marker must occupy a separate paragraph."
    End If
    Set private_FindMarker = target
End Function

Private Function private_Paragraph( _
    ByVal template As String, _
    ByVal properties As Object, _
    ByVal values As Object _
) As Object
    Dim text As String
    Dim diagnostic As String
    Dim result As Object

    If Not m_text.TryRender(template, values, text, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    text = VBA.Replace$(VBA.Replace$(text, VBA.vbCrLf, VBA.vbCr), VBA.vbLf, VBA.vbCr)
    If VBA.InStr(text, VBA.Chr$(7)) > 0 Then
        Err.Raise 5, , "Word cell-end characters are not allowed in DSL text."
    End If
    Set result = private_Dictionary()
    result.Add "text", text
    If VBA.Len(text) = 0 Then
        result.Add "count", 1
    Else
        result.Add "count", UBound(VBA.Split(text, VBA.vbCr)) + 1
    End If
    result.Add "style", properties
    Set private_Paragraph = result
End Function

Private Function private_FontSize(ByVal value As String) As Double
    Dim rx As Object

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^[0-9]+([.][0-9]+)?$"
    If Not rx.Test(value) Then
        Err.Raise 5, , "Font size must be a number of points: " & value
    End If
    private_FontSize = VBA.Val(value)
    If private_FontSize < 1 Or private_FontSize > 1638 Then
        Err.Raise 5, , "Word font size must be between 1 and 1638 points."
    End If
    If private_FontSize * 2 <> VBA.Fix(private_FontSize * 2) Then
        Err.Raise 5, , "Word font size must use half-point increments."
    End If
End Function