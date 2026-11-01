Attribute VB_Name = "ex_WordTableXml"
Option Explicit

Private Const W_NS As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
Private Const PKG_NS As String = "http://schemas.microsoft.com/office/2006/xmlPackage"

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
Public Function fn_Build( _
    ByVal plan As Object, _
    ByVal images As Object _
) As String
    Dim document As Object
    Dim package As Object
    Dim body As Object
    Dim table As Object
    Dim properties As Object
    Dim grid As Object
    Dim rowNode As Object
    Dim cellNode As Object
    Dim node As Object
    Dim paragraph As Object
    Dim row As Object
    Dim cell As Object
    Dim starts As Object
    Dim covering As Object
    Dim image As Object
    Dim relationships As Object
    Dim key As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim r As Long
    Dim c As Long
    Dim width As Double
    Dim span As Long
    Dim continuation As Boolean
    Dim padding As String

    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    If Not document.LoadXML("<pkg:package xmlns:pkg='" & PKG_NS & "'/>") Then
        Err.Raise 5, , "Cannot create Word XML package."
    End If
    Set package = document.documentElement
    Set node = private_Part(package, "/_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml")
    Set relationships = private_Node(node, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships")
    private_Relationship relationships, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "word/document.xml"
    Set node = private_Part(package, "/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    Set body = private_Node(private_W(node, "document"), "w:body", W_NS)
    Set table = private_W(body, "tbl")
    Set properties = private_W(table, "tblPr")
    Set node = private_W(properties, "tblW")
    private_Attribute node, "type", "dxa"
    private_Attribute node, "w", VBA.CStr(VBA.CLng(plan("width") * 20))
    Set node = private_W(properties, "jc")
    private_Attribute node, "val", private_Value(plan("style"), "alignment", "left")
    Set node = private_W(properties, "tblBorders")
    For Each key In VBA.Array("top", "left", "bottom", "right", "insideH", "insideV")
        private_Border node, VBA.CStr(key), private_Value(plan("style"), "borders", "none")
    Next key
    Set node = private_W(properties, "tblLayout")
    private_Attribute node, "type", "fixed"
    If plan("style").Exists("cellPadding") Or _
            plan("style").Exists("cellPaddingTop") Or plan("style").Exists("cellPaddingLeft") Or _
            plan("style").Exists("cellPaddingBottom") Or plan("style").Exists("cellPaddingRight") Then
        Set node = private_W(properties, "tblCellMar")
        For Each key In VBA.Array("top", "left", "bottom", "right")
            padding = private_Value(plan("style"), "cellPadding", "0pt")
            padding = private_Value(plan("style"), "cellPadding" & VBA.UCase$(VBA.Left$(VBA.CStr(key), 1)) & VBA.Mid$(VBA.CStr(key), 2), padding)
            Set grid = private_W(node, VBA.CStr(key))
            private_Attribute grid, "w", VBA.CStr(VBA.CLng(private_Points(padding) * 20))
            private_Attribute grid, "type", "dxa"
        Next key
    End If
    Set grid = private_W(table, "tblGrid")
    For columnIndex = 1 To plan("columns").Count
        Set node = private_W(grid, "gridCol")
        private_Attribute node, "w", VBA.CStr(VBA.CLng(plan("columns")(columnIndex) * 20))
    Next columnIndex
    Set starts = VBA.CreateObject("Scripting.Dictionary")
    Set covering = VBA.CreateObject("Scripting.Dictionary")
    For Each cell In plan("cells")
        starts.Add VBA.CStr(cell("row")) & ":" & VBA.CStr(cell("column")), cell
        For r = cell("row") To cell("row") + cell("rowSpan") - 1
            For c = cell("column") To cell("column") + cell("columnSpan") - 1
                covering.Add VBA.CStr(r) & ":" & VBA.CStr(c), cell
            Next c
        Next r
    Next cell
    For rowIndex = 1 To plan("rows").Count
        Set row = plan("rows")(rowIndex)
        Set rowNode = private_W(table, "tr")
        Set properties = private_W(rowNode, "trPr")
        If row("header") Then
            Set node = private_W(properties, "tblHeader")
        End If
        If Not row("allowBreak") Then
            Set node = private_W(properties, "cantSplit")
        End If
        If row.Exists("height") Then
            Set node = private_W(properties, "trHeight")
            private_Attribute node, "val", VBA.CStr(VBA.CLng(row("height") * 20))
            private_Attribute node, "hRule", VBA.IIf(row("heightRule") = 2, "exact", "atLeast")
        End If
        columnIndex = 1
        Do While columnIndex <= plan("columns").Count
            key = VBA.CStr(rowIndex) & ":" & VBA.CStr(columnIndex)
            Set cell = covering(key)
            continuation = Not starts.Exists(key)
            span = cell("columnSpan")
            Set cellNode = private_W(rowNode, "tc")
            Set properties = private_W(cellNode, "tcPr")
            width = 0
            For c = columnIndex To columnIndex + span - 1
                width = width + plan("columns")(c)
            Next c
            Set node = private_W(properties, "tcW")
            private_Attribute node, "w", VBA.CStr(VBA.CLng(width * 20))
            private_Attribute node, "type", "dxa"
            If span > 1 Then
                Set node = private_W(properties, "gridSpan")
                private_Attribute node, "val", VBA.CStr(span)
            End If
            If cell("rowSpan") > 1 Then
                Set node = private_W(properties, "vMerge")
                If Not continuation Then
                    private_Attribute node, "val", "restart"
                End If
            End If
            If cell("style").Exists("borderBottom") Then
                Set node = private_W(properties, "tcBorders")
                private_Border node, "bottom", cell("style")("borderBottom")
            End If
            Set node = private_W(properties, "vAlign")
            private_Attribute node, "val", private_Value(cell("style"), "vertical", "center")
            If continuation Then
                Set node = private_W(cellNode, "p")
            Else
                For Each paragraph In cell("paragraphs")
                    private_Paragraph cellNode, paragraph, cell("style"), images
                Next paragraph
            End If
            columnIndex = columnIndex + span
        Loop
    Next rowIndex
    Set node = private_W(body, "p")
    If images.Count > 0 Then
        Set node = private_Part(package, "/word/_rels/document.xml.rels", "application/vnd.openxmlformats-package.relationships+xml")
        Set relationships = private_Node(node, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships")
        For Each key In images.Keys
            Set image = images(key)
            private_Relationship relationships, image("id"), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/" & image("name")
            Set node = private_Node(package, "pkg:part", PKG_NS)
            node.setAttribute "pkg:name", "/word/media/" & image("name")
            node.setAttribute "pkg:contentType", "image/png"
            Set node = private_Node(node, "pkg:binaryData", PKG_NS)
            node.Text = image("base64")
        Next key
    End If
    For Each cellNode In document.selectNodes("//*[local-name()='tc']")
        If cellNode.selectNodes("*[local-name()='p']").Length = 0 Then
            Err.Raise 5, , "A generated Word table cell has no paragraph."
        End If
    Next cellNode
    fn_Build = document.XML
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Node( _
    ByVal parent As Object, _
    ByVal name As String, _
    ByVal namespaceUri As String _
) As Object
    Set private_Node = parent.ownerDocument.createNode(1, name, namespaceUri)
    parent.appendChild private_Node
End Function

Private Function private_W( _
    ByVal parent As Object, _
    ByVal name As String _
) As Object
    Set private_W = private_Node(parent, "w:" & name, W_NS)
End Function

Private Sub private_Attribute( _
    ByVal node As Object, _
    ByVal name As String, _
    ByVal value As String _
)
    Dim attributeNode As Object

    Set attributeNode = node.ownerDocument.createNode(2, "w:" & name, W_NS)
    attributeNode.Text = value
    node.Attributes.setNamedItem attributeNode
End Sub

Private Function private_Part( _
    ByVal parent As Object, _
    ByVal name As String, _
    ByVal contentType As String _
) As Object
    Dim node As Object

    Set node = private_Node(parent, "pkg:part", PKG_NS)
    node.setAttribute "pkg:name", name
    node.setAttribute "pkg:contentType", contentType
    Set private_Part = private_Node(node, "pkg:xmlData", PKG_NS)
End Function

Private Sub private_Relationship( _
    ByVal parent As Object, _
    ByVal id As String, _
    ByVal kind As String, _
    ByVal target As String _
)
    Dim node As Object

    Set node = private_Node(parent, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    node.setAttribute "Id", id
    node.setAttribute "Type", kind
    node.setAttribute "Target", target
End Sub

Private Function private_Value( _
    ByVal values As Object, _
    ByVal key As String, _
    ByVal defaultValue As String _
) As String
    private_Value = defaultValue
    If values.Exists(key) Then
        private_Value = VBA.CStr(values(key))
    End If
End Function

Private Function private_Points(ByVal value As String) As Double
    private_Points = VBA.Val(value)
    Select Case VBA.Right$(value, 2)
        Case "mm"
            private_Points = private_Points * 72 / 25.4
        Case "cm"
            private_Points = private_Points * 72 / 2.54
    End Select
End Function

Private Sub private_Border( _
    ByVal parent As Object, _
    ByVal side As String, _
    ByVal kind As String _
)
    Dim node As Object

    Set node = private_W(parent, side)
    private_Attribute node, "val", VBA.IIf(kind = "single", "single", "nil")
    If kind = "single" Then
        private_Attribute node, "sz", "6"
        private_Attribute node, "color", "000000"
    End If
End Sub

Private Sub private_Paragraph( _
    ByVal parent As Object, _
    ByVal paragraph As Object, _
    ByVal cellStyle As Object, _
    ByVal images As Object _
)
    Dim properties As Object
    Dim key As Variant
    Dim text As Variant
    Dim lines As Variant
    Dim p As Object
    Dim pPr As Object
    Dim run As Object
    Dim rPr As Object
    Dim node As Object
    Dim image As Object

    Set properties = VBA.CreateObject("Scripting.Dictionary")
    For Each key In cellStyle.Keys
        properties(key) = cellStyle(key)
    Next key
    For Each key In paragraph("style").Keys
        properties(key) = paragraph("style")(key)
    Next key
    ' Split of an empty string returns no elements in VBA, but every Word cell needs a paragraph.
    If VBA.Len(paragraph("text")) = 0 Then
        lines = VBA.Array(VBA.vbNullString)
    Else
        lines = VBA.Split(paragraph("text"), VBA.vbCr)
    End If
    For Each text In lines
        Set p = private_W(parent, "p")
        Set pPr = private_W(p, "pPr")
        Set node = private_W(pPr, "spacing")
        private_Attribute node, "before", VBA.CStr(VBA.CLng(private_Points(private_Value(properties, "spaceBefore", "0pt")) * 20))
        private_Attribute node, "after", VBA.CStr(VBA.CLng(private_Points(private_Value(properties, "spaceAfter", "0pt")) * 20))
        If properties.Exists("lineSpacing") Then
            private_Attribute node, "line", VBA.CStr(VBA.CLng(private_Points(properties("lineSpacing")) * 20))
            private_Attribute node, "lineRule", "atLeast"
        End If
        Set node = private_W(pPr, "jc")
        private_Attribute node, "val", private_Value(properties, "horizontal", "left")
        Set run = private_W(p, "r")
        Set rPr = private_W(run, "rPr")
        Set node = private_W(rPr, "rFonts")
        private_Attribute node, "ascii", private_Value(properties, "fontName", "Times New Roman")
        private_Attribute node, "hAnsi", private_Value(properties, "fontName", "Times New Roman")
        For Each key In VBA.Array("fontBold", "fontItalic")
            If properties.Exists(key) Then
                Set node = private_W(rPr, VBA.IIf(key = "fontBold", "b", "i"))
                private_Attribute node, "val", VBA.IIf(properties(key) = "true", "1", "0")
            End If
        Next key
        Set node = private_W(rPr, "sz")
        private_Attribute node, "val", VBA.CStr(VBA.CLng(VBA.Val(private_Value(properties, "fontSize", "12")) * 2))
        Set node = private_W(rPr, "lang")
        private_Attribute node, "val", "uk-UA"
        If paragraph.Exists("image") Then
            Set image = images(paragraph("image"))
            Set node = private_W(run, "drawing")
            ' Dimensions and identifiers were validated during preparation.
            node.appendChild private_Drawing(node.ownerDocument, image, paragraph)
        Else
            Set node = private_W(run, "t")
            node.setAttribute "xml:space", "preserve"
            node.Text = VBA.CStr(text)
        End If
    Next text
End Sub

Private Function private_Drawing( _
    ByVal document As Object, _
    ByVal image As Object, _
    ByVal paragraph As Object _
) As Object
    Dim fragment As Object
    Dim cx As String
    Dim cy As String

    cx = VBA.CStr(VBA.CLng(paragraph("width") * 12700))
    cy = VBA.CStr(VBA.CLng(paragraph("height") * 12700))
    Set fragment = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    If Not fragment.LoadXML("<wp:inline xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' " & _
            "xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture' " & _
            "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><wp:extent cx='" & cx & "' cy='" & cy & "'/>" & _
            "<wp:docPr id='" & paragraph("drawingId") & "' name='Emblem'/><wp:cNvGraphicFramePr/>" & _
            "<a:graphic><a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/picture'><pic:pic>" & _
            "<pic:nvPicPr><pic:cNvPr id='0' name='Emblem'/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed='" & image("id") & "'/>" & _
            "<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='" & cx & "' cy='" & cy & "'/>" & _
            "</a:xfrm><a:prstGeom prst='rect'><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>") Then
        Err.Raise 5, , "Cannot create Word table image."
    End If
    Set private_Drawing = document.importNode(fragment.documentElement, True)
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------