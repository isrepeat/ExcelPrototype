Attribute VB_Name = "ex_UiTableStyles"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_InstallStyle( _
    ByVal workbook As Workbook, _
    ByVal xamlPath As String, _
    ByVal styleId As String, _
    ByVal styleName As String _
) As String
    Dim document As Object
    Dim root As Object
    Dim node As Object
    Dim selected As Object
    Dim element As Object
    Dim xmlAttribute As Object
    Dim seen As Object
    Dim tableStyle As TableStyle
    Dim styleElement As TableStyleElement
    Dim elementType As XlTableStyleElementType
    Dim value As String
    Dim key As String
    Dim number As Double
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo Failed
    If workbook Is Nothing Or VBA.Len(VBA.Trim$(styleId)) = 0 Or VBA.Len(VBA.Trim$(styleName)) = 0 Then
        Err.Raise 5, , "A workbook, table style ID and display name are required."
    End If
    Set document = VBA.CreateObject("Msxml2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    document.validateOnParse = False
    document.setProperty "ProhibitDTD", True
    If Not document.Load(xamlPath) Then
        Err.Raise 5, , "Cannot load table style catalog: " & xamlPath & " | " & document.parseError.reason
    End If
    Set root = document.documentElement
    If root Is Nothing Then
        Err.Raise 5, , "The table style catalog is empty: " & xamlPath
    End If
    If root.baseName <> "tableStyleCatalog" Or root.namespaceURI <> "urn:excelprototype:profiles" Or root.getAttribute("version") <> "1" Then
        Err.Raise 5, , "Expected tableStyleCatalog version 1 in the profiles namespace: " & xamlPath
    End If
    For Each node In root.childNodes
        If node.nodeType = 1 Then
            If node.baseName <> "tableStyle" Or node.namespaceURI <> root.namespaceURI Then
                Err.Raise 5, , "Unknown table style catalog node: " & node.nodeName
            End If
            If VBA.CStr(node.getAttribute("name")) = styleId Then
                If Not selected Is Nothing Then
                    Err.Raise 5, , "Duplicate table style ID: " & styleId
                End If
                Set selected = node
            End If
        End If
    Next node
    If selected Is Nothing Then
        Err.Raise 5, , "Table style ID was not found: " & styleId & " | " & xamlPath
    End If

    ' Validate the complete definition before modifying the destination workbook.
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    For Each element In selected.childNodes
        If element.nodeType = 1 Then
            If element.baseName <> "element" Or element.namespaceURI <> root.namespaceURI Then
                Err.Raise 5, , "Unknown table style element: " & element.nodeName
            End If
            elementType = private_ElementType(VBA.CStr(element.getAttribute("type")))
            key = VBA.CStr(elementType)
            If seen.Exists(key) Then
                Err.Raise 5, , "Duplicate table style element: " & element.getAttribute("type")
            End If
            seen.Add key, True
            For Each xmlAttribute In element.Attributes
                value = VBA.CStr(xmlAttribute.Text)
                Select Case xmlAttribute.nodeName
                    Case "type"
                    Case "backColor", "textColor"
                        private_Color value
                    Case "fontBold", "fontItalic"
                        If value <> "true" And value <> "false" Then
                            Err.Raise 5, , "A table style Boolean must be true or false: " & xmlAttribute.nodeName
                        End If
                    Case "stripeSize"
                        If elementType <> xlRowStripe1 And elementType <> xlRowStripe2 And _
                                elementType <> xlColumnStripe1 And elementType <> xlColumnStripe2 Then
                            Err.Raise 5, , "stripeSize requires a row or column stripe."
                        End If
                        number = ex_Helpers.fn_ParseInvariantNumber(value, "tableStyle.stripeSize")
                        If number < 1 Or number > 9 Or number <> VBA.Fix(number) Then
                            Err.Raise 5, , "Table stripe size must be an integer from 1 to 9."
                        End If
                    Case Else
                        Err.Raise 5, , "Unsupported table style attribute: " & xmlAttribute.nodeName
                End Select
            Next xmlAttribute
        End If
    Next element
    If seen.Count = 0 Then
        Err.Raise 5, , "The selected table style has no elements: " & styleId
    End If
    Set tableStyle = workbook.TableStyles.Add(styleName)
    tableStyle.ShowAsAvailableTableStyle = True
    tableStyle.ShowAsAvailablePivotTableStyle = False
    For Each element In selected.childNodes
        If element.nodeType = 1 Then
            Set styleElement = tableStyle.TableStyleElements(private_ElementType(VBA.CStr(element.getAttribute("type"))))
            For Each xmlAttribute In element.Attributes
                value = VBA.CStr(xmlAttribute.Text)
                Select Case xmlAttribute.nodeName
                    Case "backColor"
                        styleElement.Interior.Color = private_Color(value)
                    Case "textColor"
                        styleElement.Font.Color = private_Color(value)
                    Case "fontBold"
                        styleElement.Font.Bold = (value = "true")
                    Case "fontItalic"
                        styleElement.Font.Italic = (value = "true")
                    Case "stripeSize"
                        styleElement.StripeSize = VBA.CLng(ex_Helpers.fn_ParseInvariantNumber(value, "tableStyle.stripeSize"))
                End Select
            Next xmlAttribute
        End If
    Next element
    fn_InstallStyle = tableStyle.Name
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    On Error Resume Next
    If Not tableStyle Is Nothing Then
        tableStyle.Delete
    End If
    On Error GoTo 0
    VBA.Err.Raise errorNumber, "ex_UiTableStyles.fn_InstallStyle", errorText
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_ElementType(ByVal name As String) As XlTableStyleElementType
    Select Case name
        Case "WholeTable"
            private_ElementType = xlWholeTable
        Case "HeaderRow"
            private_ElementType = xlHeaderRow
        Case "TotalRow"
            private_ElementType = xlTotalRow
        Case "FirstColumn"
            private_ElementType = xlFirstColumn
        Case "LastColumn"
            private_ElementType = xlLastColumn
        Case "RowStripe1"
            private_ElementType = xlRowStripe1
        Case "RowStripe2"
            private_ElementType = xlRowStripe2
        Case "ColumnStripe1"
            private_ElementType = xlColumnStripe1
        Case "ColumnStripe2"
            private_ElementType = xlColumnStripe2
        Case Else
            Err.Raise 5, , "Unsupported table style element type: " & name
    End Select
End Function

Private Function private_Color(ByVal value As String) As Long
    Dim digits As String
    Dim i As Long

    If VBA.Len(value) <> 7 Or VBA.Left$(value, 1) <> "#" Then
        Err.Raise 5, , "Table style colors require #RRGGBB: " & value
    End If
    digits = VBA.UCase$(VBA.Mid$(value, 2))
    For i = 1 To 6
        If VBA.InStr(1, "0123456789ABCDEF", VBA.Mid$(digits, i, 1), vbBinaryCompare) = 0 Then
            Err.Raise 5, , "Invalid table style color: " & value
        End If
    Next i
    private_Color = VBA.RGB(VBA.CLng("&H" & VBA.Left$(digits, 2)), _
        VBA.CLng("&H" & VBA.Mid$(digits, 3, 2)), VBA.CLng("&H" & VBA.Right$(digits, 2)))
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------