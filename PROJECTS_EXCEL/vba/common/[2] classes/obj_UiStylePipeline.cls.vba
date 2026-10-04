VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiStylePipeline"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_rules As Collection
Private m_regions As Object
Private m_stages As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Rules() As Collection
    Set Rules = m_rules
End Property

' //
' // API
' //
Public Function Initialize(ByVal document As Object) As Boolean
    Dim stage As Object
    Dim layer As Object
    Dim node As Object
    Dim rule As Object
    Dim selector As Object
    Dim stageName As String
    Dim target As String

    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_rules = New Collection
    Me.BeginRender
    Set m_stages = VBA.CreateObject("Scripting.Dictionary")
    m_stages.CompareMode = VBA.vbTextCompare
    For Each stage In document.SelectNodes("/*[local-name()='page' or local-name()='uiDefinition']/*[local-name()='styles']/*[local-name()='stylePipelineStage']")
        stageName = ex_UiElementFactory.fn_Attribute(stage, "name")
        If VBA.Len(stageName) = 0 Or m_stages.Exists(stageName) Then ex_UiStyleDiagnostics.fn_Raise "StageName", stageName
        m_stages.Add stageName, True
        For Each layer In stage.ChildNodes
            If layer.NodeType = 1 Then
                For Each node In layer.ChildNodes
                    If node.NodeType = 1 Then
                        target = VBA.LCase$(ex_UiElementFactory.fn_Attribute(node, "target"))
                        Select Case target
                            Case "sheet", "usedrange", "row", "column", "cell", "range", "control", "controlpart", "layoutcontainer", "layoutbound", "inlinepart"
                            Case Else
                                ex_UiStyleDiagnostics.fn_Raise "Target", target
                        End Select
                        Set selector = private_ParseSelector(ex_UiElementFactory.fn_Attribute(node, "selector"))
                        private_ValidateSelector target, selector
                        Set rule = VBA.CreateObject("Scripting.Dictionary")
                        rule.Add "stage", stageName
                        rule.Add "target", target
                        rule.Add "enabled", private_Enabled(stage) And private_Enabled(layer) And private_Enabled(node)
                        rule.Add "selector", selector
                        rule.Add "styles", ex_UiElementFactory.fn_Attribute(node, "styles")
                        rule.Add "style", ex_UiElementFactory.fn_Attribute(node, "style")
                        m_rules.Add rule
                    End If
                Next node
            End If
        Next layer
    Next stage
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Set m_rules = Nothing
    Set m_regions = Nothing
    Set m_stages = Nothing
End Sub

Public Sub BeginRender()
    Set m_regions = VBA.CreateObject("Scripting.Dictionary")
    m_regions.CompareMode = VBA.vbTextCompare
End Sub

Public Sub RegisterRegion( _
    ByVal node As Object, _
    ByVal area As Range, _
    ByVal shape As Object, _
    ByVal part As String, _
    Optional ByVal columnAlias As String, _
    Optional ByVal sourceAlias As String, _
    Optional ByVal sourceAliasTemplate As String, _
    Optional ByVal tags As String _
)
    Dim region As Object
    Dim key As String

    If area Is Nothing And shape Is Nothing Then Exit Sub
    Set region = VBA.CreateObject("Scripting.Dictionary")
    Set region("node") = node
    Set region("area") = area
    Set region("shape") = shape
    region.Add "part", part
    region.Add "tags", tags
    region.Add "columnalias", columnAlias
    region.Add "sourcealias", sourceAlias
    region.Add "sourcealiastemplate", sourceAliasTemplate
    key = node.baseName & "|" & ex_UiElementFactory.fn_Attribute(node, "name") & "|" & part & "|" & columnAlias & "|" & sourceAlias & "|" & sourceAliasTemplate
    If Not area Is Nothing Then key = key & "|" & area.Address
    If Not shape Is Nothing Then key = key & "|" & shape.Name
    Set m_regions(key) = region
End Sub

Public Function ResolveRegions( _
    ByVal rule As Object, _
    ByVal sheet As Worksheet, _
    ByVal pageScope As Range _
) As Collection
    Dim result As New Collection
    Dim selector As Object
    Dim region As Object
    Dim scope As Range
    Dim selected As Object
    Dim target As String
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim colFirst As Long
    Dim colLast As Long
    Dim regionKey As Variant

    Set selector = rule("selector")
    target = rule("target")
    If selector.Exists("sheet") Then
        If VBA.StrComp(selector("sheet"), sheet.Name, VBA.vbTextCompare) <> 0 Then
            Set ResolveRegions = result
            Exit Function
        End If
    End If
    Select Case target
        Case "control", "controlpart", "layoutcontainer", "layoutbound", "inlinepart"
            For Each regionKey In m_regions.Keys
                Set region = m_regions(regionKey)
                If private_Matches(region, selector, target) Then result.Add region
            Next regionKey
        Case Else
            Set scope = pageScope
            If target = "usedrange" Then Set scope = sheet.UsedRange
            If selector.Exists("address") Then
                Set scope = sheet.Range(selector("address"))
                If scope.Rows.Count = sheet.Rows.Count Or scope.Columns.Count = sheet.Columns.Count Then
                    Set scope = Application.Intersect(scope, pageScope)
                End If
            Else
                rowFirst = scope.Row
                rowLast = scope.Row + scope.Rows.Count - 1
                colFirst = scope.Column
                colLast = scope.Column + scope.Columns.Count - 1
                If selector.Exists("row") Then private_Span selector("row"), rowFirst, rowLast, sheet.Rows.Count
                If selector.Exists("col") Then private_Span selector("col"), colFirst, colLast, sheet.Columns.Count
                If target = "cell" And (Not selector.Exists("row") Or Not selector.Exists("col")) Then ex_UiStyleDiagnostics.fn_Raise "CellSelector"
                If target = "range" Then ex_UiStyleDiagnostics.fn_Raise "RangeSelector"
                Set scope = sheet.Range(sheet.Cells(rowFirst, colFirst), sheet.Cells(rowLast, colLast))
            End If
            If Not scope Is Nothing Then
                Set selected = VBA.CreateObject("Scripting.Dictionary")
                Set selected("area") = scope
                Set selected("shape") = Nothing
                selected.Add "part", ""
                result.Add selected
            End If
    End Select
    Set ResolveRegions = result
End Function

Public Sub ValidateStage(ByVal name As String)
    If m_stages.Count = 0 And name = "default" Then Exit Sub
    If Not m_stages.Exists(name) Then ex_UiStyleDiagnostics.fn_Raise "StageMissing", name
End Sub

Public Function MatchesInline( _
    ByVal node As Object, _
    ByVal part As String, _
    ByVal selector As Object _
) As Boolean
    Dim region As Object

    Set region = VBA.CreateObject("Scripting.Dictionary")
    region.Add "node", node
    region.Add "part", part
    region.Add "tags", ""
    region.Add "columnalias", ""
    region.Add "sourcealias", ""
    region.Add "sourcealiastemplate", ""
    MatchesInline = private_Matches(region, selector, "controlpart")
End Function

' //
' // Private
' //
Private Sub private_ValidateSelector( _
    ByVal target As String, _
    ByVal selector As Object _
)
    Dim key As Variant

    Select Case target
        Case "sheet", "usedrange", "row", "column", "cell", "range"
            For Each key In selector.Keys
                Select Case key
                    Case "sheet", "row", "col", "address"
                    Case Else
                        ex_UiStyleDiagnostics.fn_Raise "RangeKey", key
                End Select
            Next key
            If target = "range" And Not selector.Exists("address") Then ex_UiStyleDiagnostics.fn_Raise "RangeSelector"
            If target = "cell" And Not selector.Exists("address") Then
                If Not selector.Exists("row") Or Not selector.Exists("col") Then ex_UiStyleDiagnostics.fn_Raise "CellSelector"
            End If
        Case "controlpart", "inlinepart"
            If Not selector.Exists("type") Or Not selector.Exists("part") Then ex_UiStyleDiagnostics.fn_Raise "PartSelector"
        Case "layoutbound"
            If Not selector.Exists("name") Then ex_UiStyleDiagnostics.fn_Raise "BoundSelector"
    End Select
End Sub

Private Function private_ParseSelector(ByVal text As String) As Object
    Dim result As Object
    Dim segment As Variant
    Dim separator As Long
    Dim key As String
    Dim value As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    For Each segment In VBA.Split(text, ";")
        segment = VBA.Trim$(VBA.CStr(segment))
        If VBA.Len(segment) > 0 Then
            separator = VBA.InStr(1, segment, "=")
            If separator < 2 Or separator = VBA.Len(segment) Then ex_UiStyleDiagnostics.fn_Raise "Selector", segment
            key = VBA.LCase$(VBA.Trim$(VBA.Left$(segment, separator - 1)))
            value = VBA.Trim$(VBA.Mid$(segment, separator + 1))
            Select Case key
                Case "sheet", "type", "name", "style", "part", "element", "tags", "elementdepth", "row", "col", "address", "columnalias", "sourcealias", "sourcealiastemplate"
                Case Else
                    ex_UiStyleDiagnostics.fn_Raise "SelectorKey", key
            End Select
            If result.Exists(key) Then ex_UiStyleDiagnostics.fn_Raise "SelectorDuplicate", key
            result.Add key, value
        End If
    Next segment
    Set private_ParseSelector = result
End Function

Private Function private_Enabled(ByVal node As Object) As Boolean
    Dim value As String

    value = VBA.LCase$(ex_UiElementFactory.fn_Attribute(node, "enabled"))
    If value <> "" And value <> "true" And value <> "false" Then ex_UiStyleDiagnostics.fn_Raise "Enabled"
    private_Enabled = (value <> "false")
End Function

Private Function private_Matches( _
    ByVal region As Object, _
    ByVal selector As Object, _
    ByVal target As String _
) As Boolean
    Dim key As Variant
    Dim actual As String
    Dim node As Object
    Dim depth As Long
    Dim parent As Object
    Dim firstDepth As Long
    Dim lastDepth As Long

    Set node = region("node")
    If VBA.Len(region("columnalias")) > 0 And Not selector.Exists("columnalias") Then Exit Function
    If VBA.Len(region("sourcealias")) > 0 And Not selector.Exists("sourcealias") Then Exit Function
    If VBA.Len(region("sourcealiastemplate")) > 0 And Not selector.Exists("sourcealiastemplate") Then Exit Function
    If target = "control" And region("part") <> "control" Then Exit Function
    If target = "layoutcontainer" Or target = "layoutbound" Then
        If target = "layoutcontainer" And region("part") <> "layout" Then Exit Function
        If target = "layoutbound" And region("part") <> "layout" And region("part") <> "control" Then Exit Function
    End If
    If target = "inlinepart" And region("part") <> "inline" Then Exit Function
    For Each key In selector.Keys
        Select Case key
            Case "sheet"
                actual = selector(key)
            Case "part"
                actual = region("part")
            Case "element"
                actual = node.baseName
            Case "tags"
                actual = ex_UiElementFactory.fn_Attribute(node, "tags") & " " & region("tags")
                If Not private_HasTags(actual, VBA.CStr(selector(key))) Then Exit Function
                actual = selector(key)
            Case "type"
                actual = ex_UiElementFactory.fn_Attribute(node, "styleOwnerType")
                If VBA.Len(actual) = 0 Then actual = ex_UiElementFactory.fn_Attribute(node, "type")
            Case "elementdepth"
                depth = 0
                Set parent = node.parentNode
                Do While Not parent Is Nothing
                    If parent.NodeType = 1 Then depth = depth + 1
                    Set parent = parent.parentNode
                Loop
                private_Span VBA.CStr(selector(key)), firstDepth, lastDepth, 2147483647, 0
                If depth < firstDepth Or depth > lastDepth Then Exit Function
                actual = selector(key)
            Case "columnalias", "sourcealias", "sourcealiastemplate"
                actual = region(key)
            Case "row", "col", "address"
                ex_UiStyleDiagnostics.fn_Raise "TargetSelector"
            Case Else
                actual = ex_UiElementFactory.fn_Attribute(node, VBA.CStr(key))
        End Select
        If VBA.StrComp(actual, selector(key), VBA.vbTextCompare) <> 0 Then Exit Function
    Next key
    private_Matches = True
End Function

Private Function private_HasTags( _
    ByVal actual As String, _
    ByVal requested As String _
) As Boolean
    Dim token As Variant

    actual = VBA.Replace(VBA.Replace(VBA.Replace(actual, VBA.vbTab, " "), VBA.vbCr, " "), VBA.vbLf, " ")
    requested = VBA.Replace(VBA.Replace(VBA.Replace(requested, VBA.vbTab, " "), VBA.vbCr, " "), VBA.vbLf, " ")
    For Each token In VBA.Split(VBA.Trim$(requested), " ")
        If VBA.Len(token) > 0 Then
            If VBA.InStr(1, " " & actual & " ", " " & token & " ", VBA.vbTextCompare) = 0 Then Exit Function
        End If
    Next token
    private_HasTags = True
End Function

Private Sub private_Span( _
    ByVal text As String, _
    ByRef first As Long, _
    ByRef last As Long, _
    ByVal maximum As Long, _
    Optional ByVal minimum As Long = 1 _
)
    Dim parts As Variant

    parts = VBA.Split(text, ":")
    If UBound(parts) > 1 Then ex_UiStyleDiagnostics.fn_Raise "Span", text
    If Not VBA.IsNumeric(parts(0)) Then ex_UiStyleDiagnostics.fn_Raise "Span", text
    first = VBA.CLng(parts(0))
    last = first
    If UBound(parts) = 1 Then
        If Not VBA.IsNumeric(parts(1)) Then ex_UiStyleDiagnostics.fn_Raise "Span", text
        last = VBA.CLng(parts(1))
    End If
    If first < minimum Or last < first Or last > maximum Then ex_UiStyleDiagnostics.fn_Raise "SpanBounds", text
End Sub