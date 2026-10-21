Attribute VB_Name = "ex_UiFormLayouts"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Validate( _
    ByVal page As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim catalogs As Object
    Dim groups As Object
    Dim fields As Object
    Dim layouts As Object
    Dim references As Object
    Dim group As Object
    Dim layout As Object
    Dim key As Variant
    Dim layoutKey As Variant
    Dim fieldKey As Variant
    Dim catalogName As String

    On Error GoTo EH
    Set catalogs = private_Index(page, "fieldCatalog")
    Set groups = private_Index(page, "formLayouts")
    For Each key In catalogs.Keys
        Set fields = private_Index(catalogs(key), "field")
    Next key
    For Each key In groups.Keys
        Set group = groups(key)
        catalogName = ex_UiElementFactory.fn_Attribute(group, "catalog")
        If Not catalogs.Exists(catalogName) Then Err.Raise 5, , "Field catalog was not found: " & catalogName
        Set fields = private_Index(catalogs(catalogName), "field")
        Set layouts = private_Index(group, "formLayout")
        For Each layoutKey In layouts.Keys
            Set layout = layouts(layoutKey)
            Set references = private_Index(layout, "fieldRef")
            For Each fieldKey In references.Keys
                If Not fields.Exists(fieldKey) Then
                    Err.Raise 5, , "Unknown fieldRef " & fieldKey & " in layout " & layoutKey
                End If
            Next fieldKey
        Next layoutKey
    Next key
    fn_Validate = True
    Exit Function
EH:
    diagnostic = "Form layouts: " & VBA.Err.Description
End Function

Public Function fn_TryPlaceColumns( _
    ByVal columns As Collection, _
    ByVal fixedColumns As Object, _
    ByVal width As Long, _
    ByRef diagnostic As String, _
    Optional ByVal groups As Object, _
    Optional ByVal groupRanges As Object _
) As Boolean
    Dim occupied As Object
    Dim field As Variant
    Dim column As Object
    Dim position As Long
    Dim nextPosition As Long
    Dim activeGroups As Object
    Dim members As Collection
    Dim group As Variant
    Dim groupName As String
    Dim groupRow As Long
    Dim candidateColumn As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    Dim bounds As Variant
    Dim firstFixedColumn As Long
    Dim memberIndex As Long
    Dim remaining As Long
    Dim runStart As Long
    Dim runLength As Long
    Dim bestLength As Long
    Dim take As Long
    Dim alignEnd As Boolean

    On Error GoTo EH_PLACE
    If width < 1 Then Err.Raise 5, , "Form workspace width must be positive."
    Set occupied = CreateObject("Scripting.Dictionary")
    firstFixedColumn = width + 1
    For Each field In fixedColumns.Keys
        If Not IsNumeric(fixedColumns(field)) Then Err.Raise 5, , "Invalid fixed column: " & field
        If CDbl(fixedColumns(field)) <> Fix(CDbl(fixedColumns(field))) Then Err.Raise 5, , "Fixed column must be an integer: " & field
        position = CLng(fixedColumns(field))
        If position < firstFixedColumn Then firstFixedColumn = position
        If position < 1 Or position > width Then Err.Raise 5, , "Fixed column is outside the workspace: " & field
        If occupied.Exists("1:" & CStr(position)) Then Err.Raise 5, , "Duplicate fixed column: " & CStr(position)
        occupied.Add "1:" & CStr(position), True
    Next field
    Set activeGroups = CreateObject("Scripting.Dictionary")
    For Each column In columns
        column("candidateColumn") = 0
        If Not column.Exists("disabled") Then
            candidateColumn = candidateColumn + 1
            column("candidateColumn") = candidateColumn
        End If
        field = CStr(column("formField"))
        If fixedColumns.Exists(field) Then
            column("column") = CLng(fixedColumns(field))
            column("row") = 1
        Else
            groupName = CStr(field)
            If Not groups Is Nothing Then
                If groups.Exists(field) Then groupName = CStr(groups(field))
            End If
            If Not activeGroups.Exists(groupName) Then
                Set members = New Collection
                activeGroups.Add groupName, members
            End If
            activeGroups(groupName).Add column
        End If
    Next column
    For Each group In activeGroups.Keys
        Set members = activeGroups(group)
        groupRow = 1
        memberIndex = 1
        Do While memberIndex <= members.Count
            remaining = members.Count - memberIndex + 1
            nextPosition = 0
            bestLength = 0
            runLength = 0
            firstColumn = 1
            lastColumn = width
            If groupRow = 1 And Not groupRanges Is Nothing Then
                If groupRanges.Exists(group) Then
                    bounds = groupRanges(group)
                    firstColumn = CLng(bounds(0))
                    lastColumn = CLng(bounds(1))
                    If firstColumn < 1 Or lastColumn > width Or lastColumn < firstColumn Then Err.Raise 5, , "Invalid form group range: " & CStr(group)
                End If
            End If
            alignEnd = (groupRow = 1 And lastColumn < firstFixedColumn)
            ' Fit the next prefix into this row; only its remaining suffix wraps below.
            For position = firstColumn To lastColumn + 1
                If position <= lastColumn And Not occupied.Exists(CStr(groupRow) & ":" & CStr(position)) Then
                    If runLength = 0 Then runStart = position
                    runLength = runLength + 1
                Else
                    take = runLength
                    If take > remaining Then take = remaining
                    If take > bestLength Or (alignEnd And take > 0 And take = bestLength) Then
                        bestLength = take
                        nextPosition = runStart
                        If alignEnd Then nextPosition = position - take
                    End If
                    runLength = 0
                End If
            Next position
            If bestLength > 0 Then
                ex_Core.fn_Diagnostic_WriteLog "TEMP.Form.PlaceGroup | Group=" & CStr(group) & _
                    " | Row=" & CStr((groupRow - 1) * 2 + 1) & " | Column=" & CStr(nextPosition) & " | Width=" & CStr(bestLength)
                For take = 1 To bestLength
                    Set column = members(memberIndex)
                    column("column") = nextPosition
                    column("row") = (groupRow - 1) * 2 + 1
                    occupied.Add CStr(groupRow) & ":" & CStr(nextPosition), True
                    nextPosition = nextPosition + 1
                    memberIndex = memberIndex + 1
                Next take
            End If
            groupRow = groupRow + 1
        Loop
    Next group
    fn_TryPlaceColumns = True
    Exit Function
EH_PLACE:
    diagnostic = "Form column map: " & Err.Description
End Function

Public Function fn_TryResolve( _
    ByVal definition As Object, _
    ByVal bindingContext As obj_UiBindingContext, _
    ByRef resolved As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As String
    Dim rawLayout As String
    Dim layoutName As String
    Dim catalogs As Object
    Dim groups As Object
    Dim fields As Object
    Dim layouts As Object
    Dim group As Object
    Dim layout As Object
    Dim node As Object
    Dim field As Object
    Dim fieldName As String
    Dim catalogName As String
    Dim span As String

    On Error GoTo EH
    Set resolved = Nothing
    If VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "fieldsSource")) > 0 Then
        fn_TryResolve = private_ResolveFields(definition, bindingContext, resolved, diagnostic)
        Exit Function
    End If
    source = ex_UiElementFactory.fn_Attribute(definition, "layoutSource")
    rawLayout = ex_UiElementFactory.fn_Attribute(definition, "layout")
    If VBA.Len(source) = 0 And VBA.Len(rawLayout) = 0 Then
        Set resolved = definition.cloneNode(True)
        fn_TryResolve = True
        Exit Function
    End If
    If VBA.Len(source) = 0 Or VBA.Len(rawLayout) = 0 Then
        Err.Raise 5, , "Form requires both layoutSource and layout."
    End If
    For Each node In definition.ChildNodes
        If node.NodeType = 1 Then Err.Raise 5, , "A layout-based form cannot contain inline fields or controls."
    Next node
    If Not ex_UiBindingRuntime.fn_TryResolveText(rawLayout, bindingContext, layoutName) Then
        Err.Raise 5, , "Cannot resolve form layout: " & rawLayout
    End If
    Set catalogs = private_Index(definition.ownerDocument.documentElement, "fieldCatalog")
    Set groups = private_Index(definition.ownerDocument.documentElement, "formLayouts")
    If Not groups.Exists(source) Then Err.Raise 5, , "Form layouts were not found: " & source
    Set group = groups(source)
    catalogName = ex_UiElementFactory.fn_Attribute(group, "catalog")
    If Not catalogs.Exists(catalogName) Then Err.Raise 5, , "Field catalog was not found: " & catalogName
    Set fields = private_Index(catalogs(catalogName), "field")
    Set layouts = private_Index(group, "formLayout")
    If Not layouts.Exists(layoutName) Then Err.Raise 5, , "Form layout was not found: " & layoutName
    Set layout = layouts(layoutName)
    Set resolved = definition.cloneNode(False)
    resolved.removeAttribute "layout"
    resolved.removeAttribute "layoutSource"
    ' Materialize one active form; the page and shared catalog remain unchanged.
    For Each node In layout.ChildNodes
        If node.NodeType = 1 Then
            fieldName = ex_UiElementFactory.fn_Attribute(node, "name")
            If Not fields.Exists(fieldName) Then Err.Raise 5, , "Unknown fieldRef: " & fieldName
            Set field = fields(fieldName).cloneNode(True)
            span = ex_UiElementFactory.fn_Attribute(node, "columnSpan")
            If VBA.Len(span) > 0 Then field.setAttribute "columnSpan", span
            resolved.appendChild field
        End If
    Next node
    fn_TryResolve = True
    Exit Function
EH:
    Set resolved = Nothing
    diagnostic = "Form layout: " & VBA.Err.Description
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_ResolveFields( _
    ByVal definition As Object, _
    ByVal bindingContext As obj_UiBindingContext, _
    ByRef resolved As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim catalogs As Object
    Dim fields As Object
    Dim configured As Object
    Dim scalar As Variant
    Dim isObject As Boolean
    Dim column As Object
    Dim field As Object
    Dim node As Object
    Dim seen As Object
    Dim name As String
    Dim catalogName As String
    Dim position As Long
    Dim previousPosition As Long
    Dim width As Long
    Dim mapped As Boolean
    Dim row As Long
    Dim lastRow As Long
    Dim candidateColumns As Object

    On Error GoTo EH
    If VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "layoutSource")) > 0 Or _
            VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "layout")) > 0 Then
        Err.Raise 5, , "Form fieldsSource cannot be combined with a named layout."
    End If
    For Each node In definition.ChildNodes
        If node.NodeType = 1 Then Err.Raise 5, , "A fieldsSource form cannot contain inline fields."
    Next node
    catalogName = ex_UiElementFactory.fn_Attribute(definition, "fieldCatalog")
    Set catalogs = private_Index(definition.ownerDocument.documentElement, "fieldCatalog")
    If Not catalogs.Exists(catalogName) Then Err.Raise 5, , "Field catalog was not found: " & catalogName
    Set fields = private_Index(catalogs(catalogName), "field")
    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
        ex_UiElementFactory.fn_Attribute(definition, "fieldsSource"), _
        bindingContext, scalar, configured, isObject _
    ) Then Err.Raise 5, , "Cannot resolve form fieldsSource."
    If Not isObject Then Err.Raise 5, , "Form fieldsSource must be a collection."
    If Not TypeOf configured Is Collection Then Err.Raise 5, , "Form fieldsSource must be a collection."
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    Set resolved = definition.cloneNode(False)
    resolved.removeAttribute "fieldCatalog"
    resolved.removeAttribute "fieldsSource"
    mapped = False
    For Each column In configured
        If column.Exists("column") Then mapped = True
    Next column
    If mapped Then
        width = ex_UiElementFactory.fn_Long(definition, "columnSpan", 0)
        If width < 1 Then Err.Raise 5, , "Mapped forms require a reserved columnSpan."
        If ex_UiElementFactory.fn_Attribute(definition, "orientation") <> "horizontal" Or _
                ex_UiElementFactory.fn_Attribute(definition, "labelPosition") <> "top" Then
            Err.Raise 5, , "Mapped forms require horizontal fields with top labels."
        End If
        resolved.setAttribute "placement", "mapped"
        resolved.setAttribute "candidateLayout", "packed"
        Set candidateColumns = CreateObject("Scripting.Dictionary")
        lastRow = 1
        For Each column In configured
            If Not column.Exists("row") Then Err.Raise 5, , "Missing row in form column map."
            If CLng(column("row")) < 1 Then Err.Raise 5, , "Invalid form row."
            If CLng(column("row")) > lastRow Then lastRow = CLng(column("row"))
            If Not column.Exists("candidateColumn") Then Err.Raise 5, , "Missing candidate column in form map."
            position = CLng(column("candidateColumn"))
            If position = 0 And column.Exists("disabled") Then
                ' Disabled display-only fields have no candidate projection column.
            Else
                If position < 1 Or position > configured.Count Then Err.Raise 5, , "Invalid candidate column in form map."
                If candidateColumns.Exists(position) Then Err.Raise 5, , "Duplicate candidate column in form map."
                candidateColumns.Add position, True
            End If
        Next column
        For row = 1 To lastRow
            previousPosition = 0
            For position = 1 To width
                For Each column In configured
                    If Not column.Exists("column") Then Err.Raise 5, , "Incomplete form column map."
                    If CLng(column("column")) = position And CLng(column("row")) = row Then
                        name = CStr(column("formField"))
                        If Not fields.Exists(name) Then Err.Raise 5, , "Unknown form field: " & name
                        If seen.Exists(name) Then Err.Raise 5, , "Duplicate form field: " & name
                        If previousPosition = position Then Err.Raise 5, , "Duplicate form column: " & CStr(position)
                        seen.Add name, True
                        Set field = fields(name).cloneNode(True)
                        field.setAttribute "columnSpan", "1"
                        field.setAttribute "column", CStr(position)
                        field.setAttribute "row", CStr(row)
                        field.setAttribute "candidateColumn", CStr(column("candidateColumn"))
                        field.setAttribute "label", CStr(column("caption"))
                        If column.Exists("disabled") Then
                            field.setAttribute "disabled", "true"
                            field.setAttribute "fieldStyle", CStr(column("disabledStyle"))
                            field.setAttribute "readOnly", "true"
                            field.setAttribute "required", "false"
                            field.setAttribute "value", "{Binding Path=Disabled" & name & "}"
                            field.removeAttribute "onChange"
                            field.removeAttribute "onFocus"
                        End If
                        resolved.appendChild field
                        previousPosition = position
                    End If
                Next column
            Next position
        Next row
        If seen.Count <> configured.Count Then Err.Raise 5, , "Form column map exceeds its reserved columnSpan."
        private_ResolveFields = True
        Exit Function
    End If
    For Each column In configured
        name = VBA.CStr(column("formField"))
        If Not fields.Exists(name) Then Err.Raise 5, , "Unknown form field: " & name
        If seen.Exists(name) Then Err.Raise 5, , "Duplicate form field: " & name
        seen.Add name, True
        Set field = fields(name).cloneNode(True)
        field.setAttribute "columnSpan", "1"
        field.setAttribute "label", VBA.CStr(column("caption"))
        resolved.appendChild field
    Next column
    private_ResolveFields = True
    Exit Function
EH:
    Set resolved = Nothing
    diagnostic = "Form fields: " & VBA.Err.Description
End Function

Private Function private_Index( _
    ByVal parent As Object, _
    ByVal tag As String _
) As Object
    Dim index As Object
    Dim node As Object
    Dim name As String

    Set index = VBA.CreateObject("Scripting.Dictionary")
    For Each node In parent.ChildNodes
        If node.NodeType = 1 Then
            If VBA.CStr(node.namespaceURI) = "urn:excelprototype:profiles" And VBA.CStr(node.baseName) = tag Then
                name = ex_UiElementFactory.fn_Attribute(node, "name")
                If VBA.Len(VBA.Trim$(name)) = 0 Then Err.Raise 5, , "Missing " & tag & " name."
                If index.Exists(name) Then Err.Raise 5, , "Duplicate " & tag & " name: " & name
                Set index(name) = node
            End If
        End If
    Next node
    Set private_Index = index
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------