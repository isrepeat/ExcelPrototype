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