VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_LookupProfile"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_source As obj_TableSource
Private m_columns As Collection
Private m_display As Collection
Private m_search As Collection
Private m_apply As Collection
Private m_keyField As String
Private m_minChars As Long
Private m_maxResults As Long

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
Public Property Get Source() As obj_TableSource
    Set Source = m_source
End Property

Public Property Get Columns() As Collection
    Set Columns = m_columns
End Property

Public Property Get DisplayColumns() As Collection
    Set DisplayColumns = m_display
End Property

Public Property Get SearchFields() As Collection
    Set SearchFields = m_search
End Property

Public Property Get KeyField() As String
    KeyField = m_keyField
End Property

Public Property Get MinChars() As Long
    MinChars = m_minChars
End Property

Public Property Get MaxResults() As Long
    MaxResults = m_maxResults
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal configPrefix As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim document As Object
    Dim root As Object
    Dim node As Object
    Dim item As Object
    Dim names As Object
    Dim sourceNames As Object
    Dim targets As Object
    Dim headers As Object
    Dim captions As Object
    Dim fileSystem As Object
    Dim captionKey As String
    Dim configuredCaption As String
    Dim sourcePath As String
    Dim name As String
    Dim number As String

    On Error GoTo EH
    If m_isInitialized Or m_isDisposed Then
        diagnostic = "Lookup profile is already initialized or disposed."
        Exit Function
    End If
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set document = private_ConfigDocument(configPrefix)
    Set root = document.documentElement
    If root.nodeName <> "lookup" Then Err.Raise 5, , "Expected lookup profile root."
    m_keyField = private_Required(root, "key")
    number = private_Required(root, "minChars")
    If Not number Like "#*" Or Not VBA.IsNumeric(number) Then Err.Raise 5, , "Invalid minChars."
    m_minChars = VBA.CLng(number)
    If m_minChars < 1 Or VBA.CStr(m_minChars) <> number Then Err.Raise 5, , "minChars must be a positive integer."
    number = private_Required(root, "maxResults")
    If Not VBA.IsNumeric(number) Then Err.Raise 5, , "Invalid maxResults."
    m_maxResults = VBA.CLng(number)
    If m_maxResults < 1 Or m_maxResults > 1000 Or VBA.CStr(m_maxResults) <> number Then Err.Raise 5, , "maxResults must be an integer from 1 to 1000."
    Set m_source = New obj_TableSource
    If Not m_source.Initialize() Then Err.Raise 5, , "Cannot initialize lookup source."
    Set node = root.selectSingleNode("source")
    If node Is Nothing Then Err.Raise 5, , "Lookup source is required."
    sourcePath = private_Required(node, "path")
    If VBA.Mid$(sourcePath, 2, 1) <> ":" And VBA.Left$(sourcePath, 2) <> "\\" Then sourcePath = fileSystem.BuildPath(ThisWorkbook.Path, sourcePath)
    m_source.WorkbookPath = fileSystem.GetAbsolutePathName(sourcePath)
    m_source.SheetName = ex_UiElementFactory.fn_Attribute(node, "sheet")
    m_source.TableName = ex_UiElementFactory.fn_Attribute(node, "table")
    m_source.RangeAddress = ex_UiElementFactory.fn_Attribute(node, "range")
    If Not m_source.TryValidate(diagnostic) Then Err.Raise 5, , diagnostic
    Set m_columns = New Collection
    Set m_display = New Collection
    Set m_search = New Collection
    Set m_apply = New Collection
    Set names = VBA.CreateObject("Scripting.Dictionary")
    names.CompareMode = VBA.vbTextCompare
    Set sourceNames = VBA.CreateObject("Scripting.Dictionary")
    sourceNames.CompareMode = VBA.vbTextCompare
    Set headers = VBA.CreateObject("Scripting.Dictionary")
    headers.CompareMode = VBA.vbTextCompare
    Set targets = VBA.CreateObject("Scripting.Dictionary")
    targets.CompareMode = VBA.vbTextCompare
    For Each node In root.selectNodes("columns/column")
        Set item = VBA.CreateObject("Scripting.Dictionary")
        item("name") = private_Required(node, "name")
        item("header") = private_Required(node, "header")
        item("caption") = ex_UiElementFactory.fn_Attribute(node, "caption")
        captionKey = ex_UiElementFactory.fn_Attribute(node, "captionKey")
        If VBA.Len(captionKey) > 0 Then
            If Not ex_Core.fn_TryGetWorkbookConfigValue(captionKey, configuredCaption) Then Err.Raise 5, , "Required configuration text not found: " & captionKey
            item("caption") = configuredCaption
        End If
        name = item("name")
        If names.Exists(name) Then Err.Raise 5, , "Duplicate lookup field: " & name
        If headers.Exists(item("header")) Then Err.Raise 5, , "Duplicate source header: " & item("header")
        names.Add name, True
        sourceNames.Add name, True
        headers.Add item("header"), True
        m_columns.Add item
        If VBA.Len(item("caption")) > 0 Then m_display.Add item
    Next node
    If root.selectNodes("display/column").Length > 0 Then
        Set m_display = New Collection
        Set captions = VBA.CreateObject("Scripting.Dictionary")
        captions.CompareMode = VBA.vbTextCompare
        For Each node In root.selectNodes("display/column")
            Set item = VBA.CreateObject("Scripting.Dictionary")
            name = private_Required(node, "name")
            item("name") = name
            item("caption") = private_Required(node, "caption")
            If captions.Exists(name) Then Err.Raise 5, , "Duplicate display field: " & name
            captions.Add name, True
            If Not names.Exists(name) Then names.Add name, True
            m_display.Add item
        Next node
    End If
    If Not sourceNames.Exists(m_keyField) Then Err.Raise 5, , "Lookup key field not declared: " & m_keyField
    If m_display.Count = 0 Then Err.Raise 5, , "At least one display column is required."
    For Each node In root.selectNodes("search/field")
        Set item = VBA.CreateObject("Scripting.Dictionary")
        name = private_Required(node, "name")
        If Not sourceNames.Exists(name) Then Err.Raise 5, , "Search field not declared in source: " & name
        item("name") = name
        item("match") = private_Required(node, "match")
        If item("match") <> "contains" And item("match") <> "startsWith" Then Err.Raise 5, , "Unknown lookup match: " & item("match")
        m_search.Add item
    Next node
    If m_search.Count = 0 Then Err.Raise 5, , "At least one search field is required."
    For Each node In root.selectNodes("apply/map")
        Set item = VBA.CreateObject("Scripting.Dictionary")
        name = private_Required(node, "field")
        If Not names.Exists(name) Then Err.Raise 5, , "Apply field not declared: " & name
        item("field") = name
        item("required") = ex_UiElementFactory.fn_Attribute(node, "required")
        If VBA.Len(item("required")) = 0 Then item("required") = "true"
        If item("required") <> "true" And item("required") <> "false" Then Err.Raise 5, , "Map required must be true or false."
        item("target") = private_Required(node, "target")
        If VBA.InStr(1, item("target"), ".") <= 1 Then Err.Raise 5, , "Apply target requires Source.Path."
        If targets.Exists(item("target")) Then Err.Raise 5, , "Duplicate apply target: " & item("target")
        targets.Add item("target"), True
        item("policy") = private_Required(node, "policy")
        Select Case item("policy")
            Case "overwrite", "fillIfEmpty", "ignoreEmpty"
            Case Else
                Err.Raise 5, , "Unknown apply policy: " & item("policy")
        End Select
        m_apply.Add item
    Next node
    diagnostic = VBA.vbNullString
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = "Lookup profile: " & VBA.Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_source Is Nothing Then m_source.Dispose
    Set m_source = Nothing
    Set m_columns = Nothing
    Set m_display = Nothing
    Set m_search = Nothing
    Set m_apply = Nothing
End Sub

Public Function HeaderFor(ByVal fieldName As String) As String
    Dim column As Object

    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Lookup profile is not initialized."
    For Each column In m_columns
        If VBA.StrComp(column("name"), fieldName, VBA.vbTextCompare) = 0 Then
            HeaderFor = column("header")
            Exit Function
        End If
    Next column
    Err.Raise 5, , "Lookup field not declared: " & fieldName
End Function

Public Function TryApply( _
    ByVal record As obj_LookupRecord, _
    ByVal bindings As obj_UiBindingContext, _
    ByRef diagnostic As String _
) As Boolean
    Dim updates As Object
    Dim mapping As Object
    Dim value As Variant
    Dim oldValue As Variant
    Dim oldObject As Object
    Dim isObject As Boolean
    Dim target As String
    Dim separator As Long
    Dim include As Boolean

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Lookup profile is not initialized."
    If record Is Nothing Or bindings Is Nothing Then Err.Raise 5, , "Candidate and bindings are required."
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates.CompareMode = VBA.vbTextCompare
    For Each mapping In m_apply
        If Not record.TryGetValue(mapping("field"), value) Then
            If mapping("required") = "true" Then Err.Raise 5, , "Candidate field not found: " & mapping("field")
            GoTo NextMapping
        End If
        target = mapping("target")
        separator = VBA.InStr(1, target, ".")
        If Not bindings.TryGetValue(VBA.Left$(target, separator - 1), VBA.Mid$(target, separator + 1), oldValue, oldObject, isObject) Then Err.Raise 5, , "Apply target not found: " & target
        If isObject Then Err.Raise 5, , "Apply target must be scalar: " & target
        include = True
        Select Case mapping("policy")
            Case "fillIfEmpty"
                include = private_IsBlank(oldValue)
            Case "ignoreEmpty"
                include = Not private_IsBlank(value)
        End Select
        If include Then updates(target) = value
NextMapping:
    Next mapping
    TryApply = bindings.TryApplyValues(updates, diagnostic)
    Exit Function
EH:
    diagnostic = "Apply candidate: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_IsBlank(ByVal value As Variant) As Boolean
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        private_IsBlank = True
    ElseIf Not VBA.IsError(value) Then
        private_IsBlank = (VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(value))) = 0)
    End If
End Function

Private Function private_ConfigDocument(ByVal configPrefix As String) As Object
    Dim document As Object
    Dim root As Object
    Dim source As Object
    Dim attributeName As Variant
    Dim value As String

    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    Set root = document.createElement("lookup")
    document.appendChild root
    For Each attributeName In VBA.Array("key", "minChars", "maxResults")
        root.setAttribute VBA.CStr(attributeName), private_ConfigValue(configPrefix & "." & attributeName)
    Next attributeName
    Set source = document.createElement("source")
    root.appendChild source
    source.setAttribute "path", private_ConfigValue(configPrefix & ".source.path")
    For Each attributeName In VBA.Array("sheet", "table", "range")
        If ex_Core.fn_TryGetWorkbookConfigValue(configPrefix & ".source." & attributeName, value) Then
            source.setAttribute VBA.CStr(attributeName), value
        End If
    Next attributeName
    private_ConfigSection document, root, configPrefix, "columns", "column", _
        VBA.Array("name", "header"), VBA.Array("captionKey")
    private_ConfigSection document, root, configPrefix, "display", "column", _
        VBA.Array("name", "caption"), VBA.Array()
    private_ConfigSection document, root, configPrefix, "search", "field", _
        VBA.Array("name", "match"), VBA.Array()
    private_ConfigSection document, root, configPrefix, "apply", "map", _
        VBA.Array("field", "target", "policy", "required"), VBA.Array()
    Set private_ConfigDocument = document
End Function

Private Sub private_ConfigSection( _
    ByVal document As Object, _
    ByVal root As Object, _
    ByVal configPrefix As String, _
    ByVal sectionName As String, _
    ByVal itemName As String, _
    ByVal requiredAttributes As Variant, _
    ByVal optionalAttributes As Variant _
)
    Dim section As Object
    Dim node As Object
    Dim attributeName As Variant
    Dim countText As String
    Dim count As Long
    Dim index As Long
    Dim itemPrefix As String
    Dim value As String

    countText = private_ConfigValue(configPrefix & "." & sectionName & ".count")
    If Not VBA.IsNumeric(countText) Then Err.Raise 5, , "Invalid configuration count: " & sectionName
    count = VBA.CLng(countText)
    If count < 0 Or count > 1000 Or VBA.CStr(count) <> countText Then Err.Raise 5, , "Invalid configuration count: " & sectionName
    Set section = document.createElement(sectionName)
    root.appendChild section
    For index = 1 To count
        itemPrefix = configPrefix & "." & sectionName & "." & VBA.CStr(index) & "."
        Set node = document.createElement(itemName)
        section.appendChild node
        For Each attributeName In requiredAttributes
            node.setAttribute VBA.CStr(attributeName), private_ConfigValue(itemPrefix & attributeName)
        Next attributeName
        For Each attributeName In optionalAttributes
            If ex_Core.fn_TryGetWorkbookConfigValue(itemPrefix & attributeName, value) Then
                node.setAttribute VBA.CStr(attributeName), value
            End If
        Next attributeName
    Next index
End Sub

Private Function private_ConfigValue(ByVal key As String) As String
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, private_ConfigValue) Then
        Err.Raise 5, , "Required configuration key not found: " & key
    End If
End Function

Private Function private_Required( _
    ByVal node As Object, _
    ByVal attributeName As String _
) As String
    private_Required = ex_UiElementFactory.fn_Attribute(node, attributeName)
    If VBA.Len(private_Required) = 0 Then Err.Raise 5, , "Required lookup attributeName: " & node.nodeName & "." & attributeName
End Function