Attribute VB_Name = "ex_TDB_BasisConfig"
Option Explicit

Private m_snapshot As obj_TableSnapshotSet
Private m_configuration As Object
Private m_orderTemplates As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_orderTemplates = Nothing
    Set m_configuration = Nothing
    Set m_snapshot = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Validate(ByVal snapshot As obj_TableSnapshotSet)
    Dim configuration As Object

    Set configuration = private_Build(snapshot)
    Set m_configuration = configuration
    Set m_snapshot = snapshot
End Sub

Public Function fn_Render( _
    ByVal profileId As String, _
    ByVal row As Object _
) As String
    Dim configuration As Object
    Dim values As Object
    Dim bindings As Object
    Dim code As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim mode As String
    Dim matched As Boolean
    Dim original As String

    Set configuration = private_Active()
    Set bindings = configuration("Bindings")
    Set values = VBA.CreateObject("Scripting.Dictionary")
    For Each code In bindings.Keys
        values.Add VBA.CStr(code), VBA.CStr(row(bindings(code)))
    Next code
    For Each code In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.valueMaps.fields"), "|")
        If Not values.Exists(code) Then
            private_Fail "UnknownField", VBA.CStr(code)
        End If
        mode = ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.valueMaps." & code & ".mode")
        If mode <> "closed" And mode <> "open" Then
            private_Fail "InvalidBinding", VBA.CStr(code)
        End If
        original = VBA.CStr(values(code))
        matched = False
        For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.valueMaps." & code & ".items"), "|")
            parts = VBA.Split(VBA.CStr(entry), "~")
            If UBound(parts) <> 1 Then
                private_Fail "InvalidBinding", VBA.CStr(code)
            End If
            If VBA.CStr(parts(0)) = original Then
                If matched Then
                    private_Fail "InvalidBinding", VBA.CStr(code)
                End If
                values(code) = VBA.CStr(parts(1))
                matched = True
            End If
        Next entry
        If Not matched And mode = "closed" Then
            private_Fail "InvalidBinding", VBA.CStr(code) & "=" & original
        End If
    Next code
    fn_Render = ex_TDB_BasisText.fn_Render(configuration("Templates"), _
        ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.profiles." & profileId), values)
End Function

Public Function fn_Order( _
    ByVal number As String, _
    ByVal dateText As String _
) As String
    Dim values As Object
    Dim fields As Object
    Dim profiles As New Collection
    Dim templates As Object
    Dim snapshot As obj_TableSnapshotSet

    Set values = VBA.CreateObject("Scripting.Dictionary")
    values.Add "order.number", number
    values.Add "order.date", dateText
    values.Add "ownUnit.genitive", ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.ownUnitGenitive")
    Set snapshot = ex_TableSnapshots.fn_Active()
    If Not m_configuration Is Nothing And m_snapshot Is snapshot Then
        Set templates = m_configuration("Templates")
    Else
        If m_orderTemplates Is Nothing Then
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            Set m_orderTemplates = ex_TDB_BasisText.fn_Prepare(fields, profiles)
        End If
        Set templates = m_orderTemplates
    End If
    fn_Order = ex_TDB_BasisText.fn_Render(templates, _
        ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.ownOrder"), values)
End Function

Public Function fn_Caption(ByVal internalField As String) As String
    Dim configuration As Object
    Dim fields As Object
    Dim bindings As Object
    Dim code As Variant

    Set configuration = private_Active()
    Set fields = configuration("Fields")
    Set bindings = configuration("Bindings")
    For Each code In bindings.Keys
        If VBA.CStr(bindings(code)) = internalField Then
            fn_Caption = fields(code)("Name")
            Exit Function
        End If
    Next code
    private_Fail "MissingBinding", internalField
End Function
' --------------------------------------

Public Function fn_Definition( _
    ByVal sourceName As String, _
    ByVal value As String _
) As Object
    Dim configuration As Object
    Dim definitions As Object

    Set configuration = private_Active()
    Set definitions = configuration("Definitions")(sourceName)
    If Not definitions.Exists(value) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    Set fn_Definition = definitions(value)
End Function

Public Function fn_Options( _
    ByVal rulesName As String, _
    ByVal group As String _
) As String
    Dim configuration As Object
    Dim options As Object

    Set configuration = private_Active()
    Set options = configuration("Options")(rulesName)
    If options.Exists(group) Then
        fn_Options = VBA.CStr(options(group))
    End If
End Function

Public Function fn_Allowed( _
    ByVal rulesName As String, _
    ByVal group As String, _
    ByVal identifier As String _
) As Boolean
    Dim configuration As Object
    Dim groups As Object
    Dim identifiers As Object

    Set configuration = private_Active()
    Set groups = configuration("Allowed")(rulesName)
    If Not groups.Exists(group) Then Exit Function
    Set identifiers = groups(group)
    fn_Allowed = identifiers.Exists(identifier)
End Function
' } // namespace API
' --------------------------------------

Private Function private_Active() As Object
    Dim snapshot As obj_TableSnapshotSet

    Set snapshot = ex_TableSnapshots.fn_Active()
    If snapshot Is Nothing Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    End If
    If Not m_snapshot Is snapshot Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    End If
    Set private_Active = m_configuration
End Function

Private Sub private_AdaptDefinition( _
    ByVal configuration As Object, _
    ByVal definition As Object _
)
    Dim field As Variant
    Dim text As String

    For Each field In VBA.Split(VBA.CStr(definition("Fields")), "|")
        If VBA.Len(text) > 0 Then
            text = text & "|"
        End If
        text = text & private_Field(configuration, VBA.CStr(field), True)
    Next field
    definition("Fields") = text
    definition("Required") = private_Condition(configuration, VBA.CStr(definition("Required")))
End Sub

Private Function private_Build(ByVal snapshot As obj_TableSnapshotSet) As Object
    Dim configuration As Object
    Dim fields As Object
    Dim bindings As Object
    Dim seenBindings As Object
    Dim bindingTypes As Object
    Dim row As Object
    Dim profile As Object
    Dim tokens As Collection
    Dim code As String
    Dim entry As Variant
    Dim parts As Variant
    Dim token As Variant
    Dim internalField As String
    Dim condition As String
    Dim visible As String
    Dim key As Variant
    Dim profiles As New Collection
    Dim identifiers As Object
    Dim mapMode As String
    Dim mapInputs As Object
    Dim mapOutputs As Object
    Dim officialIds As Object
    Dim rule As Object
    Dim references As Object
    Dim name As Variant

    Set configuration = VBA.CreateObject("Scripting.Dictionary")
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    Set bindings = VBA.CreateObject("Scripting.Dictionary")
    Set seenBindings = VBA.CreateObject("Scripting.Dictionary")
    Set bindingTypes = VBA.CreateObject("Scripting.Dictionary")
    Set references = VBA.CreateObject("Scripting.Dictionary")
    For Each name In VBA.Array("Fields", "Documents", "Officials", "DocumentRules", "OfficialRules")
        references.Add VBA.CStr(name), private_Read(VBA.CStr(name), snapshot)
    Next name
    configuration.Add "References", references
    For Each row In references("Fields")
        code = VBA.CStr(row("Code"))
        If Not private_Code(code) Or fields.Exists(code) Then
            private_Fail "InvalidSchema", code
        End If
        If VBA.Len(VBA.CStr(row("Name"))) = 0 Or VBA.Len(VBA.CStr(row("Description"))) = 0 Then
            private_Fail "InvalidSchema", code
        End If
        If Not private_Has("text|date|choice", VBA.CStr(row("Type"))) Then
            private_Fail "InvalidType", code & "=" & row("Type")
        End If
        If Not private_Has("input|computed", VBA.CStr(row("Usage"))) Then
            private_Fail "InvalidType", code & "=" & row("Usage")
        End If
        fields.Add code, row
    Next row
    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.model.bindingTypes"), "|")
        parts = VBA.Split(VBA.CStr(entry), "=")
        If UBound(parts) <> 1 Then
            private_Fail "InvalidBinding", VBA.CStr(entry)
        End If
        If VBA.Left$(VBA.CStr(parts(0)), 5) <> "Form." Or Not private_Has("text|date|choice", VBA.CStr(parts(1))) Then
            private_Fail "InvalidBinding", VBA.CStr(entry)
        End If
        internalField = VBA.Mid$(VBA.CStr(parts(0)), 6)
        If bindingTypes.Exists(internalField) Then
            private_Fail "InvalidBinding", VBA.CStr(entry)
        End If
        bindingTypes.Add internalField, VBA.CStr(parts(1))
    Next entry
    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.model.bindings"), "|")
        parts = VBA.Split(VBA.CStr(entry), "=")
        If UBound(parts) <> 1 Then
            private_Fail "InvalidBinding", VBA.CStr(entry)
        End If
        code = VBA.CStr(parts(0))
        If Not fields.Exists(code) Or bindings.Exists(code) Then
            private_Fail "InvalidBinding", code
        End If
        If VBA.Left$(VBA.CStr(parts(1)), 5) <> "Form." Then
            private_Fail "InvalidBinding", VBA.CStr(entry)
        End If
        internalField = VBA.Mid$(VBA.CStr(parts(1)), 6)
        If Not bindingTypes.Exists(internalField) Then
            private_Fail "InvalidType", VBA.CStr(entry)
        End If
        If fields(code)("Type") <> bindingTypes(internalField) Then
            private_Fail "InvalidType", VBA.CStr(entry)
        End If
        If Not private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), internalField) Or seenBindings.Exists(internalField) Then
            private_Fail "InvalidBinding", VBA.CStr(entry)
        End If
        If fields(code)("Usage") = "input" Then
            If Not private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.editorFields"), internalField) Then
                private_Fail "InvalidBinding", VBA.CStr(entry)
            End If
        End If
        If fields(code)("Type") = "date" Then
            If Not private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.dateFields"), internalField) Then
                private_Fail "InvalidType", code
            End If
        ElseIf private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.dateFields"), internalField) Then
            private_Fail "InvalidType", code
        End If
        seenBindings.Add internalField, True
        bindings.Add code, internalField
    Next entry
    For Each key In fields.Keys
        If Not bindings.Exists(key) Then
            private_Fail "MissingBinding", VBA.CStr(key)
        End If
    Next key
    configuration.Add "Fields", fields
    configuration.Add "Bindings", bindings
    For Each key In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.valueMaps.fields"), "|")
        code = VBA.CStr(key)
        If Not fields.Exists(code) Then
            private_Fail "UnknownField", code
        End If
        mapMode = ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.valueMaps." & code & ".mode")
        If mapMode <> "closed" And mapMode <> "open" Then
            private_Fail "InvalidBinding", code
        End If
        Set mapInputs = VBA.CreateObject("Scripting.Dictionary")
        Set mapOutputs = VBA.CreateObject("Scripting.Dictionary")
        For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.valueMaps." & code & ".items"), "|")
            parts = VBA.Split(VBA.CStr(entry), "~")
            If UBound(parts) <> 1 Then
                private_Fail "InvalidBinding", code
            End If
            If mapInputs.Exists(VBA.CStr(parts(0))) Then
                private_Fail "InvalidBinding", code
            End If
            mapInputs.Add VBA.CStr(parts(0)), True
            mapOutputs(VBA.CStr(parts(1))) = True
        Next entry
        If mapMode = "closed" Then
            fields(code).Add "AllowedValues", mapOutputs
        End If
    Next key
    Set identifiers = VBA.CreateObject("Scripting.Dictionary")
    For Each profile In references("Documents")
        code = VBA.CStr(profile("Id"))
        If VBA.Len(code) = 0 Or identifiers.Exists(code) Then
            private_Fail "InvalidProfile", code
        End If
        identifiers.Add code, True
        profiles.Add code
        visible = VBA.vbNullString
        For Each token In VBA.Split(VBA.CStr(profile("Fields")), "|")
            internalField = private_Field(configuration, VBA.CStr(token), True)
            visible = visible & "|" & VBA.CStr(token)
        Next token
        condition = private_Condition(configuration, VBA.CStr(profile("Required")))
        ex_TDB_BasisChain.fn_CheckRequiredSyntax condition
        Set tokens = private_Identifiers(VBA.CStr(profile("Required")))
        For Each token In tokens
            If Not private_Has(visible, VBA.CStr(token)) Then
                private_Fail "InvalidProfile", VBA.CStr(profile("Id")) & ": " & VBA.CStr(token)
            End If
        Next token
    Next profile
    For Each rule In references("DocumentRules")
        If Not identifiers.Exists(VBA.CStr(rule("DocumentId"))) Then
            private_Fail "InvalidProfile", VBA.CStr(rule("DocumentId"))
        End If
    Next rule
    Set officialIds = VBA.CreateObject("Scripting.Dictionary")
    For Each row In references("Officials")
        code = VBA.CStr(row("Id"))
        If VBA.Len(code) = 0 Or officialIds.Exists(code) Or VBA.Len(VBA.CStr(row("Text"))) = 0 Then
            private_Fail "InvalidSchema", "Officials: " & code
        End If
        officialIds.Add code, True
    Next row
    For Each rule In references("OfficialRules")
        If Not officialIds.Exists(VBA.CStr(rule("OfficialId"))) Then
            private_Fail "InvalidProfile", VBA.CStr(rule("OfficialId"))
        End If
    Next rule
    configuration.Add "Templates", ex_TDB_BasisText.fn_Prepare(fields, profiles)
    For Each profile In references("Documents")
        private_AdaptDefinition configuration, profile
    Next profile
    private_PrepareIndexes configuration
    Set private_Build = configuration
End Function

Private Sub private_PrepareIndexes(ByVal configuration As Object)
    Dim definitions As Object
    Dim options As Object
    Dim allowed As Object
    Dim names As Object
    Dim groups As Object
    Dim identifiers As Object
    Dim groupOptions As Object
    Dim references As Object
    Dim sourceName As Variant
    Dim rulesName As String
    Dim idField As String
    Dim row As Object
    Dim rule As Object
    Dim group As Variant
    Dim text As String
    Dim name As String

    Set definitions = VBA.CreateObject("Scripting.Dictionary")
    Set options = VBA.CreateObject("Scripting.Dictionary")
    Set allowed = VBA.CreateObject("Scripting.Dictionary")
    Set references = configuration("References")
    For Each sourceName In VBA.Array("Documents", "Officials")
        Set names = VBA.CreateObject("Scripting.Dictionary")
        names.CompareMode = VBA.vbTextCompare
        For Each row In references(sourceName)
            name = VBA.CStr(row("Name"))
            If names.Exists(name) Then
                ex_TDB_History.fn_Fail "InvalidChain"
            End If
            names.Add name, row
        Next row
        definitions.Add VBA.CStr(sourceName), names
        If sourceName = "Documents" Then
            rulesName = "DocumentRules"
            idField = "DocumentId"
        Else
            rulesName = "OfficialRules"
            idField = "OfficialId"
        End If
        Set groups = VBA.CreateObject("Scripting.Dictionary")
        For Each rule In references(rulesName)
            group = VBA.CStr(rule("Group"))
            If Not groups.Exists(group) Then
                Set identifiers = VBA.CreateObject("Scripting.Dictionary")
                groups.Add group, identifiers
            End If
            Set identifiers = groups(group)
            identifiers(VBA.CStr(rule(idField))) = True
        Next rule
        allowed.Add rulesName, groups
        Set groupOptions = VBA.CreateObject("Scripting.Dictionary")
        For Each group In groups.Keys
            Set identifiers = groups(group)
            text = VBA.vbNullString
            For Each row In references(sourceName)
                If identifiers.Exists(VBA.CStr(row("Id"))) Then
                    If VBA.Len(text) > 0 Then
                        text = text & ","
                    End If
                    text = text & VBA.CStr(row("Name"))
                End If
            Next row
            groupOptions.Add group, text
        Next group
        options.Add rulesName, groupOptions
    Next sourceName
    configuration.Add "Definitions", definitions
    configuration.Add "Options", options
    configuration.Add "Allowed", allowed
End Sub

Private Function private_Field( _
    ByVal configuration As Object, _
    ByVal code As String, _
    ByVal requireInput As Boolean _
) As String
    Dim fields As Object
    Dim bindings As Object

    Set fields = configuration("Fields")
    Set bindings = configuration("Bindings")
    If Not fields.Exists(code) Then
        private_Fail "UnknownField", code
    End If
    If requireInput Then
        If fields(code)("Usage") <> "input" Then
            private_Fail "InvalidProfile", code
        End If
    End If
    private_Field = bindings(code)
End Function

Private Function private_Condition( _
    ByVal configuration As Object, _
    ByVal text As String _
) As String
    Dim token As Variant
    Dim position As Long
    Dim start As Long
    Dim character As String
    Dim output As String

    position = 1
    Do While position <= VBA.Len(text)
        character = VBA.Mid$(text, position, 1)
        If character Like "[A-Za-z]" Then
            start = position
            Do While position <= VBA.Len(text)
                If Not VBA.Mid$(text, position, 1) Like "[A-Za-z0-9_.]" Then
                    Exit Do
                End If
                position = position + 1
            Loop
            output = output & private_Field(configuration, VBA.Mid$(text, start, position - start), True)
        Else
            If VBA.InStr("()&|/ " & VBA.vbTab, character) = 0 Then
                private_Fail "InvalidProfile", text
            End If
            output = output & character
            position = position + 1
        End If
    Loop
    private_Condition = output
End Function

Private Function private_Identifiers(ByVal text As String) As Collection
    Dim expression As Object
    Dim match As Object
    Dim result As New Collection

    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Pattern = "[A-Za-z][A-Za-z0-9_.]*"
    expression.Global = True
    For Each match In expression.Execute(text)
        result.Add match.Value
    Next match
    Set private_Identifiers = result
End Function

Private Function private_Code(ByVal value As String) As Boolean
    Dim segment As Variant
    Dim index As Long

    For Each segment In VBA.Split(value, ".")
        If Not VBA.Left$(VBA.CStr(segment), 1) Like "[A-Za-z]" Then
            Exit Function
        End If
        For index = 1 To VBA.Len(VBA.CStr(segment))
            If Not VBA.Mid$(VBA.CStr(segment), index, 1) Like "[A-Za-z0-9_]" Then
                Exit Function
            End If
        Next index
    Next segment
    private_Code = (VBA.Len(value) > 0)
End Function

Private Function private_Has( _
    ByVal values As String, _
    ByVal value As String _
) As Boolean
    private_Has = (VBA.Len(value) > 0 And VBA.InStr(1, "|" & values & "|", "|" & value & "|", VBA.vbBinaryCompare) > 0)
End Function

Private Sub private_Fail( _
    ByVal key As String, _
    ByVal value As String _
)
    VBA.Err.Raise 5, "Basis.Config", VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.model.text." & key), "{value}", value)
End Sub

Private Function private_SourceRows( _
    ByVal name As String, _
    ByVal snapshot As obj_TableSnapshotSet _
) As Collection
    Dim source As New obj_TableSource
    Dim query As New obj_TableQuery
    Dim table As obj_DataTable
    Dim result As New Collection
    Dim row As Object
    Dim field As Variant
    Dim fields As Variant
    Dim prefix As String
    Dim diagnostic As String
    Dim fileSystem As Object
    Dim path As String
    Dim index As Long
    Dim value As Variant
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    prefix = "basisChain.references." & name
    If Not source.Initialize() Or Not query.Initialize() Then
        private_Fail "InvalidSchema", name
    End If
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.path")
    If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
        path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    source.WorkbookPath = fileSystem.GetAbsolutePathName(path)
    source.SheetName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.sheet")
    source.TableName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.table")
    fields = VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".fields"), "|")
    For Each field In fields
        query.AddColumn ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)
    Next field
    If Not snapshot.TryQuery(source, query, table, diagnostic) Then
        VBA.Err.Raise 5, , diagnostic
    End If
    For index = 1 To table.RowCount
        Set row = VBA.CreateObject("Scripting.Dictionary")
        For Each field In fields
            value = table.ValueAt(index, ex_TableQuery.fn_HeaderIndex(table.Headers, ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)))
            If VBA.IsError(value) Or VBA.IsNull(value) Then
                private_Fail "InvalidSchema", name
            End If
            row.Add VBA.CStr(field), VBA.Trim$(VBA.CStr(value))
        Next field
        If VBA.Len(VBA.CStr(row(fields(0)))) > 0 Then
            result.Add row
        End If
    Next index
    Set private_SourceRows = result
Cleanup:
    If Not table Is Nothing Then
        table.Dispose
    End If
    query.Dispose
    source.Dispose
    If errorNumber <> 0 Then
        VBA.Err.Raise errorNumber, "Basis.Config", errorDescription
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_Read( _
    ByVal name As String, _
    ByVal snapshot As obj_TableSnapshotSet _
) As Collection
    Dim result As New Collection
    Dim sourceRows As Collection
    Dim sourceRow As Object
    Dim row As Object
    Dim key As Variant
    Dim field As Variant
    Dim identifier As Variant
    Dim prefix As String
    Dim fieldNames As String
    Dim idField As String
    Dim matches As Long

    prefix = "basisChain.model." & name
    Select Case name
        Case "Fields"
            fieldNames = "Name|Description|Type|Usage"
            For Each key In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".keys"), "|")
                Set row = VBA.CreateObject("Scripting.Dictionary")
                row.Add "Code", VBA.CStr(key)
                For Each field In VBA.Split(fieldNames, "|")
                    row.Add VBA.CStr(field), ex_TravelDocumentsBuilder.fn_Config(prefix & "." & key & "." & field)
                Next field
                result.Add row
            Next key
        Case "DocumentRules", "OfficialRules"
            If name = "DocumentRules" Then
                idField = "DocumentId"
            Else
                idField = "OfficialId"
            End If
            For Each key In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".keys"), "|")
                For Each identifier In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & "." & key), "|")
                    Set row = VBA.CreateObject("Scripting.Dictionary")
                    row.Add "Group", VBA.CStr(key)
                    row.Add idField, VBA.CStr(identifier)
                    result.Add row
                Next identifier
            Next key
        Case "Documents", "Officials"
            Set sourceRows = private_SourceRows(name, snapshot)
            For Each key In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".keys"), "|")
                matches = 0
                For Each sourceRow In sourceRows
                    If VBA.StrComp(VBA.CStr(sourceRow("Name")), ex_TravelDocumentsBuilder.fn_Config(prefix & "." & key & ".Name"), VBA.vbTextCompare) = 0 Then
                        matches = matches + 1
                        Set row = VBA.CreateObject("Scripting.Dictionary")
                        For Each field In sourceRow.Keys
                            row.Add VBA.CStr(field), sourceRow(field)
                        Next field
                    End If
                Next sourceRow
                If matches <> 1 Then
                    private_Fail "InvalidSchema", name & ": " & VBA.CStr(key)
                End If
                row.Add "Id", VBA.CStr(key)
                If name = "Documents" Then
                    For Each field In VBA.Split("Required|Fields|Forms", "|")
                        row.Add VBA.CStr(field), ex_TravelDocumentsBuilder.fn_Config(prefix & "." & key & "." & field)
                    Next field
                End If
                result.Add row
            Next key
        Case Else
            private_Fail "InvalidSchema", name
    End Select
    Set private_Read = result
End Function