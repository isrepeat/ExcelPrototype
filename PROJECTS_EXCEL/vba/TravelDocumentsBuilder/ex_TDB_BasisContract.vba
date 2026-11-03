Attribute VB_Name = "ex_TDB_BasisContract"
Option Explicit

Private Const CONTRACT_VERSION As String = "1"
Private Const CONTRACT_FEATURES As String = "fields.v1|fragments.v1|conditions.v1"

Private m_snapshot As obj_TableSnapshotSet
Private m_contract As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_contract = Nothing
    Set m_snapshot = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Validate(ByVal snapshot As obj_TableSnapshotSet)
    Dim contract As Object

    Set contract = private_Build(snapshot)
    Set m_contract = contract
    Set m_snapshot = snapshot
End Sub

Public Sub fn_AdaptDefinition(ByVal definition As Object)
    Dim contract As Object
    Dim field As Variant
    Dim text As String

    Set contract = private_Active()
    For Each field In VBA.Split(VBA.CStr(definition("Fields")), "|")
        If VBA.Len(text) > 0 Then
            text = text & "|"
        End If
        text = text & private_Field(contract, VBA.CStr(field), True)
    Next field
    definition("Fields") = text
    definition("Required") = private_Condition(contract, VBA.CStr(definition("Required")))
End Sub

Public Function fn_Template( _
    ByVal text As String, _
    ByVal row As Object _
) As String
    fn_Template = private_Expand(private_Active(), text, row, 0)
End Function

Public Function fn_Caption(ByVal internalField As String) As String
    Dim contract As Object
    Dim fields As Object
    Dim bindings As Object
    Dim code As Variant

    Set contract = private_Active()
    Set fields = contract("Fields")
    Set bindings = contract("Bindings")
    For Each code In bindings.Keys
        If VBA.CStr(bindings(code)) = internalField Then
            fn_Caption = fields(code)("Name")
            Exit Function
        End If
    Next code
    private_Fail "MissingBinding", internalField
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Active() As Object
    Dim snapshot As obj_TableSnapshotSet

    Set snapshot = ex_TableSnapshots.fn_Active()
    If snapshot Is Nothing Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    End If
    If Not m_snapshot Is snapshot Then
        fn_Validate snapshot
    End If
    Set private_Active = m_contract
End Function

Private Function private_Build(ByVal snapshot As obj_TableSnapshotSet) As Object
    Dim contract As Object
    Dim fields As Object
    Dim fragments As Object
    Dim bindings As Object
    Dim seenBindings As Object
    Dim bindingTypes As Object
    Dim cases As Object
    Dim group As Collection
    Dim metadata As Collection
    Dim row As Object
    Dim profile As Object
    Dim tokens As Collection
    Dim code As String
    Dim feature As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim token As Variant
    Dim internalField As String
    Dim condition As String
    Dim signature As String
    Dim visible As String
    Dim key As Variant

    Set contract = VBA.CreateObject("Scripting.Dictionary")
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    Set fragments = VBA.CreateObject("Scripting.Dictionary")
    Set bindings = VBA.CreateObject("Scripting.Dictionary")
    Set seenBindings = VBA.CreateObject("Scripting.Dictionary")
    Set bindingTypes = VBA.CreateObject("Scripting.Dictionary")
    Set cases = VBA.CreateObject("Scripting.Dictionary")
    Set metadata = private_Read("Contract", snapshot)
    If metadata.Count <> 1 Then
        private_Fail "InvalidSchema", "Contract"
    End If
    If ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.version") <> CONTRACT_VERSION Then
        private_Fail "UnsupportedVersion", ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.version")
    End If
    If metadata(1)("Version") <> CONTRACT_VERSION Then
        private_Fail "UnsupportedVersion", VBA.CStr(metadata(1)("Version"))
    End If
    For Each feature In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.features"), "|")
        If Not private_Has(CONTRACT_FEATURES, VBA.CStr(feature)) Then
            private_Fail "UnsupportedFeature", VBA.CStr(feature)
        End If
    Next feature
    For Each feature In VBA.Split(VBA.CStr(metadata(1)("Features")), "|")
        If Not private_Has(CONTRACT_FEATURES, VBA.CStr(feature)) Then
            private_Fail "UnsupportedFeature", VBA.CStr(feature)
        End If
    Next feature
    For Each feature In VBA.Split(CONTRACT_FEATURES, "|")
        If Not private_Has(VBA.CStr(metadata(1)("Features")), VBA.CStr(feature)) Then
            private_Fail "UnsupportedFeature", VBA.CStr(feature)
        End If
        If Not private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.features"), VBA.CStr(feature)) Then
            private_Fail "UnsupportedFeature", VBA.CStr(feature)
        End If
    Next feature
    For Each row In private_Read("Fields", snapshot)
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
    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.bindingTypes"), "|")
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
    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.bindings"), "|")
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
    contract.Add "Fields", fields
    contract.Add "Bindings", bindings
    contract.Add "Fragments", fragments
    For Each row In private_Read("Fragments", snapshot)
        code = VBA.CStr(row("Code"))
        If Not private_Code(code) Or fields.Exists(code) Then
            private_Fail "InvalidFragment", code
        End If
        If Not private_Has("default|nonempty|equals|except", VBA.CStr(row("Mode"))) Then
            private_Fail "InvalidFragment", code
        End If
        If VBA.Len(VBA.CStr(row("Name"))) = 0 Or VBA.Len(VBA.CStr(row("Description"))) = 0 Then
            private_Fail "InvalidFragment", code
        End If
        If row("Mode") = "default" Then
            If VBA.Len(VBA.CStr(row("Source"))) > 0 Or VBA.Len(VBA.CStr(row("Value"))) > 0 Then
                private_Fail "InvalidFragment", code
            End If
        Else
            internalField = private_Field(contract, VBA.CStr(row("Source")), False)
        End If
        signature = code & "|" & row("Mode") & "|" & row("Source") & "|" & row("Value")
        If cases.Exists(signature) Then
            private_Fail "InvalidFragment", signature
        End If
        cases.Add signature, True
        If Not fragments.Exists(code) Then
            Set group = New Collection
            fragments.Add code, group
        End If
        fragments(code).Add row
    Next row
    For Each key In fragments.Keys
        private_CheckCases fragments(key), VBA.CStr(key)
        private_CheckTemplate contract, "{" & VBA.CStr(key) & "}", "|", 0
    Next key
    For Each profile In private_Read("Documents", snapshot)
        visible = VBA.vbNullString
        For Each token In VBA.Split(VBA.CStr(profile("Fields")), "|")
            internalField = private_Field(contract, VBA.CStr(token), True)
            visible = visible & "|" & VBA.CStr(token)
        Next token
        condition = private_Condition(contract, VBA.CStr(profile("Required")))
        ex_TDB_BasisChain.fn_CheckRequiredSyntax condition
        Set tokens = private_Identifiers(VBA.CStr(profile("Required")))
        For Each token In tokens
            If Not private_Has(visible, VBA.CStr(token)) Then
                private_Fail "InvalidProfile", VBA.CStr(profile("Id")) & ": " & VBA.CStr(token)
            End If
        Next token
        private_CheckTemplate contract, VBA.CStr(profile("Template")), "|", 0
    Next profile
    Set private_Build = contract
End Function

Private Sub private_CheckCases( _
    ByVal rows As Collection, _
    ByVal code As String _
)
    Dim values As Object
    Dim row As Object
    Dim source As String
    Dim candidate As Variant
    Dim sentinel As String
    Dim matches As Long

    Set values = VBA.CreateObject("Scripting.Dictionary")
    values.Add VBA.vbNullString, True
    For Each row In rows
        If row("Mode") <> "default" Then
            If VBA.Len(source) = 0 Then
                source = VBA.CStr(row("Source"))
            ElseIf source <> VBA.CStr(row("Source")) Then
                private_Fail "InvalidFragment", code
            End If
            If row("Mode") = "nonempty" And VBA.Len(VBA.CStr(row("Value"))) > 0 Then
                private_Fail "InvalidFragment", code
            End If
            values(VBA.CStr(row("Value"))) = True
        End If
    Next row
    sentinel = "_"
    Do While values.Exists(sentinel)
        sentinel = sentinel & "_"
    Loop
    values.Add sentinel, True
    For Each candidate In values.Keys
        matches = 0
        For Each row In rows
            If private_Matches(VBA.CStr(row("Mode")), VBA.CStr(candidate), VBA.CStr(row("Value"))) Then
                matches = matches + 1
            End If
        Next row
        If matches > 1 Then
            private_Fail "InvalidFragment", code
        End If
    Next candidate
End Sub

Private Function private_Matches( _
    ByVal mode As String, _
    ByVal value As String, _
    ByVal target As String _
) As Boolean
    Select Case mode
        Case "nonempty"
            private_Matches = (VBA.Len(value) > 0)
        Case "equals"
            private_Matches = (value = target)
        Case "except"
            private_Matches = (VBA.Len(value) > 0 And value <> target)
    End Select
End Function

Private Sub private_CheckTemplate( _
    ByVal contract As Object, _
    ByVal text As String, _
    ByVal trail As String, _
    ByVal depth As Long _
)
    Dim token As Variant
    Dim row As Object
    Dim fields As Object
    Dim fragments As Object

    If depth > 64 Then
        private_Fail "InvalidFragment", trail
    End If
    Set fields = contract("Fields")
    Set fragments = contract("Fragments")
    For Each token In private_Tokens(text)
        If Not fields.Exists(token) Then
            If Not fragments.Exists(token) Then
                private_Fail "UnknownToken", VBA.CStr(token)
            End If
            If private_Has(trail, VBA.CStr(token)) Then
                private_Fail "InvalidFragment", VBA.CStr(token)
            End If
            For Each row In fragments(token)
                private_CheckTemplate contract, VBA.CStr(row("Template")), trail & token & "|", depth + 1
            Next row
        End If
    Next token
End Sub

Private Function private_Expand( _
    ByVal contract As Object, _
    ByVal text As String, _
    ByVal values As Object, _
    ByVal depth As Long _
) As String
    Dim token As Variant
    Dim fields As Object
    Dim fragments As Object
    Dim row As Object
    Dim selected As Object
    Dim fallback As Object
    Dim source As String
    Dim value As String
    Dim matches As Boolean
    Dim replacement As String

    If depth > 64 Then
        private_Fail "InvalidFragment", text
    End If
    Set fields = contract("Fields")
    Set fragments = contract("Fragments")
    For Each token In private_Tokens(text)
        If fields.Exists(token) Then
            replacement = "{Value." & private_Field(contract, VBA.CStr(token), False) & "}"
        Else
            If Not fragments.Exists(token) Then
                private_Fail "UnknownToken", VBA.CStr(token)
            End If
            Set selected = Nothing
            Set fallback = Nothing
            For Each row In fragments(token)
                matches = False
                If row("Mode") = "default" Then
                    If Not fallback Is Nothing Then
                        private_Fail "InvalidFragment", VBA.CStr(token)
                    End If
                    Set fallback = row
                Else
                    source = private_Field(contract, VBA.CStr(row("Source")), False)
                    value = VBA.CStr(values(source))
                    matches = private_Matches(VBA.CStr(row("Mode")), value, VBA.CStr(row("Value")))
                End If
                If matches Then
                    If Not selected Is Nothing Then
                        private_Fail "InvalidFragment", VBA.CStr(token)
                    End If
                    Set selected = row
                End If
            Next row
            If selected Is Nothing Then
                Set selected = fallback
            End If
            replacement = VBA.vbNullString
            If Not selected Is Nothing Then
                replacement = private_Expand(contract, VBA.CStr(selected("Template")), values, depth + 1)
            End If
        End If
        text = VBA.Replace$(text, "{" & token & "}", replacement)
    Next token
    private_Expand = text
End Function

Private Function private_Field( _
    ByVal contract As Object, _
    ByVal code As String, _
    ByVal requireInput As Boolean _
) As String
    Dim fields As Object
    Dim bindings As Object

    Set fields = contract("Fields")
    Set bindings = contract("Bindings")
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
    ByVal contract As Object, _
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
            output = output & private_Field(contract, VBA.Mid$(text, start, position - start), True)
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

Private Function private_Tokens(ByVal text As String) As Collection
    Dim result As New Collection
    Dim position As Long
    Dim closing As Long
    Dim character As String
    Dim code As String

    position = 1
    Do While position <= VBA.Len(text)
        character = VBA.Mid$(text, position, 1)
        If character = "{" Then
            closing = VBA.InStr(position + 1, text, "}")
            If closing = 0 Then
                private_Fail "UnknownToken", text
            End If
            code = VBA.Mid$(text, position + 1, closing - position - 1)
            If Not private_Code(code) Then
                private_Fail "UnknownToken", code
            End If
            result.Add code
            position = closing + 1
        ElseIf character = "}" Then
            private_Fail "UnknownToken", text
        Else
            position = position + 1
        End If
    Loop
    Set private_Tokens = result
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
    VBA.Err.Raise 5, "Basis.Contract", VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.contract.text." & key), "{value}", value)
End Sub

Private Function private_Read( _
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
    Set private_Read = result
Cleanup:
    If Not table Is Nothing Then
        table.Dispose
    End If
    query.Dispose
    source.Dispose
    If errorNumber <> 0 Then
        VBA.Err.Raise errorNumber, "Basis.Contract", errorDescription
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume Cleanup
End Function