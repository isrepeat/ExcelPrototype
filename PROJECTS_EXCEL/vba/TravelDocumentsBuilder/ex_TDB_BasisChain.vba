Attribute VB_Name = "ex_TDB_BasisChain"
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
Public Function fn_Read(ByVal bindingContext As obj_UiBindingContext) As Collection
    Dim scalar As Variant
    Dim source As Object
    Dim isObject As Boolean
    Dim row As Object

    If Not bindingContext.TryGetValue("Data", "BasisChain", scalar, source, isObject) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If Not isObject Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If Not TypeOf source Is Collection Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    For Each row In source
        If VBA.Len(VBA.CStr(row("RowId"))) = 0 Then row("RowId") = ex_TDB_History.fn_NewId()
    Next row
    Set fn_Read = source
End Function

Public Sub fn_CheckRequiredSyntax(ByVal specification As String)
    Dim row As Object
    Dim field As Variant

    Set row = ex_TDB_History.fn_NewRow()
    For Each field In row.Keys
        row(field) = "1"
    Next field
    private_ValidateRequired row, specification, False
End Sub

Public Function fn_Value( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal field As String _
) As String
    Dim scalar As Variant
    Dim source As Object
    Dim isObject As Boolean

    If Not bindingContext.TryGetValue("Form", field, scalar, source, isObject) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If isObject Or VBA.IsError(scalar) Or VBA.IsNull(scalar) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    fn_Value = ex_TableQuery.fn_NormalizeText(VBA.CStr(scalar))
End Function

Public Function fn_Draft(ByVal bindingContext As obj_UiBindingContext) As Object
    Dim row As Object
    Dim field As Variant

    Set row = ex_TDB_History.fn_NewRow()
    row("RowId") = fn_Value(bindingContext, "BasisDraftId")
    If VBA.Len(VBA.CStr(row("RowId"))) = 0 Then row("RowId") = ex_TDB_History.fn_NewId()
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        row(field) = fn_Value(bindingContext, VBA.CStr(field))
    Next field
    fn_Validate row, True
    If Not VBA.CBool(fn_Value(bindingContext, "AutoFillPersonnel")) Then
        row("BasisRenderedText") = fn_Value(bindingContext, "BasisPreview")
        If VBA.Len(VBA.CStr(row("BasisRenderedText"))) = 0 Then
            VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("basisChain.text.ManualRequired")
        End If
    End If
    Set fn_Draft = row
End Function

Public Sub fn_Validate( _
    ByVal row As Object, _
    Optional ByVal capture As Boolean = False _
)
    Dim copy As Object
    Dim key As Variant

    Set copy = VBA.CreateObject("Scripting.Dictionary")
    For Each key In row.Keys
        copy.Add VBA.CStr(key), row(key)
    Next key
    If Not capture And VBA.Len(VBA.CStr(copy("BasisTemplate"))) > 0 Then
        If copy("BasisTemplate") = "xml.v1" And VBA.Len(VBA.CStr(copy("BasisRenderedText"))) = 0 Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        private_ValidateRequired copy, VBA.CStr(copy("BasisRequiredFields"))
        If VBA.Len(VBA.CStr(copy("BasisTypeId"))) = 0 Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        If VBA.Len(VBA.CStr(copy("BasisInstitutionKey"))) = 0 Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
    Else
        private_Validate copy
    End If
    If capture Then
        For Each key In copy.Keys
            row(key) = copy(key)
        Next key
    End If
End Sub

Private Sub private_Validate(ByVal row As Object)
    Dim definition As Object
    Dim official As Object
    Dim field As Variant
    Dim fields As String
    Dim required As String
    Dim group As String
    Dim label As String
    Dim incomingNumber As String
    Dim incomingDate As String
    Dim numberField As Variant

    group = VBA.CStr(row("BasisInstitutionGroup"))
    If VBA.Len(group) = 0 Then
        group = private_Group(VBA.CStr(row("BasisInstitution")))
    End If
    Set definition = ex_TDB_BasisConfig.fn_Definition("Documents", VBA.CStr(row("BasisType")))
    If Not ex_TDB_BasisConfig.fn_Allowed("DocumentRules", group, VBA.CStr(definition("Id"))) Then
        private_Required "BasisType"
    End If
    row("BasisTypeId") = definition("Id")
    row("BasisTemplate") = "xml.v1"
    row("BasisRequiredFields") = definition("Required")
    row("BasisTypeGenitive") = definition("Genitive")
    fields = VBA.CStr(definition("Fields"))
    required = VBA.CStr(definition("Required"))
    If VBA.Len(VBA.CStr(row("BasisInstitutionName"))) = 0 Then
        private_Required "BasisInstitutionQuery"
    End If
    If VBA.Len(VBA.CStr(row("BasisInstitutionKey"))) = 0 Then
        ' Legacy records keep their saved institution text until selected again.
        row("BasisInstitutionKey") = row("BasisInstitutionName")
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.editorFields"), "|")
        If Not private_Has(fields, VBA.CStr(field)) Then
            row(field) = VBA.vbNullString
        End If
    Next field
    private_ValidateRequired row, required, False
    For Each numberField In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.numberFields"), "|")
        row(numberField) = VBA.Trim$(VBA.CStr(row(numberField)))
        If VBA.Left$(row(numberField), 1) = "№" Then
            row(numberField) = VBA.Trim$(VBA.Mid$(row(numberField), 2))
        End If
    Next numberField
    private_ValidateRequired row, required, False
    If (VBA.Len(VBA.CStr(row("BasisNumber"))) > 0) <> (VBA.Len(VBA.CStr(row("BasisDate"))) > 0) Then
        private_Required "BasisDate"
    End If
    If (VBA.Len(VBA.CStr(row("BasisRefNumber"))) > 0) <> (VBA.Len(VBA.CStr(row("BasisRefDate"))) > 0) Then
        private_Required "BasisRefDate"
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.dateFields"), "|")
        If VBA.Len(VBA.CStr(row(field))) > 0 Then
            On Error GoTo InvalidDate
            row(field) = VBA.Format$(ex_TDB_History.fn_Date(row(field)), "dd.mm.yyyy")
            On Error GoTo 0
        End If
    Next field
    incomingNumber = VBA.CStr(row("BasisIncomingNumber"))
    incomingDate = VBA.CStr(row("BasisIncomingDate"))
    If (VBA.Len(incomingNumber) > 0) <> (VBA.Len(incomingDate) > 0) Then
        private_Required "BasisIncomingNumber"
    End If
    row("BasisIssuerId") = VBA.vbNullString
    row("BasisIssuerText") = VBA.vbNullString
    If VBA.Len(VBA.CStr(row("BasisIssuer"))) > 0 Then
        Set official = ex_TDB_BasisConfig.fn_Definition("Officials", VBA.CStr(row("BasisIssuer")))
        If Not ex_TDB_BasisConfig.fn_Allowed("OfficialRules", private_OfficialGroup(VBA.CStr(row("BasisInstitution")), group), VBA.CStr(official("Id"))) Then
            VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("basisChain.text.InvalidIssuer")
        End If
        row("BasisIssuerId") = official("Id")
        row("BasisIssuerText") = official("Text")
    End If
    If VBA.Len(VBA.CStr(row("BasisActing"))) > 0 Then
        If Not private_Contains(ex_TravelDocumentsBuilder.fn_Config("basisChain.acting.options"), VBA.CStr(row("BasisActing"))) Then
            private_Required "BasisActing"
        End If
        If row("BasisActing") = ex_TravelDocumentsBuilder.fn_Config("basisChain.acting.option") And VBA.Len(VBA.CStr(row("BasisIssuerText"))) = 0 Then
            private_Required "BasisIssuer"
        End If
    End If
    label = VBA.CStr(row("BasisForm"))
    If VBA.Len(label) > 0 Then
        If Not private_Has(VBA.Replace$(VBA.CStr(definition("Forms")), ",", "|"), label) Then
            private_Required "BasisForm"
        End If
    End If
    If VBA.Len(VBA.CStr(row("BasisScope"))) > 0 Then
        If Not private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.scope.values"), VBA.CStr(row("BasisScope"))) Then
            private_Required "BasisScope"
        End If
    End If
    If VBA.Len(VBA.CStr(row("BasisRefScope"))) > 0 Then
        If Not private_Has(ex_TravelDocumentsBuilder.fn_Config("basisChain.scope.values"), VBA.CStr(row("BasisRefScope"))) Then
            private_Required "BasisRefScope"
        End If
    End If
    row("BasisRenderedText") = ex_TDB_BasisConfig.fn_Render(VBA.CStr(row("BasisTypeId")), row)
    Exit Sub
InvalidDate:
    VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("basisChain.text.InvalidDate")
End Sub

Public Function fn_Build(ByVal bindingContext As obj_UiBindingContext) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim copy As Object
    Dim field As Variant
    Dim pending As Boolean

    For Each row In fn_Read(bindingContext)
        Set copy = ex_TDB_History.fn_NewRow()
        For Each field In VBA.Split("RowId|" & ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
            copy(field) = row(field)
        Next field
        fn_Validate copy
        result.Add copy
    Next row
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.pendingFields"), "|")
        If VBA.Len(fn_Value(bindingContext, VBA.CStr(field))) > 0 Then
            pending = True
        End If
    Next field
    If fn_Value(bindingContext, "BasisDraftBaseline") = private_DraftStamp(bindingContext) Then pending = False
    If VBA.CBool(fn_Value(bindingContext, "BasisManualEdited")) Then pending = True
    If pending Then
        result.Add fn_Draft(bindingContext)
    End If
    If result.Count > VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("basisChain.maxItems")) Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("basisChain.text.Limit")
    End If
    Set fn_Build = result
End Function

Public Function fn_Join( _
    ByVal chain As Collection, _
    ByVal orderNumber As String, _
    ByVal orderDate As Date, _
    Optional ByVal ownText As String = VBA.vbNullString, _
    Optional ByVal includeOrder As Boolean = True _
) As String
    Dim row As Object
    Dim text As String
    Dim field As Variant

    If includeOrder Then
        fn_Join = ownText
        If VBA.Len(fn_Join) = 0 Then
            fn_Join = ex_TDB_BasisConfig.fn_Order(orderNumber, VBA.Format$(orderDate, "dd.mm.yyyy"))
        End If
    End If
    For Each row In chain
        text = fn_Format(row)
        If VBA.Len(fn_Join) > 0 Then
            fn_Join = fn_Join & ex_TravelDocumentsBuilder.fn_Config("basisChain.separator") & VBA.Space$(1)
        End If
        fn_Join = fn_Join & text
    Next row
    If VBA.Len(fn_Join) > 0 Then fn_Join = fn_Join & ex_TravelDocumentsBuilder.fn_Config("basisChain.terminator")
End Function

Public Function fn_SavedOwnText( _
    ByVal chain As Collection, _
    ByVal approved As String _
) As String
    Dim ending As String
    Dim index As Long

    If approved = fn_Join(chain, VBA.vbNullString, 0, VBA.vbNullString, False) Then Exit Function
    ending = ex_TravelDocumentsBuilder.fn_Config("basisChain.terminator")
    For index = chain.Count To 1 Step -1
        ending = ex_TravelDocumentsBuilder.fn_Config("basisChain.separator") & VBA.Space$(1) & fn_Format(chain(index)) & ending
    Next index
    If VBA.Len(approved) <= VBA.Len(ending) Or VBA.Right$(approved, VBA.Len(ending)) <> ending Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    fn_SavedOwnText = VBA.Left$(approved, VBA.Len(approved) - VBA.Len(ending))
End Function

Public Function fn_FromGroup(ByVal group As Collection) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim parentId As String
    Dim found As Boolean
    Dim count As Long

    parentId = group(1)("RowId")
    For Each row In group
        If row("RowType") = "Basis" Then
            count = count + 1
        End If
    Next row
    Do
        found = False
        For Each row In group
            If row("RowType") = "Basis" And row("ParentRowId") = parentId Then
                result.Add row
                parentId = row("RowId")
                found = True
                If result.Count > count Then
                    ex_TDB_History.fn_Fail "InvalidChain"
                End If
                Exit For
            End If
        Next row
    Loop While found
    If result.Count <> count Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    Set fn_FromGroup = result
End Function

Public Sub fn_ClearDraft(ByVal bindingContext As obj_UiBindingContext)
    Dim field As Variant

    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        If Not bindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
    Next field
    If Not bindingContext.SetValue("Form", "BasisDraftId", VBA.vbNullString) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If Not bindingContext.SetValue("Form", "BasisPreview", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not bindingContext.SetValue("Form", "BasisInstitutionQuery", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not fn_RefreshInstitutions(bindingContext) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    fn_MarkDraftBaseline bindingContext
End Sub

Public Sub fn_PrepareNextDraft(ByVal bindingContext As obj_UiBindingContext)
    Dim field As Variant

    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.afterAdd.clearFields"), "|")
        If Not bindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
    Next field
    If Not bindingContext.SetValue("Form", "BasisDraftId", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not bindingContext.SetValue("Form", "BasisRenderedText", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not bindingContext.SetValue("Form", "BasisPreview", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not fn_RefreshInstitutions(bindingContext) Then ex_TDB_History.fn_Fail "InvalidChain"
    fn_MarkDraftBaseline bindingContext
    If Not bindingContext.SetValue("Form", "BasisPreview", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidChain"
End Sub

Public Sub fn_MarkDraftBaseline(ByVal bindingContext As obj_UiBindingContext)
    If Not bindingContext.SetValue("Form", "BasisManualEdited", False) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not bindingContext.SetValue("Form", "BasisDraftBaseline", private_DraftStamp(bindingContext)) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
End Sub

Public Function fn_RefreshInstitutions(ByVal bindingContext As obj_UiBindingContext) As Boolean
    Dim entry As Variant
    Dim parts As Variant
    Dim institutions As String
    Dim group As String
    Dim types As String
    Dim roles As String
    Dim field As Variant
    Dim definition As Object
    Dim fields As String
    Dim forms As String
    Dim visibility As String
    Dim snapshot As obj_TableSnapshotSet

    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.institutions"), "|")
        parts = VBA.Split(VBA.CStr(entry), "~")
        If UBound(parts) <> 1 Then ex_TDB_History.fn_Fail "InvalidChain"
        private_Append institutions, VBA.CStr(parts(1))
    Next entry
    If Not bindingContext.SetValue("Data", "BasisInstitutions", institutions) Then
        Exit Function
    End If
    group = fn_Value(bindingContext, "BasisInstitutionGroup")
    If VBA.Len(group) = 0 Then
        group = private_Group(fn_Value(bindingContext, "BasisInstitution"))
    End If
    Set snapshot = ex_TableSnapshots.fn_Active()
    If VBA.Len(group) > 0 And Not snapshot Is Nothing Then
        types = ex_TDB_BasisConfig.fn_Options("DocumentRules", group)
        roles = ex_TDB_BasisConfig.fn_Options("OfficialRules", private_OfficialGroup(fn_Value(bindingContext, "BasisInstitution"), group))
    End If
    If Not private_Contains(types, fn_Value(bindingContext, "BasisType")) Then
        If Not bindingContext.SetValue("Form", "BasisType", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    If Not private_Contains(roles, fn_Value(bindingContext, "BasisIssuer")) Then
        If Not bindingContext.SetValue("Form", "BasisIssuer", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    If Not bindingContext.SetValue("Data", "BasisTypes", types) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Data", "BasisIssuers", roles) Then
        Exit Function
    End If
    visibility = "Collapsed"
    If VBA.Len(group) > 0 Then
        visibility = "Visible"
    End If
    If Not bindingContext.SetValue("Form", "BasisInstitutionVisibility", visibility) Then
        Exit Function
    End If
    If VBA.Len(fn_Value(bindingContext, "BasisType")) > 0 Then
        Set definition = ex_TDB_BasisConfig.fn_Definition("Documents", fn_Value(bindingContext, "BasisType"))
        fields = VBA.CStr(definition("Fields"))
        forms = VBA.CStr(definition("Forms"))
    End If
    If Not private_Contains(forms, fn_Value(bindingContext, "BasisForm")) Then
        If Not bindingContext.SetValue("Form", "BasisForm", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    If Not bindingContext.SetValue("Data", "BasisForms", forms) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Data", "BasisActingOptions", ex_TravelDocumentsBuilder.fn_Config("basisChain.acting.options")) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Data", "BasisScopes", VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.scope.values"), "|", ",")) Then
        Exit Function
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.editorFields"), "|")
        If Not snapshot Is Nothing And VBA.Len(fields) > 0 Then
            If Not bindingContext.SetValue("Text", VBA.CStr(field), ex_TDB_BasisConfig.fn_Caption(VBA.CStr(field))) Then
                Exit Function
            End If
        End If
        visibility = "Collapsed"
        If private_Has(fields, VBA.CStr(field)) Then
            visibility = "Visible"
        Else
            If Not bindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then
                Exit Function
            End If
        End If
        If Not bindingContext.SetValue("Form", VBA.CStr(field) & "Visibility", visibility) Then
            Exit Function
        End If
    Next field
    fn_RefreshInstitutions = fn_Preview(bindingContext)
End Function

Public Function fn_Preview(ByVal bindingContext As obj_UiBindingContext) As Boolean
    Dim row As Object
    Dim text As String

    On Error GoTo Incomplete
    If Not VBA.CBool(fn_Value(bindingContext, "AutoFillPersonnel")) Then
        fn_Preview = True
        Exit Function
    End If
    If fn_Value(bindingContext, "BasisDraftBaseline") = private_DraftStamp(bindingContext) Then
        fn_Preview = bindingContext.SetValue("Form", "BasisPreview", VBA.vbNullString)
        Exit Function
    End If
    If VBA.Len(fn_Value(bindingContext, "BasisType")) = 0 Then
        fn_Preview = bindingContext.SetValue("Form", "BasisPreview", VBA.vbNullString)
        Exit Function
    End If
    Set row = fn_Draft(bindingContext)
    text = fn_Format(row)
Publish:
    fn_Preview = bindingContext.SetValue("Form", "BasisPreview", text)
    Exit Function
Incomplete:
    text = VBA.Err.Description
    Resume Publish
End Function

Public Function fn_Format(ByVal row As Object) As String
    Dim text As String
    Dim field As Variant

    If VBA.Len(VBA.CStr(row("BasisRenderedText"))) > 0 Then
        fn_Format = VBA.CStr(row("BasisRenderedText"))
        Exit Function
    End If
    text = VBA.CStr(row("BasisTemplate"))
    If text = "xml.v1" Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If VBA.Len(text) = 0 Then
        text = ex_TravelDocumentsBuilder.fn_Config("basisChain.documentTemplate")
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        text = VBA.Replace$(text, "{Value." & field & "}", VBA.CStr(row(field)))
    Next field
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields") & "|Incoming|Acting|Form|Scope|Number|Date|RefScope", "|")
        text = VBA.Replace$(text, "{" & field & "}", private_DocumentField(row, VBA.CStr(field)))
    Next field
    Do While VBA.InStr(text, "  ") > 0
        text = VBA.Replace$(text, "  ", " ")
    Loop
    If VBA.InStr(text, "{") > 0 Or VBA.InStr(text, "}") > 0 Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    fn_Format = VBA.Trim$(text)
End Function

Public Function fn_SearchInstitutions( _
    ByVal institution As String, _
    ByVal searchText As String _
) As obj_LookupResult
    Dim tableSource As New obj_TableSource
    Dim tableQuery As New obj_TableQuery
    Dim snapshot As obj_TableSnapshotSet
    Dim dataTable As obj_DataTable
    Dim lookupResult As New obj_LookupResult
    Dim lookupRecord As obj_LookupRecord
    Dim records As New Collection
    Dim columns As New Collection
    Dim fields As Object
    Dim selectedHeaders As Object
    Dim header As String
    Dim column As Object
    Dim fileSystem As Object
    Dim field As Variant
    Dim entry As Variant
    Dim parts As Variant
    Dim value As Variant
    Dim prefix As String
    Dim path As String
    Dim diagnostic As String
    Dim rowIndex As Long
    Dim limit As Long
    Dim hasMore As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim spans As Variant
    Dim columnIndex As Long

    On Error GoTo Failed
    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.institutions"), "|")
        parts = VBA.Split(VBA.CStr(entry), "~")
        If UBound(parts) <> 1 Then ex_TDB_History.fn_Fail "InvalidChain"
        If VBA.CStr(parts(1)) = institution Then prefix = "basisChain.lookup." & parts(0)
    Next entry
    If VBA.Len(prefix) = 0 Then ex_TDB_History.fn_Fail "InvalidChain"
    spans = VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.columnSpans"), "|")
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.fields"), "|")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column.Add "name", VBA.CStr(field)
        column.Add "caption", ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)
        If columnIndex > UBound(spans) Then ex_TDB_History.fn_Fail "InvalidChain"
        column.Add "columnSpan", VBA.CLng(spans(columnIndex))
        columnIndex = columnIndex + 1
        columns.Add column
    Next field
    If columnIndex <> UBound(spans) + 1 Then ex_TDB_History.fn_Fail "InvalidChain"
    limit = VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.maxResults"))
    If ex_TDB_Candidates.fn_IsBrowse(searchText) Then
        limit = ex_TDB_Candidates.fn_BrowseLimit(searchText)
    End If
    If ex_TravelDocumentsBuilder.fn_Config(prefix & ".searchField") = "Key" Then
        searchText = VBA.Trim$(VBA.Replace$(searchText, "A", "А", 1, -1, VBA.vbTextCompare))
    End If
    If ex_TDB_Candidates.fn_IsBrowse(searchText) Or VBA.Len(ex_TableQuery.fn_NormalizeText(searchText)) >= _
            VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.minChars")) Then
        Set snapshot = ex_TableSnapshots.fn_Active()
        If snapshot Is Nothing Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
        If Not tableSource.Initialize() Or Not tableQuery.Initialize() Then VBA.Err.Raise 5, , "Cannot initialize institution lookup."
        Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
        path = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.path")
        If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
            path = fileSystem.BuildPath(ThisWorkbook.Path, path)
        End If
        tableSource.WorkbookPath = fileSystem.GetAbsolutePathName(path)
        tableSource.SheetName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.sheet")
        tableSource.TableName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.table")
        Set selectedHeaders = VBA.CreateObject("Scripting.Dictionary")
        selectedHeaders.CompareMode = VBA.vbTextCompare
        For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.sourceFields"), "|")
            header = ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)
            If Not selectedHeaders.Exists(header) Then
                selectedHeaders.Add header, True
                tableQuery.AddColumn header
            End If
        Next field
        If Not ex_TDB_Candidates.fn_IsBrowse(searchText) Then
            tableQuery.AddCondition ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & _
                ex_TravelDocumentsBuilder.fn_Config(prefix & ".searchField")), QueryContains, searchText, QueryText, True
        End If
        tableQuery.Limit = limit + 1
        If Not snapshot.TryQuery(tableSource, tableQuery, dataTable, diagnostic) Then VBA.Err.Raise 5, , diagnostic
        hasMore = (dataTable.RowCount > limit)
        For rowIndex = 1 To dataTable.RowCount
            If rowIndex > limit Then Exit For
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.sourceFields"), "|")
                value = dataTable.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex( _
                    dataTable.Headers, ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)))
                If VBA.IsError(value) Then ex_TDB_History.fn_Fail "InvalidChain"
                If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString
                fields.Add VBA.CStr(field), VBA.Trim$(VBA.CStr(value))
            Next field
            ' Numbered units have configured names even without a custom full title.
            If ex_TravelDocumentsBuilder.fn_Config(prefix & ".searchField") = "Key" Then
                If VBA.Len(VBA.CStr(fields("Name"))) = 0 Then
                    fields("Name") = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config(prefix & ".nameTemplate"), _
                        "{Key}", VBA.CStr(fields("Key")))
                    If VBA.Len(VBA.CStr(fields("Genitive"))) = 0 Then
                        fields("Genitive") = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config(prefix & ".genitiveTemplate"), _
                            "{Key}", VBA.CStr(fields("Key")))
                    End If
                End If
            End If
            If VBA.Len(VBA.CStr(fields("Group"))) = 0 Then
                VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("basisChain.lookup.text.MissingGroup")
            End If
            fields.Add "InstitutionKind", institution
            Set lookupRecord = New obj_LookupRecord
            If Not lookupRecord.Initialize(VBA.CStr(fields("Key")), fields) Then VBA.Err.Raise 5, , diagnostic
            records.Add lookupRecord
        Next rowIndex
    End If
    If Not lookupResult.Initialize(records, columns, hasMore, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    Set fn_SearchInstitutions = lookupResult
Cleanup:
    If Not dataTable Is Nothing Then dataTable.Dispose
    tableQuery.Dispose
    tableSource.Dispose
    If errorNumber <> 0 Then VBA.Err.Raise errorNumber, "Basis.Institutions", errorDescription
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_IncludeOrder(ByVal bindingContext As obj_UiBindingContext) As Boolean
    Dim value As Variant
    Dim objectValue As Object
    Dim isObject As Boolean

    If Not bindingContext.TryGetValue("Form", "BasisIncludeOrder", value, objectValue, isObject) Then
        VBA.Err.Raise 5, , "Required binding not found: Form.BasisIncludeOrder"
    End If
    If isObject Or VBA.VarType(value) <> VBA.vbBoolean Then
        VBA.Err.Raise 5, , "Invalid Boolean binding: Form.BasisIncludeOrder"
    End If
    fn_IncludeOrder = VBA.CBool(value)
End Function

Public Function fn_Publish(ByVal bindingContext As obj_UiBindingContext) As Boolean
    Dim chain As Collection
    Dim text As String
    Dim number As String
    Dim dateText As String
    Dim records As New Collection
    Dim columns As New Collection
    Dim row As Object
    Dim column As Object
    Dim fields As Object
    Dim record As obj_LookupRecord
    Dim result As New obj_LookupResult
    Dim diagnostic As String
    Dim index As Long
    Dim emptyChain As New Collection
    Dim field As Variant
    Dim scalar As Variant
    Dim previous As Object
    Dim isObject As Boolean

    Set chain = fn_Read(bindingContext)
    number = fn_Value(bindingContext, "AcceptedOrderNumber")
    dateText = fn_Value(bindingContext, "AcceptedOrderDate")
    If VBA.Len(fn_Value(bindingContext, "BasisOwnText")) > 0 Then
        text = fn_Value(bindingContext, "BasisOwnText") & ex_TravelDocumentsBuilder.fn_Config("basisChain.terminator")
    ElseIf VBA.Len(number) > 0 And VBA.Len(dateText) > 0 Then
        text = fn_Join(emptyChain, number, ex_TDB_History.fn_Date(dateText))
    Else
        text = ex_TDB_BasisConfig.fn_Order(ex_TravelDocumentsBuilder.fn_Config("basisChain.text.OrderNumberPending"), _
            ex_TravelDocumentsBuilder.fn_Config("basisChain.text.OrderDatePending")) & _
            ex_TravelDocumentsBuilder.fn_Config("basisChain.terminator")
    End If
    Set column = VBA.CreateObject("Scripting.Dictionary")
    column.Add "name", "BasisText"
    column.Add "caption", ex_TravelDocumentsBuilder.fn_Config("text.BasisTitle")
    column.Add "columnSpan", 2
    columns.Add column
    Set column = VBA.CreateObject("Scripting.Dictionary")
    column.Add "name", "Remove"
    column.Add "caption", VBA.vbNullString
    column.Add "columnSpan", 1
    columns.Add column
    For index = 0 To chain.Count
        If index > 0 Then
            Set row = chain(index)
            text = fn_Format(row)
        End If
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields.Add "BasisText", VBA.Replace$(VBA.Replace$( _
            ex_TravelDocumentsBuilder.fn_Config("basisChain.itemTemplate"), _
            "{index}", VBA.CStr(index + 1)), "{text}", text)
        fields.Add "Remove", VBA.vbNullString
        fields.Add "BasisRowId", VBA.vbNullString
        fields.Add "ActionCaption", VBA.vbNullString
        fields.Add "ActionVisibility", "Visible"
        fields.Add "ActionStyle", ex_TravelDocumentsBuilder.fn_Config("basisChain.styles.checkbox")
        fields.Add "Checked", fn_IncludeOrder(bindingContext)
        fields.Add "CheckedGlyph", ex_TravelDocumentsBuilder.fn_Config("text.CheckboxChecked")
        fields.Add "UncheckedGlyph", ex_TravelDocumentsBuilder.fn_Config("text.CheckboxUnchecked")
        fields.Add "TextStyle", ex_TravelDocumentsBuilder.fn_Config("basisChain.styles.normal")
        If index = 0 And Not fn_IncludeOrder(bindingContext) Then
            fields("TextStyle") = ex_TravelDocumentsBuilder.fn_Config("basisChain.styles.excluded")
        End If
        If index > 0 Then
            fields("BasisRowId") = VBA.CStr(row("RowId"))
            fields("Checked") = False
            fields("ActionVisibility") = "Collapsed"
            fields("ActionStyle") = ex_TravelDocumentsBuilder.fn_Config("basisChain.styles.remove")
            fields("CheckedGlyph") = ex_TravelDocumentsBuilder.fn_Config("text.RemoveBasis")
            fields("UncheckedGlyph") = ex_TravelDocumentsBuilder.fn_Config("text.RemoveBasis")
        End If
        Set record = New obj_LookupRecord
        If Not record.Initialize(VBA.CStr(index), fields) Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        records.Add record
    Next index
    If Not result.Initialize(records, columns, False, diagnostic) Then
        VBA.Err.Raise 5, , diagnostic
    End If
    If bindingContext.TryGetValue("Data", "BasisView", scalar, previous, isObject) Then
        If Not isObject Then
            Set previous = Nothing
        End If
    End If
    If Not bindingContext.SetObject("Data", "BasisView", result) Then
        Exit Function
    End If
    If Not previous Is Nothing Then
        previous.Dispose
    End If
    fn_Publish = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_DraftStamp(ByVal bindingContext As obj_UiBindingContext) As String
    Dim field As Variant
    Dim value As String

    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.pendingFields"), "|")
        value = fn_Value(bindingContext, VBA.CStr(field))
        private_DraftStamp = private_DraftStamp & VBA.CStr(VBA.Len(value)) & ":" & value
    Next field
End Function

Private Function private_DocumentField( _
    ByVal row As Object, _
    ByVal field As String _
) As String
    Dim definition As Object

    Select Case field
        Case "Number"
            If VBA.Len(VBA.CStr(row("BasisNumber"))) > 0 Then
                private_DocumentField = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.numberTemplate"), "{number}", row("BasisNumber"))
            End If
        Case "Date"
            If VBA.Len(VBA.CStr(row("BasisDate"))) > 0 Then
                private_DocumentField = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.dateTemplate"), "{date}", row("BasisDate"))
            End If
        Case "RefScope"
            If VBA.Len(VBA.CStr(row("BasisRefScope"))) > 0 And row("BasisRefScope") <> ex_TravelDocumentsBuilder.fn_Config("basisChain.scope.none") Then
                private_DocumentField = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.scopeTemplate"), "{scope}", row("BasisRefScope"))
            End If
        Case "BasisProceedingNumber"
            If VBA.Len(VBA.CStr(row(field))) > 0 Then
                private_DocumentField = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.proceedingTemplate"), "{number}", row(field))
            End If
        Case "BasisType"
            private_DocumentField = VBA.CStr(row("BasisType"))
            If row("BasisForm") = ex_TravelDocumentsBuilder.fn_Config("basisChain.forms.copy") Or _
                    row("BasisForm") = ex_TravelDocumentsBuilder.fn_Config("basisChain.forms.extract") Then
                private_DocumentField = VBA.CStr(row("BasisTypeGenitive"))
            End If
        Case "Incoming"
            If VBA.Len(VBA.CStr(row("BasisIncomingNumber"))) > 0 Then
                private_DocumentField = VBA.Replace$(VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.incomingTemplate"), _
                    "{number}", VBA.CStr(row("BasisIncomingNumber"))), "{date}", VBA.CStr(row("BasisIncomingDate")))
            End If
        Case "Acting"
            If VBA.CStr(row("BasisActing")) = ex_TravelDocumentsBuilder.fn_Config("basisChain.acting.option") Then
                private_DocumentField = ex_TravelDocumentsBuilder.fn_Config("basisChain.acting.text")
            End If
        Case "Form"
            If row("BasisForm") = ex_TravelDocumentsBuilder.fn_Config("basisChain.forms.copy") Then
                private_DocumentField = ex_TravelDocumentsBuilder.fn_Config("basisChain.forms.copyText")
            End If
            If row("BasisForm") = ex_TravelDocumentsBuilder.fn_Config("basisChain.forms.extract") Then
                private_DocumentField = ex_TravelDocumentsBuilder.fn_Config("basisChain.forms.extractText")
            End If
        Case "Scope"
            If VBA.Len(VBA.CStr(row("BasisScope"))) > 0 And row("BasisScope") <> ex_TravelDocumentsBuilder.fn_Config("basisChain.scope.none") Then
                private_DocumentField = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.scopeTemplate"), "{scope}", row("BasisScope"))
            End If
        Case "BasisIssuer", "BasisIssuerText"
            private_DocumentField = VBA.CStr(row("BasisIssuerText"))
            If VBA.Len(private_DocumentField) = 0 And VBA.Len(VBA.CStr(row("BasisIssuer"))) > 0 Then
                Set definition = ex_TDB_BasisConfig.fn_Definition("Officials", VBA.CStr(row("BasisIssuer")))
                private_DocumentField = VBA.CStr(definition("Text"))
            End If
        Case Else
            private_DocumentField = VBA.CStr(row(field))
    End Select
End Function

Private Function private_Group(ByVal label As String) As String
    Dim entry As Variant
    Dim parts As Variant

    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.institutions"), "|")
        parts = VBA.Split(VBA.CStr(entry), "~")
        If UBound(parts) <> 1 Then ex_TDB_History.fn_Fail "InvalidChain"
        If VBA.CStr(parts(1)) = label Then
            private_Group = VBA.CStr(parts(0))
            Exit Function
        End If
    Next entry
End Function

Private Sub private_Required(ByVal field As String)
    VBA.Err.Raise 5, , VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("generation.message.Required"), _
        "{value}", ex_TravelDocumentsBuilder.fn_Config("text." & field))
End Sub

Private Sub private_ValidateRequired( _
    ByVal row As Object, _
    ByVal specification As String, _
    Optional ByVal allowLegacy As Boolean = True _
)
    Dim position As Long
    Dim missing As String
    Dim fulfilled As Boolean

    If VBA.Len(VBA.Trim$(specification)) = 0 Then
        Exit Sub
    End If
    ' Only saved snapshots may retain the original separator-based grammar.
    If allowLegacy And VBA.InStr(specification, "(") = 0 And VBA.InStr(specification, ")") = 0 Then
        private_ValidateLegacyRequired row, specification
        Exit Sub
    End If
    position = 1
    fulfilled = private_RequiredExpression(row, specification, position, missing, 0)
    private_RequiredSkipSpace specification, position
    If position <= VBA.Len(specification) Then
        private_RequiredSyntax specification
    End If
    If Not fulfilled Then
        private_Required missing
    End If
End Sub

Private Function private_RequiredExpression( _
    ByVal row As Object, _
    ByVal specification As String, _
    ByRef position As Long, _
    ByRef missing As String, _
    ByVal depth As Long _
) As Boolean
    Dim fulfilled As Boolean
    Dim nextValue As Boolean
    Dim nextMissing As String
    Dim operation As String
    Dim nextOperation As String

    If depth > 64 Then
        private_RequiredSyntax specification
    End If
    fulfilled = private_RequiredTerm(row, specification, position, missing, depth)
    Do
        private_RequiredSkipSpace specification, position
        nextOperation = VBA.Mid$(specification, position, 1)
        If VBA.Len(nextOperation) = 0 Or nextOperation = ")" Then
            Exit Do
        End If
        Select Case nextOperation
            Case "&", "|", "/"
                ' Mixed operators at one level require explicit nested parentheses.
                If VBA.Len(operation) > 0 And operation <> nextOperation Then
                    private_RequiredSyntax specification
                End If
            Case Else
                private_RequiredSyntax specification
        End Select
        operation = nextOperation
        position = position + 1
        nextMissing = VBA.vbNullString
        nextValue = private_RequiredTerm(row, specification, position, nextMissing, depth)
        If operation = "/" Then
            fulfilled = fulfilled Or nextValue
            If fulfilled Then
                missing = VBA.vbNullString
            End If
        Else
            If fulfilled And Not nextValue Then
                missing = nextMissing
            End If
            fulfilled = fulfilled And nextValue
        End If
    Loop
    private_RequiredExpression = fulfilled
End Function

Private Function private_RequiredTerm( _
    ByVal row As Object, _
    ByVal specification As String, _
    ByRef position As Long, _
    ByRef missing As String, _
    ByVal depth As Long _
) As Boolean
    Dim field As String
    Dim character As String
    Dim start As Long

    private_RequiredSkipSpace specification, position
    If VBA.Mid$(specification, position, 1) = "(" Then
        position = position + 1
        private_RequiredTerm = private_RequiredExpression(row, specification, position, missing, depth + 1)
        private_RequiredSkipSpace specification, position
        If VBA.Mid$(specification, position, 1) <> ")" Then
            private_RequiredSyntax specification
        End If
        position = position + 1
        Exit Function
    End If
    start = position
    character = VBA.Mid$(specification, position, 1)
    If Not character Like "[A-Za-z]" Then
        private_RequiredSyntax specification
    End If
    Do While position <= VBA.Len(specification)
        character = VBA.Mid$(specification, position, 1)
        If Not character Like "[A-Za-z0-9_]" Then
            Exit Do
        End If
        position = position + 1
    Loop
    field = VBA.Mid$(specification, start, position - start)
    If Not row.Exists(field) Then
        private_RequiredSyntax specification
    End If
    private_RequiredTerm = (VBA.Len(VBA.Trim$(VBA.CStr(row(field)))) > 0)
    If Not private_RequiredTerm Then
        missing = field
    End If
End Function

Private Sub private_RequiredSkipSpace( _
    ByVal specification As String, _
    ByRef position As Long _
)
    Dim character As String

    Do While position <= VBA.Len(specification)
        character = VBA.Mid$(specification, position, 1)
        If character <> " " And character <> VBA.vbTab Then
            Exit Do
        End If
        position = position + 1
    Loop
End Sub

Private Sub private_RequiredSyntax(ByVal specification As String)
    VBA.Err.Raise 5, "Basis.RequiredFields", VBA.Replace$( _
        ex_TravelDocumentsBuilder.fn_Config("basisChain.text.InvalidRequiredSpec"), "{value}", specification)
End Sub

Private Sub private_ValidateLegacyRequired( _
    ByVal row As Object, _
    ByVal specification As String _
)
    Dim requirement As Variant
    Dim alternative As Variant
    Dim field As Variant
    Dim fulfilled As Boolean
    Dim complete As Boolean
    Dim firstField As String

    For Each requirement In VBA.Split(specification, "|")
        If VBA.Len(VBA.CStr(requirement)) > 0 Then
            fulfilled = False
            firstField = VBA.Split(VBA.Split(VBA.CStr(requirement), "/")(0), "&")(0)
            For Each alternative In VBA.Split(VBA.CStr(requirement), "/")
                complete = True
                For Each field In VBA.Split(VBA.CStr(alternative), "&")
                    If VBA.Len(VBA.Trim$(VBA.CStr(row(field)))) = 0 Then
                        complete = False
                    End If
                Next field
                If complete Then
                    fulfilled = True
                End If
            Next alternative
            If Not fulfilled Then
                private_Required firstField
            End If
        End If
    Next requirement
End Sub

Private Function private_Has( _
    ByVal values As String, _
    ByVal value As String _
) As Boolean
    private_Has = (VBA.Len(value) > 0 And VBA.InStr(1, "|" & values & "|", "|" & value & "|", VBA.vbTextCompare) > 0)
End Function

Private Function private_Contains( _
    ByVal values As String, _
    ByVal value As String _
) As Boolean
    private_Contains = private_Has(VBA.Replace$(values, ",", "|"), value)
End Function

Private Sub private_Append( _
    ByRef values As String, _
    ByVal value As String _
)
    If VBA.Len(values) > 0 Then
        values = values & ","
    End If
    values = values & value
End Sub

Private Function private_OfficialGroup( _
    ByVal institution As String, _
    ByVal sourceGroup As String _
) As String
    Dim entry As Variant
    Dim parts As Variant
    Dim kind As String

    kind = private_Group(institution)
    private_OfficialGroup = sourceGroup
    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.model.officialGroupOverrides"), "|")
        parts = VBA.Split(VBA.CStr(entry), "~")
        If UBound(parts) <> 1 Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        If VBA.CStr(parts(0)) = kind Then
            private_OfficialGroup = VBA.CStr(parts(1))
        End If
    Next entry
End Function