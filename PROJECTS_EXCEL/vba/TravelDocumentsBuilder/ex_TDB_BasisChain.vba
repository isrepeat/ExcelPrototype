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

    If Not bindingContext.TryGetValue("Data", "BasisChain", scalar, source, isObject) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If Not isObject Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    If Not TypeOf source Is Collection Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
    Set fn_Read = source
End Function

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
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        row(field) = fn_Value(bindingContext, VBA.CStr(field))
    Next field
    fn_Validate row
    Set fn_Draft = row
End Function

Public Sub fn_Validate(ByVal row As Object)
    Dim field As Variant
    Dim roles As String

    row("BasisNumber") = VBA.Trim$(VBA.CStr(row("BasisNumber")))
    If VBA.Left$(row("BasisNumber"), 1) = "№" Then
        row("BasisNumber") = VBA.Trim$(VBA.Mid$(row("BasisNumber"), 2))
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        If field <> "BasisIssuer" And VBA.Len(VBA.Trim$(VBA.CStr(row(field)))) = 0 Then
            VBA.Err.Raise 5, , VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("generation.message.Required"), _
                "{value}", ex_TravelDocumentsBuilder.fn_Config("text." & field))
        End If
    Next field
    roles = private_Roles(VBA.CStr(row("BasisInstitution")))
    If VBA.Len(roles) > 0 And VBA.Len(VBA.CStr(row("BasisIssuer"))) > 0 Then
        If VBA.InStr(1, "," & roles & ",", "," & row("BasisIssuer") & ",", VBA.vbTextCompare) = 0 Then
            VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("basisChain.text.InvalidIssuer")
        End If
    End If
    On Error GoTo InvalidDate
    row("BasisDate") = VBA.Format$(ex_TDB_History.fn_Date(row("BasisDate")), "dd.mm.yyyy")
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
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        If VBA.Len(fn_Value(bindingContext, VBA.CStr(field))) > 0 Then
            pending = True
        End If
    Next field
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
    ByVal orderDate As Date _
) As String
    Dim row As Object
    Dim text As String
    Dim field As Variant

    fn_Join = VBA.Replace$(VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.orderTemplate"), _
        "{number}", orderNumber), "{date}", VBA.Format$(orderDate, "dd.mm.yyyy"))
    For Each row In chain
        text = ex_TravelDocumentsBuilder.fn_Config("basisChain.documentTemplate")
        For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
            text = VBA.Replace$(text, "{" & field & "}", VBA.CStr(row(field)))
        Next field
        text = VBA.Trim$(VBA.Replace$(text, "  ", " "))
        fn_Join = fn_Join & ex_TravelDocumentsBuilder.fn_Config("basisChain.separator") & VBA.Space$(1) & text
    Next row
    fn_Join = fn_Join & ex_TravelDocumentsBuilder.fn_Config("basisChain.terminator")
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
    If Not fn_RefreshInstitutions(bindingContext) Then
        ex_TDB_History.fn_Fail "InvalidChain"
    End If
End Sub

Public Function fn_RefreshInstitutions( _
    ByVal bindingContext As obj_UiBindingContext, _
    Optional ByVal institutionChanged As Boolean = False _
) As Boolean
    Dim entry As Variant
    Dim parts As Variant
    Dim institutions As String
    Dim roles As String
    Dim issuer As String
    Dim institution As String

    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.institutions"), "|")
        parts = VBA.Split(VBA.CStr(entry), "~")
        If UBound(parts) <> 1 Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        If VBA.Len(institutions) > 0 Then
            institutions = institutions & ","
        End If
        institutions = institutions & parts(1)
    Next entry
    institution = fn_Value(bindingContext, "BasisInstitution")
    roles = private_Roles(institution)
    issuer = fn_Value(bindingContext, "BasisIssuer")
    If institutionChanged And VBA.Len(roles) = 0 Then
        If Not bindingContext.SetValue("Form", "BasisIssuer", VBA.vbNullString) Then
            Exit Function
        End If
    End If
    If VBA.Len(institution) = 0 Or (VBA.Len(roles) > 0 And VBA.Len(issuer) > 0) Then
        If VBA.Len(institution) = 0 Or VBA.InStr(1, "," & roles & ",", "," & issuer & ",", VBA.vbTextCompare) = 0 Then
            If Not bindingContext.SetValue("Form", "BasisIssuer", VBA.vbNullString) Then
                Exit Function
            End If
        End If
    End If
    If Not bindingContext.SetValue("Data", "BasisInstitutions", institutions) Then
        Exit Function
    End If
    fn_RefreshInstitutions = bindingContext.SetValue("Data", "BasisIssuers", roles)
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
    If VBA.Len(number) > 0 And VBA.Len(dateText) > 0 Then
        text = fn_Join(emptyChain, number, ex_TDB_History.fn_Date(dateText))
    Else
        text = ex_TravelDocumentsBuilder.fn_Config("basisChain.text.ChooseOrder")
    End If
    Set column = VBA.CreateObject("Scripting.Dictionary")
    column.Add "name", "BasisText"
    column.Add "caption", ex_TravelDocumentsBuilder.fn_Config("text.BasisTitle")
    column.Add "columnSpan", 3
    columns.Add column
    For index = 0 To chain.Count
        If index > 0 Then
            Set row = chain(index)
            text = ex_TravelDocumentsBuilder.fn_Config("basisChain.documentTemplate")
            For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
                text = VBA.Replace$(text, "{" & field & "}", VBA.CStr(row(field)))
            Next field
            text = VBA.Trim$(VBA.Replace$(text, "  ", " "))
        End If
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields.Add "BasisText", text
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

Private Function private_Roles(ByVal institution As String) As String
    Dim entry As Variant
    Dim parts As Variant

    For Each entry In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("basisChain.institutions"), "|")
        parts = VBA.Split(VBA.CStr(entry), "~")
        If UBound(parts) <> 1 Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        If VBA.StrComp(institution, VBA.CStr(parts(1)), VBA.vbTextCompare) = 0 Then
            private_Roles = ex_TravelDocumentsBuilder.fn_Config("basisChain.roles." & parts(0))
            Exit Function
        End If
    Next entry
End Function