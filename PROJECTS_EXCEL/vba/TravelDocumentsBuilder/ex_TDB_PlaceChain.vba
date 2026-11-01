Attribute VB_Name = "ex_TDB_PlaceChain"
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
Public Function fn_Draft(ByVal bindingContext As obj_UiBindingContext) As Object
    Dim values As Object
    Dim field As Variant
    Dim scalar As Variant
    Dim objectValue As Object
    Dim isObject As Boolean

    Set values = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split("PlaceKind|DocumentTarget|AutoFillPersonnel|VacationAbroad|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
        If Not bindingContext.TryGetValue("Form", VBA.CStr(field), scalar, objectValue, isObject) Then VBA.Err.Raise 5, , VBA.CStr(field)
        If isObject Or VBA.IsError(scalar) Or VBA.IsNull(scalar) Then VBA.Err.Raise 5, , VBA.CStr(field)
        values.Add VBA.CStr(field), scalar
    Next field
    Set fn_Draft = fn_FromValues(values)
End Function

Public Function fn_FromValues(ByVal values As Object) As Object
    Dim row As Object
    Dim field As Variant
    Dim address As String
    Dim caption As String
    Dim isAbroad As Boolean
    Dim prefix As String

    Set row = ex_TDB_History.fn_NewRow()
    isAbroad = ex_TDB_Places.fn_IsAbroad(VBA.CStr(values("VacationAbroad")))
    prefix = ex_TDB_Places.fn_Prefix(isAbroad)
    For Each field In ex_TDB_Places.fn_FormFields()
        row(field) = ex_TableQuery.fn_NormalizeText(VBA.CStr(values(field)))
    Next field
    row("PlaceAbroad") = values("VacationAbroad")
    row("Destination") = ex_TableQuery.fn_NormalizeText(VBA.CStr(values("DocumentTarget")))
    If values("PlaceKind") = ex_TravelDocumentsBuilder.fn_Config("placeChain.military.kind") Then
        If VBA.Len(row("Destination")) = 0 Then private_Required "text.VacationTicket.DocumentTargetLabel"
        ' Preserve the place kind in the existing structured history fields.
        row("PlaceType") = ex_TravelDocumentsBuilder.fn_Config("placeChain.military.type")
        row("PlaceName") = row("Destination")
        If VBA.StrComp(VBA.Left$(row("Destination"), VBA.Len(row("PlaceType"))), row("PlaceType"), VBA.vbTextCompare) <> 0 Then
            row("Destination") = row("PlaceType") & VBA.Space$(1) & row("Destination")
        End If
        Set fn_FromValues = row
        Exit Function
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".requiredFields"), "|")
        If VBA.Len(row(field)) = 0 Then
            If VBA.CBool(values("AutoFillPersonnel")) Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.Stale")
            private_Required "text." & field
        End If
    Next field
    If Not isAbroad Then
        If VBA.InStr(1, "|" & ex_TravelDocumentsBuilder.fn_Config("lookup.Places.types") & "|", _
                "|" & row("PlaceType") & "|", VBA.vbBinaryCompare) = 0 Then
            caption = ex_TravelDocumentsBuilder.fn_Config("text.PlaceType")
            VBA.Err.Raise 5, , VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("generation.message.Invalid"), "{value}", caption)
        End If
    End If
    address = ex_TDB_Places.fn_RowAddress(row)
    If VBA.CBool(values("AutoFillPersonnel")) Then
        If VBA.StrComp(address, row("Destination"), VBA.vbBinaryCompare) <> 0 Then
            VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.Stale")
        End If
    ElseIf isAbroad Then
        row("Destination") = address
    End If
    If VBA.Len(row("Destination")) = 0 Then private_Required "text.VacationTicket.DocumentTargetLabel"
    Set fn_FromValues = row
End Function

Public Function fn_Build( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal values As Object _
) As Collection
    Dim source As Object
    Dim scalar As Variant
    Dim isObject As Boolean
    Dim result As New Collection
    Dim item As Object
    Dim copy As Object
    Dim field As Variant
    Dim hasDraft As Boolean

    If Not bindingContext.TryGetValue("Data", "PlaceChain", scalar, source, isObject) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not isObject Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not TypeOf source Is Collection Then ex_TDB_History.fn_Fail "InvalidChain"
    For Each item In source
        Set copy = ex_TDB_History.fn_NewRow()
        For Each field In VBA.Split("RowId|Destination|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
            copy(field) = item(field)
        Next field
        For Each field In VBA.Split("Destination|" & ex_TravelDocumentsBuilder.fn_Config( _
                ex_TDB_Places.fn_Prefix(ex_TDB_Places.fn_IsAbroad(VBA.CStr(copy("PlaceAbroad")))) & ".requiredFields"), "|")
            If fn_Kind(copy) <> ex_TravelDocumentsBuilder.fn_Config("placeChain.military.kind") Or field = "Destination" Then
                If VBA.Len(VBA.Trim$(VBA.CStr(copy(field)))) = 0 Then ex_TDB_History.fn_Fail "InvalidChain"
            End If
        Next field
        result.Add copy
    Next item
    For Each field In VBA.Split("DocumentTarget|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
        If field <> "PlaceAbroad" Then
            If VBA.Len(VBA.Trim$(VBA.CStr(values(field)))) > 0 Then hasDraft = True
        End If
    Next field
    If hasDraft Then result.Add fn_FromValues(values)
    If result.Count = 0 Then private_Required "text.VacationTicket.DocumentTargetLabel"
    If result.Count > VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("placeChain.maxItems")) Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("placeChain.text.Limit")
    End If
    Set fn_Build = result
End Function

Public Function fn_Join(ByVal chain As Collection) As String
    Dim row As Object

    For Each row In chain
        If VBA.Len(fn_Join) > 0 Then fn_Join = fn_Join & ex_TravelDocumentsBuilder.fn_Config("placeChain.separator") & VBA.Space$(1)
        fn_Join = fn_Join & VBA.CStr(row("Destination"))
    Next row
End Function

Public Function fn_Kind(ByVal row As Object) As String
    If row("PlaceType") = ex_TravelDocumentsBuilder.fn_Config("placeChain.military.type") Then
        fn_Kind = ex_TravelDocumentsBuilder.fn_Config("placeChain.military.kind")
    Else
        fn_Kind = ex_TravelDocumentsBuilder.fn_Config("generation.defaults.PlaceKind")
    End If
End Function

Public Function fn_FromGroup(ByVal group As Collection) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim root As Object
    Dim field As Variant
    Dim parentId As String
    Dim found As Boolean
    Dim count As Long

    parentId = group(1)("RowId")
    For Each row In group
        If row("RowType") = "Place" Then count = count + 1
    Next row
    Do
        found = False
        For Each row In group
            If row("RowType") = "Place" And row("ParentRowId") = parentId Then
                result.Add row
                parentId = row("RowId")
                found = True
                If result.Count > count Then ex_TDB_History.fn_Fail "InvalidChain"
                Exit For
            End If
        Next row
    Loop While found
    If result.Count <> count Then ex_TDB_History.fn_Fail "InvalidChain"
    Set root = group(1)
    If result.Count = 0 And VBA.Len(VBA.Trim$(VBA.CStr(root("PlaceName")))) > 0 Then
        ' A single place is stored on Event; its in-memory item has no persisted child ID.
        Set row = ex_TDB_History.fn_NewRow()
        For Each field In VBA.Split("Destination|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
            row(field) = root(field)
        Next field
        row("RowType") = "Place"
        result.Add row
    End If
    Set fn_FromGroup = result
End Function

Public Sub fn_ClearDraft(ByVal bindingContext As obj_UiBindingContext)
    Dim updates As Object
    Dim field As Variant
    Dim diagnostic As String

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split("DocumentTarget|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
        updates("Form." & field) = VBA.vbNullString
    Next field
    If Not bindingContext.TryApplyValues(updates, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If Not bindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then ex_TDB_History.fn_Fail "InvalidRows"
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_Required(ByVal captionKey As String)
    VBA.Err.Raise 5, , VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("generation.message.Required"), _
        "{value}", ex_TravelDocumentsBuilder.fn_Config(captionKey))
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------