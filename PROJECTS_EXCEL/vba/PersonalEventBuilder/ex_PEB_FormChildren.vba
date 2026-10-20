Attribute VB_Name = "ex_PEB_FormChildren"
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
Public Sub fn_Attach(ByVal rows As Collection)
    Dim ids As Object
    Dim row As Object
    Dim root As Object
    Dim seen As Object

    Set ids = CreateObject("Scripting.Dictionary")
    For Each row In rows
        ids.Add row("RowId"), row
        If row("RowType") = "Event" Then Set row("FormExtensionRows") = New Collection
    Next row
    For Each row In rows
        If row("RowType") <> "Event" Then
            Set root = row
            Set seen = CreateObject("Scripting.Dictionary")
            Do While root("RowType") <> "Event"
                If seen.Exists(root("RowId")) Then Err.Raise 5, , "Cyclic child row."
                seen.Add root("RowId"), True
                If Not ids.Exists(root("ParentRowId")) Then Err.Raise 5, , "Missing child parent."
                Set root = ids(root("ParentRowId"))
            Loop
            root("FormExtensionRows").Add row
        End If
    Next row
End Sub

Public Function fn_Plan( _
    ByVal table As ListObject, _
    ByVal rows As Collection, _
    ByVal root As Object, _
    ByVal children As Collection, _
    Optional ByVal rebindFormExtensions As Boolean = False _
) As Object
    Dim plan As Object
    Dim ids As Object
    Dim removed As Object
    Dim pending As New Collection
    Dim oldChildren As New Collection
    Dim newChildren As New Collection
    Dim row As Object
    Dim child As Object
    Dim parent As Object
    Dim fields As Object
    Dim key As Variant
    Dim field As Variant
    Dim mapping As Variant
    Dim index As Long
    Dim identifier As String
    Dim serial As Long
    Dim changed As Boolean
    Dim before As Variant
    Dim formats() As Variant
    Dim r As Long
    Dim c As Long
    Dim people As Object
    Dim uniformFormat As Variant
    Dim rebound As Object

    Set plan = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")
    Set removed = CreateObject("Scripting.Dictionary")
    For Each row In rows
        ids.Add row("RowId"), row
    Next row
    For Each row In rows
        Set parent = row
        Do While parent("RowType") <> "Event"
            If Not ids.Exists(parent("ParentRowId")) Then Err.Raise 5, , "Invalid child parent: " & parent("ParentRowId")
            Set parent = ids(parent("ParentRowId"))
        Loop
        If row("RowType") <> "Event" And parent("RowId") = root("RowId") Then
            removed.Add row("RowId"), row
            row("__data") = table.ListRows(row("__tableIndex")).Range.Value2
            oldChildren.Add row
        Else
            pending.Add row
        End If
    Next row
    If Not ids.Exists(root("RowId")) Then pending.Add root
    Set rebound = CreateObject("Scripting.Dictionary")
    If rebindFormExtensions Then Set rebound = private_Rebind(children, root, ids, removed)
    For Each child In children
        serial = serial + 1
        identifier = child("RowId")
        If Len(identifier) = 0 Then
            index = serial
            Do
                identifier = root("RowId") & "/C" & CStr(index)
                index = index + 1
            Loop While ids.Exists(identifier) Or rebound.Exists(identifier)
        ElseIf Not removed.Exists(identifier) And Not rebound.Exists(identifier) Then
            Err.Raise 5, , "Child does not belong to this event: " & identifier
        End If
        If Not removed.Exists(identifier) Then ids.Add identifier, child
        Set fields = CreateObject("Scripting.Dictionary")
        For Each field In Split(ex_PEB_Export.fn_Config("history.fields"), "|")
            mapping = Split(field, "=")
            fields(mapping(0)) = vbNullString
        Next field
        For Each key In child.Keys
            If fields.Exists(key) Then fields(key) = child(key)
        Next key
        fields("RowId") = identifier
        If removed.Exists(identifier) Then fields("__data") = removed(identifier)("__data")
        child("RowId") = identifier
        fields("RowOrder") = CStr(serial)
        fields("ParentRowId") = child("ParentRowId")
        If Len(fields("ParentRowId")) = 0 Then fields("ParentRowId") = root("RowId")
        child("ParentRowId") = fields("ParentRowId")
        fields("AcceptedOrderNumber") = root("AcceptedOrderNumber")
        fields("AcceptedOrderDate") = root("AcceptedOrderDate")
        fields("EventId") = root("EventId")
        fields("EventDate") = root("EventDate")
        If fields("RowType") <> "Document" And fields("RowType") <> "Report" And fields("RowType") <> "TVO" Then Err.Raise 5, , "Invalid extension row type."
        newChildren.Add fields
        pending.Add fields
    Next child
    If oldChildren.Count <> newChildren.Count Then
        changed = True
    Else
        For r = 1 To newChildren.Count
            For Each field In Split(ex_PEB_Export.fn_Config("history.fields"), "|")
                mapping = Split(field, "=")
                If CStr(oldChildren(r)(mapping(0))) <> CStr(newChildren(r)(mapping(0))) Then changed = True
            Next field
        Next r
    End If
    ' Validate parent ownership and cycles before preparing any revision or writing rows.
    Set ids = CreateObject("Scripting.Dictionary")
    Set people = CreateObject("Scripting.Dictionary")
    people.Add root("PersonId"), True
    For Each row In pending
        If ids.Exists(row("RowId")) Then Err.Raise 5, , "Duplicate extension row id."
        ids.Add row("RowId"), row
    Next row
    For Each row In newChildren
        If row("RowType") = "TVO" Then
            If people.Exists(row("PersonId")) Then Err.Raise 5, , ex_PEB_Export.fn_Config("extensions.duplicateTvo")
            people.Add row("PersonId"), True
        End If
        Set parent = row
        Set fields = CreateObject("Scripting.Dictionary")
        Do While parent("RowType") <> "Event"
            If fields.Exists(parent("RowId")) Then Err.Raise 5, , "Cyclic extension parent."
            fields.Add parent("RowId"), True
            If Not ids.Exists(parent("ParentRowId")) Then Err.Raise 5, , "Missing extension parent."
            Set parent = ids(parent("ParentRowId"))
            If parent("RowType") <> "Event" Then
                If row("RowType") <> "TVO" Or parent("RowType") <> "TVO" Then Err.Raise 5, , "Only TVO chains may be nested."
            End If
        Loop
        If parent("RowId") <> root("RowId") Then Err.Raise 5, , "Extension belongs to another event."
    Next row
    If Not changed Then Exit Function
    plan.Add "Rows", pending
    plan.Add "Old", oldChildren
    plan.Add "New", newChildren
    plan.Add "Changed", changed
    plan.Add "Count", table.ListRows.Count
    If table.ListRows.Count > 0 Then
        before = table.DataBodyRange.Value2
        plan.Add "Before", before
        uniformFormat = table.DataBodyRange.NumberFormat
        If Not IsNull(uniformFormat) Then
            plan.Add "UniformFormat", uniformFormat
        Else
            ReDim formats(1 To table.ListRows.Count, 1 To table.ListColumns.Count)
            For r = 1 To table.ListRows.Count
                For c = 1 To table.ListColumns.Count
                    formats(r, c) = table.DataBodyRange.Cells(r, c).NumberFormat
                Next c
            Next r
            plan.Add "Formats", formats
        End If
    End If
    Set fn_Plan = plan
End Function

Public Sub fn_Write( _
    ByVal table As ListObject, _
    ByVal plan As Object _
)
    Dim removed As Object
    Dim row As Object
    Dim index As Long
    Dim data As Variant
    Dim field As Variant
    Dim mapping As Variant
    Dim target As ListRow

    If plan Is Nothing Then Exit Sub
    Set removed = CreateObject("Scripting.Dictionary")
    For Each row In plan("Old")
        removed.Add row("RowId"), True
    Next row
    For index = table.ListRows.Count To 1 Step -1
        If removed.Exists(CStr(table.ListRows(index).Range.Cells(1, table.ListColumns("RowId").Index).Value2)) Then table.ListRows(index).Delete
    Next index
    For Each row In plan("New")
        data = fn_Data(table, row)
        Set target = table.ListRows.Add
        target.Range.NumberFormat = "@"
        target.Range.Value2 = data
    Next row
End Sub

Public Function fn_Data( _
    ByVal table As ListObject, _
    ByVal row As Object _
) As Variant
    Dim data As Variant
    Dim field As Variant
    Dim mapping As Variant

    If row.Exists("__data") Then
        data = row("__data")
    Else
        ReDim data(1 To 1, 1 To table.ListColumns.Count)
    End If
    For Each field In Split(ex_PEB_Export.fn_Config("history.fields"), "|")
        mapping = Split(field, "=")
        data(1, table.ListColumns(mapping(1)).Index) = row(mapping(0))
    Next field
    fn_Data = data
End Function

Public Sub fn_Rollback( _
    ByVal table As ListObject, _
    ByVal plan As Object _
)
    Dim r As Long
    Dim c As Long
    Dim formats As Variant

    If plan Is Nothing Then Exit Sub
    Do While table.ListRows.Count > plan("Count")
        table.ListRows(table.ListRows.Count).Delete
    Loop
    Do While table.ListRows.Count < plan("Count")
        table.ListRows.Add
    Loop
    If plan("Count") > 0 Then
        table.DataBodyRange.NumberFormat = "@"
        table.DataBodyRange.Value2 = plan("Before")
        If plan.Exists("UniformFormat") Then
            table.DataBodyRange.NumberFormat = plan("UniformFormat")
        Else
            formats = plan("Formats")
            For r = 1 To table.ListRows.Count
                For c = 1 To table.ListColumns.Count
                    table.DataBodyRange.Cells(r, c).NumberFormat = formats(r, c)
                Next c
            Next r
        End If
    End If
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Rebind( _
    ByVal children As Collection, _
    ByVal root As Object, _
    ByVal ids As Object, _
    ByVal owned As Object _
) As Object
    Dim mapping As Object
    Dim allocated As Object
    Dim parents As Object
    Dim child As Object
    Dim identifier As String
    Dim replacement As String
    Dim parentId As String
    Dim index As Long

    Set mapping = CreateObject("Scripting.Dictionary")
    Set allocated = CreateObject("Scripting.Dictionary")
    Set parents = CreateObject("Scripting.Dictionary")

    ' Allocate copies without changing rows belonging to the source event.
    For Each child In children
        identifier = CStr(child("RowId"))
        If Len(identifier) > 0 Then
            replacement = identifier
            If Not owned.Exists(identifier) Then
                If Not ids.Exists(identifier) Then Err.Raise 5, , "Missing source extension: " & identifier
                If ids(identifier)("RowType") = "Event" Then Err.Raise 5, , "An event cannot be copied as an extension."
                Do
                    index = index + 1
                    replacement = root("RowId") & "/C" & CStr(index)
                Loop While ids.Exists(replacement) Or allocated.Exists(replacement)
                allocated.Add replacement, True
            End If
            mapping.Add identifier, replacement
        End If
    Next child

    ' Preserve nested TVO links; replace only the event-level parent.
    For Each child In children
        parentId = CStr(child("ParentRowId"))
        If Len(parentId) = 0 Or parentId = root("RowId") Then
            parents.Add CStr(parents.Count + 1), root("RowId")
        ElseIf mapping.Exists(parentId) Then
            parents.Add CStr(parents.Count + 1), mapping(parentId)
        Else
            If Not ids.Exists(parentId) Then Err.Raise 5, , "Missing source extension parent: " & parentId
            If ids(parentId)("RowType") <> "Event" Then Err.Raise 5, , "Source extension parent is outside the copied form."
            parents.Add CStr(parents.Count + 1), root("RowId")
        End If
    Next child
    index = 0
    For Each child In children
        index = index + 1
        identifier = CStr(child("RowId"))
        If Len(identifier) > 0 Then child("RowId") = mapping(identifier)
        child("ParentRowId") = parents(CStr(index))
    Next child
    Set private_Rebind = allocated
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------