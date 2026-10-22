Attribute VB_Name = "ex_PEB_MovementConflicts"
Option Explicit

#Const TVO_DEPRECATED = True

Private m_target As String
Private m_states As Object
Private m_choices As Collection
Private m_items As Collection
Private m_overwrite As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    m_target = VBA.vbNullString
    Set m_states = Nothing
    Set m_choices = Nothing
    Set m_items = Nothing
    m_overwrite = False
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Begin( _
    ByVal target As String, _
    ByVal choices As Collection _
)
    Dim book As Workbook
    Dim table As ListObject
    Dim data As Variant
    Dim index As Long
    Dim failure As String
    Dim policy As String

    On Error GoTo EH
    fn_Module_Dispose
    policy = ex_PEB_Export.fn_Config("export.Movement.manualChanges.policy")
    If policy <> "Overwrite" And policy <> "Resolve" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ConflictPolicyInvalid")
    End If
    m_overwrite = (policy = "Overwrite")
    m_target = target
    Set m_choices = choices
    Set m_items = New Collection
    Set m_states = VBA.CreateObject("Scripting.Dictionary")
    If m_overwrite Then
        Exit Sub
    End If
    Set book = private_Book()
    Set table = fn_Table(book, "history.fieldState")
    If Not table.DataBodyRange Is Nothing Then
        data = table.DataBodyRange.Value2
        For index = 1 To UBound(data, 1)
            If data(index, 2) = "Movement" And data(index, 3) = target Then
                m_states(private_Key(data(index, 1), data(index, 4), data(index, 5))) = _
                    VBA.Array(data(index, 6), data(index, 7), data(index, 8))
            End If
        Next index
    End If
    Set table = fn_Table(book, "history.conflicts")
    Exit Sub
EH:
    failure = VBA.Err.Description
    fn_Module_Dispose
    Err.Raise 5, , failure
End Sub

Public Function fn_Items() As Collection
    If m_items Is Nothing Then
        Set m_items = New Collection
    End If
    Set fn_Items = m_items
End Function

Public Function fn_Overwrite() As Boolean
    If m_states Is Nothing Then
        Exit Function
    End If
    fn_Overwrite = m_overwrite
End Function

Public Function fn_Resolve( _
    ByVal action As Object, _
    ByVal side As String, _
    ByVal fields As Object, _
    ByVal cells As Object, _
    ByVal basisHeader As String, _
    ByVal hashHeader As String _
) As Boolean
    Dim field As Variant
    Dim state As Variant
    Dim item As Object
    Dim choice As Object
    Dim key As String
    Dim actual As String
    Dim generated As String
    Dim manual As Boolean
    Dim resolved As Boolean
    Dim validHash As Boolean
    Dim previousFields As Object
    Dim completeState As Boolean
#If TVO_DEPRECATED Then
    Dim legacyFields As Object
    Dim fieldKey As String
#End If

    If m_states Is Nothing Then
        Exit Function
    End If
    ' Overwrite mode keeps generated fields intact and ignores all previous manual choices.
    If m_overwrite Then
        fn_Resolve = True
        Exit Function
    End If
    validHash = (cells(hashHeader) = ex_PEB_MovementLinks.fn_FieldsHash(fields, cells))
#If TVO_DEPRECATED Then
    ' Older exports did not own TVO columns; migrate only empty, untracked cells.
    If Not validHash And side = "Opening" Then
        Set legacyFields = VBA.CreateObject("Scripting.Dictionary")
        For Each field In fields.Keys
            fieldKey = private_Key(action("RowId"), side, field)
            If Not ex_PEB_MovementRegistry.fn_IsTvoHeader(VBA.CStr(field)) Or m_states.Exists(fieldKey) Or _
                    VBA.Len(ex_PEB_Export.fn_Text(cells(field))) > 0 Then
                legacyFields.Add field, fields(field)
            End If
        Next field
        validHash = (cells(hashHeader) = ex_PEB_MovementLinks.fn_FieldsHash(legacyFields, cells))
    End If
#End If
    Set previousFields = VBA.CreateObject("Scripting.Dictionary")
    completeState = True
    For Each field In fields.Keys
        key = private_Key(action("RowId"), side, field)
        If m_states.Exists(key) Then
            state = m_states(key)
            previousFields.Add field, private_Decode(VBA.CStr(state(2)))
        Else
            completeState = False
        End If
    Next field
    If completeState Then
        If cells(hashHeader) <> ex_PEB_MovementLinks.fn_FieldsHash(previousFields) Then
            ex_PEB_MovementLinks.fn_Fail "AppliedMovementChanged", action("RowId") & "/" & hashHeader
        End If
    End If
    fn_Resolve = True
    For Each field In fields.Keys
        key = private_Key(action("RowId"), side, field)
        actual = ex_PEB_MovementLinks.fn_ValueToken(cells(field))
        generated = ex_PEB_MovementLinks.fn_ValueToken(fields(field))
        manual = False
        If m_states.Exists(key) Then
            state = m_states(key)
            manual = (actual <> state(2))
            If Not manual And state(1) = generated Then
                If state(1) <> state(2) Then
                    If field <> basisHeader Then
                        ex_PEB_MovementLinks.fn_Fail "AppliedEventChanged", action("RowId") & "/" & field
                    End If
                    fields(field) = cells(field)
                End If
            End If
        ElseIf Not validHash Then
            ' A legacy hash cannot identify the changed field. Only the basis can be explicitly accepted.
            manual = (actual <> generated Or field = basisHeader)
        End If
        If manual Then
            If field <> basisHeader Then
                ex_PEB_MovementLinks.fn_Fail "AppliedMovementChanged", action("RowId") & "/" & field
            End If
            Set item = VBA.CreateObject("Scripting.Dictionary")
            item.Add "EventRowId", action("RowId")
            item.Add "Revision", action("Revision")
            item.Add "HistoryStamp", action("HistoryStamp")
            item.Add "Target", m_target
            item.Add "Side", side
            item.Add "Field", field
            item.Add "GeneratedValue", fields(field)
            item.Add "MovementValue", cells(field)
            item.Add "Choice", VBA.vbNullString
            item.Add "LinkHash", cells(hashHeader)
            resolved = False
            If Not m_choices Is Nothing Then
                For Each choice In m_choices
                    If private_Same(item, choice) Then
                        If choice("Choice") = "Movement" Or choice("Choice") = "EventsHistory" Then
                            item("Choice") = choice("Choice")
                            resolved = True
                            If choice("Choice") = "Movement" Then
                                fields(field) = cells(field)
                            End If
                            Exit For
                        End If
                    End If
                Next choice
            End If
            m_items.Add item
            If Not resolved Then
                fn_Resolve = False
            End If
        End If
    Next field
    fn_Resolve = True
End Function

Public Function fn_Pending(ByVal identifier As String) As Boolean
    Dim item As Object

    If m_items Is Nothing Then
        Exit Function
    End If
    For Each item In m_items
        If item("EventRowId") = identifier And VBA.Len(item("Choice")) = 0 Then
            fn_Pending = True
            Exit Function
        End If
    Next item
End Function

Public Function fn_Table( _
    ByVal book As Workbook, _
    ByVal prefix As String _
) As ListObject
    Dim table As ListObject
    Dim headers As Variant
    Dim index As Long

    On Error GoTo EH
    Set table = book.Worksheets(ex_PEB_Export.fn_Config(prefix & ".sheet")).ListObjects( _
        ex_PEB_Export.fn_Config(prefix & ".table"))
    headers = VBA.Split(ex_PEB_Export.fn_Config(prefix & ".fields"), "|")
    If table.ListColumns.Count <> UBound(headers) + 1 Then
        Err.Raise 5
    End If
    For index = 0 To UBound(headers)
        If table.ListColumns(index + 1).Name <> headers(index) Then
            Err.Raise 5
        End If
    Next index
    If book.ReadOnly Or table.Parent.ProtectContents Then
        Err.Raise 5
    End If
    Set fn_Table = table
    Exit Function
EH:
    Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.ConflictSchema"), "{table}", _
        ex_PEB_Export.fn_Config(prefix & ".sheet") & "/" & ex_PEB_Export.fn_Config(prefix & ".table"))
End Function

Public Sub fn_Record( _
    ByVal book As Workbook, _
    ByVal actions As Collection, _
    ByVal target As String _
)
    Dim action As Object
    Dim fields As Object
    Dim original As Object
    Dim table As ListObject
    Dim row As ListRow
    Dim item As Object
    Dim side As Variant
    Dim field As Variant
    Dim name As String
    Dim data(1 To 1, 1 To 8) As Variant
    Dim decision(1 To 1, 1 To 12) As Variant
    Dim key As String
    Dim previous As Variant
    Dim unchanged As Boolean
    Dim pending As New Collection
    Dim decisions As New Collection
    Dim startedAt As Double

    startedAt = VBA.Timer
    Set table = fn_Table(book, "history.fieldState")
    If m_overwrite Then
        private_LoadStates table
    End If
    For Each action In actions
        If action("EventStatus") <> "Cancelled" Then
            For Each side In VBA.Array("Opening", "Closing")
                name = "Writes"
                If side = "Closing" Then
                    name = "ClosingWrites"
                End If
                If (side = "Opening" And action("Operation") <> "close") Or _
                        (side = "Closing" And action("Operation") <> "append") Then
                    Set fields = action(name)
                    Set original = fields
                    If action.Exists("Generated" & name) Then
                        Set original = action("Generated" & name)
                    End If
                    For Each field In fields.Keys
                        data(1, 1) = action("RowId")
                        data(1, 2) = "Movement"
                        data(1, 3) = target
                        data(1, 4) = side
                        data(1, 5) = field
                        data(1, 6) = action("Revision")
                        data(1, 7) = ex_PEB_MovementLinks.fn_ValueToken(original(field))
                        data(1, 8) = ex_PEB_MovementLinks.fn_ValueToken(fields(field))
                        key = private_Key(action("RowId"), side, field)
                        unchanged = False
                        If Not m_states Is Nothing Then
                            If m_states.Exists(key) Then
                                previous = m_states(key)
                                unchanged = (VBA.CStr(previous(0)) = VBA.CStr(data(1, 6)) And _
                                    previous(1) = data(1, 7) And previous(2) = data(1, 8))
                            End If
                        End If
                        If Not unchanged Then
                            pending.Add data
                        End If
                    Next field
                End If
            Next side
        End If
    Next action
    private_WriteRows table, pending
    Set table = fn_Table(book, "history.conflicts")
    If Not m_items Is Nothing Then
        For Each item In m_items
            If VBA.Len(item("Choice")) > 0 Then
                decision(1, 1) = item("EventRowId") & "/" & item("Revision") & "/" & VBA.CStr(table.ListRows.Count + decisions.Count + 1)
                decision(1, 2) = item("EventRowId")
                decision(1, 3) = item("Revision")
                decision(1, 4) = "Movement"
                decision(1, 5) = target
                decision(1, 6) = item("Side")
                decision(1, 7) = item("Field")
                decision(1, 8) = item("GeneratedValue")
                decision(1, 9) = item("MovementValue")
                decision(1, 10) = item("Choice")
                decision(1, 11) = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
                decision(1, 12) = Application.UserName & " / " & VBA.Environ$("USERNAME")
                decisions.Add decision
            End If
        Next item
    End If
    private_WriteRows table, decisions
    ex_Core.fn_Diagnostic_WritePerf "Movement.Journal.FieldStatesAndDecisions", startedAt
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Decode(ByVal token As String) As Variant
    Select Case VBA.Left$(token, 2)
        Case "s:"
            private_Decode = VBA.Mid$(token, 3)
        Case "n:"
            private_Decode = VBA.Val(VBA.Mid$(token, 3))
        Case Else
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.ConflictStateInvalid")
    End Select
End Function

Private Function private_Key( _
    ByVal identifier As Variant, _
    ByVal side As Variant, _
    ByVal field As Variant _
) As String
    private_Key = ex_PEB_MovementLinks.fn_Token(identifier) & ex_PEB_MovementLinks.fn_Token(side) & _
        ex_PEB_MovementLinks.fn_Token(field)
End Function

Private Function private_Same( _
    ByVal item As Object, _
    ByVal other As Object _
) As Boolean
    Dim key As Variant

    For Each key In VBA.Array("EventRowId", "Revision", "HistoryStamp", "Target", "Side", "Field", _
            "GeneratedValue", "MovementValue", "LinkHash")
        If ex_PEB_MovementLinks.fn_ValueToken(item(key)) <> ex_PEB_MovementLinks.fn_ValueToken(other(key)) Then
            Exit Function
        End If
    Next key
    private_Same = True
End Function

Private Function private_Book() As Workbook
    Dim candidate As Workbook
    Dim path As String

    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path"))
    For Each candidate In Application.Workbooks
        If VBA.StrComp(candidate.FullName, path, VBA.vbTextCompare) = 0 Then
            Set private_Book = candidate
            Exit Function
        End If
    Next candidate
    Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryOpenRequired"), "{value}", path)
End Function
' --------------------------------------

Private Sub private_WriteRows( _
    ByVal table As ListObject, _
    ByVal rows As Collection _
)
    Dim data() As Variant
    Dim item As Variant
    Dim area As Range
    Dim first As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim extension As Range
    Dim other As ListObject
    Dim canResize As Boolean

    If rows.Count = 0 Then
        Exit Sub
    End If
    first = table.ListRows.Count + 1
    ReDim data(1 To rows.Count, 1 To table.ListColumns.Count)
    For rowIndex = 1 To rows.Count
        item = rows(rowIndex)
        For columnIndex = 1 To table.ListColumns.Count
            data(rowIndex, columnIndex) = item(1, columnIndex)
        Next columnIndex
    Next rowIndex
    ' Resize once when the extension is empty; retain Excel insertion semantics otherwise.
    canResize = Not table.ShowTotals
    If table.Range.Row + table.Range.Rows.Count + rows.Count - 1 > table.Parent.Rows.Count Then
        canResize = False
    End If
    If canResize Then
        Set extension = table.Range.Rows(table.Range.Rows.Count).Offset(1).Resize(rows.Count)
        canResize = (Application.CountA(extension) = 0 And Not VBA.IsNull(extension.MergeCells))
        If canResize Then
            canResize = Not extension.MergeCells
        End If
        For Each other In table.Parent.ListObjects
            If Not other Is table Then
                If Not Application.Intersect(extension, other.Range) Is Nothing Then
                    canResize = False
                End If
            End If
        Next other
    End If
    If canResize Then
        table.Resize table.Range.Resize(table.Range.Rows.Count + rows.Count)
    Else
        For rowIndex = 1 To rows.Count
            table.ListRows.Add
        Next rowIndex
    End If
    Set area = table.DataBodyRange.Rows(first).Resize(rows.Count)
    area.NumberFormat = "@"
    area.Value2 = data
End Sub

Private Sub private_LoadStates(ByVal table As ListObject)
    Dim data As Variant
    Dim index As Long

    If table.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    data = table.DataBodyRange.Value2
    For index = 1 To UBound(data, 1)
        If data(index, 2) = "Movement" And data(index, 3) = m_target Then
            m_states(private_Key(data(index, 1), data(index, 4), data(index, 5))) = _
                VBA.Array(data(index, 6), data(index, 7), data(index, 8))
        End If
    Next index
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------