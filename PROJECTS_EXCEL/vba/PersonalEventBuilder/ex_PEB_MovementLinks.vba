Attribute VB_Name = "ex_PEB_MovementLinks"
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
Public Function fn_Headers() As Object
    Dim result As Object
    Dim names As Object
    Dim key As Variant
    Dim field As Variant
    Dim name As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    Set names = VBA.CreateObject("Scripting.Dictionary")
    names.CompareMode = VBA.vbTextCompare
    For Each key In VBA.Array("OpeningId", "OpeningHash", "ClosingId", "ClosingHash")
        name = ex_PEB_Export.fn_Config("export.Movement.links." & key)
        If names.Exists(name) Then
            fn_Fail "DuplicateLinkHeader", name
        End If
        If VBA.Len(name) > 255 Then
            fn_Fail "InvalidLinkHeader", name
        End If
        For Each field In VBA.Array("PersonId", "PersonName", "EventName", "ArrivalDate", "DepartureDate", _
                "Destination", "ArrivalBasis", "ArrivalOrder", "DepartureBasis", "DepartureOrder", _
                "TicketDocument", "ArrivalFoodSupport", "DepartureFoodSupport")
            If VBA.StrComp(name, ex_PEB_Export.fn_Config("export.Movement.headers." & field), VBA.vbTextCompare) = 0 Then
                fn_Fail "LinkBusinessHeaderCollision", name
            End If
        Next field
        names.Add name, True
        result.Add VBA.CStr(key), name
    Next key
    Set fn_Headers = result
End Function

Public Function fn_EventStamp(ByVal stored As Object) As String
    Dim document As Object
    Dim field As Variant
    Dim stamp As String

    stamp = "2|" & ex_PEB_History.fn_Stamp(stored)
    stamp = stamp & fn_Token(stored("MetaDocumentTables").Count)
    For Each document In stored("MetaDocumentTables")
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs.MetaDocumentTables.fields"), "|")
            stamp = stamp & fn_Token(ex_PEB_Export.fn_Required(document, VBA.CStr(field)))
        Next field
    Next document
    fn_EventStamp = fn_Hash(stamp)
End Function

Public Function fn_Hash(ByVal text As String) As String
    Dim first As Double
    Dim second As Double
    Dim code As Long
    Dim index As Long

    first = 1
    second = 7
    For index = 1 To VBA.Len(text)
        code = VBA.AscW(VBA.Mid$(text, index, 1)) And &HFFFF&
        first = first * 257# + code
        first = first - VBA.Int(first / 2147483647#) * 2147483647#
        second = second * 263# + code
        second = second - VBA.Int(second / 2147483629#) * 2147483629#
    Next index
    fn_Hash = "H1-" & VBA.Right$("00000000" & VBA.Hex$(VBA.CLng(first)), 8) & _
        VBA.Right$("00000000" & VBA.Hex$(VBA.CLng(second)), 8)
End Function

Public Function fn_FieldsHash( _
    ByVal fields As Object, _
    Optional ByVal cells As Object = Nothing _
) As String
    Dim names As Variant
    Dim key As Variant
    Dim temporary As Variant
    Dim value As Variant
    Dim index As Long
    Dim other As Long
    Dim stamp As String

    names = fields.Keys
    For index = 0 To fields.Count - 2
        For other = index + 1 To fields.Count - 1
            If VBA.StrComp(names(index), names(other), VBA.vbBinaryCompare) > 0 Then
                temporary = names(index)
                names(index) = names(other)
                names(other) = temporary
            End If
        Next other
    Next index
    For Each key In names
        If cells Is Nothing Then
            value = fields(key)
        Else
            value = ex_PEB_Export.fn_Required(cells, VBA.CStr(key))
        End If
        stamp = stamp & fn_Token(VBA.CStr(key)) & fn_Token(fn_ValueToken(value))
    Next key
    fn_FieldsHash = fn_Hash(stamp)
End Function

Public Function fn_ValueToken(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsObject(value) Or VBA.IsArray(value) Then
        fn_Fail "InvalidStampValue", ""
    End If
    Select Case VBA.VarType(value)
        Case VBA.vbEmpty
            fn_ValueToken = "s:"
        Case VBA.vbDate, VBA.vbByte, VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            fn_ValueToken = "n:" & VBA.Trim$(VBA.Str$(VBA.CDbl(value)))
        Case Else
            fn_ValueToken = "s:" & VBA.CStr(value)
    End Select
End Function

Public Function fn_StructureStamp(ByVal target As ListObject) As String
    Dim column As ListColumn
    Dim stamp As String

    stamp = fn_Token(target.Parent.Parent.FullName) & fn_Token(target.Parent.Name) & fn_Token(target.Name)
    stamp = stamp & fn_Token(target.Range.Address) & fn_Token(target.ListRows.Count) & fn_Token(target.ShowTotals)
    For Each column In target.ListColumns
        stamp = stamp & fn_Token(column.Name)
    Next column
    fn_StructureStamp = fn_Hash(stamp)
End Function

Public Function fn_ReadRange(ByVal area As Range) As Object
    Dim snapshot As Object
    Dim formulaCells As Range
    Dim block As Range
    Dim masks() As Boolean
    Dim formats() As Variant
    Dim values As Variant
    Dim formulas As Variant
    Dim singleValue(1 To 1, 1 To 1) As Variant
    Dim hasFormula As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim rowCount As Long
    Dim columnCount As Long
    Dim blockRows As Long
    Dim blockColumns As Long

    rowCount = area.Rows.Count
    columnCount = area.Columns.Count
    values = area.Value2
    formulas = area.Formula
    If rowCount = 1 And columnCount = 1 Then
        singleValue(1, 1) = values
        values = singleValue
        singleValue(1, 1) = formulas
        formulas = singleValue
    End If
    ReDim masks(1 To rowCount, 1 To columnCount)
    ReDim formats(1 To rowCount, 1 To columnCount)
    hasFormula = area.HasFormula
    If VBA.IsNull(hasFormula) Then
        Set formulaCells = Application.Intersect(area, area.SpecialCells(xlCellTypeFormulas))
        For Each block In formulaCells.Areas
            firstRow = block.Row - area.Row + 1
            firstColumn = block.Column - area.Column + 1
            blockRows = block.Rows.Count
            blockColumns = block.Columns.Count
            For rowIndex = firstRow To firstRow + blockRows - 1
                For columnIndex = firstColumn To firstColumn + blockColumns - 1
                    masks(rowIndex, columnIndex) = True
                Next columnIndex
            Next rowIndex
        Next block
    ElseIf VBA.CBool(hasFormula) Then
        For rowIndex = 1 To rowCount
            For columnIndex = 1 To columnCount
                masks(rowIndex, columnIndex) = True
            Next columnIndex
        Next rowIndex
    End If
    For columnIndex = 1 To columnCount
        private_ReadFormats area.Columns(columnIndex), formats, columnIndex, 1
    Next columnIndex
    Set snapshot = VBA.CreateObject("Scripting.Dictionary")
    snapshot.Add "Address", area.Address
    snapshot.Add "Values", values
    snapshot.Add "Formulas", formulas
    snapshot.Add "FormulaMasks", masks
    snapshot.Add "Formats", formats
    Set fn_ReadRange = snapshot
End Function

Public Function fn_Token(ByVal value As Variant) As String
    Dim text As String

    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsObject(value) Or VBA.IsArray(value) Then
        fn_Fail "InvalidStampValue", ""
    End If
    text = VBA.CStr(value)
    fn_Token = VBA.CStr(VBA.VarType(value)) & ":" & VBA.Len(text) & ":" & text
End Function

Public Function fn_PlanStamp(ByVal plan As Object) As String
    Dim update As Object
    Dim action As Object
    Dim writes As Object
    Dim mapping As Variant
    Dim key As Variant
    Dim links As Object
    Dim stamp As String

    Set links = fn_Headers()
    stamp = fn_Token(plan("OrderIdentity")) & fn_Token(ex_PEB_Export.fn_Config("export.Movement.dateFormat"))
    For Each key In links.Keys
        stamp = stamp & fn_Token(key) & fn_Token(links(key))
    Next key
    For Each action In plan("Actions")
        stamp = stamp & fn_Token(action("RowId")) & fn_Token(action("HistoryStamp")) & fn_Token(action("Status"))
        If action.Exists("Operation") Then
            stamp = stamp & fn_Token(action("Operation"))
        End If
        For Each key In VBA.Array("SourceRow", "ClosingIndex", "TargetIndex", "PlaceholderIndex", "IsUpdate", "Revision", "EventStatus")
            If action.Exists(key) Then
                stamp = stamp & fn_Token(key) & fn_Token(action(key))
            End If
        Next key
        For Each mapping In VBA.Array("Writes", "ClosingWrites", "PlaceholderWrites", "Joins", "GeneratedWrites", "GeneratedClosingWrites")
            If action.Exists(mapping) Then
                Set writes = action(mapping)
                For Each key In writes.Keys
                    stamp = stamp & fn_Token(key) & fn_Token(writes(key))
                Next key
            End If
        Next mapping
        If action.Exists("UpdateCells") Then
            For Each update In action("UpdateCells")
                For Each key In VBA.Array("RowIndex", "Header", "Before", "After")
                    stamp = stamp & fn_Token(key) & fn_Token(update(key))
                Next key
            Next update
        End If
    Next action
    fn_PlanStamp = stamp
End Function

Public Sub fn_Fail( _
    ByVal code As String, _
    ByVal detail As String _
)
    Err.Raise 5, , VBA.Replace$(VBA.Replace$(ex_PEB_Export.fn_Config("text.MovementPlanError"), "{code}", code), "{detail}", detail)
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_ReadFormats( _
    ByVal area As Range, _
    ByRef formats() As Variant, _
    ByVal columnIndex As Long, _
    ByVal firstRow As Long _
)
    Dim format As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim half As Long

    format = area.NumberFormat
    rowCount = area.Rows.Count
    If Not VBA.IsNull(format) Then
        For rowIndex = firstRow To firstRow + rowCount - 1
            formats(rowIndex, columnIndex) = format
        Next rowIndex
    Else
        If rowCount = 1 Then
            fn_Fail "InvalidNumberFormat", area.Address
        End If
        half = rowCount \ 2
        private_ReadFormats area.Resize(half, 1), formats, columnIndex, firstRow
        private_ReadFormats area.Cells(half + 1, 1).Resize(rowCount - half, 1), formats, columnIndex, firstRow + half
    End If
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------