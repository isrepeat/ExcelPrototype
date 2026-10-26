Attribute VB_Name = "ex_PEB_MovementTvo"
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
Public Function fn_IsStarted( _
    ByVal startOrder As Variant, _
    ByVal startDate As Variant, _
    ByRef startDateKey As Double, _
    ByVal detail As String _
) As Boolean
    startDateKey = 0
    ' Admission alone does not establish an acting-duty interval.
    If Len(ex_PEB_Export.fn_Text(startOrder)) = 0 Then
        Exit Function
    End If
    If Len(ex_PEB_Export.fn_Text(startDate)) = 0 Then
        ex_PEB_MovementLinks.fn_Fail "TvoMissingStartDate", detail
    End If
    startDateKey = CDbl(ex_PEB_Export.fn_ParseDate(startDate, "StartDate"))
    fn_IsStarted = True
End Function

Public Function fn_ReturnRows( _
    ByVal values As Object, _
    ByVal period As Object, _
    ByRef found As Boolean _
) As Collection
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim table As obj_DataTable
    Dim model As New Collection
    Dim fields As Variant
    Dim field As Variant
    Dim row As Object
    Dim index As Long
    Dim diagnostic As String
    Dim failure As String

    On Error GoTo EH
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        ex_PEB_MovementLinks.fn_Fail "TvoSourceInitializationFailed", ""
    End If
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("reporter.tvo.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("reporter.tvo.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("reporter.tvo.source.range")
    fields = private_Fields()
    For Each field In fields
        query.AddColumn private_Header(CStr(field))
    Next field
    ' Return restoration reads live source data, independently of reporter snapshot dates.
    service.Backend = QueryAuto
    If Not service.TryExecute(source, query, table, diagnostic, False) Then
        Err.Raise 5, , diagnostic
    End If
    For index = 1 To table.RowCount
        Set row = CreateObject("Scripting.Dictionary")
        For Each field In fields
            row(field) = table.ValueAt(index, ex_TableQuery.fn_HeaderIndex(table.Headers, private_Header(CStr(field))))
        Next field
        row("Index") = index
        model.Add row
    Next index
    Set fn_ReturnRows = private_Chain(model, ex_PEB_Export.fn_Text(period("PositionCode")), _
        ex_PEB_History.fn_Date(values("EventDate")), ex_PEB_Export.fn_Text(values("AcceptedOrderNumber")), _
        ex_PEB_MovementSource.fn_OrderDate(period("DepartureDate")), found)
Cleanup:
    If Not table Is Nothing Then
        table.Dispose
    End If
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If Len(failure) > 0 Then
        On Error GoTo 0
        Err.Raise 5, , failure
    End If
    Exit Function
EH:
    failure = Err.Description
    Resume Cleanup
End Function

Public Sub fn_Prepare( _
    ByVal plan As Object, _
    ByVal target As ListObject _
)
    Dim table As ListObject
    Dim model As New Collection
    Dim originals As New Collection
    Dim entries As New Collection
    Dim state As Object
    Dim row As Object
    Dim original As Object
    Dim action As Object
    Dim values As Object
    Dim chain As Collection
    Dim item As Object
    Dim field As Variant
    Dim entry As Object
    Dim cell As Range
    Dim extension As Range
    Dim other As ListObject
    Dim found As Boolean
    Dim index As Long
    Dim eventDate As Date
    Dim operation As String
    Dim basis As String
    Dim originalCount As Long
    Dim actors As Object
    Dim existing As Object
    Dim lowerDate As Double
    Dim existingStartDate As Double

    Set table = private_Target(target.Parent.Parent)
    If table.Parent.ProtectContents Or Not table.ShowHeaders Then
        ex_PEB_MovementLinks.fn_Fail "InvalidTvoTarget", table.Name
    End If
    originalCount = table.ListRows.Count
    For index = 1 To originalCount
        Set row = CreateObject("Scripting.Dictionary")
        Set original = CreateObject("Scripting.Dictionary")
        For Each field In private_Fields()
            row(field) = table.DataBodyRange.Cells(index, table.ListColumns(private_Header(CStr(field))).Index).Value2
            original(field) = row(field)
        Next field
        row("Index") = index
        row("Blank") = (Application.WorksheetFunction.CountA(table.ListRows(index).Range) = 0)
        model.Add row
        originals.Add original
    Next index
    For Each action In plan("Actions")
        If action("Status") = "Conflict" Then
            GoTo NextAction
        End If
        If Not action.Exists("TvoValues") Then
            GoTo NextAction
        End If
        Set values = action("TvoValues")
        If action("Status") = "AlreadyApplied" Then
            GoTo NextAction
        End If
        operation = action("Operation")
        eventDate = ex_PEB_History.fn_Date(values("EventDate"))
        lowerDate = 0
        If action.Exists("ClosingIndex") Then
            If action("ClosingIndex") > 0 And action("ClosingIndex") <= target.ListRows.Count Then
                lowerDate = ex_PEB_MovementSource.fn_OrderDate(target.DataBodyRange.Cells( _
                    action("ClosingIndex"), _
                    target.ListColumns(ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")).Index).Value2)
            End If
        End If
        ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_TVO_EVENT | RowId=" & CStr(action("RowId")) & _
            " | Person=" & CStr(values("PersonName")) & " | Position=" & CStr(values("PositionCode")) & _
            " | Operation=" & operation & " | Date=" & Format$(eventDate, "yyyy-mm-dd")
        Set chain = private_Chain(model, CStr(values("PositionCode")), eventDate, CStr(values("AcceptedOrderNumber")), lowerDate, found, operation = "close")
        If action("EventStatus") = "Cancelled" Then
            If chain.Count > 0 Then
                ex_PEB_MovementLinks.fn_Fail "TvoCancellationRequiresReview", action("RowId")
            End If
            GoTo NextAction
        End If
        ' A repeated export must not close a replacement segment opened by the same transition.
        If action("IsUpdate") And chain.Count > 0 Then
            ex_PEB_MovementLinks.fn_Fail "TvoRevisionRequiresReview", action("RowId")
        End If
        If operation = "append" And chain.Count > 0 Then
            If Not private_Matches(chain, values("MetaTvoTables"), model, eventDate) Then
                ex_PEB_MovementLinks.fn_Fail "TvoAssignmentAlreadyOpen", action("RowId")
            End If
            GoTo NextAction
        End If
        If operation = "close" Or operation = "transition" Then
            basis = private_Basis(action, "ClosingWrites", "ArrivalBasis")
            For Each row In chain
                If ex_PEB_MovementSource.fn_OrderDate(row("EndDate")) > CDbl(eventDate) Then
                    ex_PEB_MovementLinks.fn_Fail "TvoHasLaterHandover", action("RowId")
                End If
                row("EndDate") = eventDate
                row("EndOrder") = CStr(values("AcceptedOrderNumber"))
                row("EndBasis") = basis
                Set row("Action") = action
            Next row
        End If
        If operation = "append" Or operation = "transition" Then
            If chain.Count = 0 Then
                Set chain = values("MetaTvoTables")
            End If
            If operation = "transition" And chain.Count = 0 And Not found Then
                Set chain = ex_PEB_ReturnTvo.fn_LegacyRows(private_LegacyPeriod(target, action), CStr(values("PersonId")))
            End If
            basis = private_Basis(action, "Writes", "DepartureBasis")
            Set actors = CreateObject("Scripting.Dictionary")
            actors.CompareMode = vbTextCompare
            For Each item In chain
                If private_AdmissionOnly(model, item, eventDate) Then
                    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_TVO_ADMISSION_RETAINED | RowId=" & CStr(action("RowId")) & _
                        " | PersonId=" & CStr(item("PersonId")) & " | Position=" & CStr(item("PositionCode"))
                    GoTo NextDuty
                End If
                If CStr(item("PersonId")) = CStr(values("PersonId")) Or actors.Exists(CStr(item("PersonId"))) Then
                    ex_PEB_MovementLinks.fn_Fail "InvalidTvoChain", action("RowId")
                End If
                actors.Add CStr(item("PersonId")), True
                For Each existing In model
                    If StrComp(ex_PEB_Export.fn_Text(existing("PositionCode")), CStr(item("PositionCode")), vbTextCompare) = 0 Then
                        If fn_IsStarted(existing("StartOrder"), existing("StartDate"), existingStartDate, CStr(item("PositionCode"))) Then
                            If ex_PEB_MovementSource.fn_OrderDate(existing("EndDate")) = 0 Or _
                                    ex_PEB_MovementSource.fn_OrderDate(existing("EndDate")) > CDbl(eventDate) Then
                                ex_PEB_MovementLinks.fn_Fail "OverlappingTvoAssignment", CStr(item("PositionCode"))
                            End If
                        End If
                    End If
                Next existing
                Set row = private_NewRow(model, originals)
                row("PersonId") = item("PersonId")
                row("PersonName") = item("PersonName")
                row("PositionCode") = item("PositionCode")
                row("OwnPositionCode") = private_Staff(CStr(item("PersonId")), "lookup.Personnel.columns.6.header")
                row("PersonRank") = private_Staff(CStr(item("PersonId")), "lookup.Personnel.columns.3.header")
                row("StartDate") = eventDate
                row("StartOrder") = CStr(values("AcceptedOrderNumber"))
                row("AccessOrder") = CStr(values("AcceptedOrderNumber"))
                row("StartBasis") = basis
                If values.Exists("DateToShort") Then
                    row("PlannedEndDate") = values("DateToShort")
                End If
                Set row("Action") = action
NextDuty:
            Next item
        End If
NextAction:
    Next action
    For index = 1 To model.Count
        Set row = model(index)
        Set original = originals(index)
        For Each field In private_Fields()
            If ex_PEB_MovementLinks.fn_ValueToken(row(field)) <> ex_PEB_MovementLinks.fn_ValueToken(original(field)) Then
                Set cell = table.HeaderRowRange.Cells(1, table.ListColumns(private_Header(CStr(field))).Index).Offset(index)
                If cell.HasFormula Then
                    ex_PEB_MovementLinks.fn_Fail "FormulaOverwrite", cell.Address
                End If
                Set entry = CreateObject("Scripting.Dictionary")
                Set entry("Cell") = cell
                Set entry("Action") = row("Action")
                entry("Before") = cell.Value2
                entry("After") = row(field)
                entry("Format") = cell.NumberFormat
                entry("Header") = private_Header(CStr(field))
                entry("Added") = (Len(ex_PEB_Export.fn_Text(original("PersonId"))) = 0)
                entries.Add entry
            End If
        Next field
    Next index
    Set state = CreateObject("Scripting.Dictionary")
    Set state("Table") = table
    Set state("Entries") = entries
    state("OriginalAddress") = table.Range.Address
    state("OriginalCount") = originalCount
    state("FinalCount") = model.Count
    state("Stamp") = private_TableStamp(table)
    state("Changed") = False
    If model.Count > originalCount Then
        If table.ShowTotals Then
            ex_PEB_MovementLinks.fn_Fail "AppendWithTotalsUnsupported", table.Name
        End If
        Set extension = table.Range.Rows(table.Range.Rows.Count).Offset(1).Resize(model.Count - originalCount)
        If IsNull(extension.MergeCells) Then
            ex_PEB_MovementLinks.fn_Fail "MergedExpansionCell", extension.Address
        End If
        If extension.MergeCells Or Application.WorksheetFunction.CountA(extension) <> 0 Then
            ex_PEB_MovementLinks.fn_Fail "OccupiedExpansionCell", extension.Address
        End If
        For Each other In table.Parent.ListObjects
            If Not other Is table Then
                If Not Application.Intersect(extension, other.Range) Is Nothing Then
                    ex_PEB_MovementLinks.fn_Fail "ExpansionIntersectsTable", other.Name
                End If
            End If
        Next other
        Set state("Extension") = ex_PEB_MovementLinks.fn_ReadRange(extension)
    End If
    Set plan("TvoPlan") = state
End Sub

Public Function fn_Stamp(ByVal plan As Object) As String
    Dim state As Object
    Dim entry As Object
    Dim text As String

    If Not plan.Exists("TvoPlan") Then
        Exit Function
    End If
    Set state = plan("TvoPlan")
    text = state("Stamp") & CStr(state("FinalCount"))
    For Each entry In state("Entries")
        text = text & ex_PEB_MovementLinks.fn_Token(entry("Cell").Address) & _
            ex_PEB_MovementLinks.fn_Token(entry("Header")) & ex_PEB_MovementLinks.fn_ValueToken(entry("Before")) & _
            ex_PEB_MovementLinks.fn_ValueToken(entry("After"))
    Next entry
    fn_Stamp = text
End Function

Public Function fn_ChangeCount(ByVal plan As Object) As Long
    If plan.Exists("TvoPlan") Then
        fn_ChangeCount = plan("TvoPlan")("Entries").Count
    End If
End Function

Public Sub fn_Apply(ByVal plan As Object)
    Dim state As Object
    Dim table As ListObject
    Dim entry As Object
    Dim cell As Range
    Dim expected As Variant
    Dim currentFormat As Variant

    Set state = plan("TvoPlan")
    Set table = state("Table")
    If private_TableStamp(table) <> state("Stamp") Then
        ex_PEB_MovementLinks.fn_Fail "TvoTargetChangedSincePreview", table.Name
    End If
    If state("Entries").Count = 0 Then
        Exit Sub
    End If
    state("Changed") = True
    If state("FinalCount") > state("OriginalCount") Then
        table.Resize table.Range.Resize(state("FinalCount") + 1)
    End If
    For Each entry In state("Entries")
        Set cell = entry("Cell")
        expected = entry("After")
        currentFormat = cell.NumberFormat
        If VarType(expected) = vbString Then
            cell.NumberFormat = "@"
        End If
        If VarType(expected) = vbDate Then
            expected = CDbl(expected)
        End If
        cell.Value2 = expected
        cell.NumberFormat = currentFormat
        If cell.HasFormula Or ex_PEB_MovementLinks.fn_ValueToken(cell.Value2) <> ex_PEB_MovementLinks.fn_ValueToken(entry("After")) Then
            ex_PEB_MovementLinks.fn_Fail "WrittenPlanVerificationFailed", cell.Address
        End If
    Next entry
End Sub

Public Function fn_Restore(ByVal plan As Object) As String
    Dim state As Object
    Dim table As ListObject
    Dim entry As Object
    Dim cell As Range
    Dim extension As Object
    Dim formats As Variant
    Dim area As Range
    Dim row As Long
    Dim column As Long

    If Not plan.Exists("TvoPlan") Then
        Exit Function
    End If
    Set state = plan("TvoPlan")
    If Not state("Changed") Then
        Exit Function
    End If
    Set table = state("Table")
    On Error Resume Next
    For Each entry In state("Entries")
        Err.Clear
        Set cell = entry("Cell")
        cell.NumberFormat = "@"
        cell.Value2 = entry("Before")
        cell.NumberFormat = entry("Format")
        If Err.Number <> 0 Then
            fn_Restore = fn_Restore & " " & Err.Description
        End If
        Err.Clear
        If ex_PEB_MovementLinks.fn_ValueToken(cell.Value2) <> ex_PEB_MovementLinks.fn_ValueToken(entry("Before")) Then
            fn_Restore = fn_Restore & " " & cell.Address
        End If
    Next entry
    Err.Clear
    table.Resize table.Parent.Range(state("OriginalAddress"))
    If Err.Number <> 0 Then
        fn_Restore = fn_Restore & " " & Err.Description
    End If
    If state.Exists("Extension") Then
        Err.Clear
        Set extension = state("Extension")
        Set area = table.Parent.Range(extension("Address"))
        formats = extension("Formats")
        area.ClearContents
        If Err.Number <> 0 Then
            fn_Restore = fn_Restore & " " & Err.Description
        End If
        For row = 1 To area.Rows.Count
            For column = 1 To area.Columns.Count
                Err.Clear
                area.Cells(row, column).NumberFormat = formats(row, column)
                If Err.Number <> 0 Then
                    fn_Restore = fn_Restore & " " & Err.Description
                End If
            Next column
        Next row
    End If
    Err.Clear
    If table.Range.Address <> state("OriginalAddress") Then
        fn_Restore = fn_Restore & " " & table.Name
    End If
    If Err.Number <> 0 Then
        fn_Restore = fn_Restore & " " & Err.Description
    End If
    On Error GoTo 0
End Function

Public Sub fn_AddDiff( _
    ByVal report As Object, _
    ByVal plan As Object _
)
    Dim entry As Object
    Dim row As Object
    Dim action As Object

    For Each entry In plan("TvoPlan")("Entries")
        Set action = entry("Action")
        Set row = CreateObject("Scripting.Dictionary")
        row("PersonName") = action("PersonName")
        row("EventRowId") = action("RowId")
        row("EventName") = action("EventName")
        row("Worksheet") = plan("TvoPlan")("Table").Parent.Name
        row("TargetRow") = CStr(entry("Cell").Row)
        row("Field") = entry("Header")
        row("Before") = ex_PEB_Export.fn_Text(entry("Before"))
        row("After") = ex_PEB_Export.fn_Text(entry("After"))
        If Len(row("Before")) = 0 Then
            row("Before") = ex_PEB_Export.fn_Config("movementDiff.empty")
        End If
        If Len(row("After")) = 0 Then
            row("After") = ex_PEB_Export.fn_Config("movementDiff.empty")
        End If
        row("Action") = ex_PEB_Export.fn_Config("movementDiff.updated")
        If entry("Added") Then
            row("Action") = ex_PEB_Export.fn_Config("movementDiff.added")
        End If
        report("Rows").Add row
    Next entry
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Fields() As Variant
    private_Fields = Split("PersonRank|PersonName|PersonId|OwnPositionCode|PositionCode|StartDate|EndDate|StartOrder|EndOrder|StartBasis|EndBasis|PlannedEndDate|AccessOrder", "|")
End Function

Private Function private_Target(ByVal book As Workbook) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim sheetName As String
    Dim tableName As String

    sheetName = ex_PEB_Export.fn_Config("reporter.tvo.source.sheet")
    tableName = ex_PEB_Export.fn_Config("returnTvo.table.name")
    For Each sheet In book.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            For Each table In sheet.ListObjects
                If StrComp(table.Name, tableName, vbTextCompare) = 0 Then
                    Set private_Target = table
                    Exit Function
                End If
            Next table
        End If
    Next sheet
    ex_PEB_MovementLinks.fn_Fail "MissingTvoTable", sheetName & " / " & tableName
End Function

Private Function private_Header(ByVal field As String) As String
    private_Header = ex_PEB_Export.fn_Config("returnTvo.table.headers." & field)
End Function

Private Function private_Chain( _
    ByVal model As Collection, _
    ByVal positionCode As String, _
    ByVal eventDate As Date, _
    ByVal orderNumber As String, _
    ByVal lowerDate As Double, _
    ByRef found As Boolean, _
    Optional ByVal includeClosed As Boolean = True _
) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim selected As Object
    Dim seen As Object
    Dim startDate As Double
    Dim endDate As Double
    Dim closed As Boolean
    Dim selectedClosed As Boolean

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    found = False
    Do While Len(positionCode) > 0
        If seen.Exists(positionCode) Then
            ex_PEB_MovementLinks.fn_Fail "TvoCycle", positionCode
        End If
        seen.Add positionCode, True
        Set selected = Nothing
        selectedClosed = False
        For Each row In model
            If StrComp(ex_PEB_Export.fn_Text(row("PositionCode")), positionCode, vbTextCompare) = 0 Then
                If Not fn_IsStarted(row("StartOrder"), row("StartDate"), startDate, positionCode) Then
                    ' An admission is represented in the registry but restores no acting duties.
                    found = True
                    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_TVO_NOT_STARTED | Position=" & positionCode & _
                        " | PersonId=" & ex_PEB_Export.fn_Text(row("PersonId")) & _
                        " | AccessOrder=" & ex_PEB_Export.fn_Text(row("AccessOrder"))
                    GoTo NextAssignment
                End If
                endDate = ex_PEB_MovementSource.fn_OrderDate(row("EndDate"))
                If endDate <> 0 And endDate < startDate Then
                    ex_PEB_MovementLinks.fn_Fail "InvalidTvoDates", positionCode
                End If
                closed = (endDate = CDbl(eventDate) And ex_PEB_Export.fn_Text(row("EndOrder")) = orderNumber)
                If startDate <= CDbl(eventDate) And (endDate = 0 Or endDate > lowerDate Or closed) Then
                    found = True
                    If endDate = 0 Or endDate > CDbl(eventDate) Or (closed And includeClosed) Then
                        If selected Is Nothing Then
                            Set selected = row
                            selectedClosed = closed
                        ElseIf selectedClosed And Not closed Then
                            Set selected = row
                            selectedClosed = False
                        ElseIf Not selectedClosed And closed Then
                            ' Prefer the still-open segment after a same-day transition.
                        Else
                            ex_PEB_MovementLinks.fn_Fail "AmbiguousTvoAssignment", positionCode
                        End If
                    End If
                End If
            End If
NextAssignment:
        Next row
        If selected Is Nothing Then
            Exit Do
        End If
        If Len(ex_PEB_Export.fn_Text(selected("PersonId"))) = 0 Or Len(ex_PEB_Export.fn_Text(selected("PersonName"))) = 0 Then
            ex_PEB_MovementLinks.fn_Fail "IncompleteTvoAssignment", positionCode
        End If
        result.Add selected
        positionCode = ex_PEB_Export.fn_Text(selected("OwnPositionCode"))
        If Len(positionCode) = 0 Then
            ex_PEB_MovementLinks.fn_Fail "TvoMissingOwnPosition", CStr(selected("PersonId"))
        End If
    Loop
    Set private_Chain = result
End Function

Private Function private_AdmissionOnly( _
    ByVal model As Collection, _
    ByVal item As Object, _
    ByVal eventDate As Date _
) As Boolean
    Dim row As Object
    Dim admitted As Boolean
    Dim startDate As Double
    Dim endDate As Double

    ' A segment already accepted in the registry may continue through a status transition.
    If item.Exists("StartOrder") Then
        If Len(ex_PEB_Export.fn_Text(item("StartOrder"))) > 0 Then
            Exit Function
        End If
    End If
    For Each row In model
        If StrComp(ex_PEB_Export.fn_Text(row("PersonId")), CStr(item("PersonId")), vbTextCompare) = 0 And _
                StrComp(ex_PEB_Export.fn_Text(row("PositionCode")), CStr(item("PositionCode")), vbTextCompare) = 0 Then
            If Len(ex_PEB_Export.fn_Text(row("StartOrder"))) = 0 Then
                admitted = True
            ElseIf fn_IsStarted(row("StartOrder"), row("StartDate"), startDate, CStr(item("PositionCode"))) Then
                endDate = ex_PEB_MovementSource.fn_OrderDate(row("EndDate"))
                If startDate <= CDbl(eventDate) And (endDate = 0 Or endDate > CDbl(eventDate)) Then
                    Exit Function
                End If
            End If
        End If
    Next row
    private_AdmissionOnly = admitted
End Function

Private Function private_Matches( _
    ByVal actual As Collection, _
    ByVal requested As Collection, _
    ByVal model As Collection, _
    ByVal eventDate As Date _
) As Boolean
    Dim index As Long
    Dim effective As New Collection
    Dim item As Object

    For Each item In requested
        If Not private_AdmissionOnly(model, item, eventDate) Then
            effective.Add item
        End If
    Next item
    If effective.Count = 0 Then
        private_Matches = True
        Exit Function
    End If
    If actual.Count <> effective.Count Then
        Exit Function
    End If
    For index = 1 To actual.Count
        If CStr(actual(index)("PersonId")) <> CStr(effective(index)("PersonId")) Then
            Exit Function
        End If
        If StrComp(CStr(actual(index)("PositionCode")), CStr(effective(index)("PositionCode")), vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next index
    private_Matches = True
End Function

Private Function private_NewRow( _
    ByVal model As Collection, _
    ByVal originals As Collection _
) As Object
    Dim row As Object
    Dim original As Object
    Dim field As Variant

    For Each row In model
        If row("Blank") Then
            row("Blank") = False
            Set private_NewRow = row
            Exit Function
        End If
    Next row
    Set row = CreateObject("Scripting.Dictionary")
    Set original = CreateObject("Scripting.Dictionary")
    For Each field In private_Fields()
        row(field) = vbNullString
        original(field) = vbNullString
    Next field
    row("Index") = model.Count + 1
    row("Blank") = False
    model.Add row
    originals.Add original
    Set private_NewRow = row
End Function

Private Function private_Basis( _
    ByVal action As Object, _
    ByVal mapping As String, _
    ByVal field As String _
) As String
    Dim header As String

    header = ex_PEB_Export.fn_Config("export.Movement.headers." & field)
    If action(mapping).Exists(header) Then
        private_Basis = ex_PEB_Export.fn_Text(action(mapping)(header))
    End If
End Function

Private Function private_LegacyPeriod( _
    ByVal target As ListObject, _
    ByVal action As Object _
) As Object
    Dim period As Object
    Dim field As Variant
    Dim index As Long

    If Not action.Exists("ClosingIndex") Then
        Exit Function
    End If
    index = action("ClosingIndex")
    If index < 1 Or index > target.ListRows.Count Then
        Exit Function
    End If
    Set period = CreateObject("Scripting.Dictionary")
    For Each field In Array("TvoName", "TvoPersonId", "TvoPositionCode")
        period(field) = target.DataBodyRange.Cells(index, target.ListColumns( _
            ex_PEB_Export.fn_Config("returnTvo.headers." & field)).Index).Value2
    Next field
    Set private_LegacyPeriod = period
End Function

Private Function private_Staff( _
    ByVal personId As String, _
    ByVal fieldKey As String _
) As String
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim table As obj_DataTable
    Dim diagnostic As String
    Dim available As Boolean
    Dim failure As String

    On Error GoTo EH
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        ex_PEB_MovementLinks.fn_Fail "TvoStaffInitializationFailed", personId
    End If
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("lookup.Personnel.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("lookup.Personnel.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("lookup.Personnel.source.range")
    query.AddColumn ex_PEB_Export.fn_Config(fieldKey)
    query.AddCondition ex_PEB_Export.fn_Config("lookup.Personnel.columns.2.header"), QueryEquals, personId, QueryText, True
    query.Limit = 2
    If Not ex_TableDataRuntime.fn_TryQueryLoaded(source, query, available, table, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If Not available Then
        If Not service.TryExecute(source, query, table, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
    End If
    If table.RowCount <> 1 Then
        ex_PEB_MovementLinks.fn_Fail "AmbiguousTvoStaff", personId
    End If
    private_Staff = ex_PEB_Export.fn_Text(table.ValueAt(1, 1))
    If Len(private_Staff) = 0 Then
        ex_PEB_MovementLinks.fn_Fail "IncompleteTvoStaff", personId
    End If
Cleanup:
    If Not table Is Nothing Then
        table.Dispose
    End If
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If Len(failure) > 0 Then
        On Error GoTo 0
        Err.Raise 5, , failure
    End If
    Exit Function
EH:
    failure = Err.Description
    Resume Cleanup
End Function

Private Function private_TableStamp(ByVal table As ListObject) As String
    Dim cell As Range
    Dim text As String

    text = ex_PEB_MovementLinks.fn_StructureStamp(table)
    For Each cell In table.Range.Cells
        text = text & ex_PEB_MovementLinks.fn_ValueToken(cell.Value2) & _
            ex_PEB_MovementLinks.fn_ValueToken(cell.Formula)
    Next cell
    private_TableStamp = text
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------