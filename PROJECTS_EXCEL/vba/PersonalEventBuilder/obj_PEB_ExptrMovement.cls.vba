VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_ExptrMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_path As String
Private m_sheet As String
Private m_table As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    Dim enabled As String

    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    enabled = ex_PEB_Export.fn_Config("export.Movement.enabled")
    If enabled = "false" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ExportMovementUnavailable")
    End If
    If enabled <> "true" Then
        Err.Raise 5, , "Movement enabled must be true or false."
    End If
    m_path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Movement.path"))
    m_sheet = ex_PEB_Export.fn_Config("export.Movement.sheet")
    m_table = ex_PEB_Export.fn_Config("export.Movement.table")
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
End Sub

Public Function TryExport( _
    ByVal values As Object, _
    ByRef diagnostic As String, _
    Optional ByRef resultCode As String = "", _
    Optional ByVal targetWorkbook As Object = Nothing _
) As Boolean
    Dim book As Object
    Dim table As Object
    Dim targetRow As Object
    Dim openBook As Workbook
    Dim sheet As Worksheet
    Dim candidateTable As ListObject
    Dim environment As String
    Dim writes As Object
    Dim identities As Object
    Dim joins As Object
    Dim prefix As String
    Dim operation As String
    Dim openHeader As String
    Dim columns As Object
    Dim column As Object
    Dim key As Variant
    Dim rowIndex As Long
    Dim found As Long
    Dim selected As Long
    Dim changed As Boolean
    Dim saved As Boolean
    Dim saveBook As Boolean
    Dim closingWrites As Object
    Dim closingRow As Object
    Dim appendedRow As Object
    Dim previousValues As Object
    Dim previousFormats As Object

    On Error GoTo EH
    diagnostic = VBA.vbNullString
    resultCode = "ExportError"
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Movement exporter is not initialized."
    End If
    prefix = "profiles." & ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "MovementPlanId"))
    operation = ex_PEB_MovementRegistry.fn_Value(prefix & ".operation")
    If operation <> "append" And operation <> "close" And operation <> "transition" Then
        Err.Raise 5, , "Unknown Movement operation: " & operation
    End If
    ex_PEB_MovementRegistry.fn_ValidateTexts values
    Set writes = private_Mappings(prefix & ".fields", values)
    Set identities = private_Mappings(prefix & ".identity", values)
    If operation = "close" Or operation = "transition" Then
        Set joins = private_Mappings(prefix & ".join", values)
        openHeader = ex_PEB_MovementRegistry.fn_Value(prefix & ".openHeader")
        Set closingWrites = writes
        If operation = "transition" Then
            Set closingWrites = private_Mappings("profiles." & ex_PEB_MovementRegistry.fn_Value(prefix & ".closeProfile") & ".fields", values)
        End If
    Else
        Set joins = private_Mappings(prefix & ".guard.join", values)
        openHeader = ex_PEB_MovementRegistry.fn_Value(prefix & ".guard.openHeader")
    End If
    ' Resolve an already open target without creating another Excel instance.
    environment = ex_PEB_Export.fn_Config("environment.name")
    If targetWorkbook Is Nothing Then
        For Each openBook In Application.Workbooks
            If environment <> "TEST" Or VBA.StrComp(openBook.FullName, m_path, VBA.vbTextCompare) = 0 Then
                For Each sheet In openBook.Worksheets
                    For Each candidateTable In sheet.ListObjects
                        If VBA.StrComp(candidateTable.Name, m_table, VBA.vbTextCompare) = 0 Then
                            If Not table Is Nothing Then
                                Err.Raise 5, , ex_PEB_Export.fn_Config("text.MovementOpenAmbiguous")
                            End If
                            Set table = candidateTable
                            Set book = openBook
                        End If
                    Next candidateTable
                Next sheet
            End If
        Next openBook
        If table Is Nothing Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.MovementOpenRequired"), "{table}", m_table)
        End If
        saveBook = True
    Else
        If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then Err.Raise 5, , "Borrowed export workbooks require TEST."
        If VBA.StrComp(targetWorkbook.FullName, m_path, VBA.vbTextCompare) <> 0 Then
            Err.Raise 5, , "Borrowed workbook differs from the configured Movement target."
        End If
        Set book = targetWorkbook
        Set table = book.Worksheets(m_sheet).ListObjects(m_table)
    End If
    If book.ReadOnly Then
        Err.Raise 5, , "Movement workbook is locked or read-only: " & m_path
    End If
    If book.Date1904 Then
        Err.Raise 5, , "Movement workbook must use the 1900 date system."
    End If
    If table.Parent.ProtectContents Then
        Err.Raise 5, , "Movement worksheet is protected."
    End If
    Set columns = VBA.CreateObject("Scripting.Dictionary")
    columns.CompareMode = VBA.vbTextCompare
    For Each column In table.ListColumns
        columns.Add VBA.CStr(column.Name), VBA.CLng(column.Index)
    Next column
    private_ValidateHeaders writes, columns
    private_ValidateHeaders identities, columns
    private_ValidateHeaders joins, columns
    If Not columns.Exists(openHeader) Then
        Err.Raise 5, , "Movement open header not found: " & openHeader
    End If
    If operation = "close" Or operation = "transition" Then
        private_ValidateHeaders joins, columns
        private_ValidateHeaders closingWrites, columns
        If Not columns.Exists(openHeader) Then
            Err.Raise 5, , "Movement open header not found: " & openHeader
        End If
        If Not closingWrites.Exists(openHeader) Then
            Err.Raise 5, , "Movement close mapping must populate: " & openHeader
        End If
        If VBA.Len(ex_PEB_Export.fn_Text(closingWrites(openHeader))) = 0 Then
            Err.Raise 5, , "Movement close value is empty: " & openHeader
        End If
    End If

    ' Проверяем идентичность событий до изменения таблицы.
    For rowIndex = 1 To table.ListRows.Count
        Set targetRow = table.ListRows(rowIndex)
        If private_Matches(targetRow, identities, columns) Then
            resultCode = "DuplicateEvent"
            Err.Raise 5, , "Movement event already exists."
        End If
        If operation = "close" Or operation = "transition" Then
            If private_Matches(targetRow, joins, columns) Then
                If VBA.Len(ex_PEB_Export.fn_Text(targetRow.Range.Cells(1, columns(openHeader)).Value2)) = 0 Then
                    found = found + 1
                    selected = rowIndex
                End If
            End If
        ElseIf private_Matches(targetRow, joins, columns) Then
            If VBA.Len(ex_PEB_Export.fn_Text(targetRow.Range.Cells(1, columns(openHeader)).Value2)) = 0 Then
                found = found + 1
            End If
        End If
    Next rowIndex
    If operation = "append" And found > 0 Then
        resultCode = "OpenDepartureExists"
        Err.Raise 5, , "Movement departure requires an explicit transition from an existing open event."
    End If
    If operation = "close" Or operation = "transition" Then
        If found <> 1 Then
            resultCode = "AmbiguousOpenDeparture"
            If found = 0 Then resultCode = "OpenDepartureNotFound"
            Err.Raise 5, , "Movement arrival requires exactly one matching open departure; found " & VBA.CStr(found)
        End If
        Set targetRow = table.ListRows(selected)
        private_ValidateChronology targetRow, values, columns, prefix, resultCode
        For Each key In closingWrites.Keys
            If targetRow.Range.Cells(1, columns(key)).HasFormula Then
                Err.Raise 5, , "Movement arrival cannot overwrite a formula: " & VBA.CStr(key)
            End If
            If VBA.Len(ex_PEB_Export.fn_Text(targetRow.Range.Cells(1, columns(key)).Value2)) > 0 Then
                Err.Raise 5, , "Movement arrival would overwrite populated data: " & VBA.CStr(key)
            End If
        Next key
        Set closingRow = targetRow
        Set previousValues = VBA.CreateObject("Scripting.Dictionary")
        Set previousFormats = VBA.CreateObject("Scripting.Dictionary")
        For Each key In closingWrites.Keys
            previousValues(key) = closingRow.Range.Cells(1, columns(key)).Value2
            previousFormats(key) = closingRow.Range.Cells(1, columns(key)).NumberFormat
        Next key
    End If
    changed = True
    If operation = "transition" Then
        For Each key In closingWrites.Keys
            If VBA.VarType(closingWrites(key)) = VBA.vbString Then
                closingRow.Range.Cells(1, columns(key)).NumberFormat = "@"
            End If
            closingRow.Range.Cells(1, columns(key)).Value2 = closingWrites(key)
        Next key
    End If
    If operation <> "close" Then
        Set appendedRow = table.ListRows.Add
        Set targetRow = appendedRow
    End If
    For Each key In writes.Keys
        If VBA.VarType(writes(key)) = VBA.vbString Then
            targetRow.Range.Cells(1, columns(key)).NumberFormat = "@"
        End If
        targetRow.Range.Cells(1, columns(key)).Value2 = writes(key)
    Next key
    If saveBook Then book.Save
    saved = True
    resultCode = "DepartureCreated"
    If operation = "close" Then resultCode = "ArrivalClosed"
    If operation = "transition" Then
        resultCode = "StatusChanged"
    End If
    TryExport = True
Cleanup:
    If changed And Not saved Then
        diagnostic = diagnostic & " " & ex_PEB_Export.fn_Config("text.ExportNotSaved")
    End If
    Exit Function
EH:
    diagnostic = "Movement export: " & VBA.Err.Description
    ' Restore the open workbook when an export fails before saving.
    If changed And Not saved Then
        On Error Resume Next
        If Not appendedRow Is Nothing Then
            appendedRow.Delete
        End If
        If Not previousValues Is Nothing Then
            For Each key In previousValues.Keys
                closingRow.Range.Cells(1, columns(key)).Value2 = previousValues(key)
                closingRow.Range.Cells(1, columns(key)).NumberFormat = previousFormats(key)
            Next key
        End If
        If VBA.Err.Number <> 0 Then
            diagnostic = diagnostic & " Rollback: " & VBA.Err.Description
        End If
        On Error GoTo 0
    End If
    Resume Cleanup
End Function

' //
' // Private
' //
Private Function private_Mappings( _
    ByVal prefix As String, _
    ByVal values As Object _
) As Object
    Dim result As Object
    Dim index As Long
    Dim count As Long
    Dim field As String
    Dim header As String
    Dim value As Variant

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    count = ex_PEB_MovementRegistry.fn_Count(prefix & ".count")
    If count = 0 Then
        Err.Raise 5, , "Movement mappings cannot be empty: " & prefix
    End If
    For index = 1 To count
        field = ex_PEB_MovementRegistry.fn_Value(prefix & "." & VBA.CStr(index) & ".field")
        header = ex_PEB_MovementRegistry.fn_Value(prefix & "." & VBA.CStr(index) & ".header")
        If VBA.Right$(prefix, 7) = ".fields" Then
            value = ex_PEB_MovementRegistry.fn_WriteValue(values, field, header)
        Else
            value = ex_PEB_Export.fn_Required(values, field)
        End If
        If VBA.InStr(prefix, ".fields") = 0 Then
            If VBA.Len(ex_PEB_Export.fn_Text(value)) = 0 Then
                Err.Raise 5, , "Movement match field is empty: " & field
            End If
        End If
        If result.Exists(header) Then
            Err.Raise 5, , "Duplicate Movement mapping: " & header
        End If
        result.Add header, value
    Next index
    Set private_Mappings = result
End Function

Private Sub private_ValidateHeaders( _
    ByVal mappings As Object, _
    ByVal columns As Object _
)
    Dim key As Variant

    For Each key In mappings.Keys
        If Not columns.Exists(key) Then
            Err.Raise 5, , "Movement target header not found: " & VBA.CStr(key)
        End If
    Next key
End Sub

Private Function private_Matches( _
    ByVal row As Object, _
    ByVal mappings As Object, _
    ByVal columns As Object _
) As Boolean
    Dim key As Variant
    Dim actual As Variant
    Dim expected As Variant

    For Each key In mappings.Keys
        actual = row.Range.Cells(1, columns(key)).Value2
        expected = mappings(key)
        If VBA.CStr(key) = ex_PEB_Export.fn_Config("export.Movement.headers.Destination") Then
            If VBA.StrComp( _
                    ex_PEB_MovementRegistry.fn_Destination(actual), _
                    ex_PEB_MovementRegistry.fn_Destination(expected), _
                    VBA.vbTextCompare) <> 0 Then Exit Function
        ElseIf VBA.VarType(expected) = VBA.vbDate Then
            If Not VBA.IsNumeric(actual) Then
                Exit Function
            End If
            If VBA.CDbl(actual) <> VBA.CDbl(expected) Then
                Exit Function
            End If
        ElseIf VBA.StrComp(ex_PEB_Export.fn_Text(actual), ex_PEB_Export.fn_Text(expected), VBA.vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next key
    private_Matches = True
End Function

Private Sub private_ValidateChronology( _
    ByVal row As Object, _
    ByVal values As Object, _
    ByVal columns As Object, _
    ByVal prefix As String, _
    ByRef resultCode As String _
)
    Dim header As String
    Dim field As String
    Dim departure As Variant
    Dim arrival As Variant

    header = ex_PEB_MovementRegistry.fn_Value(prefix & ".departureDateHeader")
    field = ex_PEB_MovementRegistry.fn_Value(prefix & ".arrivalDateField")
    If Not columns.Exists(header) Then
        Err.Raise 5, , "Movement departure date header not found: " & header
    End If
    departure = row.Range.Cells(1, columns(header)).Value2
    arrival = ex_PEB_Export.fn_Required(values, field)
    If Not VBA.IsNumeric(departure) Or VBA.VarType(arrival) <> VBA.vbDate Then
        Err.Raise 5, , "Movement requires typed departure and arrival dates."
    End If
    If VBA.CDbl(arrival) < VBA.CDbl(departure) Then
        resultCode = "ArrivalBeforeDeparture"
        Err.Raise 5, , "Movement arrival precedes departure."
    End If
End Sub