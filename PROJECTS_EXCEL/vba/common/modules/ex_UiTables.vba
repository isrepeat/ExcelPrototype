Attribute VB_Name = "ex_UiTables"
Option Explicit

Private Const MARKER_PREFIX As String = "_pxTable_"

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_CreatePlan( _
    ByVal definition As Object, _
    ByVal target As Range, _
    ByVal rawTable As obj_UiRawTable, _
    ByVal showHeaders As Boolean _
) As Object
    Dim plan As Object
    Dim area As Range
    Dim owner As String
    Dim mode As String
    Dim index As String
    Dim headerRow As Long
    Dim tableName As String
    Dim headers As Object
    Dim column As Long
    Dim header As Variant
    Dim conversionRange As Range

    mode = ex_UiElementFactory.fn_Attribute(definition, "representation")
    If VBA.Len(mode) = 0 Then mode = "Raw"
    If mode <> "Raw" And mode <> "Smart" Then private_Fail "Invalid table representation: " & mode
    index = ex_UiElementFactory.fn_Attribute(definition, "tableIndex")
    If VBA.Len(index) = 0 Then index = "1"
    owner = ex_UiElementFactory.fn_Attribute(definition, "name") & "|" & index
    tableName = ex_UiElementFactory.fn_Attribute(definition, "excelTableName")
    If mode = "Raw" And VBA.Len(tableName) > 0 Then private_Fail "excelTableName requires Smart representation."
    headerRow = 1
    If VBA.Len(rawTable.Title) > 0 Then headerRow = 2
    Set area = target
    If showHeaders And VBA.IsArray(rawTable.Headers) And rawTable.RowCount > 0 Then
        Set conversionRange = target.Cells(headerRow, 1).Resize(rawTable.RowCount + 1, rawTable.ColumnCount)
    End If
    If mode = "Smart" Then
        If VBA.Len(owner) > 200 Then private_Fail "Table owner exceeds 200 characters."
        If Not showHeaders Or Not VBA.IsArray(rawTable.Headers) Then private_Fail "Smart table requires visible headers: " & owner
        If VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "onSelect")) > 0 Or _
                VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "selectedItem")) > 0 Then
            private_Fail "Smart table does not support positional selection bindings: " & owner
        End If
        Set headers = VBA.CreateObject("Scripting.Dictionary")
        headers.CompareMode = VBA.vbTextCompare
        For column = 1 To rawTable.ColumnCount
            header = rawTable.HeaderAt(column)
            If VBA.VarType(header) <> VBA.vbString Then private_Fail "Smart table header must be text: " & owner
            If VBA.Len(VBA.Trim$(header)) = 0 Or VBA.Len(header) > 255 Then private_Fail "Invalid Smart table header: " & owner
            If headers.Exists(VBA.Trim$(header)) Then private_Fail "Duplicate Smart table header: " & header
            headers.Add VBA.Trim$(header), True
        Next column
        Set area = target.Cells(headerRow, 1).Resize(Application.Max(1, rawTable.RowCount) + 1, rawTable.ColumnCount)
        Set conversionRange = area
    End If
    Set plan = VBA.CreateObject("Scripting.Dictionary")
    plan.Add "owner", owner
    plan.Add "marker", MARKER_PREFIX & private_Hash(target.Parent.CodeName & "|" & owner)
    plan.Add "smart", mode = "Smart"
    plan.Add "range", Empty
    Set plan("range") = area
    plan.Add "layoutRange", Empty
    Set plan("layoutRange") = target
    plan.Add "conversionRange", Empty
    If Not conversionRange Is Nothing Then Set plan("conversionRange") = conversionRange
    plan.Add "selectable", VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "onSelect")) > 0 Or _
        VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "selectedItem")) > 0
    plan.Add "name", tableName
    Set fn_CreatePlan = plan
End Function

Public Sub fn_Prepare( _
    ByVal sheet As Worksheet, _
    ByVal plans As Collection _
)
    Dim plan As Object
    Dim other As Object
    Dim names As Object
    Dim owners As Object
    Dim owned As Object
    Dim marker As Name
    Dim table As ListObject
    Dim area As Range
    Dim otherArea As Range
    Dim key As Variant
    Dim tableName As String
    Dim suffix As Long
    Dim candidate As ListObject
    Dim keep As Boolean
    Dim previousRange As Range
    Dim parts As Collection
    Dim part As Range
    Dim smartPlans As New Collection

    On Error GoTo EH_PREPARE
    Set names = VBA.CreateObject("Scripting.Dictionary")
    names.CompareMode = VBA.vbTextCompare
    Set owners = VBA.CreateObject("Scripting.Dictionary")
    Set owned = private_OwnedTables(sheet)
    ' Сначала проверяем весь план, до изменения объектов листа.
    For Each plan In plans
        If owners.Exists(plan("marker")) Then private_Fail "Duplicate table owner: " & plan("owner")
        owners.Add plan("marker"), Empty
        Set owners(plan("marker")) = plan
        Set marker = private_FindMarker(sheet, plan("marker"))
        If Not marker Is Nothing Then
        End If
        If plan("smart") Then
            smartPlans.Add plan
            tableName = plan("name")
            If VBA.Len(tableName) = 0 Then
                tableName = "tbGenerated_" & private_SafePart(sheet.Name) & "_" & _
                    private_SafePart(VBA.Replace(plan("owner"), "|", "_")) & "_" & private_Hash(sheet.CodeName & "|" & plan("owner"))
                suffix = 0
                Do
                    Set candidate = private_FindTable(sheet.Parent, tableName)
                    If candidate Is Nothing And Not names.Exists(tableName) And Not private_DefinedNameExists(sheet.Parent, tableName) Then Exit Do
                    If Not candidate Is Nothing Then
                        If owned.Exists(plan("marker")) Then
                            Set table = owned(plan("marker"))
                            If candidate Is table Then Exit Do
                        End If
                    End If
                    suffix = suffix + 1
                    tableName = "tbGenerated_" & private_Hash(sheet.CodeName & "|" & plan("owner")) & "_" & VBA.CStr(suffix)
                Loop
            End If
            private_ValidateName tableName
            If private_DefinedNameExists(sheet.Parent, tableName) Then private_Fail "Excel name already exists: " & tableName
            If names.Exists(tableName) Then private_Fail "Duplicate Excel table name: " & tableName
            names.Add tableName, True
            Set candidate = private_FindTable(sheet.Parent, tableName)
            If Not candidate Is Nothing Then
                Set marker = private_FindMarker(sheet, plan("marker"))
                If marker Is Nothing Then private_Fail "Excel table name already exists: " & tableName
            End If
            plan("name") = tableName
        End If
    Next plan
    For Each plan In smartPlans
        Set area = plan("layoutRange")
        For Each other In plans
            If Not plan Is other Then
                Set otherArea = other("layoutRange")
                If Not Application.Intersect(area, otherArea) Is Nothing Then private_Fail "Table ranges overlap."
            End If
        Next other
        For Each table In sheet.ListObjects
            If Not Application.Intersect(area, table.Range) Is Nothing Then
                If Not private_IsPlannedTable(table, smartPlans) Then
                    private_Fail "Table range overlaps an unmanaged Excel table: " & table.Name
                End If
            End If
        Next table
    Next plan
    ' Сначала освобождаем перемещённые и исчезнувшие таблицы, затем расширяем оставшиеся.
    For Each key In owned.Keys
        Set table = owned(key)
        keep = False
        If owners.Exists(key) Then
            Set plan = owners(key)
            Set area = plan("range")
            keep = plan("smart") And area.Row = table.Range.Row And area.Column = table.Range.Column
        End If
        If Not keep Then
            Set previousRange = table.Range
            table.Unlist
            previousRange.Clear
            Set marker = private_FindMarker(sheet, VBA.CStr(key))
            If Not marker Is Nothing Then marker.Delete
        End If
    Next key
    For Each plan In plans
        If plan("smart") Then
            Set marker = private_FindMarker(sheet, plan("marker"))
            Set table = Nothing
            If Not marker Is Nothing Then Set table = private_MarkerTable(marker, sheet)
            If Not table Is Nothing Then
                Set area = plan("range")
                Set previousRange = table.Range
                table.ShowTotals = False
                If table.AutoFilter.FilterMode Then table.AutoFilter.ShowAllData
                ' Уменьшаем все старые области до расширения соседних таблиц.
                Set part = Application.Intersect(table.Range, area)
                If table.Range.Address <> part.Address Then table.Resize part
                Set parts = New Collection
                private_Subtract previousRange, Application.Intersect(previousRange, area), parts
                For Each part In parts
                    part.Clear
                Next part
                table.Range.ClearFormats
            End If
        End If
    Next plan
    For Each plan In smartPlans
        Set marker = private_FindMarker(sheet, plan("marker"))
        Set table = Nothing
        If Not marker Is Nothing Then Set table = private_MarkerTable(marker, sheet)
        If Not table Is Nothing Then
            Set area = plan("range")
            If table.Range.Address <> area.Address Then table.Resize area
        End If
    Next plan
    Exit Sub
EH_PREPARE:
    Err.Raise Err.Number, "ex_UiTables.fn_Prepare", Err.Description
End Sub

Public Sub fn_Apply(ByVal plan As Object)
    Dim area As Range
    Dim table As ListObject
    Dim marker As Name
    Dim sheet As Worksheet

    If Not plan("smart") Then Exit Sub
    Set area = plan("range")
    Set sheet = area.Parent
    Set marker = private_FindMarker(sheet, plan("marker"))
    If Not marker Is Nothing Then Set table = private_MarkerTable(marker, sheet)
    If table Is Nothing And Not marker Is Nothing Then Set table = private_FindTable(sheet.Parent, plan("name"))
    If table Is Nothing Then
        Set table = sheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=area, XlListObjectHasHeaders:=xlYes)
        ' Маркер создаём сразу, чтобы последующая ошибка не оставила бесхозную таблицу.
        Set marker = sheet.Names.Add(Name:=plan("marker"), RefersTo:="=1", Visible:=False)
        marker.Comment = table.Name
    End If
    If table.Name <> plan("name") Then table.Name = plan("name")
    marker.Comment = table.Name
    table.TableStyle = ""
    table.ShowTotals = False
    table.ShowHeaders = True
End Sub

Public Function fn_TryEnsureSmart( _
    ByVal tableRange As Range, _
    ByVal owner As String, _
    ByRef diagnostic As String, _
    Optional ByVal excelTableName As String _
) As Boolean
    Dim plan As Object
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim marker As Name
    Dim headers As Object
    Dim cell As Range
    Dim value As Variant
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean
    Dim existing As ListObject
    Dim automaticName As Boolean
    Dim suffix As Long
    Dim previousExpand As Boolean
    Dim previousFill As Boolean

    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    previousExpand = Application.AutoCorrect.AutoExpandListRange
    previousFill = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo Failed
    diagnostic = VBA.vbNullString
    If tableRange Is Nothing Then private_Fail "Table range is required."
    If tableRange.Areas.Count <> 1 Or tableRange.Rows.Count < 2 Then private_Fail "Smart range requires headers and at least one body row."
    If VBA.Len(owner) = 0 Or VBA.Len(owner) > 200 Then private_Fail "Table owner must contain 1 to 200 characters."
    Set sheet = tableRange.Parent
    Set headers = VBA.CreateObject("Scripting.Dictionary")
    headers.CompareMode = VBA.vbTextCompare
    For Each cell In tableRange.Rows(1).Cells
        value = cell.Value2
        If VBA.VarType(value) <> VBA.vbString Then private_Fail "Smart table header must be text."
        If VBA.Len(VBA.Trim$(value)) = 0 Or VBA.Len(value) > 255 Then private_Fail "Invalid Smart table header."
        If headers.Exists(VBA.Trim$(value)) Then private_Fail "Duplicate Smart table header: " & value
        headers.Add VBA.Trim$(value), True
    Next cell
    Set plan = VBA.CreateObject("Scripting.Dictionary")
    plan.Add "owner", owner
    plan.Add "marker", MARKER_PREFIX & private_Hash(sheet.CodeName & "|" & owner)
    plan.Add "smart", True
    plan.Add "range", Empty
    Set plan("range") = tableRange
    automaticName = VBA.Len(excelTableName) = 0
    If automaticName Then excelTableName = "tbGenerated_" & private_SafePart(sheet.Name) & "_" & _
        private_SafePart(VBA.Replace(owner, "|", "_")) & "_" & private_Hash(sheet.CodeName & "|" & owner)
    private_ValidateName excelTableName
    plan.Add "name", excelTableName
    Set marker = private_FindMarker(sheet, plan("marker"))
    If Not marker Is Nothing Then
        Set existing = private_MarkerTable(marker, sheet)
        If Not existing Is Nothing Then
            If automaticName Then excelTableName = existing.Name
            plan("name") = excelTableName
            If existing.Range.Address <> tableRange.Address Then private_Fail "Existing table range does not match."
        End If
    End If
    Set table = private_FindTable(sheet.Parent, excelTableName)
    If automaticName And existing Is Nothing Then
        Do While Not table Is Nothing Or private_DefinedNameExists(sheet.Parent, excelTableName)
            suffix = suffix + 1
            excelTableName = "tbGenerated_" & private_Hash(sheet.CodeName & "|" & owner) & "_" & VBA.CStr(suffix)
            Set table = private_FindTable(sheet.Parent, excelTableName)
        Loop
        plan("name") = excelTableName
    End If
    If private_DefinedNameExists(sheet.Parent, excelTableName) Then private_Fail "Excel name already exists: " & excelTableName
    If Not table Is Nothing Then
        If marker Is Nothing Then private_Fail "Excel table name already exists: " & excelTableName
        If Not existing Is table Then private_Fail "Excel table name belongs to another table."
    End If
    For Each table In sheet.ListObjects
        If Not Application.Intersect(tableRange, table.Range) Is Nothing Then
            If marker Is Nothing Then private_Fail "Range overlaps an existing Excel table."
            If Not existing Is table Then private_Fail "Range overlaps another Excel table."
            If table.Range.Address <> tableRange.Address Then private_Fail "Existing table range does not match."
        End If
    Next table
    value = tableRange.MergeCells
    If VBA.IsNull(value) Then private_Fail "Smart table cannot contain merged cells."
    If value Then private_Fail "Smart table cannot contain merged cells."
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.AutoCorrect.AutoExpandListRange = False
    Application.AutoCorrect.AutoFillFormulasInLists = False
    fn_Apply plan
    fn_TryEnsureSmart = True
CleanExit:
    Application.AutoCorrect.AutoExpandListRange = previousExpand
    Application.AutoCorrect.AutoFillFormulasInLists = previousFill
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    Exit Function
Failed:
    diagnostic = Err.Description
    Resume CleanExit
End Function

Public Sub fn_ClearRange( _
    ByVal area As Range, _
    Optional ByVal contentsOnly As Boolean = False _
)
    Dim owned As Object
    Dim parts As Collection
    Dim nextParts As Collection
    Dim part As Range
    Dim overlap As Range
    Dim table As ListObject

    Set owned = private_OwnedTables(area.Parent)
    Set parts = New Collection
    parts.Add area
    For Each table In area.Parent.ListObjects
        If Not Application.Intersect(area, table.Range) Is Nothing Then
            If Not private_IsOwned(table, owned) Then private_Fail "UI clear overlaps an unmanaged Excel table: " & table.Name
            Set nextParts = New Collection
            For Each part In parts
                Set overlap = Application.Intersect(part, table.Range)
                If overlap Is Nothing Then
                    nextParts.Add part
                Else
                    private_Subtract part, overlap, nextParts
                End If
            Next part
            Set parts = nextParts
        End If
    Next table
    For Each part In parts
        part.UnMerge
        If contentsOnly Then
            part.ClearContents
        Else
            part.Clear
        End If
    Next part
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_Subtract( _
    ByVal area As Range, _
    ByVal cut As Range, _
    ByVal output As Collection _
)
    Dim top As Long
    Dim left As Long
    Dim bottom As Long
    Dim right As Long

    top = cut.Row - area.Row
    left = cut.Column - area.Column
    bottom = area.Row + area.Rows.Count - cut.Row - cut.Rows.Count
    right = area.Column + area.Columns.Count - cut.Column - cut.Columns.Count
    If top > 0 Then output.Add area.Cells(1, 1).Resize(top, area.Columns.Count)
    If bottom > 0 Then output.Add area.Cells(top + cut.Rows.Count + 1, 1).Resize(bottom, area.Columns.Count)
    If left > 0 Then output.Add area.Cells(top + 1, 1).Resize(cut.Rows.Count, left)
    If right > 0 Then output.Add area.Cells(top + 1, left + cut.Columns.Count + 1).Resize(cut.Rows.Count, right)
End Sub

Private Function private_OwnedTables(ByVal sheet As Worksheet) As Object
    Dim result As Object
    Dim marker As Name
    Dim table As ListObject
    Dim key As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    For Each marker In sheet.Parent.Names
        key = VBA.Mid$(marker.Name, VBA.InStrRev(marker.Name, "!") + 1)
        If VBA.Left$(key, VBA.Len(MARKER_PREFIX)) = MARKER_PREFIX Then
            Set table = private_MarkerTable(marker, sheet)
            If Not table Is Nothing Then
                If table.Parent Is sheet Then
                    result.Add key, Empty
                    Set result(key) = table
                End If
            End If
        End If
    Next marker
    Set private_OwnedTables = result
End Function

Private Function private_FindMarker( _
    ByVal sheet As Worksheet, _
    ByVal key As String _
) As Name
    Dim marker As Name

    For Each marker In sheet.Parent.Names
        If VBA.Mid$(marker.Name, VBA.InStrRev(marker.Name, "!") + 1) = key Then
            Set private_FindMarker = marker
            Exit Function
        End If
    Next marker
End Function

Private Function private_MarkerTable( _
    ByVal marker As Name, _
    ByVal sheet As Worksheet _
) As ListObject
    Dim table As ListObject
    Dim tableName As String

    tableName = private_MarkerTableName(marker)
    If VBA.Len(tableName) = 0 Then Exit Function
    For Each table In sheet.ListObjects
        If VBA.StrComp(table.Name, tableName, VBA.vbTextCompare) = 0 Then
            Set private_MarkerTable = table
            Exit Function
        End If
    Next table
End Function

Private Function private_MarkerTableName(ByVal marker As Name) As String
    private_MarkerTableName = marker.Comment
End Function

Private Function private_IsOwned( _
    ByVal table As ListObject, _
    ByVal owned As Object _
) As Boolean
    Dim key As Variant
    Dim candidate As ListObject

    For Each key In owned.Keys
        Set candidate = owned(key)
        If candidate Is table Then
            private_IsOwned = True
            Exit Function
        End If
    Next key
End Function

Private Function private_IsPlannedTable( _
    ByVal table As ListObject, _
    ByVal plans As Collection _
) As Boolean
    Dim plan As Object

    For Each plan In plans
        If VBA.StrComp(table.Name, plan("name"), VBA.vbTextCompare) = 0 Then
            private_IsPlannedTable = True
            Exit Function
        End If
    Next plan
End Function

Private Function private_FindTable( _
    ByVal workbook As Workbook, _
    ByVal tableName As String _
) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject

    For Each sheet In workbook.Worksheets
        For Each table In sheet.ListObjects
            If VBA.StrComp(table.Name, tableName, VBA.vbTextCompare) = 0 Then
                Set private_FindTable = table
                Exit Function
            End If
        Next table
    Next sheet
End Function

Private Function private_DefinedNameExists( _
    ByVal workbook As Workbook, _
    ByVal value As String _
) As Boolean
    Dim item As Name
    Dim localName As String

    For Each item In workbook.Names
        localName = VBA.Mid$(item.Name, VBA.InStrRev(item.Name, "!") + 1)
        If VBA.StrComp(localName, value, VBA.vbTextCompare) = 0 Then
            private_DefinedNameExists = True
            Exit Function
        End If
    Next item
End Function

Private Function private_Hash(ByVal value As String) As String
    Dim index As Long
    Dim hash As Double

    For index = 1 To VBA.Len(value)
        hash = hash * 31# + (VBA.AscW(VBA.Mid$(value, index, 1)) And &HFFFF&)
        hash = hash - VBA.Fix(hash / 2147483647#) * 2147483647#
    Next index
    private_Hash = VBA.Hex$(VBA.CLng(hash))
End Function

Private Function private_SafePart(ByVal value As String) As String
    Dim index As Long
    Dim character As String
    Dim result As String

    For index = 1 To Application.Min(40, VBA.Len(value))
        character = VBA.Mid$(value, index, 1)
        If Not character Like "[A-Za-z0-9_]" Then character = "_"
        result = result & character
    Next index
    private_SafePart = result
End Function

Private Sub private_ValidateName(ByVal value As String)
    Dim regex As Object

    Set regex = VBA.CreateObject("VBScript.RegExp")
    regex.Pattern = "^[A-Za-z_][A-Za-z0-9_]*$"
    If VBA.Len(value) > 255 Or Not regex.Test(value) Then private_Fail "Invalid Excel table name: " & value
    regex.IgnoreCase = True
    regex.Pattern = "^([A-Z]{1,3}[0-9]+|R[0-9]*C[0-9]*|R|C)$"
    If regex.Test(value) Then private_Fail "Excel table name resembles a cell reference: " & value
End Sub

Private Sub private_Fail(ByVal diagnostic As String)
    Err.Raise VBA.vbObjectError + 2281, "ex_UiTables", diagnostic
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------