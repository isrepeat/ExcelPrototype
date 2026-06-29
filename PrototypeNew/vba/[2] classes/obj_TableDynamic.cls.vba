VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableDynamic"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_SectionTitle As String
Private m_Columns As list__obj_Column
Private m_Rows As list__obj_Row
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_Columns = New list__obj_Column
    Set m_Rows = New list__obj_Row
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get SectionTitle() As String
    SectionTitle = m_SectionTitle
End Property

Public Property Let SectionTitle(ByVal value As String)
    m_SectionTitle = VBA.CStr(value)
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_Columns.Count
End Property

Public Property Get RowCount() As Long
    RowCount = m_Rows.Count
End Property

Public Property Get Columns() As list__obj_Column
    Set Columns = m_Columns
End Property

Public Property Get Rows() As list__obj_Row
    Set Rows = m_Rows
End Property

Public Property Get HeaderText() As String
    Dim i As Long
    Dim colObj As obj_Column
    Dim joined As String

    For i = 1 To m_Columns.Count
        Set colObj = m_Columns.Item(i)
        If i > 1 Then joined = joined & " | "
        joined = joined & colObj.Name
    Next i

    HeaderText = joined
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Set m_Columns = Nothing
    Set m_Rows = Nothing
    On Error GoTo 0
End Sub

Public Function PushColumn(ByVal tableColumn As obj_Column) As Boolean
    Dim newColumn As obj_Column

    If tableColumn Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_TableDynamic: column is not specified."
#End If
        Exit Function
    End If

    Set newColumn = New obj_Column
    newColumn.Name = tableColumn.Name
    newColumn.Position = m_Columns.Count + 1
    newColumn.FormatKind = tableColumn.FormatKind

    If VBA.Len(newColumn.Name) = 0 Then
        newColumn.Name = "Col" & VBA.CStr(newColumn.Position)
    End If

    If Not private_CopyColumnAliases(tableColumn, newColumn) Then Exit Function
    PushColumn = m_Columns.Add(newColumn)
End Function

Public Function InsertColumnAt( _
    ByVal tableColumn As obj_Column, _
    ByVal oneBasedIndex As Long, _
    Optional ByVal cellDesc As String = VBA.vbNullString _
) As Boolean
    Dim oldColumns As list__obj_Column
    Dim rebuiltColumns As list__obj_Column
    Dim newColumn As obj_Column
    Dim existingColumn As obj_Column
    Dim rowObj As obj_Row
    Dim oldCount As Long
    Dim i As Long
    Dim oldIndex As Long

    If tableColumn Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_TableDynamic: column is not specified for insert."
#End If
        Exit Function
    End If

    oldCount = m_Columns.Count
    If oneBasedIndex <= 0 Then oneBasedIndex = 1
    If oneBasedIndex > oldCount + 1 Then oneBasedIndex = oldCount + 1

    Set oldColumns = m_Columns
    Set rebuiltColumns = New list__obj_Column

    Set newColumn = New obj_Column
    newColumn.Name = VBA.Trim$(tableColumn.Name)
    newColumn.FormatKind = tableColumn.FormatKind
    If VBA.Len(newColumn.Name) = 0 Then newColumn.Name = "Col" & VBA.CStr(oneBasedIndex)
    If Not private_CopyColumnAliases(tableColumn, newColumn) Then Exit Function

    oldIndex = 1
    For i = 1 To oldCount + 1
        If i = oneBasedIndex Then
            newColumn.Position = i
            If Not rebuiltColumns.Add(newColumn) Then Exit Function
        Else
            Set existingColumn = oldColumns.Item(oldIndex)
            If existingColumn Is Nothing Then Exit Function
            existingColumn.Position = i
            If Not rebuiltColumns.Add(existingColumn) Then Exit Function
            oldIndex = oldIndex + 1
        End If
    Next i

    Set m_Columns = rebuiltColumns

    For i = 1 To m_Rows.Count
        Set rowObj = m_Rows.Item(i)
        If rowObj Is Nothing Then GoTo ContinueRow
        If Not rowObj.InsertCellAt(oneBasedIndex, VBA.vbNullString, cellDesc) Then Exit Function
ContinueRow:
    Next i

    InsertColumnAt = True
End Function

Public Function PushRow(ByVal tableRow As obj_Row) As Boolean
    Dim requiredCols As Long

    If tableRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_TableDynamic: row is not specified."
#End If
        Exit Function
    End If

    requiredCols = tableRow.CellCount
    If requiredCols > m_Columns.Count Then
        If Not private_EnsureColumns(requiredCols) Then Exit Function
    End If

    PushRow = m_Rows.Add(tableRow)
End Function

Public Function InsertRowAt( _
    ByVal tableRow As obj_Row, _
    ByVal oneBasedIndex As Long _
) As Boolean
    Dim oldRows As list__obj_Row
    Dim rebuiltRows As list__obj_Row
    Dim existingRow As obj_Row
    Dim oldCount As Long
    Dim oldIndex As Long
    Dim i As Long
    Dim requiredCols As Long

    If tableRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_TableDynamic: row is not specified for insert."
#End If
        Exit Function
    End If

    requiredCols = tableRow.CellCount
    If requiredCols > m_Columns.Count Then
        If Not private_EnsureColumns(requiredCols) Then Exit Function
    End If

    oldCount = m_Rows.Count
    If oneBasedIndex <= 0 Then oneBasedIndex = 1
    If oneBasedIndex > oldCount + 1 Then oneBasedIndex = oldCount + 1

    Set oldRows = m_Rows
    Set rebuiltRows = New list__obj_Row
    oldIndex = 1

    For i = 1 To oldCount + 1
        If i = oneBasedIndex Then
            If Not rebuiltRows.Add(tableRow) Then Exit Function
        Else
            Set existingRow = oldRows.Item(oldIndex)
            If existingRow Is Nothing Then Exit Function
            If Not rebuiltRows.Add(existingRow) Then Exit Function
            oldIndex = oldIndex + 1
        End If
    Next i

    Set m_Rows = rebuiltRows
    InsertRowAt = True
End Function

Public Function TryGetColumnIndexByAlias(ByVal aliasName As String, ByRef outIndex As Long) As Boolean
    Dim colIndex As Long
    Dim columnObj As obj_Column

    outIndex = 0
    aliasName = VBA.Trim$(VBA.CStr(aliasName))
    If VBA.Len(aliasName) = 0 Then Exit Function

    For colIndex = 1 To m_Columns.Count
        Set columnObj = m_Columns.Item(colIndex)
        If columnObj Is Nothing Then GoTo ContinueColumn
        If columnObj.HasAlias(aliasName) Then
            outIndex = colIndex
            TryGetColumnIndexByAlias = True
            Exit Function
        End If
ContinueColumn:
    Next colIndex
End Function

Public Function TryGetColumnIndexByName(ByVal columnName As String, ByRef outIndex As Long) As Boolean
    Dim colIndex As Long
    Dim columnObj As obj_Column
    Dim expectedName As String

    outIndex = 0
    expectedName = private_NormalizeText(columnName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For colIndex = 1 To m_Columns.Count
        Set columnObj = m_Columns.Item(colIndex)
        If columnObj Is Nothing Then GoTo ContinueColumn
        If VBA.StrComp(private_NormalizeText(columnObj.Name), expectedName, VBA.vbTextCompare) = 0 Then
            outIndex = colIndex
            TryGetColumnIndexByName = True
            Exit Function
        End If
ContinueColumn:
    Next colIndex
End Function

Public Function TryFindRowIndexByDesc(ByVal descToken As String, ByRef outIndex As Long) As Boolean
    Dim rowIndex As Long
    Dim rowObj As obj_Row
    Dim normalizedToken As String
    Dim rowDesc As String

    outIndex = 0
    normalizedToken = VBA.Trim$(VBA.CStr(descToken))
    If VBA.Len(normalizedToken) = 0 Then Exit Function

    For rowIndex = 1 To m_Rows.Count
        Set rowObj = m_Rows.Item(rowIndex)
        If rowObj Is Nothing Then GoTo ContinueRow

        rowDesc = VBA.Trim$(rowObj.Desc)
        If VBA.Len(rowDesc) = 0 Then GoTo ContinueRow
        If VBA.InStr(1, rowDesc, normalizedToken, VBA.vbTextCompare) > 0 Then
            outIndex = rowIndex
            TryFindRowIndexByDesc = True
            Exit Function
        End If
ContinueRow:
    Next rowIndex
End Function

Public Function TryGetRowByDesc(ByVal descToken As String, ByRef outRow As obj_Row) As Boolean
    Dim rowIndex As Long

    Set outRow = Nothing
    If Not Me.TryFindRowIndexByDesc(descToken, rowIndex) Then Exit Function
    If rowIndex <= 0 Then Exit Function

    Set outRow = m_Rows.Item(rowIndex)
    TryGetRowByDesc = Not outRow Is Nothing
End Function

' //
' // Internal
' //
Private Function private_EnsureColumns(ByVal requiredCount As Long) As Boolean
    Dim i As Long
    Dim autoColumn As obj_Column

    If requiredCount <= m_Columns.Count Then
        private_EnsureColumns = True
        Exit Function
    End If

    For i = m_Columns.Count + 1 To requiredCount
        Set autoColumn = New obj_Column
        autoColumn.Position = i
        autoColumn.Name = "Col" & VBA.CStr(i)
        m_Columns.Add autoColumn
    Next i

    private_EnsureColumns = True
End Function

Private Function private_CopyColumnAliases(ByVal sourceColumn As obj_Column, ByVal targetColumn As obj_Column) As Boolean
    Dim aliasItem As Variant
    Dim aliases As Collection

    If sourceColumn Is Nothing Then Exit Function
    If targetColumn Is Nothing Then Exit Function

    targetColumn.ClearAliases
    Set aliases = sourceColumn.Aliases
    If aliases Is Nothing Then
        private_CopyColumnAliases = True
        Exit Function
    End If

    For Each aliasItem In aliases
        If Not targetColumn.AddAlias(VBA.CStr(aliasItem)) Then Exit Function
    Next aliasItem

    private_CopyColumnAliases = True
End Function

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.CStr(valueText)
    valueText = VBA.Replace$(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace$(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace$(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace$(valueText, VBA.ChrW$(160), " ")
    valueText = VBA.Replace$(valueText, "  ", " ")
    valueText = VBA.Replace$(valueText, "  ", " ")
    valueText = VBA.Trim$(valueText)
    private_NormalizeText = VBA.LCase$(valueText)
End Function
