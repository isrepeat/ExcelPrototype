Attribute VB_Name = "ex_TableFilters"
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
Public Function fn_Suspend( _
    ByVal tables As Collection, _
    ByRef states As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim table As ListObject
    Dim state As Object
    Dim restoreError As String
    Dim seen As Object
    Dim key As String
    Dim context As String

    On Error GoTo EH
    Set states = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    diagnostic = vbNullString
    ' Capture every table before clearing the first one.
    For Each table In tables
        context = "Capture/" & table.Parent.Name & "/" & table.Name
        key = table.Parent.Parent.FullName & vbNullChar & table.Parent.CodeName & vbNullChar & table.Name
        If Not seen.Exists(key) Then
            Set state = fn_Capture(table)
            states.Add state
            seen(key) = True
        End If
    Next table
    For Each state In states
        Set table = state("Table")
        context = "Clear/" & table.Parent.Name & "/" & table.Name
        fn_Clear table
    Next state
    fn_Suspend = True
    Exit Function
EH:
    diagnostic = "TableFilters." & context & ": " & Err.Description
    restoreError = fn_Restore(states)
    If Len(restoreError) > 0 Then
        diagnostic = diagnostic & vbCrLf & restoreError
    End If
End Function

Public Sub fn_Clear(ByVal table As ListObject)
    Dim autoFilter As Excel.AutoFilter

    Set autoFilter = table.AutoFilter
    If autoFilter Is Nothing Then
        Exit Sub
    End If
    If autoFilter.FilterMode Then
        autoFilter.ShowAllData
    End If
End Sub

Public Sub fn_ClearCaptured(ByVal states As Collection)
    Dim state As Object
    Dim table As ListObject

    If states Is Nothing Then
        Exit Sub
    End If
    For Each state In states
        Set table = state("Table")
        fn_Clear table
    Next state
End Sub

Public Function fn_Capture(ByVal table As ListObject) As Object
    Dim state As Object
    Dim criteria As New Collection
    Dim entry As Object
    Dim filter As Object
    Dim index As Long
    Dim autoFilter As Excel.AutoFilter

    Set state = CreateObject("Scripting.Dictionary")
    Set state("Table") = table
    Set state("Criteria") = criteria
    Set autoFilter = table.AutoFilter
    ' Tables without an AutoFilter object have no active criteria to capture or clear.
    If autoFilter Is Nothing Then
        Set fn_Capture = state
        Exit Function
    End If
    For index = 1 To autoFilter.Filters.Count
        Set filter = autoFilter.Filters(index)
        If filter.On Then
            Set entry = CreateObject("Scripting.Dictionary")
            entry("Field") = index
            entry("Operator") = filter.Operator
            private_CaptureFilterCriterion filter, "Criteria1", entry
            private_CaptureFilterCriterion filter, "Criteria2", entry
            If Not entry.Exists("Criteria1") And Not entry.Exists("Criteria2") Then
                Err.Raise 5, "ex_TableFilters", "TableFilterCaptureFailed: " & table.Name
            End If
            criteria.Add entry
        End If
    Next index
    Set fn_Capture = state
End Function

Public Function fn_Restore(ByVal states As Collection) As String
    Dim state As Object
    Dim entry As Object
    Dim table As ListObject

    If states Is Nothing Then
        Exit Function
    End If
    On Error Resume Next
    For Each state In states
        Set table = state("Table")
        For Each entry In state("Criteria")
            Err.Clear
            If entry.Exists("Criteria1") And entry.Exists("Criteria2") Then
                table.Range.AutoFilter _
                    Field:=entry("Field"), _
                    Criteria1:=entry("Criteria1"), _
                    Operator:=entry("Operator"), _
                    Criteria2:=entry("Criteria2")
            ElseIf entry.Exists("Criteria2") Then
                table.Range.AutoFilter _
                    Field:=entry("Field"), _
                    Operator:=entry("Operator"), _
                    Criteria2:=entry("Criteria2")
            ElseIf entry("Operator") <> 0 Then
                table.Range.AutoFilter _
                    Field:=entry("Field"), _
                    Criteria1:=entry("Criteria1"), _
                    Operator:=entry("Operator")
            Else
                table.Range.AutoFilter _
                    Field:=entry("Field"), _
                    Criteria1:=entry("Criteria1")
            End If
            If Err.Number <> 0 Then
                fn_Restore = fn_Restore & " " & table.Name & ": " & Err.Description
            End If
        Next entry
    Next state
    On Error GoTo 0
End Function

' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_CaptureFilterCriterion( _
    ByVal filter As Object, _
    ByVal name As String, _
    ByVal entry As Object _
)
    ' Excel raises 1004 for an absent criterion, including Criteria1 in grouped date filters.
    On Error GoTo NoCriterion
    If IsObject(CallByName(filter, name, VbGet)) Then
        Set entry(name) = CallByName(filter, name, VbGet)
    Else
        entry(name) = CallByName(filter, name, VbGet)
    End If
NoCriterion:
End Sub

' --------------------------------------
' } // namespace Private
' --------------------------------------