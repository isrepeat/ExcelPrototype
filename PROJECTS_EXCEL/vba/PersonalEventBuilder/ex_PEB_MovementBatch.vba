Attribute VB_Name = "ex_PEB_MovementBatch"
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
Public Function fn_ExportOrder( _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByRef diagnostic As String, _
    Optional ByVal choices As Collection = Nothing, _
    Optional ByRef fieldConflicts As Collection, _
    Optional ByRef exportDiff As Object _
) As Boolean
    Dim target As ListObject
    Dim plan As Object
    Dim action As Object
    Dim snapshotStarted As Boolean
    Dim applied As Long
    Dim skipped As Long
    Dim successTemplate As String
    Dim skippedTemplate As String
    Dim warning As String
    Dim stage As String
    Dim startedAt As Double
    Dim totalStartedAt As Double

    On Error GoTo EH
    totalStartedAt = VBA.Timer
    diagnostic = VBA.vbNullString
    Set exportDiff = Nothing
    ex_PEB_Export.fn_BeginConfigSnapshot
    snapshotStarted = True
    private_Stage "Export.ResolveTarget", stage, startedAt
    If ex_PEB_Export.fn_Config("export.Movement.enabled") <> "true" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ExportMovementUnavailable")
    End If
    Set target = private_Target()
    ex_PEB_MovementConflicts.fn_Begin target.Parent.Parent.FullName, choices
    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_TARGET | FullName=" & target.Parent.Parent.FullName & _
        " | Path=" & target.Parent.Parent.Path & " | Sheet=" & target.Parent.Name & " | Table=" & target.Name & _
        " | Rows=" & target.ListRows.Count & " | Columns=" & target.ListColumns.Count
    successTemplate = ex_PEB_Export.fn_Config("text.MovementBatchSuccess")
    skippedTemplate = ex_PEB_Export.fn_Config("text.MovementBatchAlreadyApplied")
    private_ValidateTarget target
    private_Stage "Export.BuildPlan", stage, startedAt
    If Not ex_PEB_MovementPlanner.fn_BuildOrder(orderNumber, orderDate, target, plan, diagnostic) Then
        GoTo Cleanup
    End If
    If Not plan("IsReady") Then
        diagnostic = private_Conflicts(plan)
        GoTo Cleanup
    End If
    private_Stage "Export.Apply", stage, startedAt
    If Not ex_PEB_MovementWriter.fn_Apply(plan, target, diagnostic, exportDiff) Then
        GoTo Cleanup
    End If
    exportDiff("Summary") = VBA.Replace$(exportDiff("Summary"), "{order}", orderNumber)
    exportDiff("Summary") = VBA.Replace$(exportDiff("Summary"), "{date}", _
        VBA.Format$(ex_PEB_History.fn_Date(orderDate), ex_PEB_Export.fn_Config("export.Movement.dateFormat")))
    warning = diagnostic
    For Each action In plan("Actions")
        If action("Status") = "AlreadyApplied" Then
            skipped = skipped + 1
        Else
            applied = applied + 1
        End If
    Next action
    diagnostic = successTemplate
    If applied = 0 Then
        diagnostic = skippedTemplate
    End If
    diagnostic = VBA.Replace$(diagnostic, "{applied}", VBA.CStr(applied))
    diagnostic = VBA.Replace$(diagnostic, "{skipped}", VBA.CStr(skipped))
    If VBA.Len(warning) > 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & warning
    End If
    fn_ExportOrder = True
Cleanup:
    Set fieldConflicts = ex_PEB_MovementConflicts.fn_Items()
    ex_PEB_MovementConflicts.fn_Module_Dispose
    private_Stage VBA.vbNullString, stage, startedAt
    ex_Core.fn_Diagnostic_WritePerf "Movement.Export.TotalIncludingCalculation", totalStartedAt
    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_EXPORT_COMPLETED | Success=" & VBA.CStr(fn_ExportOrder) & " | Diagnostic=" & diagnostic
    If Not ex_Core.fn_Diagnostic_Flush(True) Then
        Debug.Print "Movement diagnostic flush failed."
    End If
    If snapshotStarted Then
        ex_PEB_Export.fn_EndConfigSnapshot
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Movement." & stage, VBA.Err.Number, VBA.Err.Source, diagnostic
    Resume Cleanup
End Function

' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Target() As ListObject
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim selected As ListObject
    Dim sheetName As String
    Dim tableName As String
    Dim message As String

    sheetName = ex_PEB_Export.fn_Config("export.Movement.sheet")
    tableName = ex_PEB_Export.fn_Config("export.Movement.table")
    For Each book In Application.Workbooks
        For Each sheet In book.Worksheets
            If VBA.StrComp(sheet.Name, sheetName, VBA.vbTextCompare) = 0 Then
                For Each table In sheet.ListObjects
                    If VBA.StrComp(table.Name, tableName, VBA.vbTextCompare) = 0 Then
                        If Not selected Is Nothing Then
                            message = ex_PEB_Export.fn_Config("text.MovementBatchTargetAmbiguous")
                            message = VBA.Replace$(message, "{sheet}", sheetName)
                            Err.Raise 5, , VBA.Replace$(message, "{table}", tableName)
                        End If
                        Set selected = table
                    End If
                Next table
            End If
        Next sheet
    Next book
    If selected Is Nothing Then
        message = ex_PEB_Export.fn_Config("text.MovementBatchTargetMissing")
        message = VBA.Replace$(message, "{sheet}", sheetName)
        Err.Raise 5, , VBA.Replace$(message, "{table}", tableName)
    End If
    Set private_Target = selected
End Function

Private Sub private_ValidateTarget(ByVal target As ListObject)
    Dim book As Workbook
    Dim selected As ListObject

    If target Is Nothing Then
        ex_PEB_MovementLinks.fn_Fail "MissingTarget", ""
    End If
    Set book = target.Parent.Parent
    Set selected = private_Target()
    If Not target Is selected Then
        ex_PEB_MovementLinks.fn_Fail "TargetSelectionChanged", book.FullName
    End If
    If VBA.StrComp(target.Parent.Name, ex_PEB_Export.fn_Config("export.Movement.sheet"), VBA.vbTextCompare) <> 0 Then
        ex_PEB_MovementLinks.fn_Fail "WrongTargetWorksheet", target.Parent.Name
    End If
    If VBA.StrComp(target.Name, ex_PEB_Export.fn_Config("export.Movement.table"), VBA.vbTextCompare) <> 0 Then
        ex_PEB_MovementLinks.fn_Fail "WrongTargetTable", target.Name
    End If
    If book Is ThisWorkbook Then
        ex_PEB_MovementLinks.fn_Fail "TargetIsBuilderWorkbook", book.FullName
    End If
    If VBA.StrComp(book.FullName, ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path")), VBA.vbTextCompare) = 0 Then
        ex_PEB_MovementLinks.fn_Fail "TargetIsHistoryWorkbook", book.FullName
    End If
    If book.ReadOnly Then
        ex_PEB_MovementLinks.fn_Fail "ReadOnlyTarget", book.FullName
    End If
    If book.Date1904 Then
        ex_PEB_MovementLinks.fn_Fail "UnsupportedDateSystem", book.FullName
    End If
    If book.MultiUserEditing Then
        ex_PEB_MovementLinks.fn_Fail "SharedTargetUnsupported", book.FullName
    End If
    If target.Parent.ProtectContents Then
        ex_PEB_MovementLinks.fn_Fail "ProtectedTarget", target.Parent.Name
    End If
    If Not target.ShowHeaders Then
        ex_PEB_MovementLinks.fn_Fail "HiddenTableHeadersUnsupported", target.Name
    End If
    If book.FileFormat <> 50 And book.FileFormat <> 51 And book.FileFormat <> 52 Then
        ex_PEB_MovementLinks.fn_Fail "UnsupportedTargetFormat", book.FullName
    End If
End Sub

Private Function private_Conflicts(ByVal plan As Object) As String
    Dim action As Object
    Dim text As String

    text = VBA.Replace$(ex_PEB_Export.fn_Config("text.MovementBatchConflicts"), "{count}", VBA.CStr(plan("ConflictCount")))
    For Each action In plan("Actions")
        If action("Status") = "Conflict" Then
            text = text & VBA.vbCrLf & action("RowId") & ": " & action("Diagnostic")
        End If
    Next action
    private_Conflicts = text
End Function

Private Sub private_Stage( _
    ByVal nextStage As String, _
    ByRef stage As String, _
    ByRef startedAt As Double _
)
    If VBA.Len(stage) > 0 Then
        ex_Core.fn_Diagnostic_WritePerf "Movement." & stage, startedAt
    End If
    stage = nextStage
    startedAt = VBA.Timer
    If VBA.Len(stage) > 0 Then
        ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_STAGE_STARTED | Stage=" & stage
    End If
    If Not ex_Core.fn_Diagnostic_Flush(True) Then
        Debug.Print "Movement diagnostic flush failed."
    End If
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------