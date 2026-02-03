Attribute VB_Name = "ex_Run"
Option Explicit

Public Sub RunQuery()
    Dim tableData As Variant
    
    tableData = ex_QueryRunner.GetResultTable()
    ex_ResultWriter.WriteTableToResultSheet tableData
End Sub
