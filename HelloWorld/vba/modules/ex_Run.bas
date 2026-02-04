Attribute VB_Name = "ex_Run"
Option Explicit

Public Sub TestCompareOldNew()
    Dim resultTable As Variant
    
    ex_TestData.GenerateTestTables
    
    resultTable = ex_TableComparer.CompareSheets( _
        "Old", _
        "New", _
        "Id" _
    )
    
    ex_ResultWriter.WriteTableToResultSheet resultTable
End Sub
