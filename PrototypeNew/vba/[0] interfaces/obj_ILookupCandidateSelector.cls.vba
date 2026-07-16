VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ILookupCandidateSelector"
Option Explicit

' Selects one extension row from rows whose join key already exactly matched.
' outSelectedRowIndex = 0 means that no extension row should be merged.
Public Function TrySelectCandidateRow( _
    ByVal candidateTable As obj_TableDynamic, _
    ByVal candidateRow As obj_Row, _
    ByVal extensionTable As obj_TableDynamic, _
    ByVal matchingRowIndexes As Collection, _
    ByRef outSelectedRowIndex As Long _
) As Boolean
End Function
