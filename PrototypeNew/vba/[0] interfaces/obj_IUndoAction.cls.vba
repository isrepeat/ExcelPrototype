VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IUndoAction"
Option Explicit

Public Function GetActionId() As String
End Function

Public Function GetCaption() As String
End Function

Public Function GetScopeKey() As String
End Function

Public Function IsValid() As Boolean
End Function

Public Function Execute(ByRef outErrorText As String) As Boolean
End Function

Public Function Undo(ByRef outErrorText As String) As Boolean
End Function

Public Function Redo(ByRef outErrorText As String) As Boolean
End Function
