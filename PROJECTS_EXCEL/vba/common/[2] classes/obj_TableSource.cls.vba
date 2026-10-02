VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TableSource"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Public WorkbookPath As String
Public TableName As String
Public SheetName As String
Public RangeAddress As String

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
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    WorkbookPath = VBA.vbNullString
    TableName = VBA.vbNullString
    SheetName = VBA.vbNullString
    RangeAddress = VBA.vbNullString
End Sub

Public Function TryValidate(ByRef diagnostic As String) As Boolean
    If m_isDisposed Or Not m_isInitialized Then
        diagnostic = "Object is not initialized or is disposed."
        Exit Function
    End If
    diagnostic = VBA.vbNullString
    If VBA.Len(VBA.Trim$(WorkbookPath)) = 0 Then
        diagnostic = "WorkbookPath is required."
    ElseIf VBA.Len(TableName) > 0 And VBA.Len(RangeAddress) > 0 Then
        diagnostic = "Specify TableName or RangeAddress, not both."
    ElseIf VBA.Len(TableName) = 0 And (VBA.Len(SheetName) = 0 Or VBA.Len(RangeAddress) = 0) Then
        diagnostic = "Specify TableName or SheetName and RangeAddress."
    Else
        TryValidate = True
    End If
End Function