VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_QueryOrder"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Public ColumnName As String
Public ValueType As en_QueryType
Public Descending As Boolean
Public NormalizeText As Boolean
Public CaseSensitive As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    ValueType = QueryText
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
    ColumnName = VBA.vbNullString
    ValueType = QueryText
    Descending = False
    NormalizeText = False
    CaseSensitive = False
End Sub