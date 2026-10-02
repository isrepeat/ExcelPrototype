VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiMarkupDiagnostic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_path As String
Private m_member As String
Private m_message As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Path() As String
    Path = m_path
End Property

Public Property Get Member() As String
    Member = m_member
End Property

Public Property Get Message() As String
    Message = m_message
End Property

' //
' // API
' //
Public Sub Initialize(ByVal path As String, ByVal member As String, ByVal message As String)
    m_path = path
    m_member = member
    m_message = message
End Sub

Public Function Describe() As String
    Describe = m_path & ": " & m_message
    If VBA.Len(m_member) > 0 Then Describe = Describe & " [" & m_member & "]"
End Function

Public Sub Dispose()
    m_path = VBA.vbNullString
    m_member = VBA.vbNullString
    m_message = VBA.vbNullString
End Sub