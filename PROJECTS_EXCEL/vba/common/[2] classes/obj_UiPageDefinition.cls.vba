VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiPageDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_document As Object
Private m_xamlPath As String
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Document() As Object
    Set Document = m_document
End Property

Public Property Get XamlPath() As String
    XamlPath = m_xamlPath
End Property

' //
' // API
' //
Public Function Initialize(ByVal document As Object, ByVal xamlPath As String) As Boolean
    m_isDisposed = False
    If document Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(xamlPath)) = 0 Then Exit Function

    Set m_document = document
    m_xamlPath = xamlPath
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    Set m_document = Nothing
    m_xamlPath = VBA.vbNullString
End Sub