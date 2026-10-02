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

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_document As Object
Private m_xamlPath As String

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
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If document Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(xamlPath)) = 0 Then Exit Function

    Set m_document = document
    m_xamlPath = xamlPath
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_document = Nothing
    m_xamlPath = VBA.vbNullString
End Sub