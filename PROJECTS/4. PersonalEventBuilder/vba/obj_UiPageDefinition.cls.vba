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

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize(ByVal document As Object, ByVal xamlPath As String) As Boolean
    If document Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(xamlPath)) = 0 Then Exit Function

    Set m_document = document
    m_xamlPath = xamlPath
    fn_Initialize = True
End Function

Public Property Get fn_Document() As Object
    Set fn_Document = m_document
End Property

Public Property Get fn_XamlPath() As String
    fn_XamlPath = m_xamlPath
End Property

Public Sub fn_Dispose()
    Set m_document = Nothing
    m_xamlPath = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------