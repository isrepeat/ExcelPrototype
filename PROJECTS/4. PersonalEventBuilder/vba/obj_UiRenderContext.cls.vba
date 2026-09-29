VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiRenderContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_targetWorksheet As Worksheet
Private m_uiPageDefinition As obj_UiPageDefinition
Private m_uiFolderPath As String
Private m_controls As Collection

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize( _
    ByVal targetWorksheet As Worksheet, _
    ByVal uiPageDefinition As obj_UiPageDefinition, _
    ByVal uiFolderPath As String _
) As Boolean
    If targetWorksheet Is Nothing Then Exit Function
    If uiPageDefinition Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(uiFolderPath)) = 0 Then Exit Function

    Set m_targetWorksheet = targetWorksheet
    Set m_uiPageDefinition = uiPageDefinition
    m_uiFolderPath = uiFolderPath
    Set m_controls = New Collection
    fn_Initialize = True
End Function

Public Sub fn_AddControl(ByVal uiControl As Object)
    If uiControl Is Nothing Then Exit Sub
    If m_controls Is Nothing Then Set m_controls = New Collection
    m_controls.Add uiControl
End Sub

Public Property Get fn_TargetWorksheet() As Worksheet
    Set fn_TargetWorksheet = m_targetWorksheet
End Property

Public Property Get fn_PageDefinition() As obj_UiPageDefinition
    Set fn_PageDefinition = m_uiPageDefinition
End Property

Public Property Get fn_UiFolderPath() As String
    fn_UiFolderPath = m_uiFolderPath
End Property

Public Sub fn_Dispose()
    Dim uiControl As Object

    If Not m_controls Is Nothing Then
        For Each uiControl In m_controls
            uiControl.fn_Dispose
        Next uiControl
    End If
    Set m_controls = Nothing
    Set m_uiPageDefinition = Nothing
    Set m_targetWorksheet = Nothing
    m_uiFolderPath = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------