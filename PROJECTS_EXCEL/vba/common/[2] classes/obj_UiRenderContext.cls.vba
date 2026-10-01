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
Private m_uiBindingContext As obj_UiBindingContext

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get TargetWorksheet() As Worksheet
    Set TargetWorksheet = m_targetWorksheet
End Property

Public Property Get PageDefinition() As obj_UiPageDefinition
    Set PageDefinition = m_uiPageDefinition
End Property

Public Property Get BindingContext() As obj_UiBindingContext
    Set BindingContext = m_uiBindingContext
End Property

Public Property Get UiFolderPath() As String
    UiFolderPath = m_uiFolderPath
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal targetWorksheet As Worksheet, _
    ByVal uiPageDefinition As obj_UiPageDefinition, _
    ByVal uiFolderPath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    If targetWorksheet Is Nothing Or uiPageDefinition Is Nothing Or uiBindingContext Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(uiFolderPath)) = 0 Then Exit Function
    Set m_targetWorksheet = targetWorksheet
    Set m_uiPageDefinition = uiPageDefinition
    m_uiFolderPath = uiFolderPath
    Set m_uiBindingContext = uiBindingContext
    Set m_controls = New Collection
    Initialize = True
End Function

Public Sub Dispose()
    Dim uiControl As Object
    For Each uiControl In m_controls
        uiControl.Dispose
    Next uiControl
    Set m_controls = Nothing
    Set m_uiBindingContext = Nothing
    Set m_uiPageDefinition = Nothing
    Set m_targetWorksheet = Nothing
    m_uiFolderPath = VBA.vbNullString
End Sub

Public Sub AddControl(ByVal uiControl As Object)
    If uiControl Is Nothing Then Exit Sub
    m_controls.Add uiControl
End Sub