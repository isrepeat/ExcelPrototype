VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiCellBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_worksheetName As String
Private m_cellAddress As String
Private m_bindingContext As obj_UiBindingContext
Private m_sourceName As String
Private m_bindingPath As String
Private m_command As obj_UiCommand
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get WorksheetName() As String
    WorksheetName = m_worksheetName
End Property

Public Property Get CellAddress() As String
    CellAddress = m_cellAddress
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal worksheetName As String, _
    ByVal cellAddress As String, _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal sourceName As String, _
    ByVal bindingPath As String, _
    ByVal command As obj_UiCommand _
) As Boolean
    If VBA.Len(VBA.Trim$(worksheetName)) = 0 Or _
       VBA.Len(VBA.Trim$(cellAddress)) = 0 Then Exit Function
    If bindingContext Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(sourceName)) = 0 Or _
       VBA.Len(VBA.Trim$(bindingPath)) = 0 Then Exit Function
    m_worksheetName = worksheetName
    m_cellAddress = cellAddress
    Set m_bindingContext = bindingContext
    m_sourceName = sourceName
    m_bindingPath = bindingPath
    Set m_command = command
    m_isDisposed = False
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    Set m_command = Nothing
    Set m_bindingContext = Nothing
    m_worksheetName = VBA.vbNullString
    m_cellAddress = VBA.vbNullString
    m_sourceName = VBA.vbNullString
    m_bindingPath = VBA.vbNullString
End Sub

Public Function HandleCellChange(ByVal target As Range) As Boolean
    If m_isDisposed Or target Is Nothing Then Exit Function
    If VBA.StrComp(target.Parent.Name, m_worksheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.StrComp(target.Address(False, False), m_cellAddress, VBA.vbTextCompare) <> 0 Then Exit Function
    If Not m_bindingContext.TrySetPathValue( _
            m_sourceName, m_bindingPath, target.Value2) Then Exit Function
    If Not m_command Is Nothing Then
        If Not m_command.Execute() Then Exit Function
    End If
    HandleCellChange = True
End Function