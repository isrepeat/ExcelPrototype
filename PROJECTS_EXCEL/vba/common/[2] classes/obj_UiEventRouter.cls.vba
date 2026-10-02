VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiEventRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_shapes As Object
Private m_cells As Object
Private m_seed As Long

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
Public Sub Initialize()
    Set m_shapes = VBA.CreateObject("Scripting.Dictionary")
    m_shapes.CompareMode = VBA.vbTextCompare
    Set m_cells = VBA.CreateObject("Scripting.Dictionary")
    m_cells.CompareMode = VBA.vbTextCompare
End Sub

Public Function NextId() As Long
    m_seed = m_seed + 1
    NextId = m_seed
End Function

Public Sub RegisterShape(ByVal name As String, ByVal handler As obj_IUiEventHandler)
    Set m_shapes(name) = handler
End Sub

Public Sub RegisterCell(ByVal cell As Range, ByVal handler As obj_IUiEventHandler)
    Dim key As String

    key = cell.Address(False, False)
    Set m_cells(key) = handler
End Sub

Public Function DispatchShape(ByVal name As String) As Boolean
    Dim handler As obj_IUiEventHandler

    If Not m_shapes.Exists(name) Then Exit Function
    Set handler = m_shapes(name)
    DispatchShape = handler.HandleEvent("click", name)
End Function

Public Function DispatchCells(ByVal target As Range) As Boolean
    Dim cell As Range
    Dim handler As obj_IUiEventHandler
    Dim payload As Variant

    For Each cell In target.Cells
        If m_cells.Exists(cell.Address(False, False)) Then
            Set handler = m_cells(cell.Address(False, False))
            Set payload = cell
            If Not handler.HandleEvent("change", payload) Then Exit Function
            DispatchCells = True
        End If
    Next cell
End Function

Public Sub UnregisterShape(ByVal name As String)
    If m_shapes.Exists(name) Then m_shapes.Remove name
End Sub

Public Sub Broadcast(ByVal kind As String)
    Dim key As Variant
    Dim handler As obj_IUiEventHandler

    For Each key In m_shapes.Keys
        Set handler = m_shapes(key)
        handler.HandleEvent kind, Empty
    Next key
End Sub

Public Sub Dispose()
    Set m_shapes = Nothing
    Set m_cells = Nothing
End Sub