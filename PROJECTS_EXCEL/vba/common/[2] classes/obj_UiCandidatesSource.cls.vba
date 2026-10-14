VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiCandidatesSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiTableSource
Implements obj_IUiSelectionSource

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_table As obj_UiRawTable
Private m_selectionSource As obj_IUiSelectionSource

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
Private Property Get obj_IUiTableSource_TableCount() As Long
    If m_isInitialized And Not m_isDisposed Then obj_IUiTableSource_TableCount = 1
End Property

' //
' // Interface
' //
Private Function obj_IUiTableSource_GetTable(ByVal index As Long) As Object
    If Not m_isInitialized Or m_isDisposed Or index <> 1 Then Exit Function
    Set obj_IUiTableSource_GetTable = m_table
End Function

Private Function obj_IUiSelectionSource_GetItemAt(ByVal rowIndex As Long) As Object
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If rowIndex < 1 Or rowIndex > m_table.RowCount Then Exit Function
    Set obj_IUiSelectionSource_GetItemAt = m_selectionSource.GetItemAt(rowIndex)
End Function

' //
' // API
' //
Public Function Initialize( _
    ByVal table As obj_UiRawTable, _
    ByVal selectionSource As obj_IUiSelectionSource _
) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    If table Is Nothing Or selectionSource Is Nothing Then Exit Function
    Set m_table = table
    Set m_selectionSource = selectionSource
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_table Is Nothing Then m_table.Dispose
    Set m_table = Nothing
    Set m_selectionSource = Nothing
End Sub