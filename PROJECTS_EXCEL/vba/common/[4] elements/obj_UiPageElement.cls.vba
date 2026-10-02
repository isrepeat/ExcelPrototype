VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiPageElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElement
Implements obj_IUiContainer

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_grid As obj_IUiElement

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_grid = New obj_UiGridElement
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiElement_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    If m_isDisposed Or m_isInitialized Then
        diagnostic = "Element is already initialized or disposed."
        Exit Function
    End If
    obj_IUiElement_Configure = m_grid.Configure(definition, context, source, diagnostic)
    m_isInitialized = obj_IUiElement_Configure
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim element As obj_IUiElement

    Set element = m_grid
    obj_IUiElement_Measure = element.Measure(rows, columns, diagnostic)
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim element As obj_IUiElement

    Set element = m_grid
    obj_IUiElement_Arrange = element.Arrange(row, column, diagnostic)
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = m_grid.Render(diagnostic)
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = m_grid.Validate(errors)
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub

Private Sub obj_IUiContainer_AddChild(ByVal child As obj_IUiElement)
    Dim container As obj_IUiContainer

    Set container = m_grid
    container.AddChild child
End Sub

' //
' // API
' //
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_grid Is Nothing Then m_grid.Dispose
    Set m_grid = Nothing
End Sub