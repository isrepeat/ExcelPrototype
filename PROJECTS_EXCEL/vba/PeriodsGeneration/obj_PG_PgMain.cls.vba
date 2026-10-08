VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PG_PgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IPage

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_controller As obj_PG_PgMainController

' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Интерфейс
' //
Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If

    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, VBA.Trim$(profileId)) Then
        Me.Dispose
        Exit Function
    End If

    Set m_controller = New obj_PG_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then
        Me.Dispose
        Exit Function
    End If

    obj_IPage_Initialize = True
    m_isInitialized = obj_IPage_Initialize
End Function

Private Function obj_IPage_Render() As Boolean
    Dim oldScreenUpdating As Boolean

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    obj_IPage_Render = m_pageBase.Render()
Cleanup:
    Application.ScreenUpdating = oldScreenUpdating
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If target Is Nothing Then
        Exit Function
    End If

    obj_IPage_HandleCellChange = True
End Function

Private Sub obj_IPage_Dispose()
    Me.Dispose
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
    If Not m_controller Is Nothing Then
        m_controller.Dispose
    End If
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then
        m_pageBase.Dispose
    End If
    Set m_pageBase = Nothing
End Sub