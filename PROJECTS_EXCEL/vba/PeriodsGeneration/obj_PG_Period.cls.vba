VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_Period"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_Rank As String
Private m_PersonName As String
Private m_taxId As String
Private m_Position As String
Private m_EventName As String
Private m_LastEventName As String
Private m_HasDisplayableEvent As Boolean
Private m_IsOpen As Boolean
Private m_DepartureOrder As String
Private m_dateFrom As Date
Private m_dateTo As Date
Private m_ArrivalOrder As String

' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Свойства
' //
Public Property Get Rank() As String
    private_EnsureReady
    Rank = m_Rank
End Property

Public Property Let Rank(ByVal value As String)
    private_EnsureReady
    m_Rank = value
End Property

Public Property Get PersonName() As String
    private_EnsureReady
    PersonName = m_PersonName
End Property

Public Property Let PersonName(ByVal value As String)
    private_EnsureReady
    m_PersonName = value
End Property

Public Property Get taxId() As String
    private_EnsureReady
    taxId = m_taxId
End Property

Public Property Let taxId(ByVal value As String)
    private_EnsureReady
    m_taxId = value
End Property

Public Property Get Position() As String
    private_EnsureReady
    Position = m_Position
End Property

Public Property Let Position(ByVal value As String)
    private_EnsureReady
    m_Position = value
End Property

Public Property Get EventName() As String
    private_EnsureReady
    EventName = m_EventName
End Property

Public Property Let EventName(ByVal value As String)
    private_EnsureReady
    m_EventName = value
End Property

Public Property Get LastEventName() As String
    private_EnsureReady
    LastEventName = m_LastEventName
End Property

Public Property Let LastEventName(ByVal value As String)
    private_EnsureReady
    m_LastEventName = value
End Property

Public Property Get HasDisplayableEvent() As Boolean
    private_EnsureReady
    HasDisplayableEvent = m_HasDisplayableEvent
End Property

Public Property Let HasDisplayableEvent(ByVal value As Boolean)
    private_EnsureReady
    m_HasDisplayableEvent = value
End Property

Public Property Get IsOpen() As Boolean
    private_EnsureReady
    IsOpen = m_IsOpen
End Property

Public Property Let IsOpen(ByVal value As Boolean)
    private_EnsureReady
    m_IsOpen = value
End Property

Public Property Get DepartureOrder() As String
    private_EnsureReady
    DepartureOrder = m_DepartureOrder
End Property

Public Property Let DepartureOrder(ByVal value As String)
    private_EnsureReady
    m_DepartureOrder = value
End Property

Public Property Get dateFrom() As Date
    private_EnsureReady
    dateFrom = m_dateFrom
End Property

Public Property Let dateFrom(ByVal value As Date)
    private_EnsureReady
    m_dateFrom = value
End Property

Public Property Get dateTo() As Date
    private_EnsureReady
    dateTo = m_dateTo
End Property

Public Property Let dateTo(ByVal value As Date)
    private_EnsureReady
    m_dateTo = value
End Property

Public Property Get ArrivalOrder() As String
    private_EnsureReady
    ArrivalOrder = m_ArrivalOrder
End Property

Public Property Let ArrivalOrder(ByVal value As String)
    private_EnsureReady
    m_ArrivalOrder = value
End Property
' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
End Sub

' //
' // Вспомогательные методы
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_Period", "Service is not initialized."
    End If
End Sub