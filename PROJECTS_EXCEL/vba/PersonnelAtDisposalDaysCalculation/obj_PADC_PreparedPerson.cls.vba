VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_PreparedPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_fullName As String
Private m_taxId As String
Private m_startDay As Long
Private m_tripFlag As Variant
Private m_intervals As Collection
Private m_validationError As String

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
Public Property Get FullName() As String
    private_EnsureReady
    FullName = m_fullName
End Property

Public Property Get TaxId() As String
    private_EnsureReady
    TaxId = m_taxId
End Property

Public Property Get StartDay() As Long
    private_EnsureReady
    StartDay = m_startDay
End Property

Public Property Get TripFlag() As Variant
    private_EnsureReady
    TripFlag = m_tripFlag
End Property

Public Property Get Intervals() As Collection
    private_EnsureReady
    Set Intervals = m_intervals
End Property

Public Property Get ValidationError() As String
    private_EnsureReady
    ValidationError = m_validationError
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal fullName As String, _
    ByVal taxId As String, _
    ByVal startDay As Long, _
    ByVal tripFlag As Variant, _
    ByVal intervals As Collection, _
    ByVal validationError As String _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If intervals Is Nothing Then
        Exit Function
    End If
    m_fullName = fullName
    m_taxId = taxId
    m_startDay = startDay
    m_tripFlag = tripFlag
    Set m_intervals = intervals
    m_validationError = validationError
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_intervals = Nothing
End Sub

' //
' // Private
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_PreparedPerson", "Prepared person is not initialized."
    End If
End Sub