VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_configuration As obj_PG_Configuration
Private m_startDate As Date
Private m_endDate As Date

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

Public Property Get StartDate() As Date
    private_EnsureReady
    StartDate = m_startDate
End Property

Public Property Get EndDate() As Date
    private_EnsureReady
    EndDate = m_endDate
End Property

' //
' // API
' //

Public Function Initialize( _
    ByVal pageBase As obj_PageBase, _
    ByVal configuration As obj_PG_Configuration _
) As Boolean
    Dim value As Variant
    Dim valueObject As Object
    Dim isObject As Boolean

    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If pageBase Is Nothing Or configuration Is Nothing Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    Set m_configuration = configuration
    If Not m_pageBase.BindingContext.TryGetValue("Form", "StartDate", value, valueObject, isObject) Then
        GoTo Invalid
    End If
    If isObject Or Not private_TryDate(value, m_startDate) Then
        GoTo Invalid
    End If
    If Not m_pageBase.BindingContext.TryGetValue("Form", "EndDate", value, valueObject, isObject) Then
        GoTo Invalid
    End If
    If isObject Or Not private_TryDate(value, m_endDate) Then
        GoTo Invalid
    End If
    If m_startDate > m_endDate Then
        GoTo Invalid
    End If
    m_isInitialized = True
    Initialize = True
    Exit Function
Invalid:
    ex_WindowsUi.fn_ShowMessage m_configuration.GetText("message.InvalidPeriod"), vbExclamation
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_pageBase = Nothing
    Set m_configuration = Nothing
End Sub

' //
' // Вспомогательные методы
' //

Private Function private_TryDate( _
    ByVal value As Variant, _
    ByRef result As Date _
) As Boolean
    Dim parts As Variant
    Dim day As Long
    Dim month As Long
    Dim year As Long

    On Error GoTo Invalid
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        Exit Function
    End If
    If VarType(value) = vbDate Or IsNumeric(value) Then
        result = VBA.DateValue(VBA.CDate(value))
    Else
        parts = VBA.Split(VBA.Trim$(VBA.CStr(value)), ".")
        If UBound(parts) <> 2 Then
            Exit Function
        End If
        If Len(parts(0)) = 0 Or Len(parts(1)) = 0 Or Len(parts(2)) <> 4 Then
            Exit Function
        End If
        If Len(parts(0)) > 2 Or Len(parts(1)) > 2 Then
            Exit Function
        End If
        If parts(0) Like "*[!0-9]*" Or parts(1) Like "*[!0-9]*" Or parts(2) Like "*[!0-9]*" Then
            Exit Function
        End If
        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
            Exit Function
        End If
        day = VBA.CLng(parts(0))
        month = VBA.CLng(parts(1))
        year = VBA.CLng(parts(2))
        result = VBA.DateSerial(year, month, day)
        If VBA.Day(result) <> day Or VBA.Month(result) <> month Or VBA.Year(result) <> year Then
            Exit Function
        End If
    End If
    If result < VBA.DateSerial(1900, 3, 1) Then
        Exit Function
    End If
    private_TryDate = True
Invalid:
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_Parameters", "Service is not initialized."
    End If
End Sub