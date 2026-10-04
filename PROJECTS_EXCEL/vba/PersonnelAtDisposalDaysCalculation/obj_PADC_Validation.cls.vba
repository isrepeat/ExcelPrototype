VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_Validation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PADC_Configuration
Private Const FORMAT_DATE As String = "dd.mm.yyyy"
Private Const DATE_TEXT_PATTERN As String = "##.##.####"
Private Const DATE_1904_OFFSET As Long = 1462
Private Const MIN_DATE_SERIAL As Long = 61
Private Const MAX_DATE_SERIAL As Long = 2958465
Private Const NBSP_CODE As Long = 160

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
Public Function Initialize( _
    ByVal configuration As obj_PADC_Configuration _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_configuration = Nothing
End Sub

Public Function RequiredText( _
    ByVal value As Variant, _
    ByVal context As String _
) As String
    private_EnsureReady
    If VBA.IsError(value) Or VBA.IsNull(value) Then
        m_configuration.Fail context & m_configuration.GetText("legacy.MSG_CELL_ERROR")
    End If
    RequiredText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
    If VBA.Len(RequiredText) = 0 Then
        m_configuration.Fail context & m_configuration.GetText("legacy.MSG_REQUIRED_VALUE")
    End If
End Function

Public Function ReadFlag( _
    ByVal value As Variant, _
    ByVal context As String _
) As Boolean
    Dim flagText As String

    private_EnsureReady
    If Not VBA.IsError(value) And Not VBA.IsNull(value) Then
        If VBA.Len(VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))) = 0 Then
            ReadFlag = False
            Exit Function
        End If
    End If
    flagText = VBA.LCase$(Me.RequiredText(value, context & m_configuration.GetText("legacy.MSG_SUBTRACT_BUSINESS_TRIPS")))
    If private_IsEventInList(flagText, m_configuration.GetText("flag.includeTrips")) Then
        ReadFlag = False
    ElseIf private_IsEventInList(flagText, m_configuration.GetText("flag.excludeTrips")) Then
        ReadFlag = True
    Else
        m_configuration.Fail context & m_configuration.GetText("legacy.MSG_INVALID_TRIP_FLAG")
    End If
End Function

Public Function ReadDay( _
    ByVal value As Variant, _
    ByVal date1904 As Boolean, _
    ByVal context As String _
) As Long
    Dim text As String
    Dim parsed As Date
    Dim serial As Double

    private_EnsureReady
    On Error GoTo InvalidDate
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        GoTo InvalidDate
    End If
    If VBA.VarType(value) = vbString Then
        text = VBA.Trim$(VBA.CStr(value))
        If Not text Like DATE_TEXT_PATTERN Then
            GoTo InvalidDate
        End If
        parsed = VBA.DateSerial(VBA.CInt(VBA.Right$(text, 4)), VBA.CInt(VBA.Mid$(text, 4, 2)), VBA.CInt(VBA.Left$(text, 2)))
        If VBA.Format$(parsed, FORMAT_DATE) <> text Then
            GoTo InvalidDate
        End If
        serial = VBA.CDbl(parsed)
    Else
        If VBA.VarType(value) = vbBoolean Or Not VBA.IsNumeric(value) Then
            GoTo InvalidDate
        End If
        serial = VBA.Int(VBA.CDbl(value))
        If date1904 Then
            serial = serial + DATE_1904_OFFSET
        End If
    End If
    If serial < MIN_DATE_SERIAL Or serial > MAX_DATE_SERIAL Then
        GoTo InvalidDate
    End If
    ReadDay = VBA.CLng(serial)
    Exit Function
InvalidDate:
    m_configuration.Fail context & m_configuration.GetText("legacy.MSG_INVALID_DATE")
End Function

Public Function IsExcluded( _
    ByVal eventName As String, _
    ByVal subtractTrips As Boolean _
) As Boolean
    private_EnsureReady
    IsExcluded = private_IsEventInList(eventName, m_configuration.GetText("event.alwaysExcluded"))
    If Not IsExcluded And subtractTrips Then
        IsExcluded = private_IsEventInList(eventName, m_configuration.GetText("event.optionalExcluded"))
    End If
End Function

Public Function MatchText(ByVal value As Variant) As String
    private_EnsureReady
    If VBA.IsError(value) Or VBA.IsNull(value) Then
        Exit Function
    End If
    MatchText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
End Function

' //
' // Private
' //
Private Function private_IsEventInList( _
    ByVal eventName As String, _
    ByVal eventList As String _
) As Boolean
    eventName = VBA.Trim$(eventName)
    If VBA.Len(eventName) = 0 Then
        Exit Function
    End If
    private_IsEventInList = VBA.InStr(1, "|" & eventList & "|", "|" & eventName & "|", vbTextCompare) > 0
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_Validation", "Service is not initialized."
    End If
End Sub