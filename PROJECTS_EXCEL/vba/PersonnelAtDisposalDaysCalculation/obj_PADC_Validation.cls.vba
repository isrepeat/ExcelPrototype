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
        m_configuration.Fail context & m_configuration.GetText("text.CellError")
    End If
    RequiredText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
    If VBA.Len(RequiredText) = 0 Then
        m_configuration.Fail context & m_configuration.GetText("text.RequiredValue")
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
    If VBA.VarType(value) = vbDate Then
        serial = VBA.Int(VBA.CDbl(value))
    ElseIf VBA.VarType(value) = vbString Then
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
    m_configuration.Fail context & m_configuration.GetText("text.InvalidDate")
End Function

Public Function IsEligibleOpenStatus(ByVal openStatus As String) As Boolean
    private_EnsureReady
    If VBA.Len(openStatus) = 0 Then
        openStatus = "{blank}"
    End If
    IsEligibleOpenStatus = private_IsEventInList(openStatus, m_configuration.GetText("event.eligibleOpenStatuses"))
End Function

Public Function ResolveEventRules( _
    ByVal personnelCode As String, _
    ByVal openStatus As String _
) As Object
    Dim decisions As Object
    Dim owners As Object
    Dim ruleId As Variant
    Dim ruleName As String
    Dim ruleKey As String
    Dim statuses As String
    Dim operation As Variant
    Dim eventName As String
    Dim excluded As Boolean
    Dim diagnostic As String

    private_EnsureReady
    Set decisions = VBA.CreateObject("Scripting.Dictionary")
    Set owners = VBA.CreateObject("Scripting.Dictionary")
    decisions.CompareMode = vbTextCompare
    owners.CompareMode = vbTextCompare
    If VBA.Len(openStatus) = 0 Then
        openStatus = "{blank}"
    End If
    For Each ruleId In VBA.Split(m_configuration.GetText("event.rules.ids"), "|")
        ruleName = VBA.Trim$(VBA.CStr(ruleId))
        ruleKey = "event.rules." & ruleName
        If private_IsEventInList(personnelCode, m_configuration.GetText(ruleKey & ".codes")) Then
            statuses = m_configuration.GetText(ruleKey & ".openStatuses")
            If private_IsEventInList("*", statuses) Or private_IsEventInList(openStatus, statuses) Then
                For Each operation In VBA.Split(m_configuration.GetText(ruleKey & ".events"), "|")
                    operation = VBA.Trim$(VBA.CStr(operation))
                    eventName = VBA.Trim$(VBA.Mid$(operation, 2))
                    excluded = (VBA.Left$(operation, 1) = "-")
                    If decisions.Exists(eventName) Then
                        If decisions(eventName) <> excluded Then
                            diagnostic = m_configuration.GetText("text.EventRuleConflict")
                            diagnostic = VBA.Replace$(diagnostic, "{first}", VBA.CStr(owners(eventName)))
                            diagnostic = VBA.Replace$(diagnostic, "{second}", ruleName)
                            diagnostic = VBA.Replace$(diagnostic, "{event}", eventName)
                            m_configuration.Fail diagnostic
                        End If
                    Else
                        decisions.Add eventName, excluded
                        owners.Add eventName, ruleName
                    End If
                Next operation
            End If
        End If
    Next ruleId
    Set ResolveEventRules = decisions
End Function

Public Function DisplayOpenStatus( _
    ByVal openStatus As String, _
    ByVal departure As Long, _
    ByVal chronology As Collection _
) As String
    Dim item As Variant
    Dim previousDeparture As Long
    Dim allowedPredecessors As String
    Dim contiguous As Boolean
    Dim isCandidate As Boolean

    private_EnsureReady
    DisplayOpenStatus = openStatus
    If VBA.StrComp(openStatus, m_configuration.GetText("event.tripChain.target"), vbTextCompare) <> 0 Then
        Exit Function
    End If
    allowedPredecessors = m_configuration.GetText("event.tripChain.predecessors")
    For Each item In chronology
        If item(0) < departure And item(0) >= previousDeparture Then
            isCandidate = (item(1) = departure) And private_IsEventInList(VBA.CStr(item(2)), allowedPredecessors)
            If item(0) = previousDeparture Then
                contiguous = contiguous And isCandidate
            Else
                previousDeparture = item(0)
                contiguous = isCandidate
            End If
        End If
    Next item
    If contiguous Then
        DisplayOpenStatus = m_configuration.GetText("event.tripChain.label")
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
    Dim entry As Variant

    eventName = VBA.Trim$(eventName)
    If VBA.Len(eventName) = 0 Then
        Exit Function
    End If
    For Each entry In VBA.Split(eventList, "|")
        If VBA.StrComp(eventName, VBA.Trim$(VBA.CStr(entry)), vbTextCompare) = 0 Then
            private_IsEventInList = True
            Exit Function
        End If
    Next entry
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_Validation", "Service is not initialized."
    End If
End Sub