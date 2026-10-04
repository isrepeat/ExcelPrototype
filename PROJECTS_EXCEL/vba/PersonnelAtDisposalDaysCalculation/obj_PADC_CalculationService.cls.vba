VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_CalculationService"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private COUNTED_DAYS_THRESHOLD As Long
Private PRESENT_EVENT_NAME As String
Private m_periodRangeSeparator As String
Private Const EVENT_NAMES_SEPARATOR As String = " / "
Private Const PERIOD_LABEL_OPEN As String = " ("
Private Const PERIOD_LABEL_CLOSE As String = ")"
Private Const PERIODS_SEPARATOR As String = " | "
Private Const DICTIONARY_PROG_ID As String = "Scripting.Dictionary"
Private Const FORMAT_DATE As String = "dd.mm.yyyy"

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
    ByVal thresholdDays As Long, _
    ByVal presentEventName As String, _
    ByVal periodRangeSeparator As String _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If thresholdDays < 1 Or VBA.Len(presentEventName) = 0 Or VBA.Len(periodRangeSeparator) = 0 Then
        Exit Function
    End If
    COUNTED_DAYS_THRESHOLD = thresholdDays
    PRESENT_EVENT_NAME = presentEventName
    m_periodRangeSeparator = periodRangeSeparator
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

Public Function Calculate( _
    ByVal rows As Collection, _
    ByVal firstDay As Long, _
    ByVal lastDay As Long, _
    ByRef periodsText As String, _
    ByRef thresholdDate As Variant, _
    ByVal runContext As obj_PADC_RunContext _
) As Long
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "PADC.Calculate", "Calculation service is not initialized."
    End If
    If runContext Is Nothing Then
        VBA.Err.Raise vbObjectError + 2100, "PADC.Calculate", "Run context is required."
    End If
    Calculate = private_BuildCountedPeriods(rows, firstDay, lastDay, periodsText, thresholdDate, runContext)
End Function

' //
' // Private
' //
Private Function private_BuildCountedPeriods( _
    ByVal rows As Collection, _
    ByVal firstDay As Long, _
    ByVal lastDay As Long, _
    ByRef periodsText As String, _
    ByRef thresholdDate As Variant, _
    ByVal runContext As obj_PADC_RunContext _
) As Long
    Dim boundaries() As Long
    Dim unused() As Long
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim item As Variant
    Dim segmentStart As Long
    Dim segmentEnd As Long
    Dim pendingStart As Long
    Dim pendingEnd As Long
    Dim label As String
    Dim pendingLabel As String
    Dim excluded As Boolean
    Dim names As Object
    Dim key As Variant

    periodsText = vbNullString
    thresholdDate = Empty
    If firstDay >= lastDay Then
        Exit Function
    End If
    count = rows.Count * 2 + 2
    ReDim boundaries(1 To count)
    ReDim unused(1 To count)
    boundaries(1) = firstDay
    boundaries(2) = lastDay
    For i = 1 To rows.Count
        private_CheckCancel runContext, i
        item = rows(i)
        boundaries(i * 2 + 1) = item(0)
        boundaries(i * 2 + 2) = item(1)
    Next i
    private_SortIntervals boundaries, unused, 1, count, runContext
    For i = 1 To count - 1
        private_CheckCancel runContext, i
        segmentStart = boundaries(i)
        segmentEnd = boundaries(i + 1)
        If segmentStart >= segmentEnd Then
            GoTo NextSegment
        End If
        excluded = False
        Set names = VBA.CreateObject(DICTIONARY_PROG_ID)
        names.CompareMode = vbTextCompare
        For j = 1 To rows.Count
            private_CheckCancel runContext, j
            item = rows(j)
            If item(0) < segmentEnd And item(1) > segmentStart Then
                If item(3) Then
                    excluded = True
                    Exit For
                End If
                names(item(2)) = True
            End If
        Next j
        If Not excluded Then
            label = vbNullString
            For Each key In names.Keys
                If VBA.Len(label) > 0 Then
                    label = label & EVENT_NAMES_SEPARATOR
                End If
                label = label & VBA.CStr(key)
            Next key
            If names.Count = 0 Then
                label = PRESENT_EVENT_NAME
            End If
            If VBA.IsEmpty(thresholdDate) Then
                If private_BuildCountedPeriods + segmentEnd - segmentStart > COUNTED_DAYS_THRESHOLD Then
                    thresholdDate = VBA.CDate(segmentStart + COUNTED_DAYS_THRESHOLD - private_BuildCountedPeriods)
                End If
            End If
            private_BuildCountedPeriods = private_BuildCountedPeriods + segmentEnd - segmentStart
            If pendingEnd = segmentStart And pendingLabel = label Then
                pendingEnd = segmentEnd
            Else
                If VBA.Len(pendingLabel) > 0 Then
                    private_AppendCountedPeriod periodsText, pendingStart, pendingEnd, pendingLabel
                End If
                pendingStart = segmentStart
                pendingEnd = segmentEnd
                pendingLabel = label
            End If
        ElseIf VBA.Len(pendingLabel) > 0 Then
            private_AppendCountedPeriod periodsText, pendingStart, pendingEnd, pendingLabel
            pendingLabel = vbNullString
        End If
NextSegment:
    Next i
    If VBA.Len(pendingLabel) > 0 Then
        private_AppendCountedPeriod periodsText, pendingStart, pendingEnd, pendingLabel
    End If
End Function

Private Sub private_AppendCountedPeriod( _
    ByRef periodsText As String, _
    ByVal firstDay As Long, _
    ByVal endExclusive As Long, _
    ByVal eventName As String _
)
    If VBA.Len(periodsText) > 0 Then
        periodsText = periodsText & PERIODS_SEPARATOR
    End If
    periodsText = periodsText & eventName & PERIOD_LABEL_OPEN & _
        VBA.Format$(VBA.CDate(firstDay), FORMAT_DATE) & m_periodRangeSeparator & _
        VBA.Format$(VBA.CDate(endExclusive - 1), FORMAT_DATE) & PERIOD_LABEL_CLOSE
End Sub

Private Sub private_SortIntervals( _
    ByRef a() As Long, _
    ByRef b() As Long, _
    ByVal low As Long, _
    ByVal high As Long, _
    ByVal runContext As obj_PADC_RunContext _
)
    Dim i As Long
    Dim j As Long
    Dim pivot As Long
    Dim temp As Long

    i = low
    j = high
    pivot = a(low + (high - low) \ 2)
    Do While i <= j
        private_CheckCancel runContext, i
        Do While a(i) < pivot
            i = i + 1
            private_CheckCancel runContext, i
        Loop
        Do While a(j) > pivot
            j = j - 1
            private_CheckCancel runContext, j
        Loop
        If i <= j Then
            temp = a(i)
            a(i) = a(j)
            a(j) = temp
            temp = b(i)
            b(i) = b(j)
            b(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then
        private_SortIntervals a, b, low, j, runContext
    End If
    If i < high Then
        private_SortIntervals a, b, i, high, runContext
    End If
End Sub

Private Sub private_CheckCancel( _
    ByVal runContext As obj_PADC_RunContext, _
    ByVal index As Long _
)
    runContext.CheckCancel index
End Sub