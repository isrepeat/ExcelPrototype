VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PADC_CalculationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_alwaysExcludedEvents As Object
Private m_optionalExcludedEvents As Object
Private m_thresholdDays As Long

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
    ByVal alwaysExcludedEvents As Collection, _
    ByVal optionalExcludedEvents As Collection, _
    ByVal thresholdDays As Long _
) As Boolean
    Dim eventName As Variant

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If alwaysExcludedEvents Is Nothing Or optionalExcludedEvents Is Nothing Then
        Exit Function
    End If
    If thresholdDays < 1 Then
        Exit Function
    End If

    Set m_alwaysExcludedEvents = VBA.CreateObject("Scripting.Dictionary")
    m_alwaysExcludedEvents.CompareMode = VBA.vbTextCompare
    For Each eventName In alwaysExcludedEvents
        m_alwaysExcludedEvents(VBA.Trim$(VBA.CStr(eventName))) = True
    Next eventName
    Set m_optionalExcludedEvents = VBA.CreateObject("Scripting.Dictionary")
    m_optionalExcludedEvents.CompareMode = VBA.vbTextCompare
    For Each eventName In optionalExcludedEvents
        m_optionalExcludedEvents(VBA.Trim$(VBA.CStr(eventName))) = True
    Next eventName
    m_thresholdDays = thresholdDays

    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If

    m_isDisposed = True
    m_isInitialized = False
    Set m_alwaysExcludedEvents = Nothing
    Set m_optionalExcludedEvents = Nothing
    m_thresholdDays = 0
End Sub

Public Function TryCalculate( _
    ByVal eventRows As Collection, _
    ByVal firstDay As Long, _
    ByVal lastDay As Long, _
    ByVal includeBusinessTrips As Boolean, _
    ByRef countedDays As Long, _
    ByRef periodsText As String, _
    ByRef thresholdDate As Variant, _
    ByRef diagnostic As String _
) As Boolean
    Dim normalizedRows As Collection
    Dim eventRow As Variant
    Dim eventName As String
    Dim isExcluded As Boolean

    diagnostic = vbNullString
    countedDays = 0
    periodsText = vbNullString
    thresholdDate = Empty
    If Not m_isInitialized Or m_isDisposed Then
        diagnostic = "The calculation service is not initialized."
        Exit Function
    End If
    If eventRows Is Nothing Then
        diagnostic = "Event rows are required."
        Exit Function
    End If
    If firstDay > lastDay Then
        diagnostic = "The calculation start date is later than the end date."
        Exit Function
    End If

    Set normalizedRows = New Collection
    For Each eventRow In eventRows
        If Not private_TryNormalizeEventRow(eventRow, firstDay, lastDay, eventName, diagnostic) Then
            Exit Function
        End If
        If VBA.Len(eventName) > 0 Then
            isExcluded = private_IsExcluded(eventName, includeBusinessTrips)
            normalizedRows.Add VBA.Array(eventRow(0), eventRow(1), eventName, isExcluded)
        End If
    Next eventRow

    countedDays = private_BuildCountedPeriods(normalizedRows, firstDay, lastDay, periodsText, thresholdDate)
    TryCalculate = True
End Function

' //
' // Private
' //
Private Function private_TryNormalizeEventRow( _
    ByVal eventRow As Variant, _
    ByVal firstDay As Long, _
    ByVal lastDay As Long, _
    ByRef eventName As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim eventFirstDay As Long
    Dim eventLastDay As Long

    If Not VBA.IsArray(eventRow) Then
        diagnostic = "An event row must be an array."
        Exit Function
    End If
    If UBound(eventRow) - LBound(eventRow) + 1 <> 3 Then
        diagnostic = "An event row must contain a start date, end date, and event name."
        Exit Function
    End If
    If Not VBA.IsNumeric(eventRow(0)) Or Not VBA.IsNumeric(eventRow(1)) Then
        diagnostic = "An event row contains an invalid date."
        Exit Function
    End If

    eventFirstDay = VBA.CLng(eventRow(0))
    eventLastDay = VBA.CLng(eventRow(1))
    If eventLastDay < eventFirstDay Then
        diagnostic = "An event ends before it starts."
        Exit Function
    End If
    If eventFirstDay < firstDay Then
        eventFirstDay = firstDay
    End If
    If eventLastDay > lastDay Then
        eventLastDay = lastDay
    End If
    eventName = VBA.Trim$(VBA.CStr(eventRow(2)))
    eventRow(0) = eventFirstDay
    eventRow(1) = eventLastDay
    private_TryNormalizeEventRow = True
End Function

Private Function private_IsExcluded( _
    ByVal eventName As String, _
    ByVal includeBusinessTrips As Boolean _
) As Boolean
    private_IsExcluded = m_alwaysExcludedEvents.Exists(eventName)
    If Not private_IsExcluded And Not includeBusinessTrips Then
        private_IsExcluded = m_optionalExcludedEvents.Exists(eventName)
    End If
End Function

Private Function private_BuildCountedPeriods( _
    ByVal eventRows As Collection, _
    ByVal firstDay As Long, _
    ByVal lastDay As Long, _
    ByRef periodsText As String, _
    ByRef thresholdDate As Variant _
) As Long
    Dim boundaries() As Long
    Dim scratch() As Long
    Dim eventRow As Variant
    Dim activeNames As Object
    Dim activeName As Variant
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim segmentFirstDay As Long
    Dim segmentLastDay As Long
    Dim label As String
    Dim pendingFirstDay As Long
    Dim pendingLastDay As Long
    Dim pendingLabel As String
    Dim excluded As Boolean

    periodsText = vbNullString
    thresholdDate = Empty
    If firstDay >= lastDay Then
        Exit Function
    End If
    count = eventRows.Count * 2 + 2
    ReDim boundaries(1 To count)
    ReDim scratch(1 To count)
    boundaries(1) = firstDay
    boundaries(2) = lastDay
    For i = 1 To eventRows.Count
        eventRow = eventRows(i)
        boundaries(i * 2 + 1) = eventRow(0)
        boundaries(i * 2 + 2) = eventRow(1)
    Next i
    private_SortIntervals boundaries, scratch, 1, count

    For i = 1 To count - 1
        segmentFirstDay = boundaries(i)
        segmentLastDay = boundaries(i + 1)
        If segmentFirstDay >= segmentLastDay Then
            GoTo NextSegment
        End If
        excluded = False
        Set activeNames = VBA.CreateObject("Scripting.Dictionary")
        activeNames.CompareMode = VBA.vbTextCompare
        For j = 1 To eventRows.Count
            eventRow = eventRows(j)
            If eventRow(0) < segmentLastDay And eventRow(1) > segmentFirstDay Then
                If eventRow(3) Then
                    excluded = True
                    Exit For
                End If
                activeNames(eventRow(2)) = True
            End If
        Next j
        If excluded Then
            If VBA.Len(pendingLabel) > 0 Then
                private_AppendPeriod periodsText, pendingFirstDay, pendingLastDay, pendingLabel
                pendingLabel = vbNullString
            End If
            GoTo NextSegment
        End If
        label = vbNullString
        For Each activeName In activeNames.Keys
            If VBA.Len(label) > 0 Then
                label = label & " / "
            End If
            label = label & VBA.CStr(activeName)
        Next activeName
        If activeNames.Count = 0 Then
            label = "In unit"
        End If
        If VBA.IsEmpty(thresholdDate) Then
            If private_BuildCountedPeriods + segmentLastDay - segmentFirstDay > m_thresholdDays Then
                thresholdDate = VBA.CDate(segmentFirstDay + m_thresholdDays - private_BuildCountedPeriods)
            End If
        End If
        private_BuildCountedPeriods = private_BuildCountedPeriods + segmentLastDay - segmentFirstDay
        If pendingLastDay = segmentFirstDay And pendingLabel = label Then
            pendingLastDay = segmentLastDay
        Else
            If VBA.Len(pendingLabel) > 0 Then
                private_AppendPeriod periodsText, pendingFirstDay, pendingLastDay, pendingLabel
            End If
            pendingFirstDay = segmentFirstDay
            pendingLastDay = segmentLastDay
            pendingLabel = label
        End If
NextSegment:
    Next i
    If VBA.Len(pendingLabel) > 0 Then
        private_AppendPeriod periodsText, pendingFirstDay, pendingLastDay, pendingLabel
    End If
End Function

Private Sub private_AppendPeriod( _
    ByRef periodsText As String, _
    ByVal firstDay As Long, _
    ByVal endExclusive As Long, _
    ByVal eventName As String _
)
    If VBA.Len(periodsText) > 0 Then
        periodsText = periodsText & " | "
    End If
    periodsText = periodsText & eventName & " (" & _
        VBA.Format$(VBA.CDate(firstDay), "dd.mm.yyyy") & "–" & _
        VBA.Format$(VBA.CDate(endExclusive - 1), "dd.mm.yyyy") & ")"
End Sub

Private Sub private_SortIntervals( _
    ByRef starts() As Long, _
    ByRef ends() As Long, _
    ByVal low As Long, _
    ByVal high As Long _
)
    Dim i As Long
    Dim j As Long
    Dim pivot As Long
    Dim temporaryValue As Long

    i = low
    j = high
    pivot = starts(low + (high - low) \ 2)
    Do While i <= j
        Do While starts(i) < pivot
            i = i + 1
        Loop
        Do While starts(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temporaryValue = starts(i)
            starts(i) = starts(j)
            starts(j) = temporaryValue
            temporaryValue = ends(i)
            ends(i) = ends(j)
            ends(j) = temporaryValue
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then
        private_SortIntervals starts, ends, low, j
    End If
    If i < high Then
        private_SortIntervals starts, ends, i, high
    End If
End Sub