VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_values As Object
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
' // Properties
' //
Public Property Get ThresholdDays() As Long
    private_EnsureReady
    ThresholdDays = m_thresholdDays
End Property

' //
' // API
' //
Public Function Initialize( _
    ByRef diagnostic As String, _
    Optional ByVal tvoMode As Boolean = False _
) As Boolean
    Dim suffix As Variant
    Dim key As String
    Dim value As String
    Dim requiredKeys As Collection
    Dim ruleId As Variant
    Dim field As Variant
    Dim ruleKey As String
    Dim operation As Variant
    Dim tvoKeys As Variant

    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    Set m_values = VBA.CreateObject("Scripting.Dictionary")
    If tvoMode Then
        Set requiredKeys = New Collection
        For Each suffix In VBA.Split("source.events.path|source.events.range|source.events.sheet|" & _
            "source.events.columns.TaxId|source.events.columns.FullName|source.events.columns.Event|" & _
            "source.events.columns.Departure|source.events.columns.Arrival|text.Cancelled|text.CellError|" & _
            "text.RequiredValue|text.InvalidDate|text.TargetPrepareFailed|text.SnapshotFailed|text.ExportEmpty|" & _
            "text.ExportFolderMissing|progress.Template|progress.Tvo|progress.Events|progress.EventIndex|progress.Calculation|progress.Publish|" & _
            "format.date|header.FullName|header.TaxId|header.Periods|tvo.text.AbsencePeriod|" & _
            "export.palette|export.font|export.fontSize|export.computedColor|export.styleId|export.styleName|export.styleSource|" & _
            "export.sheet|export.labels.Date|export.padding.rows|export.padding.columns|export.versionSuffix", "|")
            requiredKeys.Add VBA.CStr(suffix)
        Next suffix
    Else
        Set requiredKeys = private_RequiredKeys()
    End If
    For Each suffix In requiredKeys
        key = "PersonnelAtDisposalDaysCalculation::" & VBA.CStr(suffix)
        If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
            GoTo Failed
        End If
        m_values.Add VBA.CStr(suffix), value
    Next suffix
    If tvoMode Then
        tvoKeys = VBA.Split("source.tvo.sheet|source.tvo.range|source.tvo.columns.Rank|source.tvo.columns.FullName|" & _
            "source.tvo.columns.TaxId|source.tvo.columns.Position|source.tvo.columns.OwnPosition|source.tvo.columns.Duration|" & _
            "source.tvo.columns.StartOrder|source.tvo.columns.AccessOrder|" & _
            "source.tvo.columns.StartDate|source.tvo.columns.EndDate|tvo.days|tvo.title|tvo.fileName|tvo.widths|tvo.ignoredEvents|tvo.ignoredDuration|" & _
            "tvo.header.Rank|tvo.header.Position|tvo.header.OwnPosition|tvo.header.StartDate|tvo.header.EndDate|tvo.header.StreakStart|" & _
            "tvo.header.Days|tvo.header.ThresholdDate|tvo.text.Ready|tvo.text.Complete|" & _
            "tvo.text.Active|tvo.text.Absent|tvo.text.Closed|tvo.text.Reached|tvo.text.Forecast|" & _
            "tvo.text.NotStarted|tvo.text.InvalidInterval|tvo.text.OrderNotFound|tvo.text.OrderAmbiguous|" & _
            "tvo.header.Duration|tvo.header.AccessOrder|tvo.header.StartOrder|tvo.header.EndOrder|source.tvo.columns.EndOrder|" & _
            "tvo.header.PlannedEndDate|source.tvo.columns.PlannedEndDate", "|")
        For Each suffix In tvoKeys
            key = "PersonnelAtDisposalDaysCalculation::" & VBA.CStr(suffix)
            If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
                GoTo Failed
            End If
            m_values.Add VBA.CStr(suffix), value
        Next suffix
        m_isInitialized = True
        Initialize = True
        Exit Function
    End If
    For Each ruleId In VBA.Split(m_values("event.rules.ids"), "|")
        key = "PersonnelAtDisposalDaysCalculation::event.rules.ids"
        If VBA.Len(VBA.Trim$(VBA.CStr(ruleId))) = 0 Then
            GoTo Failed
        End If
        For Each field In VBA.Array("codes", "openStatuses", "events")
            ruleKey = "event.rules." & VBA.Trim$(VBA.CStr(ruleId)) & "." & VBA.CStr(field)
            key = "PersonnelAtDisposalDaysCalculation::" & ruleKey
            If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
                GoTo Failed
            End If
            For Each operation In VBA.Split(value, "|")
                operation = VBA.Trim$(VBA.CStr(operation))
                If VBA.Len(operation) = 0 Then
                    GoTo Failed
                End If
                If field = "events" Then
                    If (VBA.Left$(operation, 1) <> "+" And VBA.Left$(operation, 1) <> "-") _
                            Or VBA.Len(VBA.Trim$(VBA.Mid$(operation, 2))) = 0 Then
                        GoTo Failed
                    End If
                End If
            Next operation
            m_values.Add ruleKey, value
        Next field
    Next ruleId
    key = "PersonnelAtDisposalDaysCalculation::format.periodRangeSeparator"
    If VBA.Len(m_values("format.periodRangeSeparator")) = 0 Then
        GoTo Failed
    End If
    key = "PersonnelAtDisposalDaysCalculation::calculation.thresholdDays"
    value = m_values("calculation.thresholdDays")
    If Not VBA.IsNumeric(value) Then
        GoTo Failed
    End If
    m_thresholdDays = VBA.CLng(value)
    If m_thresholdDays < 1 Then
        GoTo Failed
    End If
    m_isInitialized = True
    Initialize = True
    Exit Function
Failed:
    diagnostic = "Required configuration key is missing or invalid: " & key
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_values = Nothing
End Sub

Public Function GetText(ByVal suffix As String) As String
    private_EnsureReady
    If Not m_values.Exists(suffix) Then
        Me.Fail "Required configuration key not loaded: " & suffix
    End If
    GetText = m_values(suffix)
End Function

Public Sub Fail(ByVal message As String)
    private_EnsureReady
    VBA.Err.Raise vbObjectError + 2100, "PersonnelAtDisposalDays", message
End Sub

' //
' // Private
' //
Private Function private_RequiredKeys() As Collection
    Dim keys As Collection

    Set keys = New Collection
    keys.Add "source.dateOverrides.fileName"
    keys.Add "source.dateOverrides.sheet"
    keys.Add "source.dateOverrides.range"
    keys.Add "source.dateOverrides.columns.FullName"
    keys.Add "source.dateOverrides.columns.Date"
    keys.Add "progress.DateOverrides"
    keys.Add "text.OverrideDateUsed"
    keys.Add "text.OverrideDateConflict"
    keys.Add "event.tripChain.label"
    keys.Add "event.tripChain.target"
    keys.Add "event.tripChain.predecessors"
    keys.Add "text.SnapshotFailed"
    keys.Add "calculation.fallbackStartDate"
    keys.Add "header.Notes"
    keys.Add "text.FallbackDateUsed"
    keys.Add "source.enrollment.sheet"
    keys.Add "source.enrollment.range"
    keys.Add "source.enrollment.columns.TaxId"
    keys.Add "source.enrollment.columns.FullName"
    keys.Add "source.enrollment.columns.Date"
    keys.Add "progress.Enrollment"
    keys.Add "progress.EnrollmentIndex"
    keys.Add "progress.Template"
    keys.Add "progress.Parameters"
    keys.Add "progress.Personnel"
    keys.Add "progress.Orders"
    keys.Add "progress.Events"
    keys.Add "progress.OrderIndex"
    keys.Add "progress.EventIndex"
    keys.Add "progress.Preparation"
    keys.Add "progress.Calculation"
    keys.Add "progress.Publish"
    keys.Add "status.Loading"
    keys.Add "status.Calculating"
    keys.Add "status.Completed"
    keys.Add "status.Cancelled"
    keys.Add "status.Cancelling"
    keys.Add "status.Exported"
    keys.Add "source.personnel.columns.TaxId"
    keys.Add "source.personnel.columns.FullName"
    keys.Add "source.personnel.columns.Rank"
    keys.Add "source.personnel.columns.Code"
    keys.Add "source.personnel.disposalCode"
    keys.Add "source.personnel.reference"
    keys.Add "source.personnel.rosterCode"
    keys.Add "source.orders.path"
    keys.Add "source.orders.sheet"
    keys.Add "source.orders.range"
    keys.Add "source.orders.columns.TaxId"
    keys.Add "source.orders.columns.FullName"
    keys.Add "source.orders.columns.Action"
    keys.Add "source.orders.columns.Date"
    keys.Add "source.orders.action"
    keys.Add "source.events.path"
    keys.Add "source.events.range"
    keys.Add "text.PersonIdentityMissing"
    keys.Add "export.header.Rank"
    keys.Add "export.header.Status"
    keys.Add "export.header.OpenStatus"
    keys.Add "export.header.StartDate"
    keys.Add "export.header.Period"
    keys.Add "export.header.UnitPeriod"
    keys.Add "export.header.Difference"
    keys.Add "export.formula.Period"
    keys.Add "export.formula.Difference"
    keys.Add "text.ParameterRow"
    keys.Add "text.ExportEmpty"
    keys.Add "text.Title"
    keys.Add "text.ExportFailed"
    keys.Add "format.date"
    keys.Add "export.font"
    keys.Add "export.padding.rows"
    keys.Add "export.padding.columns"
    keys.Add "export.computedColor"
    keys.Add "export.fileName"
    keys.Add "export.versionSuffix"
    keys.Add "text.ExportFolderMissing"
    keys.Add "export.sheet"
    keys.Add "export.table"
    keys.Add "export.title"
    keys.Add "export.widths"
    keys.Add "export.palette"
    keys.Add "export.fontSize"
    keys.Add "export.styleName"
    keys.Add "export.styleId"
    keys.Add "export.styleSource"
    keys.Add "export.header.From"
    keys.Add "export.header.Include"
    keys.Add "export.header.Through"
    keys.Add "export.labels.End"
    keys.Add "export.labels.Date"
    keys.Add "export.labels.Start"
    keys.Add "export.formula.End"
    keys.Add "export.formula.From"
    keys.Add "export.formula.Start"
    keys.Add "export.formula.Include"
    keys.Add "export.formula.Through"
    keys.Add "source.events.sheet"
    keys.Add "source.events.columns.FullName"
    keys.Add "text.PersonUnverified"
    keys.Add "text.AmbiguousName"
    keys.Add "source.events.columns.TaxId"
    keys.Add "source.events.columns.Event"
    keys.Add "source.events.columns.Departure"
    keys.Add "source.events.columns.Arrival"
    keys.Add "text.SkippedPeople"
    keys.Add "event.present"
    keys.Add "text.AlreadyRunning"
    keys.Add "text.ParameterLoadingStarted"
    keys.Add "text.SelectToolSheet"
    keys.Add "text.WrongWorkbook"
    keys.Add "text.ParameterWorksheet"
    keys.Add "text.ParameterLoadingCompleted"
    keys.Add "text.InputFile"
    keys.Add "text.EndDate"
    keys.Add "text.OperationStopped"
    keys.Add "text.RukhSource"
    keys.Add "text.ParametersEmpty"
    keys.Add "text.People"
    keys.Add "text.RukhEvents"
    keys.Add "text.WorksheetRow"
    keys.Add "text.TaxId"
    keys.Add "text.DuplicateTaxId"
    keys.Add "text.TimePoint"
    keys.Add "text.StartAfterEnd"
    keys.Add "text.FullName"
    keys.Add "text.Event"
    keys.Add "text.Departure"
    keys.Add "text.ArrivalCellError"
    keys.Add "text.Arrival"
    keys.Add "text.ArrivalBeforeDeparture"
    keys.Add "text.WritingResult"
    keys.Add "text.Rows"
    keys.Add "text.TargetPrepareFailed"
    keys.Add "text.CalculationCompletedPeople"
    keys.Add "text.RelativePathBase"
    keys.Add "text.AmbiguousPath"
    keys.Add "text.FileNotFoundOrInaccessible"
    keys.Add "text.Cancelled"
    keys.Add "text.MultipleSources"
    keys.Add "text.OpenSource"
    keys.Add "text.Table"
    keys.Add "text.TableNotFound"
    keys.Add "text.OnWorksheet"
    keys.Add "text.CheckTableConfiguration"
    keys.Add "text.TableText"
    keys.Add "text.IsMissingColumn"
    keys.Add "text.CellError"
    keys.Add "text.RequiredValue"
    keys.Add "text.InvalidDate"
    keys.Add "text.InvalidReference"
    keys.Add "event.rules.ids"
    keys.Add "text.EventRuleConflict"
    keys.Add "event.eligibleOpenStatuses"
    keys.Add "header.FullName"
    keys.Add "header.TaxId"
    keys.Add "header.StartDate"
    keys.Add "header.Periods"
    keys.Add "header.CountedDays"
    keys.Add "header.ThresholdDate"
    keys.Add "calculation.thresholdDays"
    keys.Add "format.periodRangeSeparator"
    Set private_RequiredKeys = keys
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_Configuration", "Service is not initialized."
    End If
End Sub