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
Public Function Initialize() As Boolean
    Dim suffix As Variant
    Dim key As String
    Dim value As String
    Dim requiredKeys As Collection

    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    Set m_values = VBA.CreateObject("Scripting.Dictionary")
    Set requiredKeys = private_RequiredKeys()
    For Each suffix In requiredKeys
        key = "PersonnelAtDisposalDaysCalculation::" & VBA.CStr(suffix)
        If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
            GoTo Failed
        End If
        m_values.Add VBA.CStr(suffix), value
    Next suffix
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
    ex_WindowsUi.fn_ShowMessage "Required configuration key is missing or invalid: " & key, vbExclamation
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
    keys.Add "text.DisposalDateMissing"
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
    keys.Add "header.errors.FullName"
    keys.Add "header.errors.TaxId"
    keys.Add "header.errors.Description"
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
    keys.Add "text.OperationCancelled"
    keys.Add "text.OperationStopped"
    keys.Add "text.CancellingOperation"
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
    keys.Add "text.WritingResults"
    keys.Add "text.CalculationCompletedPeople"
    keys.Add "text.RelativePathBase"
    keys.Add "text.AmbiguousPath"
    keys.Add "text.FileNotFoundOrInaccessible"
    keys.Add "text.Cancelled"
    keys.Add "text.CalculatingDays"
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
    keys.Add "event.alwaysExcluded"
    keys.Add "event.optionalExcluded"
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