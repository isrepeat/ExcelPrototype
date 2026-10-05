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
    keys.Add "legacy.SOURCE_SHEET_NAME"
    keys.Add "legacy.SOURCE_TABLE_NAME"
    keys.Add "legacy.SOURCE_COL_NAME"
    keys.Add "legacy.ROSTER_SHEET_NAME"
    keys.Add "legacy.ROSTER_TABLE_NAME"
    keys.Add "legacy.ROSTER_COL_NAME"
    keys.Add "legacy.ROSTER_COL_TAX_ID"
    keys.Add "legacy.MSG_PERSON_UNVERIFIED"
    keys.Add "legacy.MSG_AMBIGUOUS_NAME"
    keys.Add "legacy.SOURCE_COL_TAX_ID"
    keys.Add "legacy.SOURCE_COL_EVENT"
    keys.Add "legacy.SOURCE_COL_PERIOD_FROM"
    keys.Add "legacy.SOURCE_COL_PERIOD_TO"
    keys.Add "legacy.ERROR_COL_NAME"
    keys.Add "legacy.ERROR_COL_TAX_ID"
    keys.Add "legacy.ERROR_COL_DESCRIPTION"
    keys.Add "legacy.MSG_SKIPPED_PEOPLE"
    keys.Add "legacy.PARAM_COL_NAME"
    keys.Add "legacy.PARAM_COL_TAX_ID"
    keys.Add "legacy.PARAM_COL_START"
    keys.Add "legacy.PARAM_COL_TRIPS"
    keys.Add "legacy.PRESENT_EVENT_NAME"
    keys.Add "legacy.MSG_ALREADY_RUNNING"
    keys.Add "legacy.MSG_PARAMETER_LOADING_STARTED"
    keys.Add "legacy.MSG_SELECT_TOOL_SHEET"
    keys.Add "legacy.MSG_WRONG_WORKBOOK"
    keys.Add "legacy.MSG_PARAMETER_WORKSHEET"
    keys.Add "legacy.MSG_PARAMETER_LOADING_COMPLETED"
    keys.Add "legacy.MSG_INPUT_FILE"
    keys.Add "legacy.MSG_END_DATE"
    keys.Add "legacy.MSG_OPERATION_CANCELLED"
    keys.Add "legacy.MSG_OPERATION_STOPPED"
    keys.Add "legacy.MSG_CANCELLING_OPERATION"
    keys.Add "legacy.MSG_RUKH_SOURCE"
    keys.Add "legacy.MSG_PARAMETER_CLOSE_FAILED"
    keys.Add "legacy.MSG_PARAMETER_SHEET_MISSING"
    keys.Add "legacy.MSG_PARAMETERS_EMPTY"
    keys.Add "legacy.MSG_PEOPLE"
    keys.Add "legacy.MSG_RUKH_EVENTS"
    keys.Add "legacy.MSG_WORKSHEET_ROW"
    keys.Add "legacy.MSG_TAX_ID"
    keys.Add "legacy.MSG_DUPLICATE_TAX_ID"
    keys.Add "legacy.MSG_TIME_POINT"
    keys.Add "legacy.MSG_START_AFTER_END"
    keys.Add "legacy.MSG_FULL_NAME"
    keys.Add "legacy.MSG_EVENT"
    keys.Add "legacy.MSG_DEPARTURE"
    keys.Add "legacy.MSG_ARRIVAL_CELL_ERROR"
    keys.Add "legacy.MSG_ARRIVAL"
    keys.Add "legacy.MSG_ARRIVAL_BEFORE_DEPARTURE"
    keys.Add "legacy.MSG_WRITING_RESULT"
    keys.Add "legacy.MSG_ROWS"
    keys.Add "legacy.MSG_TARGET_PREPARE_FAILED"
    keys.Add "legacy.MSG_WRITING_RESULTS"
    keys.Add "legacy.MSG_CALCULATION_COMPLETED_PEOPLE"
    keys.Add "legacy.MSG_RELATIVE_PATH_BASE"
    keys.Add "legacy.MSG_AMBIGUOUS_PATH"
    keys.Add "legacy.MSG_FILE_NOT_FOUND_OR_INACCESSIBLE"
    keys.Add "legacy.MSG_CANCELLED"
    keys.Add "legacy.MSG_CALCULATING_DAYS"
    keys.Add "legacy.MSG_MULTIPLE_SOURCES"
    keys.Add "legacy.MSG_OPEN_SOURCE"
    keys.Add "legacy.MSG_TABLE"
    keys.Add "legacy.MSG_TABLE_NOT_FOUND"
    keys.Add "legacy.MSG_ON_WORKSHEET"
    keys.Add "legacy.MSG_CHECK_TABLE_CONFIGURATION"
    keys.Add "legacy.MSG_TABLE_TEXT"
    keys.Add "legacy.MSG_IS_MISSING_COLUMN"
    keys.Add "legacy.MSG_CELL_ERROR"
    keys.Add "legacy.MSG_REQUIRED_VALUE"
    keys.Add "legacy.MSG_SUBTRACT_BUSINESS_TRIPS"
    keys.Add "legacy.MSG_INVALID_TRIP_FLAG"
    keys.Add "legacy.MSG_INVALID_DATE"
    keys.Add "legacy.MSG_INVALID_REFERENCE"
    keys.Add "legacy.MSG_PARAMETER_TABLE_MISSING"
    keys.Add "event.alwaysExcluded"
    keys.Add "event.optionalExcluded"
    keys.Add "header.FullName"
    keys.Add "header.TaxId"
    keys.Add "header.StartDate"
    keys.Add "header.BusinessTrips"
    keys.Add "header.Periods"
    keys.Add "header.CountedDays"
    keys.Add "header.ThresholdDate"
    keys.Add "flag.includeTrips"
    keys.Add "flag.excludeTrips"
    keys.Add "calculation.thresholdDays"
    keys.Add "format.periodRangeSeparator"
    Set private_RequiredKeys = keys
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_Configuration", "Service is not initialized."
    End If
End Sub