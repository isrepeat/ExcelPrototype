VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_values As Object

' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

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
        key = "PeriodsGeneration::" & VBA.CStr(suffix)
        If Not ex_Core.fn_TryGetWorkbookConfigValue(key, value) Then
            GoTo Failed
        End If
        m_values.Add VBA.CStr(suffix), value
    Next suffix
    key = "PeriodsGeneration::legacy.UI_YIELD_INTERVAL"
    value = m_values("legacy.UI_YIELD_INTERVAL")
    If Not VBA.IsNumeric(value) Then
        GoTo Failed
    End If
    If VBA.CDbl(value) < 1 Or VBA.CDbl(value) <> VBA.Fix(VBA.CDbl(value)) Then
        GoTo Failed
    End If
    value = VBA.CStr(VBA.CLng(value))
    key = "PeriodsGeneration::format.date"
    If VBA.Len(m_values("format.date")) = 0 Then
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
    VBA.Err.Raise vbObjectError + 2100, "PeriodsGeneration", message
End Sub

' //
' // Вспомогательные методы
' //
Private Function private_RequiredKeys() As Collection
    Dim keys As Collection

    Set keys = New Collection
    keys.Add "legacy.SOURCE_COL_RANK"
    keys.Add "legacy.SOURCE_COL_NAME"
    keys.Add "legacy.SOURCE_COL_TAX_ID"
    keys.Add "legacy.SOURCE_COL_POSITION"
    keys.Add "legacy.SOURCE_COL_EVENT"
    keys.Add "legacy.SOURCE_COL_PERIOD_FROM"
    keys.Add "legacy.SOURCE_COL_PERIOD_TO"
    keys.Add "legacy.SOURCE_COL_DEPARTURE_ORDER"
    keys.Add "legacy.SOURCE_COL_ARRIVAL_ORDER"
    keys.Add "legacy.SOURCE_SHEET_NAME"
    keys.Add "legacy.SOURCE_TABLE_NAME"
    keys.Add "legacy.MERGEABLE_EVENTS"
    keys.Add "legacy.DISPLAYABLE_EVENTS"
    keys.Add "legacy.EVENTS_SEPARATOR"
    keys.Add "legacy.OPEN_PERIOD_TO_TEXT"
    keys.Add "legacy.UI_YIELD_INTERVAL"
    keys.Add "legacy.TARGET_COL_RANK"
    keys.Add "legacy.TARGET_COL_NAME"
    keys.Add "legacy.TARGET_COL_TAX_ID"
    keys.Add "legacy.TARGET_COL_POSITION"
    keys.Add "legacy.TARGET_COL_EVENT"
    keys.Add "legacy.TARGET_COL_DEPARTURE_ORDER"
    keys.Add "legacy.TARGET_COL_PERIOD_FROM"
    keys.Add "legacy.TARGET_COL_PERIOD_TO"
    keys.Add "legacy.TARGET_COL_ARRIVAL_ORDER"
    keys.Add "legacy.TARGET_COL_PERIOD_COUNT"
    keys.Add "format.date"
    keys.Add "text.Calculate"
    keys.Add "text.Cancel"
    keys.Add "text.RefreshPage"
    keys.Add "text.StartDateLabel"
    keys.Add "text.EndDateLabel"
    keys.Add "text.Title"
    keys.Add "message.AlreadyRunning"
    keys.Add "message.SourceMissing"
    keys.Add "message.MultipleSources"
    keys.Add "message.ColumnMissing"
    keys.Add "message.InvalidPeriod"
    keys.Add "message.Cancelled"
    keys.Add "message.Failed"
    keys.Add "message.PublishFailed"
    keys.Add "message.Done"
    keys.Add "status.Reading"
    keys.Add "status.ReadingProgress"
    keys.Add "status.Sorting"
    keys.Add "status.Merging"
    keys.Add "status.MergingProgress"
    keys.Add "status.Filtering"
    keys.Add "status.Preparing"
    keys.Add "status.Writing"
    Set private_RequiredKeys = keys
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PG_Configuration", "Service is not initialized."
    End If
End Sub