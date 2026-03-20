Attribute VB_Name = "ex_SimpleTestAction"
Option Explicit

Private Const RESULT_SHEET_NAME As String = "g_SimpleTest"
Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_COL_MARKER As Long = 1
Private Const DEV_COL_KEY As Long = 2
Private Const DEV_COL_VALUE As Long = 3

Public Sub m_RunSimpleTest()
    Dim cfg As Object

    Set cfg = mp_LoadConfigDictionary()
    ex_ModePipeline.m_RunModePipeline cfg, "ex_SimpleTestAction.m_RunMode", Nothing, False
End Sub

Public Function m_RunMode(ByVal cfg As Object, ByVal modeInput As Object, ByVal preProcessContext As Object) As Object
    Dim wsOut As Worksheet
    Dim commonKey As String
    Dim preHello As String
    Dim modeResult As Object
    Dim resultTables As Collection

    On Error GoTo EH

    Set wsOut = mp_CreateOrClearSheet(RESULT_SHEET_NAME)
    ex_Messaging.m_ClearResultTableAnchors wsOut
    ex_Messaging.m_ClearResultRowAnchors wsOut

    wsOut.Cells(1, 1).Value = "SimpleTest"
    wsOut.Cells(1, 2).Value = "Pipeline Smoke"

    commonKey = ex_ScriptIO.m_GetStringOrDefault(modeInput, "CommonKey", vbNullString)
    wsOut.Cells(2, 1).Value = "Key"
    wsOut.Cells(2, 2).Value = commonKey

    preHello = ex_ScriptIO.m_GetStringOrDefault(modeInput, "PreHello", vbNullString)
    wsOut.Cells(3, 1).Value = "PreHello"
    wsOut.Cells(3, 2).Value = preHello

    ex_OutputFormattingPipeline.m_ApplySheetPipeline wsOut, Nothing, Nothing, Nothing, "SimpleTest"
    wsOut.Activate

    Set resultTables = New Collection
    Set modeResult = CreateObject("Scripting.Dictionary")
    modeResult.CompareMode = 1
    Set modeResult("Output") = modeInput
    Set modeResult("Worksheet") = wsOut
    Set modeResult("ResultTables") = resultTables

    Set m_RunMode = modeResult
    Exit Function

EH:
    MsgBox "SimpleTest failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
    Set m_RunMode = Nothing
End Function

Private Function mp_CreateOrClearSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set mp_CreateOrClearSheet = ws
End Function

Private Function mp_LoadConfigDictionary() As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cfg As Object
    Dim dataRange As Range
    Dim r As Long
    Dim markerText As String
    Dim keyText As String

    Set ws = ws_Dev

    On Error Resume Next
    Set tbl = ws.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        Err.Raise vbObjectError + 6401, "ex_SimpleTestAction", "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'."
    End If
    If tbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 6402, "ex_SimpleTestAction", "Config table '" & DEV_CONFIG_TABLE_NAME & "' has no data rows."
    End If

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = 1
    Set dataRange = tbl.DataBodyRange

    For r = 1 To dataRange.Rows.Count
        markerText = Trim$(CStr(dataRange.Cells(r, DEV_COL_MARKER).Value))
        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then GoTo ContinueRow

        keyText = Trim$(CStr(dataRange.Cells(r, DEV_COL_KEY).Value))
        If Len(keyText) = 0 Then GoTo ContinueRow

        cfg(keyText) = CStr(dataRange.Cells(r, DEV_COL_VALUE).Value)
ContinueRow:
    Next r

    Set mp_LoadConfigDictionary = cfg
End Function

Public Function m_LogInputObjectField(ByVal fieldName As String, Optional ByVal logPath As String = "Logs\personalcard_pipeline.log") As String
    Dim inputObj As Object
    Dim valueObj As Object
    Dim info As String

    Set inputObj = ex_ScriptIO.m_GetInput()
    If Not ex_ScriptIO.m_TryGetObject(inputObj, fieldName, valueObj) Then
        info = "[POST][SIMPLE] object field '" & fieldName & "' not found"
        ex_Messaging.m_LogToFile info, logPath
        m_LogInputObjectField = info
        Exit Function
    End If

    info = "[POST][SIMPLE] object field '" & fieldName & "' type=" & TypeName(valueObj)
    ex_Messaging.m_LogToFile info, logPath

    If TypeName(valueObj) = "Collection" Then
        ex_Messaging.m_LogToFile "[POST][SIMPLE] " & fieldName & ".Count=" & CStr(valueObj.Count), logPath
    ElseIf TypeName(valueObj) = "Dictionary" Or TypeName(valueObj) = "Scripting.Dictionary" Then
        ex_Messaging.m_LogToFile "[POST][SIMPLE] " & fieldName & ".Count=" & CStr(valueObj.Count), logPath
    End If

    m_LogInputObjectField = info
End Function
