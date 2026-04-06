Attribute VB_Name = "ex_ModeSimpleTest"
Option Explicit

Private Const RESULT_SHEET_NAME As String = "g_SimpleTest"

Public Sub m_RunSimpleTest()
    Dim cfg As Object

    Set cfg = ex_ConfigProvider.m_LoadConfigDictionary("ex_ModeSimpleTest", 6401, 6402)
    ex_ModePipeline.m_RunModePipeline cfg, "ex_ModeSimpleTest.m_RunMode", Nothing, False
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
