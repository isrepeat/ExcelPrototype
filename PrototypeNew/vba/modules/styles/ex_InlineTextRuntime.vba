Attribute VB_Name = "ex_InlineTextRuntime"
Option Explicit

Private g_InlineRuns As Collection

Public Sub m_ResetInlineRuns()
    Set g_InlineRuns = Nothing
End Sub

Public Function m_RegisterInlineRuns( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal runs As Collection, _
    ByVal presentation As obj_Presentation _
) As Boolean
    Dim runInfo As Object
    Dim entry As Object
    Dim tagName As String
    Dim runStart As Long
    Dim runLength As Long
    Dim fontColor As Long
    Dim fontBold As Boolean
    Dim fontItalic As Boolean
    Dim fontUnderline As Boolean
    Dim firstCell As Range

    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then
        m_RegisterInlineRuns = True
        Exit Function
    End If
    If runs Is Nothing Then
        m_RegisterInlineRuns = True
        Exit Function
    End If
    If presentation Is Nothing Then
        m_RegisterInlineRuns = True
        Exit Function
    End If

    If g_InlineRuns Is Nothing Then Set g_InlineRuns = New Collection
    Set firstCell = targetRange.Cells(1, 1)

    For Each runInfo In runs
        tagName = LCase$(Trim$(CStr(runInfo("Tag"))))
        runStart = CLng(runInfo("Start"))
        runLength = CLng(runInfo("Length"))

        If runStart <= 0 Or runLength <= 0 Then GoTo ContinueRun
        If Not presentation.m_TryResolveInlineRunStyle( _
            tagName, fontColor, fontBold, fontItalic, fontUnderline) Then GoTo ContinueRun

        Set entry = CreateObject("Scripting.Dictionary")
        entry.CompareMode = 1
        entry("SheetName") = LCase$(ws.Name)
        entry("CellAddress") = firstCell.Address(False, False)
        entry("Start") = runStart
        entry("Length") = runLength
        entry("FontColor") = fontColor
        entry("FontBold") = fontBold
        entry("FontItalic") = fontItalic
        entry("FontUnderline") = fontUnderline

        g_InlineRuns.Add entry

ContinueRun:
    Next runInfo

    m_RegisterInlineRuns = True
End Function

Public Function m_ApplyInlineRuns(ByVal ws As Worksheet) As Boolean
    Dim entry As Object
    Dim targetCell As Range
    Dim wsKey As String
    Dim underlineValue As XlUnderlineStyle

    If ws Is Nothing Then Exit Function

    If g_InlineRuns Is Nothing Then
        m_ApplyInlineRuns = True
        Exit Function
    End If

    wsKey = LCase$(ws.Name)

    For Each entry In g_InlineRuns
        If LCase$(CStr(entry("SheetName"))) <> wsKey Then GoTo ContinueEntry

        Set targetCell = Nothing
        On Error Resume Next
        Set targetCell = ws.Range(CStr(entry("CellAddress")))
        On Error GoTo 0
        If targetCell Is Nothing Then GoTo ContinueEntry

        If CBool(entry("FontUnderline")) Then
            underlineValue = xlUnderlineStyleSingle
        Else
            underlineValue = xlUnderlineStyleNone
        End If

        On Error Resume Next
        With targetCell.Characters(CLng(entry("Start")), CLng(entry("Length"))).Font
            .Color = CLng(entry("FontColor"))
            .Bold = CBool(entry("FontBold"))
            .Italic = CBool(entry("FontItalic"))
            .Underline = underlineValue
        End With
        On Error GoTo 0

ContinueEntry:
    Next entry

    m_ApplyInlineRuns = True
End Function
