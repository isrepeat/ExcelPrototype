VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_TestWordSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_word As Object
Private m_template As Object
Private m_result As Object
Private m_work As Object
Private m_wordAnchors As obj_PEB_WordAnchors
Private m_output As String
Private m_temporary As String
Private m_insertedRange As Object
Private m_pointCount As Long
Private m_lastFailure As String
Private m_templatePath As String
Private m_eventDescription As String
Private m_personNames As Object
Private m_personColor As Long
Private m_basisColor As Long
Private m_basisPrefix As String
Private m_headingBefore As Double
Private m_headingAfter As Double
Private m_metrics As Object
Private m_finished As Boolean
Private m_started As Double
Private m_previousOptions As Variant
Private m_timingText As String
Private m_noTiming As String

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
Public Property Get OutputPath() As String
    OutputPath = m_output
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal runId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim path As String
    Dim fileSystem As Object
    Dim started As Double
    Dim section As Object
    Dim headerType As Long

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    m_started = ex_PEB_WordPerf.fn_Now()
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.word.template"))
    m_templatePath = path
    m_output = ex_PEB_Export.fn_Path(VBA.Replace$(ex_PEB_Export.fn_Config("test.word.output"), "{runId}", runId))
    m_timingText = ex_PEB_Export.fn_Config("text.TestsWordTimings")
    m_noTiming = ex_PEB_Export.fn_Config("text.TestsWordNoTiming")
    m_personColor = private_Color(ex_PEB_Export.fn_Config("test.word.style.personColor"))
    m_basisColor = private_Color(ex_PEB_Export.fn_Config("test.word.style.basisColor"))
    m_basisPrefix = ex_PEB_Export.fn_Config("text.TestsWordBasisPrefix")
    m_headingBefore = VBA.CDbl(ex_PEB_Export.fn_Config("test.word.style.headingBefore"))
    m_headingAfter = VBA.CDbl(ex_PEB_Export.fn_Config("test.word.style.headingAfter"))
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(path) Then Err.Raise 5, , "Word test template not found: " & path
    If Not fileSystem.FolderExists(fileSystem.GetParentFolderName(m_output)) Then
        Err.Raise 5, , "Word test output folder not found: " & m_output
    End If
    If VBA.LCase$(fileSystem.GetExtensionName(m_output)) <> "docx" Then Err.Raise 5, , "Word test output must be .docx."
    If fileSystem.FileExists(m_output) Then Err.Raise 5, , "Word test output already exists: " & m_output
    m_temporary = fileSystem.BuildPath(fileSystem.GetParentFolderName(m_output), fileSystem.GetTempName & ".docx")
    ex_PEB_Testing.fn_ReportPhase "WordStart"
    started = ex_PEB_WordPerf.fn_Now()
    Set m_word = VBA.CreateObject("Word.Application")
    m_word.Visible = False
    m_word.DisplayAlerts = 0
    m_word.AutomationSecurity = 3
    m_word.ScreenUpdating = False
    m_previousOptions = VBA.Array( _
        m_word.Options.CheckSpellingAsYouType, _
        m_word.Options.CheckGrammarAsYouType, _
        m_word.Options.Pagination _
    )
    m_word.Options.CheckSpellingAsYouType = False
    m_word.Options.CheckGrammarAsYouType = False
    m_word.Options.Pagination = False
    ex_PEB_WordPerf.fn_WritePerf "Start", started
    ex_PEB_Testing.fn_ReportPhase "WordTemplate"
    started = ex_PEB_WordPerf.fn_Now()
    Set m_template = m_word.Documents.Open(path, False, True, False, , , , , , , , False)
    Set m_wordAnchors = New obj_PEB_WordAnchors
    If Not m_wordAnchors.Initialize(m_template, diagnostic) Then GoTo Failed
    ex_PEB_WordPerf.fn_WritePerf "TemplateAndIndex", started
    ex_PEB_Testing.fn_ReportPhase "WordReport"
    started = ex_PEB_WordPerf.fn_Now()
    Set m_result = private_NewDocument()
    m_result.Content.Text = VBA.vbCr
    For Each section In m_result.Sections
        For headerType = 1 To 3
            section.Headers(headerType).Range.Text = VBA.vbNullString
            section.Footers(headerType).Range.Text = VBA.vbNullString
        Next headerType
    Next section
    ex_PEB_WordPerf.fn_WritePerf "ReportInitialize", started
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = "Word test session: " & VBA.Err.Description
Failed:
    Me.Dispose
End Function

Public Sub Dispose()
    Dim fileSystem As Object

    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    ' Close only documents and the Word instance owned by this run.
    On Error Resume Next
    If Not m_work Is Nothing Then m_work.Close 0
    If Not m_result Is Nothing Then m_result.Close 0
    If Not m_template Is Nothing Then m_template.Close 0
    If Not m_word Is Nothing Then
        If IsArray(m_previousOptions) Then
            m_word.Options.CheckSpellingAsYouType = m_previousOptions(0)
            m_word.Options.CheckGrammarAsYouType = m_previousOptions(1)
            m_word.Options.Pagination = m_previousOptions(2)
        End If
        m_word.Quit 0
    End If
    If Not m_wordAnchors Is Nothing Then m_wordAnchors.Dispose
    If VBA.Len(m_temporary) > 0 Then
        Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
        If fileSystem.FileExists(m_temporary) Then fileSystem.DeleteFile m_temporary
    End If
    Set m_work = Nothing
    Set m_result = Nothing
    Set m_template = Nothing
    Set m_word = Nothing
    Set m_wordAnchors = Nothing
    Set m_metrics = Nothing
    Set m_insertedRange = Nothing
    Set m_personNames = Nothing
    On Error GoTo 0
End Sub

Public Function TryExercise( _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim started As Double
    Dim field As Variant
    Dim personName As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Or m_finished Then Err.Raise 5, , "Word test session is not ready."
    If Not m_work Is Nothing Then Err.Raise 5, , "Previous Word test document has not been collected."
    Set m_metrics = VBA.CreateObject("Scripting.Dictionary")
    started = ex_PEB_WordPerf.fn_Now()
    Set m_work = private_NewDocument()
    private_RecordTime "Create", started
    started = ex_PEB_WordPerf.fn_Now()
    If Not m_wordAnchors.TryInsert(m_work, values, diagnostic) Then
        private_RecordTime "Insert", started
        Exit Function
    End If
    private_RecordTime "Insert", started
    started = ex_PEB_WordPerf.fn_Now()
    If Not m_wordAnchors.TryVerify(m_work, values, diagnostic) Then
        private_RecordTime "Verify", started
        Exit Function
    End If
    private_RecordTime "Verify", started
    If Not m_wordAnchors.TryGetInsertedRange(m_work, values, m_insertedRange, diagnostic) Then Exit Function
    Set m_personNames = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("test.word.style.personFields"), "|")
        If values.Exists(field) Then
            personName = ex_PEB_Export.fn_Text(values(field))
            If VBA.Len(personName) > 0 Then m_personNames(personName) = True
        End If
    Next field
    m_eventDescription = VBA.Replace$( _
        ex_PEB_Export.fn_Config("text.TestsWordEvent." & VBA.CStr(values("ModeDirection"))), _
        "{event}", _
        ex_PEB_EventRoutes.fn_Value(VBA.CStr(values("EventId")), "caption") _
    )
    TryExercise = True
    Exit Function
EH:
    diagnostic = "Word test insertion: " & VBA.Err.Description
End Function

Public Function TryAppend( _
    ByVal report As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Object
    Dim key As Variant
    Dim started As Double
    Dim timingText As String
    Dim timingValue As String
    Dim captionRange As Object

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Or m_finished Then Err.Raise 5, , "Word test session is not ready."
    started = ex_PEB_WordPerf.fn_Now()
    If report("WordResult") = "FAIL" Or report("WordResult") = "ERROR" Then
        m_lastFailure = VBA.CStr(report("PreviewDiagnostic"))
        If VBA.Len(m_lastFailure) = 0 Then m_lastFailure = VBA.CStr(report("Details"))
    End If
    If Not m_insertedRange Is Nothing Then
        If m_pointCount > 0 Then
            Set target = m_result.Range(m_result.Content.End - 1, m_result.Content.End - 1)
            target.Text = VBA.vbCr
        End If
        Set target = m_result.Range(m_result.Content.End - 1, m_result.Content.End - 1)
        target.Text = m_eventDescription & VBA.vbCr
        Set captionRange = m_result.Range(target.Start, target.End - 1)
        target.Font = m_insertedRange.Characters(1).Font.Duplicate
        target.Font.Bold = True
        target.ParagraphFormat = m_insertedRange.Paragraphs(1).Format.Duplicate
        target.ParagraphFormat.FirstLineIndent = 0
        target.ParagraphFormat.Alignment = 1
        target.ParagraphFormat.SpaceBefore = m_headingBefore
        target.ParagraphFormat.SpaceAfter = m_headingAfter
        target.ParagraphFormat.KeepWithNext = False
        target.ParagraphFormat.KeepTogether = False
        Set target = m_result.Range(m_result.Content.End - 1, m_result.Content.End - 1)
        target.FormattedText = m_insertedRange.FormattedText
        ' The final paragraph mark is outside the inserted range; copy its formatting explicitly.
        target.Font = m_insertedRange.Font.Duplicate
        target.ParagraphFormat = m_insertedRange.Paragraphs(1).Format.Duplicate
        private_FormatPoint target
        captionRange.Font.Bold = True
        m_pointCount = m_pointCount + 1
        Set m_insertedRange = Nothing
    End If
    If Not m_work Is Nothing Then
        m_work.Close 0
        Set m_work = Nothing
    End If
    If m_metrics Is Nothing Then Set m_metrics = VBA.CreateObject("Scripting.Dictionary")
    private_RecordTime "Append", started
    For Each key In m_metrics.Keys
        report(key) = m_metrics(key)
    Next key
    timingText = m_timingText
    For Each key In VBA.Array("Create", "Insert", "Verify", "Append")
        timingValue = m_noTiming
        If m_metrics.Exists("Word" & key & "Ms") Then timingValue = VBA.CStr(m_metrics("Word" & key & "Ms"))
        timingText = VBA.Replace$(timingText, "{" & key & "}", timingValue)
    Next key
    report("Details") = report("Details") & " " & timingText
    report("WordDocumentPath") = m_output
    Set m_metrics = Nothing
    TryAppend = True
    Exit Function
EH:
    diagnostic = "Word test report: " & VBA.Err.Description
End Function

Public Function TryFinish( _
    ByVal runState As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim started As Double
    Dim fileSystem As Object

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Word test session is not initialized."
    If m_finished Then
        TryFinish = True
        Exit Function
    End If
    started = ex_PEB_WordPerf.fn_Now()
    If m_pointCount = 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestsWordEmpty") & " " & m_lastFailure
    End If
    m_result.Content.ParagraphFormat.KeepWithNext = False
    m_result.Content.ParagraphFormat.KeepTogether = False
    m_result.SaveAs2 m_temporary, 16
    m_result.Close 0
    Set m_result = Nothing
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    fileSystem.MoveFile m_temporary, m_output
    m_temporary = VBA.vbNullString
    m_finished = True
    ex_PEB_WordPerf.fn_WritePerf "Save", started
    ex_PEB_WordPerf.fn_WritePerf "Total", m_started
    TryFinish = True
    Exit Function
EH:
    diagnostic = "Word test save: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Sub private_FormatPoint(ByVal pointRange As Object)
    Dim paragraph As Object
    Dim search As Object
    Dim personName As Variant
    Dim isBasis As Boolean
    Dim nextStart As Long

    For Each paragraph In pointRange.Paragraphs
        If VBA.Left$(VBA.Trim$(paragraph.Range.Text), VBA.Len(m_basisPrefix)) = m_basisPrefix Then isBasis = True
        paragraph.Format.KeepWithNext = False
        paragraph.Format.KeepTogether = False
        If isBasis Then
            paragraph.Range.Font.Color = m_basisColor
        Else
            For Each personName In m_personNames.Keys
                Set search = paragraph.Range.Duplicate
                Do
                    search.Find.ClearFormatting
                    If Not search.Find.Execute(VBA.CStr(personName), True, False, False, False, False, True, 0, False) Then Exit Do
                    search.Font.Color = m_personColor
                    nextStart = search.End
                    If nextStart >= paragraph.Range.End Then Exit Do
                    search.SetRange nextStart, paragraph.Range.End
                Loop
            Next personName
        End If
    Next paragraph
End Sub

Private Function private_Color(ByVal value As String) As Long
    If VBA.Len(value) <> 7 Or VBA.Left$(value, 1) <> "#" Then Err.Raise 5, , "Expected a #RRGGBB Word report color: " & value
    private_Color = VBA.RGB( _
        VBA.CLng("&H" & VBA.Mid$(value, 2, 2)), _
        VBA.CLng("&H" & VBA.Mid$(value, 4, 2)), _
        VBA.CLng("&H" & VBA.Mid$(value, 6, 2)) _
    )
End Function

Private Function private_NewDocument() As Object
    Dim document As Object
    Dim source As Object
    Dim target As Object

    ' Create from the template so named styles retain their original definitions.
    Set document = m_word.Documents.Add(m_templatePath, False, 0, False)
    Set source = m_template.Sections(1).PageSetup
    Set target = document.Sections(1).PageSetup
    target.Orientation = source.Orientation
    target.PageWidth = source.PageWidth
    target.PageHeight = source.PageHeight
    target.TopMargin = source.TopMargin
    target.BottomMargin = source.BottomMargin
    target.LeftMargin = source.LeftMargin
    target.RightMargin = source.RightMargin
    target.HeaderDistance = source.HeaderDistance
    target.FooterDistance = source.FooterDistance
    Set private_NewDocument = document
End Function

Private Sub private_RecordTime( _
    ByVal stage As String, _
    ByVal started As Double _
)
    Dim elapsed As Double

    elapsed = ex_PEB_WordPerf.fn_ElapsedMs(started)
    m_metrics("Word" & stage & "Ms") = VBA.Round(elapsed, 3)
    ex_PEB_WordPerf.fn_WritePerf stage, started
End Sub