VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_WordAnchors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_anchors As Object
Private m_sourceText As String
Private m_expectedText As String
Private m_insertDocument As Object

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
    ByVal document As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim expression As Object
    Dim whitespace As Object
    Dim match As Object
    Dim marker As String
    Dim identifier As String
    Dim pending As String
    Dim opening As Object

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    Set m_anchors = VBA.CreateObject("Scripting.Dictionary")
    m_sourceText = private_DocumentText(document)
    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "\{\s*\\\s*export\s*:[^{}]*\}"
    Set whitespace = VBA.CreateObject("VBScript.RegExp")
    whitespace.Global = True
    whitespace.Pattern = "[\s" & VBA.ChrW(160) & "]+"
    For Each match In expression.Execute(m_sourceText)
        marker = whitespace.Replace(match.Value, VBA.vbNullString)
        marker = VBA.Mid$(marker, 10, VBA.Len(marker) - 10)
        If VBA.Right$(marker, 6) = "_Begin" Then
            identifier = VBA.Left$(marker, VBA.Len(marker) - 6)
            If VBA.Len(pending) > 0 Or m_anchors.Exists(identifier) Or VBA.Len(identifier) = 0 Then
                Err.Raise 5, , "Duplicate or nested Word anchor: " & identifier
            End If
            pending = identifier
            Set opening = match
        ElseIf VBA.Right$(marker, 4) = "_End" Then
            identifier = VBA.Left$(marker, VBA.Len(marker) - 4)
            If VBA.Len(pending) = 0 Or identifier <> pending Then
                Err.Raise 5, , "Unpaired Word anchor end: " & identifier
            End If
            m_anchors.Add identifier, VBA.Array( _
                opening.Value, _
                match.Value, _
                opening.FirstIndex + opening.Length, _
                match.FirstIndex _
            )
            pending = VBA.vbNullString
        Else
            Err.Raise 5, , "Invalid Word export marker: " & match.Value
        End If
    Next match
    If VBA.Len(pending) > 0 Then Err.Raise 5, , "Unpaired Word anchor begin: " & pending
    If m_anchors.Count = 0 Then Err.Raise 5, , "Word template has no Begin/End export anchors."
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = "Word anchors: " & VBA.Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Set m_anchors = Nothing
    m_sourceText = VBA.vbNullString
    m_expectedText = VBA.vbNullString
    Set m_insertDocument = Nothing
End Sub

Public Function TryInsert( _
    ByVal document As Object, _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim identifier As String
    Dim positions As Variant
    Dim target As Object
    Dim rendered As String
    Dim beginRange As Object
    Dim endRange As Object
    Dim currentText As String
    Dim beginOffset As Long
    Dim endOffset As Long
    Dim font As Object
    Dim paragraphFormat As Object
    Dim stage As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureText As String
    Dim characterIndex As Long
    Dim character As String

    On Error GoTo EH
    stage = "ResolveAnchor"
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Word anchor index is not initialized."
    identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
    If Not m_anchors.Exists(identifier) Then Err.Raise 5, , "Word template anchor not found: " & identifier
    Set m_insertDocument = Nothing
    m_expectedText = VBA.vbNullString
    positions = m_anchors(identifier)
    ' FormattedText can normalize a copied document; verify against its own baseline.
    stage = "ReadDocumentText"
    currentText = private_DocumentText(document)
    beginOffset = VBA.InStr(1, currentText, positions(0), VBA.vbBinaryCompare)
    endOffset = VBA.InStr(1, currentText, positions(1), VBA.vbBinaryCompare)
    If beginOffset = 0 Or endOffset = 0 Then Err.Raise 5, , "Word export marker is missing: " & identifier
    If VBA.InStr(beginOffset + 1, currentText, positions(0), VBA.vbBinaryCompare) > 0 Or _
            VBA.InStr(endOffset + 1, currentText, positions(1), VBA.vbBinaryCompare) > 0 Then
        Err.Raise 5, , "Duplicate Word export marker: " & identifier
    End If
    stage = "NormalizeText"
    rendered = VBA.Replace$(VBA.CStr(values("WordText")), VBA.vbCrLf, VBA.vbCr)
    rendered = VBA.Replace$(rendered, VBA.vbLf, VBA.vbCr)
    ' Anchor sections continue the preceding heading, which ends with a colon.
    For characterIndex = 1 To VBA.Len(rendered)
        character = VBA.Mid$(rendered, characterIndex, 1)
        If VBA.UCase$(character) <> VBA.LCase$(character) Then
            Mid$(rendered, characterIndex, 1) = VBA.LCase$(character)
            Exit For
        End If
    Next characterIndex
    stage = "FindBeginMarker"
    Set beginRange = private_FindMarker(document, VBA.CStr(positions(0)))
    stage = "FindEndMarker"
    Set endRange = private_FindMarker(document, VBA.CStr(positions(1)))
    If beginRange.End > endRange.Start Then Err.Raise 5, , "Reversed Word anchor ranges: " & identifier
    stage = "CreateInsertionRange"
    Set target = document.Range(beginRange.End, endRange.Start)
    stage = "CaptureFont"
    Set font = beginRange.Characters(1).Font.Duplicate
    stage = "CaptureParagraphFormat"
    Set paragraphFormat = beginRange.Paragraphs(1).Format.Duplicate
    stage = "BuildExpectedText"
    m_expectedText = VBA.Left$(currentText, beginOffset + VBA.Len(positions(0)) - 1) & _
        rendered & VBA.Mid$(currentText, endOffset)
    stage = "WriteText"
    target.Text = rendered
    stage = "ApplyFont"
    target.Font = font
    stage = "ShowText"
    target.Font.Hidden = False
    stage = "ApplyParagraphFormat"
    target.ParagraphFormat = paragraphFormat
    Set m_insertDocument = document
    TryInsert = True
    Exit Function
EH:
    failureNumber = VBA.Err.Number
    failureSource = VBA.Err.Source
    failureText = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Word.Insertion." & stage, failureNumber, failureSource, failureText
    ex_Core.fn_Diagnostic_WriteLog "WORD_INSERTION_FAILED | Anchor=" & identifier & _
        " | Stage=" & stage & " | Characters=" & VBA.CStr(VBA.Len(rendered))
    If Not ex_Core.fn_Diagnostic_Flush(True) Then
        Debug.Print "Word insertion diagnostic flush failed."
    End If
    diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("text.WordInsertionFailed"), "{stage}", stage)
    diagnostic = VBA.Replace$(diagnostic, "{anchor}", identifier)
    diagnostic = VBA.Replace$(diagnostic, "{number}", VBA.CStr(failureNumber))
    diagnostic = VBA.Replace$(diagnostic, "{details}", failureText)
End Function

Public Function TryVerify( _
    ByVal document As Object, _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim identifier As String
    Dim positions As Variant
    Dim rendered As String
    Dim expected As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Word anchor index is not initialized."
    identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
    If Not m_anchors.Exists(identifier) Then Err.Raise 5, , "Word template anchor not found: " & identifier
    positions = m_anchors(identifier)
    rendered = VBA.Replace$(VBA.CStr(values("WordText")), VBA.vbCrLf, VBA.vbCr)
    rendered = VBA.Replace$(rendered, VBA.vbLf, VBA.vbCr)
    If Not document Is m_insertDocument Then Err.Raise 5, , "Word verification document differs from the inserted document."
    expected = m_expectedText
    If VBA.StrComp(private_DocumentText(document), expected, VBA.vbBinaryCompare) <> 0 Then
        Err.Raise 5, , "Inserted Word text or content outside the anchor differs: " & identifier
    End If
    TryVerify = True
    Exit Function
EH:
    diagnostic = "Word verification: " & VBA.Err.Description
End Function

Public Function TryGetInsertedRange( _
    ByVal document As Object, _
    ByVal values As Object, _
    ByRef insertedRange As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim identifier As String
    Dim positions As Variant
    Dim beginRange As Object
    Dim endRange As Object

    On Error GoTo EH
    Set insertedRange = Nothing
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Word anchor index is not initialized."
    identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
    positions = m_anchors(identifier)
    Set beginRange = private_FindMarker(document, VBA.CStr(positions(0)))
    Set endRange = private_FindMarker(document, VBA.CStr(positions(1)))
    Set insertedRange = document.Range(beginRange.End, endRange.Start)
    TryGetInsertedRange = True
    Exit Function
EH:
    diagnostic = "Word inserted range: " & VBA.Err.Description
End Function

Public Sub HideMarkers( _
    ByVal document As Object, _
    ByVal values As Object _
)
    Dim identifier As String
    Dim positions As Variant
    Dim marker As Object

    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word anchor index is not initialized."
    End If
    identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
    If Not m_anchors.Exists(identifier) Then
        Err.Raise 5, , "Word group anchor is missing."
    End If
    positions = m_anchors(identifier)
    Set marker = private_FindMarker(document, VBA.CStr(positions(0)))
    marker.Font.Hidden = True
    Set marker = private_FindMarker(document, VBA.CStr(positions(1)))
    marker.Font.Hidden = True
End Sub

Public Function TryRemoveMarkers( _
    ByVal document As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim identifier As Variant
    Dim positions As Variant
    Dim marker As Object
    Dim expected As String

    On Error GoTo EH_REMOVE
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordMarkersRemovalInvalid")
    End If
    expected = private_DocumentText(document)
    For Each identifier In m_anchors.Keys
        positions = m_anchors(identifier)
        expected = VBA.Replace$(expected, VBA.CStr(positions(0)), VBA.vbNullString, 1, -1, VBA.vbBinaryCompare)
        expected = VBA.Replace$(expected, VBA.CStr(positions(1)), VBA.vbNullString, 1, -1, VBA.vbBinaryCompare)
        ' Delete only marker text, preserving section paragraphs and inserted content.
        Set marker = private_FindMarker(document, VBA.CStr(positions(1)))
        marker.Delete
        Set marker = private_FindMarker(document, VBA.CStr(positions(0)))
        marker.Delete
    Next identifier
    If VBA.StrComp(private_DocumentText(document), expected, VBA.vbBinaryCompare) <> 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordMarkersRemovalInvalid")
    End If
    TryRemoveMarkers = True
    Exit Function
EH_REMOVE:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Word.Export.RemoveMarkers", VBA.Err.Number, VBA.Err.Source, diagnostic
End Function

' //
' // Private
' //
Private Function private_DocumentText(ByVal document As Object) As String
    Dim target As Object

    Set target = document.Content.Duplicate
    target.TextRetrievalMode.IncludeHiddenText = True
    private_DocumentText = target.Text
End Function

Private Function private_FindMarker( _
    ByVal document As Object, _
    ByVal marker As String _
) As Object
    Dim target As Object

    Set target = document.Content.Duplicate
    target.Find.ClearFormatting
    If Not target.Find.Execute(marker, True, False, False, False, False, True, 0, False) Then
        Err.Raise 5, , "Word export marker not found: " & marker
    End If
    If VBA.StrComp(target.Text, marker, VBA.vbBinaryCompare) <> 0 Then
        Err.Raise 5, , "Word export marker search returned different text: " & marker
    End If
    Set private_FindMarker = target
End Function