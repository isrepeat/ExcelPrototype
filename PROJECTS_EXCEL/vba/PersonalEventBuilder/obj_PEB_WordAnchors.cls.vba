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
Private m_blocks As Object
Private m_roots As Object
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
    private_IndexBlocks document
    private_IndexRoots
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
    Set m_blocks = Nothing
    Set m_roots = Nothing
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
    Dim expected As String
    Dim actual As String
    Dim difference As Long

    On Error GoTo EH_REMOVE
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordMarkersRemovalInvalid")
    End If
    expected = private_DocumentText(document)
    private_ApplyCleanup document, expected
    actual = private_DocumentText(document)
    If VBA.StrComp(actual, expected, VBA.vbBinaryCompare) <> 0 Then
        difference = 1
        Do While difference <= VBA.Len(actual) And difference <= VBA.Len(expected)
            If VBA.Mid$(actual, difference, 1) <> VBA.Mid$(expected, difference, 1) Then
                Exit Do
            End If
            difference = difference + 1
        Loop
        ex_Core.fn_Diagnostic_WriteLog "WORD_CLEANUP_MISMATCH | ExpectedLength=" & VBA.CStr(VBA.Len(expected)) & _
            " | ActualLength=" & VBA.CStr(VBA.Len(actual)) & " | FirstDifference=" & VBA.CStr(difference)
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
Private Sub private_IndexBlocks(ByVal document As Object)
    Dim expression As Object
    Dim whitespace As Object
    Dim match As Object
    Dim marker As String
    Dim identifier As String
    Dim pending As String
    Dim opening As Object
    Dim positions As Variant
    Dim otherIdentifier As Variant
    Dim otherPositions As Variant

    Set m_blocks = VBA.CreateObject("Scripting.Dictionary")
    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "\{\s*\\\s*block\s*:[^{}]*\}"
    Set whitespace = VBA.CreateObject("VBScript.RegExp")
    whitespace.Global = True
    whitespace.Pattern = "[\s" & VBA.ChrW(160) & "]+"
    For Each match In expression.Execute(m_sourceText)
        marker = whitespace.Replace(match.Value, VBA.vbNullString)
        marker = VBA.Mid$(marker, 9, VBA.Len(marker) - 9)
        If VBA.Right$(marker, 6) = "_Begin" Then
            identifier = VBA.Left$(marker, VBA.Len(marker) - 6)
            If VBA.Len(pending) > 0 Or m_blocks.Exists(identifier) Or Not m_anchors.Exists(identifier) Then
                Err.Raise 5, , "Duplicate, nested or unknown Word block: " & identifier
            End If
            pending = identifier
            Set opening = match
        ElseIf VBA.Right$(marker, 4) = "_End" Then
            identifier = VBA.Left$(marker, VBA.Len(marker) - 4)
            If VBA.Len(pending) = 0 Or identifier <> pending Then
                Err.Raise 5, , "Unpaired Word block end: " & identifier
            End If
            positions = m_anchors(identifier)
            If opening.FirstIndex + opening.Length > positions(2) - VBA.Len(positions(0)) Or _
                    match.FirstIndex < positions(3) + VBA.Len(positions(1)) Then
                Err.Raise 5, , "Word block does not enclose its export anchors: " & identifier
            End If
            ' Validate paragraph boundaries in memory, without repeated Word searches.
            private_ValidateBlockParagraph opening, whitespace
            private_ValidateBlockParagraph match, whitespace
            For Each otherIdentifier In m_anchors.Keys
                If otherIdentifier <> identifier Then
                    otherPositions = m_anchors(otherIdentifier)
                    If otherPositions(2) > opening.FirstIndex And otherPositions(3) < match.FirstIndex Then
                        Err.Raise 5, , "Word block encloses another export section: " & identifier
                    End If
                End If
            Next otherIdentifier
            m_blocks.Add identifier, VBA.Array(opening.Value, match.Value)
            pending = VBA.vbNullString
        Else
            Err.Raise 5, , "Invalid Word block marker: " & match.Value
        End If
    Next match
    If VBA.Len(pending) > 0 Then Err.Raise 5, , "Unpaired Word block begin: " & pending
End Sub

Private Sub private_IndexRoots()
    Dim expression As Object
    Dim whitespace As Object
    Dim match As Object
    Dim opening As Object
    Dim marker As String
    Dim identifier As String
    Dim pending As String
    Dim key As Variant
    Dim positions As Variant
    Dim markers As Variant
    Dim first As Long
    Dim last As Long

    Set m_roots = VBA.CreateObject("Scripting.Dictionary")
    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "\{\s*\\\s*root\s*:[^{}]*\}"
    Set whitespace = VBA.CreateObject("VBScript.RegExp")
    whitespace.Global = True
    whitespace.Pattern = "[\s" & VBA.ChrW(160) & "]+"
    For Each match In expression.Execute(m_sourceText)
        private_ValidateBlockParagraph match, whitespace
        marker = whitespace.Replace(match.Value, VBA.vbNullString)
        marker = VBA.Mid$(marker, 8, VBA.Len(marker) - 8)
        If VBA.Right$(marker, 6) = "_Begin" Then
            identifier = VBA.Left$(marker, VBA.Len(marker) - 6)
            If Len(pending) > 0 Or Len(identifier) = 0 Or m_roots.Exists(identifier) Then
                Err.Raise 5, , "Duplicate or nested Word root: " & identifier
            End If
            pending = identifier
            Set opening = match
        ElseIf VBA.Right$(marker, 4) = "_End" Then
            identifier = VBA.Left$(marker, VBA.Len(marker) - 4)
            If Len(pending) = 0 Or identifier <> pending Then
                Err.Raise 5, , "Unpaired Word root end: " & identifier
            End If
            For Each key In m_anchors.Keys
                positions = m_anchors(key)
                first = positions(2) - Len(positions(0))
                last = positions(3) + Len(positions(1))
                If m_blocks.Exists(key) Then
                    markers = m_blocks(key)
                    first = InStr(1, m_sourceText, markers(0), vbBinaryCompare) - 1
                    last = InStr(1, m_sourceText, markers(1), vbBinaryCompare) - 1 + Len(markers(1))
                End If
                If first < match.FirstIndex + match.Length And last > opening.FirstIndex Then
                    If first < opening.FirstIndex + opening.Length Or last > match.FirstIndex Then
                        Err.Raise 5, , "Word root intersects a child section: " & identifier & "/" & key
                    End If
                End If
            Next key
            m_roots.Add identifier, VBA.Array(opening.Value, match.Value)
            pending = vbNullString
        Else
            Err.Raise 5, , "Invalid Word root marker: " & match.Value
        End If
    Next match
    If Len(pending) > 0 Then
        Err.Raise 5, , "Unpaired Word root begin: " & pending
    End If
End Sub

Private Sub private_ValidateBlockParagraph( _
    ByVal match As Object, _
    ByVal whitespace As Object _
)
    Dim before As String
    Dim after As String
    Dim paragraphStart As Long
    Dim paragraphEnd As Long
    Dim markerEnd As Long

    paragraphStart = VBA.InStrRev(VBA.Left$(m_sourceText, match.FirstIndex), VBA.vbCr) + 1
    markerEnd = match.FirstIndex + match.Length + 1
    paragraphEnd = VBA.InStr(markerEnd, m_sourceText, VBA.vbCr)
    If paragraphEnd = 0 Then
        paragraphEnd = VBA.Len(m_sourceText) + 1
    End If
    before = VBA.Mid$(m_sourceText, paragraphStart, match.FirstIndex - paragraphStart + 1)
    after = VBA.Mid$(m_sourceText, markerEnd, paragraphEnd - markerEnd)
    If VBA.Len(whitespace.Replace(before & after, VBA.vbNullString)) > 0 Then
        Err.Raise 5, , "Word block marker must occupy a separate paragraph: " & match.Value
    End If
End Sub

Private Sub private_ApplyCleanup( _
    ByVal document As Object, _
    ByRef expected As String _
)
    Dim removals As Collection
    Dim merged As New Collection
    Dim ranges As New Collection
    Dim entry As Variant
    Dim current As Variant
    Dim target As Object
    Dim firstMarker As Object
    Dim lastMarker As Object
    Dim removed As String
    Dim snapshot As String
    Dim index As Long
    Dim delta As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim documentEnd As Long
    Dim startedAt As Double
    Dim fallbackCount As Long

    ' Plan and merge adjacent removals against one immutable text snapshot.
    startedAt = VBA.Timer
    snapshot = expected
    Set removals = private_PlanCleanup(snapshot)
    For Each entry In removals
        If merged.Count = 0 Then
            current = entry
        ElseIf entry(0) <= current(1) Then
            If entry(1) > current(1) Then
                current(1) = entry(1)
                current(4) = entry(4)
                current(5) = entry(5)
            End If
            merged.Remove merged.Count
        Else
            current = entry
        End If
        merged.Add current
    Next entry
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.Cleanup.Plan", startedAt

    ' Calibrate Word character positions once, then verify every candidate range.
    startedAt = VBA.Timer
    documentEnd = document.Content.End
    If merged.Count > 0 Then
        entry = merged(1)
        Set firstMarker = private_FindMarker(document, VBA.CStr(entry(2)))
        delta = firstMarker.Start - (entry(3) - 1)
    End If
    For Each entry In merged
        removed = VBA.Mid$(snapshot, entry(0), entry(1) - entry(0))
        rangeStart = entry(0) - 1 + delta
        rangeEnd = entry(1) - 1 + delta
        Set target = Nothing
        If rangeStart >= 0 And rangeEnd <= documentEnd Then
            Set target = document.Range(rangeStart, rangeEnd)
            target.TextRetrievalMode.IncludeHiddenText = True
            If VBA.StrComp(target.Text, removed, VBA.vbBinaryCompare) <> 0 Then
                Set target = Nothing
            End If
        End If
        If target Is Nothing Then
            ' Fields and objects can change the mapping; locate both marker boundaries.
            fallbackCount = fallbackCount + 1
            Set firstMarker = private_FindMarker(document, VBA.CStr(entry(2)))
            Set lastMarker = private_FindMarker(document, VBA.CStr(entry(4)))
            rangeStart = firstMarker.Start - (entry(3) - entry(0))
            rangeEnd = lastMarker.End + (entry(1) - entry(5))
            Set target = document.Range(rangeStart, rangeEnd)
            target.TextRetrievalMode.IncludeHiddenText = True
            If VBA.StrComp(target.Text, removed, VBA.vbBinaryCompare) <> 0 Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordMarkersRemovalInvalid")
            End If
            delta = firstMarker.Start - (entry(3) - 1)
        End If
        ranges.Add target
    Next entry
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.Cleanup.ResolveRanges", startedAt
    ex_Core.fn_Diagnostic_WriteLog "WORD_CLEANUP_PLAN | Ranges=" & VBA.CStr(ranges.Count) & _
        " | Fallbacks=" & VBA.CStr(fallbackCount)

    ' Delete from the end so all earlier verified ranges remain valid.
    startedAt = VBA.Timer
    For index = ranges.Count To 1 Step -1
        entry = merged(index)
        expected = VBA.Left$(expected, entry(0) - 1) & VBA.Mid$(expected, entry(1))
        Set target = ranges(index)
        target.Delete
    Next index
    ex_Core.fn_Diagnostic_WritePerf "Word.Export.Cleanup.DeleteRanges", startedAt
End Sub

Private Function private_PlanCleanup(ByVal snapshot As String) As Collection
    Dim removals As New Collection
    Dim identifier As Variant
    Dim positions As Variant
    Dim blockMarkers As Variant
    Dim opening As Variant
    Dim closing As Variant
    Dim first As Variant
    Dim last As Variant
    Dim whitespace As Object
    Dim blockText As String
    Dim entry As Variant
    Dim root As Variant
    Dim rootMarkers As Variant
    Dim rootFirst As Variant
    Dim rootLast As Variant
    Dim children As Collection
    Dim skipped As Object
    Dim populated As Boolean

    Set whitespace = VBA.CreateObject("VBScript.RegExp")
    whitespace.Global = True
    whitespace.Pattern = "[\s" & VBA.ChrW(160) & "]+"
    Set skipped = VBA.CreateObject("Scripting.Dictionary")
    ' Root headings count as metadata; only exported child content keeps a root alive.
    For Each root In m_roots.Keys
        rootMarkers = m_roots(root)
        rootFirst = private_TextRemoval(snapshot, CStr(rootMarkers(0)), True, True)
        rootLast = private_TextRemoval(snapshot, CStr(rootMarkers(1)), True, True)
        Set children = New Collection
        populated = False
        For Each identifier In m_anchors.Keys
            positions = m_anchors(identifier)
            opening = private_TextRemoval(snapshot, CStr(positions(0)), False, False)
            closing = private_TextRemoval(snapshot, CStr(positions(1)), False, False)
            If opening(0) >= rootFirst(1) And closing(1) <= rootLast(0) Then
                children.Add identifier
                blockText = Mid$(snapshot, opening(1), closing(0) - opening(1))
                If Len(whitespace.Replace(blockText, vbNullString)) > 0 Then
                    populated = True
                End If
            End If
        Next identifier
        If populated Then
            private_AddRemoval removals, rootFirst
            private_AddRemoval removals, rootLast
        Else
            entry = VBA.Array(rootFirst(0), rootLast(1), rootFirst(2), rootFirst(3), rootLast(4), rootLast(5))
            private_AddRemoval removals, entry
            For Each identifier In children
                skipped(identifier) = True
            Next identifier
        End If
    Next root
    For Each identifier In m_anchors.Keys
        If skipped.Exists(identifier) Then
            GoTo NextCleanupAnchor
        End If
        positions = m_anchors(identifier)
        opening = private_TextRemoval(snapshot, VBA.CStr(positions(0)), False, False)
        closing = private_TextRemoval(snapshot, VBA.CStr(positions(1)), False, False)
        If opening(1) > closing(0) Then
            Err.Raise 5, , "Reversed Word cleanup anchors: " & VBA.CStr(identifier)
        End If
        If m_blocks.Exists(identifier) Then
            blockMarkers = m_blocks(identifier)
            first = private_TextRemoval(snapshot, VBA.CStr(blockMarkers(0)), True, True)
            last = private_TextRemoval(snapshot, VBA.CStr(blockMarkers(1)), True, True)
            If first(1) > opening(0) Or closing(1) > last(0) Then
                Err.Raise 5, , "Word cleanup block does not enclose its export anchors."
            End If
            blockText = VBA.Mid$(snapshot, opening(1), closing(0) - opening(1))
            If VBA.Len(whitespace.Replace(blockText, VBA.vbNullString)) = 0 Then
                entry = VBA.Array(first(0), last(1), first(2), first(3), last(4), last(5))
                private_AddRemoval removals, entry
            Else
                private_AddRemoval removals, first
                private_AddRemoval removals, last
                entry = private_TextRemoval(snapshot, VBA.CStr(positions(0)), False, True)
                private_AddRemoval removals, entry
                entry = private_TextRemoval(snapshot, VBA.CStr(positions(1)), False, True)
                private_AddRemoval removals, entry
            End If
        Else
            ' Legacy export markers preserve their surrounding paragraphs.
            private_AddRemoval removals, opening
            private_AddRemoval removals, closing
        End If
NextCleanupAnchor:
    Next identifier
    Set private_PlanCleanup = removals
End Function

Private Function private_TextRemoval( _
    ByVal snapshot As String, _
    ByVal marker As String, _
    ByVal requireParagraph As Boolean, _
    ByVal removeStandalone As Boolean _
) As Variant
    Dim start As Long
    Dim finish As Long
    Dim paragraphStart As Long
    Dim paragraphEnd As Long
    Dim remaining As String
    Dim removalStart As Long
    Dim removalEnd As Long

    start = VBA.InStr(1, snapshot, marker, VBA.vbBinaryCompare)
    If start = 0 Then
        Err.Raise 5, , "Word cleanup marker is missing: " & marker
    End If
    If VBA.InStr(start + 1, snapshot, marker, VBA.vbBinaryCompare) > 0 Then
        Err.Raise 5, , "Duplicate Word cleanup marker: " & marker
    End If
    finish = start + VBA.Len(marker)
    removalStart = start
    removalEnd = finish
    If removeStandalone Then
        paragraphStart = VBA.InStrRev(VBA.Left$(snapshot, start - 1), VBA.vbCr) + 1
        paragraphEnd = VBA.InStr(finish, snapshot, VBA.vbCr)
        If paragraphEnd = 0 Then
            paragraphEnd = VBA.Len(snapshot) + 1
        End If
        remaining = VBA.Mid$(snapshot, paragraphStart, start - paragraphStart) & _
            VBA.Mid$(snapshot, finish, paragraphEnd - finish)
        remaining = VBA.Replace$(remaining, VBA.vbTab, VBA.vbNullString)
        remaining = VBA.Replace$(remaining, VBA.ChrW(160), " ")
        If VBA.Len(VBA.Trim$(remaining)) = 0 Then
            removalStart = paragraphStart
            removalEnd = paragraphEnd + 1
            ' Preserve the mandatory final document paragraph mark.
            If removalEnd > VBA.Len(snapshot) Then
                removalEnd = VBA.Len(snapshot)
            End If
        ElseIf requireParagraph Then
            Err.Raise 5, , "Word block marker must occupy a separate paragraph: " & marker
        End If
    End If
    private_TextRemoval = VBA.Array(removalStart, removalEnd, marker, start, marker, finish)
End Function

Private Sub private_AddRemoval( _
    ByVal removals As Collection, _
    ByVal entry As Variant _
)
    Dim index As Long
    Dim existing As Variant

    For index = 1 To removals.Count
        existing = removals(index)
        If entry(0) < existing(0) Then
            If entry(1) > existing(0) Then
                Err.Raise 5, , "Overlapping Word cleanup ranges."
            End If
            removals.Add entry, , index
            Exit Sub
        End If
        If entry(0) < existing(1) Then
            Err.Raise 5, , "Overlapping Word cleanup ranges."
        End If
    Next index
    removals.Add entry
End Sub

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