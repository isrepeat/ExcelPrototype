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
    m_sourceText = document.Content.Text
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

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Word anchor index is not initialized."
    identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
    If Not m_anchors.Exists(identifier) Then Err.Raise 5, , "Word template anchor not found: " & identifier
    Set m_insertDocument = Nothing
    m_expectedText = VBA.vbNullString
    positions = m_anchors(identifier)
    ' FormattedText can normalize a copied document; verify against its own baseline.
    currentText = document.Content.Text
    beginOffset = VBA.InStr(1, currentText, positions(0), VBA.vbBinaryCompare)
    endOffset = VBA.InStr(1, currentText, positions(1), VBA.vbBinaryCompare)
    If beginOffset = 0 Or endOffset = 0 Then Err.Raise 5, , "Word export marker is missing: " & identifier
    If VBA.InStr(beginOffset + 1, currentText, positions(0), VBA.vbBinaryCompare) > 0 Or _
            VBA.InStr(endOffset + 1, currentText, positions(1), VBA.vbBinaryCompare) > 0 Then
        Err.Raise 5, , "Duplicate Word export marker: " & identifier
    End If
    rendered = VBA.Replace$(VBA.CStr(values("WordText")), VBA.vbCrLf, VBA.vbCr)
    rendered = VBA.Replace$(rendered, VBA.vbLf, VBA.vbCr)
    Set beginRange = private_FindMarker(document, VBA.CStr(positions(0)))
    Set endRange = private_FindMarker(document, VBA.CStr(positions(1)))
    If beginRange.End > endRange.Start Then Err.Raise 5, , "Reversed Word anchor ranges: " & identifier
    Set target = document.Range(beginRange.End, endRange.Start)
    Set font = beginRange.Characters(1).Font.Duplicate
    Set paragraphFormat = beginRange.Paragraphs(1).Format.Duplicate
    m_expectedText = VBA.Left$(currentText, beginOffset + VBA.Len(positions(0)) - 1) & _
        rendered & VBA.Mid$(currentText, endOffset)
    target.Text = rendered
    target.Font = font
    target.ParagraphFormat = paragraphFormat
    Set m_insertDocument = document
    TryInsert = True
    Exit Function
EH:
    diagnostic = "Word insertion: " & VBA.Err.Description
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
    If VBA.StrComp(document.Content.Text, expected, VBA.vbBinaryCompare) <> 0 Then
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

' //
' // Private
' //
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