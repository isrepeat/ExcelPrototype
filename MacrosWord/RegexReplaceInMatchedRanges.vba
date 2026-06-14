Sub RegexReplaceInMatchedRanges()

    Dim rangePattern As String
    Dim replacePattern As String
    Dim replacement As String

    Dim reRanges As Object
    Dim reReplace As Object

    Dim matches As Object
    Dim m As Object

    Dim docText As String
    Dim rangeText As String
    Dim newRangeText As String

    Dim rng As Range
    Dim docStart As Long

    rangePattern = InputBox( _
        "Enter regex to identify target ranges:", _
        "Range Pattern")

    If rangePattern = "" Then Exit Sub

    replacePattern = InputBox( _
        "Enter regex to replace inside those ranges:", _
        "Replace Pattern")

    If replacePattern = "" Then Exit Sub

    replacement = InputBox( _
        "Enter replacement text ($1, $2, ... supported):", _
        "Replacement")

    docText = ActiveDocument.Content.Text
    docStart = ActiveDocument.Content.Start

    Set reRanges = CreateObject("VBScript.RegExp")
    reRanges.Global = True
    reRanges.Multiline = True
    reRanges.IgnoreCase = False
    reRanges.pattern = rangePattern

    Set reReplace = CreateObject("VBScript.RegExp")
    reReplace.Global = True
    reReplace.Multiline = True
    reReplace.IgnoreCase = False
    reReplace.pattern = replacePattern

    Set matches = reRanges.Execute(docText)

    Dim i As Long

    For i = matches.count - 1 To 0 Step -1

        Set m = matches(i)

        Set rng = ActiveDocument.Range( _
            Start:=docStart + m.FirstIndex, _
            End:=docStart + m.FirstIndex + m.Length)

        rangeText = rng.Text

        newRangeText = reReplace.Replace(rangeText, replacement)

        rng.Text = newRangeText

    Next i

    MsgBox matches.count & " range(s) processed.", _
           vbInformation, _
           "Regex Replace"

End Sub

