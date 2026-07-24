VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UaLocationInflector"
Option Explicit
Implements obj_IUaInflector

Private Const CASE_GENITIVE As String = "genitive"
Private Const CASE_ACCUSATIVE As String = "accusative"
Private Const CASE_DATIVE As String = "dative"

' Склоняет административные части украинского адреса отдельно от названия
' населённого пункта. Например, "с. Костів" остаётся без изменений, а
' "Богодухівський р-н" и "Харківська обл." получают согласованные формы.
Public Function TryInflect( _
    ByVal sourceText As String, _
    ByVal targetCase As String, _
    ByRef outText As String _
) As Boolean
    Dim segments As Variant
    Dim segmentIndex As Long
    Dim caseName As String

    outText = sourceText
    caseName = VBA.LCase$(VBA.Trim$(targetCase))
    If Not private_IsSupportedCase(caseName) Then Exit Function

    sourceText = VBA.Trim$(sourceText)
    If VBA.Len(sourceText) = 0 Then
        outText = sourceText
        TryInflect = True
        Exit Function
    End If

    segments = VBA.Split(sourceText, ",")
    For segmentIndex = LBound(segments) To UBound(segments)
        segments(segmentIndex) = private_InflectLocationSegment( _
            VBA.Trim$(VBA.CStr(segments(segmentIndex))), caseName)
    Next segmentIndex

    outText = VBA.Join(segments, ", ")
    TryInflect = True
End Function

Private Function obj_IUaInflector_TryInflect( _
    ByVal sourceText As String, _
    ByVal targetCase As String, _
    ByRef outText As String _
) As Boolean
    obj_IUaInflector_TryInflect = TryInflect(sourceText, targetCase, outText)
End Function

Private Function private_InflectLocationSegment( _
    ByVal segmentText As String, _
    ByVal caseName As String _
) As String
    Dim modifierText As String
    Dim inflectedSegment As String

    private_InflectLocationSegment = segmentText
    If VBA.Len(segmentText) = 0 Then Exit Function

    ' В источниках район и область не всегда разделены запятыми. Выделяем
    ' административные блоки по их маркерам, склоняем независимо и затем
    ' восстанавливаем канонический разделитель между блоками.
    If private_TryInflectAdministrativeBlocks( _
        segmentText, caseName, inflectedSegment) Then
        private_InflectLocationSegment = inflectedSegment
        Exit Function
    End If

    If private_TryRemoveAdministrativeSuffix( _
        segmentText, "р-н\.?|район", modifierText) Then
        private_InflectLocationSegment = private_BuildAdministrativeSegment( _
            modifierText, caseName, False)
        Exit Function
    End If

    If private_TryRemoveAdministrativeSuffix( _
        segmentText, "обл\.?|область", modifierText) Then
        private_InflectLocationSegment = private_BuildAdministrativeSegment( _
            modifierText, caseName, True)
    End If
End Function

Private Function private_TryInflectAdministrativeBlocks( _
    ByVal segmentText As String, _
    ByVal caseName As String, _
    ByRef outText As String _
) As Boolean
    Dim phraseRx As Object
    Dim matches As Object
    Dim matchIndex As Long
    Dim matchObj As Object
    Dim blocks As Collection
    Dim blockValues() As String
    Dim blockIndex As Long
    Dim cursorIndex As Long
    Dim prefixText As String
    Dim modifierWord As String
    Dim suffixText As String
    Dim replacementText As String
    Dim trailingText As String
    Dim isFeminine As Boolean

    outText = segmentText
    Set phraseRx = VBA.CreateObject("VBScript.RegExp")
    phraseRx.Global = True
    phraseRx.IgnoreCase = True
    ' Маркер административной единицы задаёт правую границу блока. Поэтому
    ' отсутствие запятых не мешает отделить район и область от населённого пункта.
    phraseRx.Pattern = "([^ \t,]+)[ \t]+(р-н\.?|район|обл\.?|область)"
    Set matches = phraseRx.Execute(segmentText)
    If matches.Count = 0 Then Exit Function

    Set blocks = New Collection
    cursorIndex = 0
    For matchIndex = 0 To matches.Count - 1
        Set matchObj = matches.Item(matchIndex)
        prefixText = private_TrimBlockDelimiter(VBA.Mid$( _
            segmentText, cursorIndex + 1, matchObj.FirstIndex - cursorIndex))
        If VBA.Len(prefixText) > 0 Then blocks.Add prefixText

        modifierWord = VBA.CStr(matchObj.SubMatches(0))
        suffixText = VBA.LCase$(VBA.CStr(matchObj.SubMatches(1)))
        isFeminine = (VBA.InStr(1, suffixText, "обл", VBA.vbTextCompare) > 0)
        replacementText = private_BuildAdministrativeSegment( _
            modifierWord, caseName, isFeminine)
        blocks.Add replacementText
        cursorIndex = matchObj.FirstIndex + matchObj.Length
    Next matchIndex

    trailingText = private_TrimBlockDelimiter( _
        VBA.Mid$(segmentText, cursorIndex + 1))
    If VBA.Len(trailingText) > 0 Then blocks.Add trailingText

    ReDim blockValues(0 To blocks.Count - 1)
    For blockIndex = 1 To blocks.Count
        blockValues(blockIndex - 1) = VBA.CStr(blocks.Item(blockIndex))
    Next blockIndex
    outText = VBA.Join(blockValues, ", ")
    private_TryInflectAdministrativeBlocks = True
End Function

Private Function private_TrimBlockDelimiter(ByVal blockText As String) As String
    blockText = VBA.Trim$(blockText)
    Do While VBA.Len(blockText) > 0 And VBA.Left$(blockText, 1) = ","
        blockText = VBA.Trim$(VBA.Mid$(blockText, 2))
    Loop
    Do While VBA.Len(blockText) > 0 And VBA.Right$(blockText, 1) = ","
        blockText = VBA.Trim$(VBA.Left$(blockText, VBA.Len(blockText) - 1))
    Loop
    private_TrimBlockDelimiter = blockText
End Function

Private Function private_TryRemoveAdministrativeSuffix( _
    ByVal segmentText As String, _
    ByVal suffixPattern As String, _
    ByRef outModifierText As String _
) As Boolean
    Dim suffixRx As Object
    Dim matches As Object

    outModifierText = VBA.vbNullString
    Set suffixRx = VBA.CreateObject("VBScript.RegExp")
    suffixRx.Global = False
    suffixRx.IgnoreCase = True
    ' VBScript.RegExp не поддерживает все современные конструкции regex,
    ' поэтому используем обычную capturing group вместо (?:...).
    suffixRx.Pattern = "^(.*?)[ \t]+(" & suffixPattern & ")[ \t]*$"

    Set matches = suffixRx.Execute(segmentText)
    If matches.Count = 0 Then Exit Function

    outModifierText = VBA.Trim$(VBA.CStr(matches(0).SubMatches(0)))
    private_TryRemoveAdministrativeSuffix = (VBA.Len(outModifierText) > 0)
End Function

Private Function private_BuildAdministrativeSegment( _
    ByVal modifierText As String, _
    ByVal caseName As String, _
    ByVal isFeminine As Boolean _
) As String
    Dim inflectedModifier As String
    Dim administrativeNoun As String

    inflectedModifier = private_InflectLastModifierWord( _
        modifierText, caseName, isFeminine)

    If isFeminine Then
        Select Case caseName
            Case CASE_GENITIVE: administrativeNoun = "області"
            Case CASE_ACCUSATIVE: administrativeNoun = "область"
            Case CASE_DATIVE: administrativeNoun = "області"
        End Select
    Else
        Select Case caseName
            Case CASE_GENITIVE: administrativeNoun = "району"
            Case CASE_ACCUSATIVE: administrativeNoun = "район"
            Case CASE_DATIVE: administrativeNoun = "району"
        End Select
    End If

    private_BuildAdministrativeSegment = _
        VBA.Trim$(inflectedModifier & " " & administrativeNoun)
End Function

Private Function private_InflectLastModifierWord( _
    ByVal modifierText As String, _
    ByVal caseName As String, _
    ByVal isFeminine As Boolean _
) As String
    Dim separatorIndex As Long
    Dim prefixText As String
    Dim modifierWord As String

    modifierText = VBA.Trim$(modifierText)
    separatorIndex = VBA.InStrRev(modifierText, " ")
    If separatorIndex > 0 Then
        prefixText = VBA.Left$(modifierText, separatorIndex)
        modifierWord = VBA.Mid$(modifierText, separatorIndex + 1)
    Else
        modifierWord = modifierText
    End If

    modifierWord = private_InflectModifierWord(modifierWord, caseName, isFeminine)
    private_InflectLastModifierWord = prefixText & modifierWord
End Function

Private Function private_InflectModifierWord( _
    ByVal modifierWord As String, _
    ByVal caseName As String, _
    ByVal isFeminine As Boolean _
) As String
    Dim lowerWord As String
    Dim inflectedLowerWord As String

    lowerWord = VBA.LCase$(modifierWord)
    inflectedLowerWord = lowerWord

    If isFeminine Then
        If VBA.Right$(lowerWord, 1) = "а" Then
            Select Case caseName
                Case CASE_GENITIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 1) & "ої"
                Case CASE_ACCUSATIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 1) & "у"
                Case CASE_DATIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 1) & "ій"
            End Select
        End If
    Else
        If VBA.Right$(lowerWord, 2) = "ий" Then
            Select Case caseName
                Case CASE_GENITIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 2) & "ого"
                Case CASE_DATIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 2) & "ому"
            End Select
        ElseIf VBA.Right$(lowerWord, 2) = "ій" Then
            Select Case caseName
                Case CASE_GENITIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 2) & "ього"
                Case CASE_DATIVE
                    inflectedLowerWord = VBA.Left$(lowerWord, VBA.Len(lowerWord) - 2) & "ьому"
            End Select
        End If
    End If

    private_InflectModifierWord = private_ApplySourceCase( _
        modifierWord, inflectedLowerWord)
End Function

Private Function private_ApplySourceCase( _
    ByVal sourceWord As String, _
    ByVal lowerWord As String _
) As String
    If sourceWord = VBA.UCase$(sourceWord) Then
        private_ApplySourceCase = VBA.UCase$(lowerWord)
    ElseIf VBA.Left$(sourceWord, 1) = VBA.UCase$(VBA.Left$(sourceWord, 1)) Then
        private_ApplySourceCase = VBA.UCase$(VBA.Left$(lowerWord, 1)) & _
            VBA.Mid$(lowerWord, 2)
    Else
        private_ApplySourceCase = lowerWord
    End If
End Function

Private Function private_IsSupportedCase(ByVal caseName As String) As Boolean
    private_IsSupportedCase = _
        (caseName = CASE_GENITIVE Or _
         caseName = CASE_ACCUSATIVE Or _
         caseName = CASE_DATIVE)
End Function
