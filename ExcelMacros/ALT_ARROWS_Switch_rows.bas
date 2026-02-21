Sub MoveRowUp()
    Dim r As Long
    r = Selection.Row
    If r <= 1 Then Exit Sub

    Rows(r).Cut
    Rows(r - 1).Insert Shift:=xlDown
    Application.CutCopyMode = False
    Rows(r - 1).Select
End Sub

Sub MoveRowDown()
    Dim r As Long
    r = Selection.Row
    If r >= Rows.Count Then Exit Sub

    Rows(r).Cut
    ' ÂÀÆÍÎ: ïîñëå âûðåçàíèÿ ñòðîêà íèæå ïîäíèìàåòñÿ ââåðõ,
    ' ïîýòîìó äëÿ øàãà "âíèç íà 1" âñòàâëÿåì â r+2
    Rows(r + 2).Insert Shift:=xlDown
    Application.CutCopyMode = False
    Rows(r + 1).Select
End Sub

Sub BindKeys()
    Application.OnKey "%{UP}", "MoveRowUp"
    Application.OnKey "%{DOWN}", "MoveRowDown"
End Sub