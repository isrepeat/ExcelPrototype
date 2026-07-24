Sub fn_RecalculateWorkbook()
    Application.Calculate
End Sub

Sub fn_RecalculateActiveSheet()
    ActiveSheet.Calculate
End Sub


Public Sub fn_TestHotkey()

    ActiveCell.Interior.Color = RGB(255, 0, 0)

End Sub

Public Sub fn_ToggleFirstTwoRows()

    On Error GoTo ExitPoint

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 0
        .SplitColumn = 0
    End With

    If ws.Rows("1:2").Hidden Then

        ws.Rows("1:2").Hidden = False

        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1

        ws.Range("A3").Select
        ActiveWindow.FreezePanes = True

    Else

        ws.Rows("1:2").Hidden = True

        ActiveWindow.ScrollRow = 3
        ActiveWindow.ScrollColumn = 1

    End If

ExitPoint:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Sub fn_DatePlusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) + 1
    End If
End Sub

Sub fn_DateMinusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) - 1
    End If
End Sub