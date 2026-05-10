Attribute VB_Name = "ex_LayoutControlFallbackRndr"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Отложенные fallback-диапазоны (рисуются после style pipeline, чтобы не затиралось стилями страницы).
' Формат каждого элемента: Array(sheetName, rowStart, colStart, rowEnd, colEnd, controlName)
Private m_PendingFallbackRanges As Collection

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_LayoutControlFallbackRndr.fn_Module_Dispose"
#End If
    Set m_PendingFallbackRanges = Nothing
End Sub

' //
' // API
' //
Public Sub fn_ResetControlFallbacks()
    Set m_PendingFallbackRanges = Nothing
End Sub


Public Sub fn_RegisterControlFallback( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal controlName As String = "" _
)
    If ws Is Nothing Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Then Exit Sub
    If rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    private_RegisterControlFallbackRange ws.Name, rowStart, colStart, rowEnd, colEnd, controlName
End Sub


Public Sub fn_ApplyPendingControlFallbacks(ByVal ws As Worksheet)
    Dim entry As Variant

    If ws Is Nothing Then Exit Sub
    If m_PendingFallbackRanges Is Nothing Then Exit Sub

    On Error GoTo EH_APPLY
    For Each entry In m_PendingFallbackRanges
        If VBA.LCase$(VBA.CStr(entry(0))) = VBA.LCase$(ws.Name) Then
            private_PaintControlFallbackRange ws, CLng(entry(1)), CLng(entry(2)), CLng(entry(3)), CLng(entry(4)), VBA.CStr(entry(5))
        End If
    Next entry
    Exit Sub

EH_APPLY:
    On Error GoTo 0
End Sub

' //
' // Internal
' //
Private Sub private_RegisterControlFallbackRange( _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal controlName As String = "" _
)
    If VBA.Len(sheetName) = 0 Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Then Exit Sub
    If rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    If m_PendingFallbackRanges Is Nothing Then
        Set m_PendingFallbackRanges = New Collection
    End If

    m_PendingFallbackRanges.Add Array(sheetName, rowStart, colStart, rowEnd, colEnd, controlName)
End Sub


Private Sub private_PaintControlFallbackRange( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal controlName As String = "" _
)
    Dim targetRange As Range
    Dim captionText As String

    On Error GoTo EH_FALLBACK
    Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    If targetRange Is Nothing Then Exit Sub

    ' Убираем условное форматирование в fallback-области, иначе заливка может визуально оставаться белой.
    On Error Resume Next
    targetRange.FormatConditions.Delete
    On Error GoTo EH_FALLBACK

    ' Сводим область контрола в одну ячейку, чтобы показать диагностический текст.
    On Error Resume Next
    targetRange.UnMerge
    On Error GoTo EH_FALLBACK
    targetRange.Merge

    captionText = VBA.Trim$(controlName)
    If VBA.Len(captionText) = 0 Then captionText = "Unconfigured control"

    targetRange.Interior.Pattern = xlSolid
    targetRange.Interior.TintAndShade = 0
    targetRange.Interior.Color = RGB(255, 87, 107)
    targetRange.Borders.LineStyle = xlContinuous
    targetRange.Borders.Weight = xlThin
    targetRange.Borders.Color = RGB(200, 90, 90)
    targetRange.Cells(1, 1).Value2 = captionText
    targetRange.HorizontalAlignment = xlLeft
    targetRange.VerticalAlignment = xlTop
    targetRange.WrapText = True
    targetRange.Font.Bold = True
    targetRange.Font.Color = RGB(156, 0, 6)
    Exit Sub

EH_FALLBACK:
    On Error GoTo 0
End Sub
