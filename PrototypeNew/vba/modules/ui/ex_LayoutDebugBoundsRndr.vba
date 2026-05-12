Attribute VB_Name = "ex_LayoutDebugBoundsRndr"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private Const DEBUG_LAYOUT_BOUNDS_ENABLED As Boolean = True

' Отложенные debug-границы layout-узлов.
' Формат записи: Array(sheetName, rowStart, colStart, rowEnd, colEnd, kind, name, debugStyle)
Private m_PendingDebugBounds As Collection

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_LayoutDebugBoundsRndr.fn_Module_Dispose"
#End If
    Set m_PendingDebugBounds = Nothing
End Sub

Public Sub fn_ResetDebugBounds()
    Set m_PendingDebugBounds = Nothing
End Sub

Public Sub fn_RegisterDebugBounds( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal nodeKind As String = "", _
    Optional ByVal nodeName As String = "", _
    Optional ByVal debugStyle As String = "" _
)
    If Not DEBUG_LAYOUT_BOUNDS_ENABLED Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Then Exit Sub
    If rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    If m_PendingDebugBounds Is Nothing Then
        Set m_PendingDebugBounds = New Collection
    End If

    m_PendingDebugBounds.Add Array( _
        ws.Name, _
        CLng(rowStart), _
        CLng(colStart), _
        CLng(rowEnd), _
        CLng(colEnd), _
        VBA.Trim$(nodeKind), _
        VBA.Trim$(nodeName), _
        VBA.Trim$(debugStyle))
End Sub

Public Sub fn_ApplyPendingDebugBounds(ByVal ws As Worksheet)
    Dim entry As Variant

    If Not DEBUG_LAYOUT_BOUNDS_ENABLED Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If m_PendingDebugBounds Is Nothing Then Exit Sub

    On Error GoTo EH_APPLY
    For Each entry In m_PendingDebugBounds
        If VBA.StrComp(VBA.CStr(entry(0)), ws.Name, VBA.vbTextCompare) = 0 Then
            private_PaintDebugFrame _
                ws, _
                CLng(entry(1)), _
                CLng(entry(2)), _
                CLng(entry(3)), _
                CLng(entry(4)), _
                VBA.CStr(entry(7))
        End If
    Next entry
    Exit Sub

EH_APPLY:
    On Error GoTo 0
End Sub

Private Sub private_PaintDebugFrame( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal debugStyle As String = "" _
)
    Dim targetRange As Range
    Dim borderWeight As Variant
    Dim borderColor As Long

    On Error GoTo EH_FRAME
    Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
    If targetRange Is Nothing Then Exit Sub

    private_ResolveDebugBorderStyle debugStyle, borderWeight, borderColor

    With targetRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = borderWeight
        .Color = borderColor
    End With
    With targetRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = borderWeight
        .Color = borderColor
    End With
    With targetRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = borderWeight
        .Color = borderColor
    End With
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = borderWeight
        .Color = borderColor
    End With
    Exit Sub

EH_FRAME:
    On Error GoTo 0
End Sub

Private Sub private_ResolveDebugBorderStyle( _
    ByVal styleText As String, _
    ByRef outWeight As Variant, _
    ByRef outColor As Long _
)
    Dim weightToken As String
    Dim colorToken As String

    outWeight = xlThin
    outColor = VBA.RGB(128, 0, 255)

    styleText = VBA.Trim$(styleText)
    If VBA.Len(styleText) = 0 Then Exit Sub

    ' Поддерживаем только borderWeight и borderColor.
    If ex_HelpersCSS.fn_TryReadStyleValue(styleText, "borderweight", weightToken) Then
        Call ex_HelpersCSS.fn_TryParseCellBorderWeight(weightToken, outWeight)
    End If

    If ex_HelpersCSS.fn_TryReadStyleValue(styleText, "bordercolor", colorToken) Then
        Call ex_HelpersCSS.fn_TryParseColor(colorToken, outColor)
    End If
End Sub
