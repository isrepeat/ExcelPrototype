VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_LabelControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_TextRaw As String
Private m_TextResolved As String
Private m_ControlLayout As obj_ControlLayout
Private m_IsConfigured As Boolean
Private m_MergeCells As Boolean
Private m_Page As obj_IPage

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    obj_IControl_Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IControl_Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    m_MergeCells = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim dataContext As Object
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_ControlBase = Nothing

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Label", "label", m_ControlName) Then Exit Sub

    m_TextRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "text"))
    If VBA.Len(VBA.Trim$(m_TextRaw)) = 0 Then
        m_TextRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "caption"))
    End If

    Set dataContext = m_ControlBase.DataContext
    If dataContext Is Nothing Then Set dataContext = m_Page
    If Not ex_BindingRuntime.fn_TryResolveTextBinding(m_TextRaw, dataContext, m_TextResolved) Then Exit Sub
    If Not private_TryReadMergeCells(controlNode, m_MergeCells) Then Exit Sub

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Label", m_ControlName, "style") Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim pageBase As obj_PageBase
    Dim previousDisplayAlerts As Boolean

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Label: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If
    previousDisplayAlerts = Application.DisplayAlerts

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Label: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Label: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), ws.Cells(m_ControlLayout.RowEnd, m_ControlLayout.ColEnd))
    On Error GoTo 0

    If m_MergeCells And targetRange.Cells.CountLarge > 1 Then
        Application.DisplayAlerts = False
        targetRange.ClearContents
        targetRange.Merge
        Application.DisplayAlerts = previousDisplayAlerts
    End If
    targetRange.Cells(1, 1).Value2 = m_TextResolved
    targetRange.HorizontalAlignment = xlHAlignLeft
    targetRange.VerticalAlignment = xlVAlignCenter
    targetRange.WrapText = False
    If Not private_ApplyPresetStyle(targetRange, m_ControlLayout.StyleName) Then Exit Sub
    If Not private_RegisterControlPart(ws, targetRange) Then Exit Sub
    Exit Sub

EH_RANGE:
    On Error Resume Next
    Application.DisplayAlerts = previousDisplayAlerts
    On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Label: failed to resolve target range for control '" & m_ControlName & "'."
#End If
End Sub

Private Function obj_IControl_Measure( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    outSpanRows = 1
    outSpanColls = 1
    obj_IControl_Measure = True
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "text", "caption", "mergecells"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function private_TryReadMergeCells( _
    ByVal controlNode As Object, _
    ByRef outValue As Boolean _
) As Boolean
    Dim rawValue As String

    outValue = False
    rawValue = VBA.LCase$(VBA.Trim$(VBA.CStr( _
        ex_XmlCore.fn_NodeAttrText(controlNode, "mergeCells"))))
    If VBA.Len(rawValue) = 0 Or rawValue = "false" Or rawValue = "0" Then
        private_TryReadMergeCells = True
        Exit Function
    End If
    If rawValue = "true" Or rawValue = "1" Then
        outValue = True
        private_TryReadMergeCells = True
        Exit Function
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "Label: mergeCells must be true or false for control '" & _
        m_ControlName & "'."
#End If
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

' //
' // API
' //
' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_ApplyPresetStyle(ByVal targetRange As Range, ByVal styleName As String) As Boolean
    If targetRange Is Nothing Then Exit Function

    Select Case VBA.LCase$(VBA.Trim$(styleName))
        Case VBA.vbNullString
            ' no-op

        Case "tablesection"
            targetRange.Interior.Color = VBA.RGB(23, 58, 94)
            targetRange.Font.Color = VBA.RGB(234, 246, 255)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(14, 34, 57)
            targetRange.Borders.Weight = xlThin

        Case "tableheadercell"
            targetRange.Interior.Color = VBA.RGB(43, 74, 107)
            targetRange.Font.Color = VBA.RGB(221, 238, 255)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(31, 54, 80)
            targetRange.Borders.Weight = xlThin

        Case "tabledatacell"
            targetRange.Interior.Color = VBA.RGB(58, 58, 58)
            targetRange.Font.Color = VBA.RGB(240, 240, 240)
            targetRange.Font.Bold = False
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(42, 42, 42)
            targetRange.Borders.Weight = xlThin

        Case "tablespacer"
            targetRange.Interior.Color = VBA.RGB(31, 31, 31)
            targetRange.Font.Color = VBA.RGB(31, 31, 31)
            targetRange.Font.Bold = False
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = VBA.RGB(31, 31, 31)
            targetRange.Borders.Weight = xlHairline

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Label: unsupported style '" & styleName & "' for control '" & m_ControlName & "'."
#End If
            Exit Function
    End Select

    private_ApplyPresetStyle = True
End Function

Private Function private_RegisterControlPart( _
    ByVal ws As Worksheet, _
    ByVal labelRange As Range _
) As Boolean
    If ws Is Nothing Then Exit Function
    If labelRange Is Nothing Then Exit Function

    private_RegisterControlPart = ex_ControlPartsRuntime.fn_RegisterControlPart( _
        ws, _
        "label", _
        m_ControlName, _
        "cell", _
        labelRange)
End Function

Private Function private_GetWorksheetByName(ByVal page As obj_PageBase, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If page Is Nothing Then Exit Function
    Set ws = page.Worksheet
    If ws Is Nothing Then Exit Function

    sheetName = VBA.LCase$(VBA.Trim$(sheetName))
    If VBA.Len(sheetName) > 0 Then
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(ws.Name)), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    End If

    Set private_GetWorksheetByName = ws
End Function
