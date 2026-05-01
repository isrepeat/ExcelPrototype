Attribute VB_Name = "ex_ControlRefreshRuntime"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private g_ControlRegistry As Object

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_ControlRefreshRuntime.m_Module_Dispose"
#End If
    On Error Resume Next
    Set g_ControlRegistry = Nothing
    On Error GoTo 0
End Sub
' //
' // API
' //
Public Sub m_ResetRegisteredControls()
    Set g_ControlRegistry = private_CreateDictionary()
End Sub


Public Sub m_RegisterControlRenderBounds( _
    ByVal controlName As String, _
    ByVal controlType As String, _
    ByVal sheetName As String, _
    ByVal uiPath As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
)
    Dim key As String
    Dim entry As Object

    controlName = VBA.Trim$(controlName)
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    sheetName = VBA.Trim$(sheetName)
    uiPath = VBA.Trim$(uiPath)

    If VBA.Len(controlName) = 0 Then Exit Sub
    If VBA.Len(controlType) = 0 Then Exit Sub
    If VBA.Len(sheetName) = 0 Then Exit Sub
    If VBA.Len(uiPath) = 0 Then Exit Sub
    If rowStart <= 0 Or colStart <= 0 Then Exit Sub
    If rowEnd < rowStart Or colEnd < colStart Then Exit Sub

    private_EnsureRegistry
    key = VBA.LCase$(controlName)

    Set entry = private_CreateDictionary()
    entry("Name") = controlName
    entry("Type") = controlType
    entry("Sheet") = sheetName
    entry("UiPath") = uiPath
    entry("RowStart") = VBA.CLng(rowStart)
    entry("ColStart") = VBA.CLng(colStart)
    entry("RowEnd") = VBA.CLng(rowEnd)
    entry("ColEnd") = VBA.CLng(colEnd)

    If g_ControlRegistry.Exists(key) Then g_ControlRegistry.Remove key
    g_ControlRegistry.Add key, entry
End Sub


Public Function m_TryRefreshStaticControl(ByVal controlName As String) As Boolean
    Dim key As String
    Dim entry As Object
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim pageRef As obj_IPage
    Dim pageBase As obj_PageBase
    Dim uiDoc As Object
    Dim controlNode As Object
    Dim escapedName As String
    Dim xPath As String
    Dim renderCtx As obj_LayoutRenderContext

    controlName = VBA.Trim$(controlName)
    If VBA.Len(controlName) = 0 Then Exit Function

    private_EnsureRegistry

    key = VBA.LCase$(controlName)
    If Not g_ControlRegistry.Exists(key) Then Exit Function

    Set entry = g_ControlRegistry(key)
    If entry Is Nothing Then Exit Function
    If Not private_IsStaticControlType(VBA.CStr(entry("Type"))) Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VBA.CStr(entry("Sheet")))
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set wb = ws.Parent
    If wb Is Nothing Then Exit Function

    Set uiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        VBA.CStr(entry("UiPath")), _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If uiDoc Is Nothing Then Exit Function

    escapedName = ex_XmlCore.m_XPathLiteral(VBA.CStr(entry("Name")))
    xPath = "/p:page//p:control[@name=" & escapedName & "] | /p:uiDefinition/p:layout//p:control[@name=" & escapedName & "]"
    Set controlNode = uiDoc.selectSingleNode(xPath)
    If controlNode Is Nothing Then Exit Function

    If Not rt_PageManager.m_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    Set pageBase = pageRef.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set renderCtx = New obj_LayoutRenderContext
    If Not renderCtx.Initialize(pageBase) Then Exit Function

    If Not ex_XmlLayoutEngine.m_RenderNodeInBounds( _
        renderCtx:=renderCtx, _
        layoutNode:=controlNode, _
        rowStart:=VBA.CLng(entry("RowStart")), _
        colStart:=VBA.CLng(entry("ColStart")), _
        rowEnd:=VBA.CLng(entry("RowEnd")), _
        colEnd:=VBA.CLng(entry("ColEnd"))) Then Exit Function

    If Not pageBase.ApplyInlineRuns() Then Exit Function

    m_TryRefreshStaticControl = True
End Function

' //
' // Internal
' //

Private Sub private_EnsureRegistry()
    If g_ControlRegistry Is Nothing Then
        Set g_ControlRegistry = private_CreateDictionary()
    End If
End Sub


Private Function private_CreateDictionary() As Object
    Set private_CreateDictionary = VBA.CreateObject("Scripting.Dictionary")
    private_CreateDictionary.CompareMode = 1
End Function


Private Function private_IsStaticControlType(ByVal controlType As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(controlType))
        Case "label", "banner", "button", "config", "select"
            private_IsStaticControlType = True
    End Select
End Function
