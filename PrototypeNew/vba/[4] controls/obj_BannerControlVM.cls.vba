VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_BannerControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_BannerViewItem As obj_BannerViewItem
Private m_ControlLayout As obj_ControlLayout
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim headerText As String
    Dim messageText As String
    Dim visibleRaw As String
    Dim isVisible As Boolean
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_BannerViewItem = Nothing
    Set m_ControlBase = Nothing

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Banner", "banner", m_ControlName) Then Exit Sub

    headerText = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "header"))
    messageText = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "message"))
    visibleRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "visible")))

    If VBA.Len(visibleRaw) = 0 Then
        isVisible = (VBA.Len(VBA.Trim$(headerText)) > 0 Or VBA.Len(VBA.Trim$(messageText)) > 0)
    Else
        isVisible = private_ParseBooleanText(visibleRaw)
    End If

    Set m_BannerViewItem = New obj_BannerViewItem
    m_BannerViewItem.Model.Header = headerText
    m_BannerViewItem.Model.Message = messageText
    m_BannerViewItem.Model.Visible = isVisible
    m_BannerViewItem.Presentation.EffectiveVisible = isVisible

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Banner", m_ControlName, "style") Then Exit Sub
    m_BannerViewItem.Presentation.StyleName = m_ControlLayout.StyleName

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim pageBase As obj_PageBase

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Banner: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Banner: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Banner: sheet '" & m_ControlLayout.LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If m_BannerViewItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Banner: view item is not configured for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If Not m_BannerViewItem.Render(pageBase, m_ControlLayout.RowStart, m_ControlLayout.ColStart, m_ControlLayout.RowEnd, m_ControlLayout.ColEnd, m_ControlName) Then Exit Sub
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "header", "message", "visible"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_BannerViewItem = Nothing
    Set m_ControlLayout = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_ParseBooleanText(ByVal rawText As String) As Boolean
    rawText = VBA.LCase$(VBA.Trim$(rawText))
    private_ParseBooleanText = (rawText = "1" Or rawText = "true" Or rawText = "yes")
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
