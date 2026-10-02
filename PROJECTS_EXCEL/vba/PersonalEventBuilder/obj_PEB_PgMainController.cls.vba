VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_tableList As obj_UiRawTableList

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
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        ex_WindowsUi.fn_ShowMessage "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set m_tableList = New obj_UiRawTableList
    If Not m_tableList.Initialize() Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "Tables", m_tableList) Then Exit Function
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_tableList Is Nothing Then m_tableList.Dispose
    Set m_tableList = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function GenerateTablesCommandHandler() As Boolean
    Dim tableIndex As Long
    Dim rowIndex As Long
    Dim values As Variant
    Dim headers As Variant
    Dim rawTable As obj_UiRawTable

    If m_tableList Is Nothing Then Exit Function
    m_tableList.Clear
    headers = VBA.Array("Candidate", "Category", "Status")
    For tableIndex = 1 To 10
        ReDim values(1 To 3, 1 To 3)
        For rowIndex = 1 To 3
            values(rowIndex, 1) = "Candidate " & VBA.CStr(tableIndex) & "." & VBA.CStr(rowIndex)
            values(rowIndex, 2) = "Group " & VBA.CStr(tableIndex)
            values(rowIndex, 3) = "Ready"
        Next rowIndex
        Set rawTable = New obj_UiRawTable
        If Not rawTable.Initialize(values, headers, "Table " & VBA.CStr(tableIndex)) Then Exit Function
        If Not m_tableList.Add(rawTable) Then Exit Function
    Next tableIndex
    GenerateTablesCommandHandler = m_pageBase.UpdatePage()
End Function

Public Function ResetCommandHandler() As Boolean
    Dim bindingContext As obj_UiBindingContext

    If m_pageBase Is Nothing Or m_tableList Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If Not bindingContext.SetValue("Form", "EventName", VBA.vbNullString) Then Exit Function
    If Not bindingContext.SetValue("Form", "Category", "Meeting") Then Exit Function
    If Not bindingContext.SetValue("Form", "Notes", VBA.vbNullString) Then Exit Function
    m_tableList.Clear
    ResetCommandHandler = m_pageBase.UpdatePage()
End Function

Public Function UpdatePageCommandHandler() As Boolean
    If m_pageBase Is Nothing Then Exit Function
    UpdatePageCommandHandler = m_pageBase.UpdatePage()
End Function

Public Function FormChangedCommandHandler() As Boolean
    Dim eventName As Variant
    Dim category As Variant
    Dim notes As Variant

    If m_pageBase Is Nothing Then Exit Function
    If Not private_TryReadFormValue("EventName", eventName) Then Exit Function
    If Not private_TryReadFormValue("Category", category) Then Exit Function
    If Not private_TryReadFormValue("Notes", notes) Then Exit Function
    ex_Core.fn_Diagnostic_WriteLog "FORM_CHANGED | EventName=" & VBA.CStr(eventName) & _
        " | Category=" & VBA.CStr(category) & " | Notes=" & VBA.CStr(notes)
    FormChangedCommandHandler = True
End Function

Public Function SubmitFormCommandHandler() As Boolean
    Dim eventName As Variant
    Dim category As Variant
    Dim notes As Variant

    Dim context As obj_UiRenderContext
    Dim errors As Collection

    Set errors = New Collection
    If Not ex_UiRuntime.fn_TryGetContext(ThisWorkbook.Worksheets("MainPage"), context) Then Exit Function
    If Not context.ValidateForm("EventDraftForm", errors) Then
        ex_WindowsUi.fn_ShowMessage VBA.CStr(errors(1)), VBA.vbExclamation, "Form validation"
        Exit Function
    End If
    If Not Me.FormChangedCommandHandler() Then Exit Function
    If Not private_TryReadFormValue("EventName", eventName) Then Exit Function
    If Not private_TryReadFormValue("Category", category) Then Exit Function
    If Not private_TryReadFormValue("Notes", notes) Then Exit Function
    If ex_WindowsUi.fn_ShowInformation( _
            "Event: " & VBA.CStr(eventName) & VBA.vbCrLf & _
            "Category: " & VBA.CStr(category) & VBA.vbCrLf & _
            "Notes: " & VBA.CStr(notes), "Form data") = 0 Then Exit Function
    SubmitFormCommandHandler = True
End Function

' //
' // Private
' //
Private Function private_TryReadFormValue( _
    ByVal fieldName As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim outObject As Object
    Dim isObject As Boolean

    If m_pageBase Is Nothing Then Exit Function
    private_TryReadFormValue = m_pageBase.BindingContext.TryGetValue( _
        "Form", fieldName, outValue, outObject, isObject)
End Function