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

Private m_pageBase As obj_PageBase
Private m_tableList As obj_UiRawTableList
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    m_isDisposed = False
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        VBA.MsgBox "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set m_tableList = New obj_UiRawTableList
    If Not m_tableList.Initialize() Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "Tables", m_tableList) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
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
    m_tableList.Dispose
    If Not m_tableList.Initialize() Then Exit Function
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

Public Function HelloWorldCommandHandler() As Boolean
    ex_Core.fn_Diagnostic_WriteLog "HELLO_WORLD_CLICKED"
    VBA.MsgBox "Hello World from PersonalEventBuilder.", VBA.vbInformation, "PersonalEventBuilder"
    HelloWorldCommandHandler = True
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

    If Not FormChangedCommandHandler() Then Exit Function
    If Not private_TryReadFormValue("EventName", eventName) Then Exit Function
    If Not private_TryReadFormValue("Category", category) Then Exit Function
    If Not private_TryReadFormValue("Notes", notes) Then Exit Function
    VBA.MsgBox "Event: " & VBA.CStr(eventName) & VBA.vbCrLf & _
        "Category: " & VBA.CStr(category) & VBA.vbCrLf & _
        "Notes: " & VBA.CStr(notes), VBA.vbInformation, "Form data"
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