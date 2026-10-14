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
Private m_lookupProfile As obj_LookupProfile
Private m_lookupService As obj_LookupService
Private m_lookupResult As obj_LookupResult
Private m_messages As Object

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
    Dim messageKey As Variant
    Dim messageText As String
    Dim diagnostic As String

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        ex_WindowsUi.fn_ShowMessage "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set m_messages = VBA.CreateObject("Scripting.Dictionary")
    For Each messageKey In VBA.Array("PersonnelError", "PersonnelMinimum", "PersonnelEmpty", _
            "PersonnelCount", "PersonnelMore", "PersonnelSelected")
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text." & messageKey, messageText) Then
            ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text." & messageKey, vbExclamation, "Configuration"
            Exit Function
        End If
        m_messages(messageKey) = messageText
    Next messageKey
    Set m_lookupProfile = New obj_LookupProfile
    If Not m_lookupProfile.Initialize("PersonalEventBuilder::lookup.Personnel", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    Set m_lookupService = New obj_LookupService
    If Not m_lookupService.Initialize(m_lookupProfile) Then Exit Function
    If Not m_lookupService.TrySearch(vbNullString, m_lookupResult, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", m_lookupResult) Then Exit Function
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_lookupService Is Nothing Then m_lookupService.Dispose
    If Not m_lookupResult Is Nothing Then m_lookupResult.Dispose
    If Not m_lookupProfile Is Nothing Then m_lookupProfile.Dispose
    Set m_messages = Nothing
    Set m_lookupService = Nothing
    Set m_lookupProfile = Nothing
    Set m_lookupResult = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function SearchPersonnelHandler() As Boolean
    Dim text As Variant
    Dim diagnostic As String
    Dim result As obj_LookupResult
    Dim updates As Object
    Dim name As Variant
    Dim status As String
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("PersonName", text) Then Exit Function
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each name In VBA.Array("PersonId", "PersonRank", "PersonPosition", "PersonUnit")
        updates("Form." & name) = vbNullString
    Next name
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failure
    If Not m_lookupService.TrySearch(VBA.CStr(text), result, diagnostic) Then
        If Not m_lookupService.TrySearch(vbNullString, result, status) Then GoTo Failure
        Set m_lookupResult = result
        m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", result
        m_pageBase.BindingContext.SetValue "Text", "PersonnelStatus", m_messages("PersonnelError")
        GoTo Failure
    End If
    Set m_lookupResult = result
    ex_Core.fn_Diagnostic_WriteLog "LOOKUP_STAGE_STARTED | Name=PublishCandidates"
    startedAt = VBA.Timer
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then Exit Function
    ex_Core.fn_Diagnostic_WritePerf "Lookup.PublishCandidates | Rows=" & VBA.CStr(result.RowCount), startedAt
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(text))) < m_lookupProfile.MinChars Then
        status = VBA.Replace$(m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(m_lookupProfile.MinChars))
    ElseIf result.RowCount = 0 Then
        status = m_messages("PersonnelEmpty")
    Else
        status = VBA.Replace$(m_messages("PersonnelCount"), "{count}", VBA.CStr(result.RowCount))
        If result.HasMore Then status = VBA.Replace$(m_messages("PersonnelMore"), "{count}", VBA.CStr(result.RowCount))
    End If
    SearchPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", status)
    Exit Function
Failure:
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
End Function

Public Function SelectPersonnelHandler(ByVal selected As Object) As Boolean
    Dim record As obj_LookupRecord
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If selected Is Nothing Then Exit Function
    If Not TypeOf selected Is obj_LookupRecord Then Exit Function
    Set record = selected
    If Not m_lookupProfile.TryApply(record, m_pageBase.BindingContext, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    SelectPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", m_messages("PersonnelSelected"))
End Function

Public Function RefreshPersonnelHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ex_TableDataRuntime.fn_ResetCaches
    m_lookupService.InvalidateCache
    RefreshPersonnelHandler = Me.SearchPersonnelHandler()
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