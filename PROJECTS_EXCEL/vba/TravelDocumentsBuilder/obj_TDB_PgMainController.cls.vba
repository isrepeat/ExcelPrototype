VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TDB_PgMainController"
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

Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    Dim key As Variant
    Dim valueText As String
    Dim diagnostic As String

    If m_isDisposed Or m_isInitialized Then Exit Function
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then Exit Function
    Set m_messages = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Array("PersonnelError", "PersonnelMinimum", "PersonnelEmpty", _
            "PersonnelCount", "PersonnelMore", "PersonnelSelected")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & VBA.CStr(key), valueText) Then Exit Function
        m_messages(VBA.CStr(key)) = valueText
    Next key
    Set m_lookupProfile = New obj_LookupProfile
    If Not m_lookupProfile.Initialize("TravelDocumentsBuilder::lookup.Personnel", diagnostic) Then Exit Function
    Set m_lookupService = New obj_LookupService
    If Not m_lookupService.Initialize(m_lookupProfile) Then Exit Function
    If Not m_lookupService.TrySearch(VBA.vbNullString, m_lookupResult, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", m_lookupResult) Then Exit Function
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub

    m_isDisposed = True
    m_isInitialized = False
    If Not m_lookupService Is Nothing Then m_lookupService.Dispose
    If Not m_lookupResult Is Nothing Then m_lookupResult.Dispose
    If Not m_lookupProfile Is Nothing Then m_lookupProfile.Dispose
    Set m_messages = Nothing
    Set m_lookupService = Nothing
    Set m_lookupResult = Nothing
    Set m_lookupProfile = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function SearchPersonnelHandler() As Boolean
    Dim searchText As Variant
    Dim result As obj_LookupResult
    Dim diagnostic As String
    Dim statusText As String

    If Not private_TryReadFormValue("PersonName", searchText) Then Exit Function
    If Not m_lookupService.TrySearch(VBA.CStr(searchText), result, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "Personnel lookup"
        Exit Function
    End If
    Set m_lookupResult = result
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then Exit Function
    statusText = private_BuildPersonnelStatus(VBA.CStr(searchText), result)
    SearchPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", statusText)
End Function

Public Function SelectPersonnelHandler(ByVal selected As Object) As Boolean
    Dim diagnostic As String

    If selected Is Nothing Or Not TypeOf selected Is obj_LookupRecord Then Exit Function
    If Not m_lookupProfile.TryApply(selected, m_pageBase.BindingContext, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "Personnel lookup"
        Exit Function
    End If
    SelectPersonnelHandler = m_pageBase.BindingContext.SetValue( _
        "Text", "PersonnelStatus", m_messages("PersonnelSelected"))
End Function

Public Function RefreshPersonnelHandler() As Boolean
    If m_lookupService Is Nothing Then Exit Function
    m_lookupService.InvalidateCache
    RefreshPersonnelHandler = Me.SearchPersonnelHandler()
End Function

Public Function DocumentTypeChangedHandler() As Boolean
    Dim documentType As Variant
    Dim prefix As String
    Dim key As Variant
    Dim valueText As String

    If Not private_TryReadFormValue("DocumentType", documentType) Then Exit Function
    Select Case VBA.CStr(documentType)
        Case "Відпускний квиток"
            prefix = "VacationTicket"
        Case "Посвідчення про відрядження"
            prefix = "BusinessTripCertificate"
        Case Else
            Exit Function
    End Select
    For Each key In VBA.Array("DocumentFormTitle", "DocumentTargetLabel", _
            "DepartureDateLabel", "ReturnDateLabel", "PurposeLabel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & prefix & "." & VBA.CStr(key), valueText) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    DocumentTypeChangedHandler = m_pageBase.UpdatePage()
End Function

Private Function private_BuildPersonnelStatus( _
    ByVal searchText As String, _
    ByVal result As obj_LookupResult _
) As String
    If VBA.Len(ex_TableQuery.fn_NormalizeText(searchText)) < m_lookupProfile.MinChars Then
        private_BuildPersonnelStatus = VBA.Replace$( _
            m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(m_lookupProfile.MinChars))
    ElseIf result.RowCount = 0 Then
        private_BuildPersonnelStatus = m_messages("PersonnelEmpty")
    ElseIf result.HasMore Then
        private_BuildPersonnelStatus = VBA.Replace$(m_messages("PersonnelMore"), _
            "{count}", VBA.CStr(result.RowCount))
    Else
        private_BuildPersonnelStatus = VBA.Replace$(m_messages("PersonnelCount"), _
            "{count}", VBA.CStr(result.RowCount))
    End If
End Function

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