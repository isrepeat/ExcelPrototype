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
    Dim key As Variant
    Dim valueText As String
    Dim diagnostic As String

    If m_isDisposed Or m_isInitialized Then Exit Function
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then Exit Function
    Set m_messages = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Array("PersonnelError", "PersonnelMinimum", "PersonnelEmpty", _
            "PersonnelCount", "PersonnelMore", "PersonnelSelected", "ConnectionsReset", "CandidateCountTemplate")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & VBA.CStr(key), valueText) Then Exit Function
        m_messages(VBA.CStr(key)) = valueText
    Next key
    Set m_lookupProfile = New obj_LookupProfile
    If Not m_lookupProfile.Initialize("TravelDocumentsBuilder::lookup.Personnel", diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=LookupProfile | " & diagnostic
        Exit Function
    End If
    Set m_lookupService = New obj_LookupService
    If Not m_lookupService.Initialize(m_lookupProfile) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=LookupService"
        Exit Function
    End If
    If Not m_lookupService.TrySearch(VBA.vbNullString, m_lookupResult, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=InitialSearch | " & diagnostic
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", m_lookupResult) Then Exit Function
    If Not private_SetCandidateCount(m_lookupResult.RowCount) Then Exit Function
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

    If Not private_TryReadFormValue("PersonnelQuery", searchText) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not m_lookupService.TrySearch(VBA.CStr(searchText), result, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_PERSONNEL_SEARCH_FAILED | " & diagnostic
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "Personnel lookup"
        Exit Function
    End If
    Set m_lookupResult = result
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then Exit Function
    If Not private_SetCandidateCount(result.RowCount) Then Exit Function
    statusText = private_BuildPersonnelStatus(VBA.CStr(searchText), result)
    SearchPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", statusText)
End Function

Public Function PersonNameChangedHandler() As Boolean
    Dim personName As Variant
    Dim fieldName As Variant

    If Not private_TryReadFormValue("PersonName", personName) Then Exit Function
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(personName))) > 0 Then
        PersonNameChangedHandler = True
        Exit Function
    End If
    For Each fieldName In VBA.Array("PersonId", "PersonRank", "PersonPosition", "PersonUnit", "PersonPositionCode", "PersonnelQuery")
        If Not m_pageBase.BindingContext.SetValue( _
                "Form", VBA.CStr(fieldName), VBA.vbNullString) Then Exit Function
    Next fieldName
    PersonNameChangedHandler = Me.SearchPersonnelHandler()
End Function

Public Function SelectPersonnelHandler(ByVal selected As Object) As Boolean
    Dim diagnostic As String
    Dim autoFillPersonnel As Variant

    If selected Is Nothing Or Not TypeOf selected Is obj_LookupRecord Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFillPersonnel) Then Exit Function
    If Not VBA.CBool(autoFillPersonnel) Then
        SelectPersonnelHandler = True
        Exit Function
    End If
    If Not m_lookupProfile.TryApply(selected, m_pageBase.BindingContext, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_PERSONNEL_APPLY_FAILED | " & diagnostic
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "Personnel lookup"
        Exit Function
    End If
    SelectPersonnelHandler = m_pageBase.BindingContext.SetValue( _
        "Text", "PersonnelStatus", m_messages("PersonnelSelected"))
End Function

Public Function RefreshPersonnelHandler() As Boolean
    Dim emptyResult As obj_LookupResult
    Dim diagnostic As String

    If m_lookupService Is Nothing Then Exit Function
    ex_TableDataRuntime.fn_ResetCaches
    If Not m_lookupService.TrySearch(VBA.vbNullString, emptyResult, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", emptyResult) Then Exit Function
    If Not m_lookupResult Is Nothing Then m_lookupResult.Dispose
    Set m_lookupResult = emptyResult
    If Not private_SetCandidateCount(emptyResult.RowCount) Then Exit Function
    RefreshPersonnelHandler = m_pageBase.BindingContext.SetValue( _
        "Text", "PersonnelStatus", m_messages("ConnectionsReset"))
End Function

Public Function DocumentTypeChangedHandler() As Boolean
    Dim documentType As Variant
    Dim prefix As String
    Dim key As Variant
    Dim valueText As String

    If Not private_TryReadFormValue("DocumentType", documentType) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetDocumentPrefix(VBA.CStr(documentType), prefix) Then Exit Function
    For Each key In VBA.Array("DocumentFormTitle", "DocumentTargetLabel", _
            "DepartureDateLabel", "ReturnDateLabel", "PurposeLabel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & prefix & "." & VBA.CStr(key), valueText) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    ' Bindings refresh captions; rebuilding the page during a click disposes the active handler.
    DocumentTypeChangedHandler = True
End Function

Public Function CreateDocumentHandler() As Boolean
    CreateDocumentHandler = private_GenerateDocument(False)
End Function

Public Function UpdateDocumentHandler() As Boolean
    UpdateDocumentHandler = private_GenerateDocument(True)
End Function

Public Function VacationDatesChangedHandler() As Boolean
    Dim value As Variant
    Dim parts As Variant
    Dim departure As Date
    Dim totalDays As Long
    Dim key As Variant
    Dim maximum As String
    Dim result As String
    Dim text As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo InvalidInput
    If Not private_TryReadFormValue("DepartureDate", value) Then GoTo InvalidInput
    text = VBA.Trim$(VBA.CStr(value))
    If Not text Like "##.##.####" Then GoTo InvalidInput
    parts = VBA.Split(text, ".")
    departure = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(departure, "dd.mm.yyyy") <> text Then GoTo InvalidInput
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.limits.maxDays", maximum) Then GoTo InvalidInput
    For Each key In VBA.Array("VacationDays", "RoadDays", "DonationDays")
        If Not private_TryReadFormValue(VBA.CStr(key), value) Then GoTo InvalidInput
        text = VBA.Trim$(VBA.CStr(value))
        If VBA.Len(text) = 0 Then
            If VBA.CStr(key) = "VacationDays" Then GoTo InvalidInput
            text = "0"
        End If
        If text Like "*[!0-9]*" Then GoTo InvalidInput
        If VBA.CDbl(text) > VBA.CDbl(maximum) Then GoTo InvalidInput
        totalDays = totalDays + VBA.CLng(text)
    Next key
    result = VBA.Format$(VBA.DateAdd("d", totalDays, departure), "dd.mm.yyyy")
InvalidInput:
    VacationDatesChangedHandler = m_pageBase.BindingContext.SetValue("Form", "ReturnDate", result)
End Function
' //
' // Private
' //
Private Function private_GenerateDocument(ByVal isUpdate As Boolean) As Boolean
    Dim vacationGenerator As obj_TDB_VacationGenerator
    Dim diagnostic As String
    Dim title As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text.GenerationTitle", title) Then Exit Function
    Set vacationGenerator = New obj_TDB_VacationGenerator
    If vacationGenerator.Initialize(m_pageBase.BindingContext, m_lookupService) Then
        private_GenerateDocument = vacationGenerator.TryGenerate(isUpdate, diagnostic)
    End If
    vacationGenerator.Dispose
    If VBA.Len(diagnostic) > 0 Then
        m_pageBase.BindingContext.SetValue "Form", "GenerationStatus", diagnostic
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, title
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    If Not vacationGenerator Is Nothing Then vacationGenerator.Dispose
    ex_Core.fn_Diagnostic_WriteLog "TDB_GENERATION_HANDLER_FAILED | " & diagnostic
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, title
End Function

Private Function private_SetCandidateCount(ByVal count As Long) As Boolean
    private_SetCandidateCount = m_pageBase.BindingContext.SetValue( _
        "Text", "CandidateCount", VBA.Replace$(m_messages("CandidateCountTemplate"), "{count}", VBA.CStr(count)))
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