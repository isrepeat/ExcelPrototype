Option Explicit

Private m_resetRunning As Boolean
Private m_resetValues As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_UiPageManager.fn_Module_Dispose
    ex_TDB_WordSession.fn_Module_Dispose
    ex_TableSnapshots.fn_Module_Dispose
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_CreatePage(ByVal pageId As String) As obj_IPage
    If VBA.StrComp(VBA.Trim$(pageId), "TravelDocumentsBuilder", VBA.vbTextCompare) <> 0 Then
        VBA.Err.Raise VBA.vbObjectError + 2230, "ex_TravelDocumentsBuilder.fn_CreatePage", _
            "The page is not registered in this workbook: " & pageId
    End If
    Set fn_CreatePage = New obj_TDB_PgMain
End Function

Public Function fn_Initialize() As Boolean
    Dim profileId As String

    ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_STARTED"
    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=ReadProfile"
        Exit Function
    End If
    fn_Initialize = ex_UiPageManager.fn_ShowPage("TravelDocumentsBuilder", profileId)
    ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_COMPLETED | Success=" & VBA.CStr(fn_Initialize)
End Function

Public Function fn_TryGetConfig( _
    ByVal keySuffix As String, _
    ByRef valueText As String _
) As Boolean
    Dim keyName As String

    keyName = "TravelDocumentsBuilder::" & keySuffix
    If Not ex_Core.fn_TryGetWorkbookConfigValue(keyName, valueText) Then
        ex_WindowsUi.fn_ShowMessage "Required configuration key not found: " & keyName, VBA.vbExclamation, "Configuration"
        Exit Function
    End If
    fn_TryGetConfig = True
End Function

Public Function fn_Config(ByVal keySuffix As String) As String
    If Not fn_TryGetConfig(keySuffix, fn_Config) Then
        VBA.Err.Raise 5, , "Required configuration key not found: TravelDocumentsBuilder::" & keySuffix
    End If
End Function

Public Function fn_LogValue(ByVal value As Variant) As String
    Dim text As String

    If VBA.IsObject(value) Then
        text = "<object>"
    ElseIf VBA.IsError(value) Then
        text = "<error>"
    ElseIf VBA.IsNull(value) Then
        text = "<null>"
    ElseIf VBA.IsEmpty(value) Then
        text = "<empty>"
    ElseIf VBA.VarType(value) = VBA.vbDate Then
        text = VBA.Format$(value, "yyyy-mm-dd hh:nn:ss")
    Else
        text = VBA.CStr(value)
    End If
    text = VBA.Replace$(text, "\", "\\")
    text = VBA.Replace$(text, VBA.vbCr, "\r")
    text = VBA.Replace$(text, VBA.vbLf, "\n")
    text = VBA.Replace$(text, VBA.vbTab, "\t")
    text = VBA.Replace$(text, "|", "\|")
    fn_LogValue = text
End Function

Public Function fn_TryGetDocumentPrefix( _
    ByVal documentType As String, _
    ByRef prefix As String _
) As Boolean
    Dim modeId As Variant
    Dim caption As String
    Dim modeIds As String
    Dim diagnostic As String

    prefix = VBA.vbNullString
    If Not fn_TryGetConfig("documentTypes.ids", modeIds) Then Exit Function
    For Each modeId In VBA.Split(modeIds, ",")
        If Not fn_TryGetConfig("text." & VBA.Trim$(VBA.CStr(modeId)) & ".DocumentFormTitle", caption) Then Exit Function
        If VBA.StrComp(VBA.Trim$(documentType), caption, VBA.vbTextCompare) = 0 Then
            prefix = VBA.Trim$(VBA.CStr(modeId))
            fn_TryGetDocumentPrefix = True
            Exit Function
        End If
    Next modeId
    If Not fn_TryGetConfig("text.UnknownDocumentType", diagnostic) Then Exit Function
    ex_WindowsUi.fn_ShowMessage VBA.Replace$(diagnostic, "{value}", documentType), VBA.vbExclamation, "Configuration"
End Function
' --------------------------------------
Public Function fn_IsMedicalCompany(ByVal vacationKind As String) As Boolean
    fn_IsMedicalCompany = VBA.StrComp(vacationKind, fn_Config("generation.kind.annualMedical.input"), VBA.vbTextCompare) = 0 Or _
        VBA.StrComp(vacationKind, fn_Config("generation.kind.treatmentMedical.input"), VBA.vbTextCompare) = 0
End Function

Public Function fn_DocumentId(ByVal eventKind As String) As String
    Dim kind As Variant

    If eventKind = fn_Config("generation.businessTrip.kind") Then
        fn_DocumentId = "BusinessTripCertificate"
        Exit Function
    End If
    For Each kind In VBA.Split(fn_Config("generation.kind.ids"), ",")
        If VBA.StrComp(eventKind, fn_Config("generation.kind." & kind & ".input"), VBA.vbTextCompare) = 0 Then
            fn_DocumentId = "VacationTicket"
            Exit Function
        End If
    Next kind
    VBA.Err.Raise 5, , VBA.Replace$(fn_Config("text.UnknownDocumentType"), "{value}", eventKind)
End Function

Public Function fn_ParseDuration( _
    ByVal documentId As String, _
    ByVal value As String _
) As Long
    Dim phrase As String
    Dim maximum As String
    Dim diagnostic As String

    value = ex_TableQuery.fn_NormalizeText(value)
    maximum = fn_Config("generation.limits.maxDays")
    If documentId = "BusinessTripCertificate" Then
        phrase = fn_Config("generation.businessTrip.indefiniteInput")
        If VBA.Len(VBA.Trim$(phrase)) = 0 Then
            VBA.Err.Raise 5, , "Empty configuration key: TravelDocumentsBuilder::generation.businessTrip.indefiniteInput"
        End If
        If VBA.StrComp(value, phrase, VBA.vbTextCompare) = 0 Then
            Exit Function
        End If
        diagnostic = VBA.Replace$(VBA.Replace$(fn_Config("generation.businessTrip.durationError"), _
            "{phrase}", phrase), "{max}", maximum)
    Else
        diagnostic = VBA.Replace$(fn_Config("generation.message.Invalid"), "{value}", _
            fn_Config("text.VacationTicket.VacationDays"))
    End If
    If VBA.Len(value) = 0 Or value Like "*[!0-9]*" Then
        VBA.Err.Raise 5, , diagnostic
    End If
    If VBA.Len(value) > 10 Then
        VBA.Err.Raise 5, , diagnostic
    End If
    If VBA.CDbl(value) <= 0 Or VBA.CDbl(value) > VBA.CDbl(maximum) Then
        VBA.Err.Raise 5, , diagnostic
    End If
    fn_ParseDuration = VBA.CLng(value)
End Function

Public Function fn_IsTreatment(ByVal vacationKind As String) As Boolean
    fn_IsTreatment = VBA.StrComp(vacationKind, fn_Config("generation.kind.treatment.input"), VBA.vbTextCompare) = 0 Or _
        VBA.StrComp(vacationKind, fn_Config("generation.kind.treatmentMedical.input"), VBA.vbTextCompare) = 0
End Function

Public Function fn_IsResetting() As Boolean
    fn_IsResetting = m_resetRunning
End Function

Public Function fn_ResetForm(ByVal restoreValues As Object) As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousEvents As Boolean
    Dim startedAt As Double
    Dim diagnostic As String
    Dim snapshotActive As Boolean

    If m_resetRunning Then Exit Function
    m_resetRunning = True
    startedAt = VBA.Timer
    previousScreenUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ex_Core.fn_BeginConfigReadSnapshot
    snapshotActive = True
    Set m_resetValues = restoreValues
    ex_Core.fn_Diagnostic_WriteLog "TDB_RESET_STARTED"
    fn_ResetForm = fn_Initialize()
Cleanup:
    On Error Resume Next
    Set m_resetValues = Nothing
    If snapshotActive Then ex_Core.fn_EndConfigReadSnapshot
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    m_resetRunning = False
    ex_Core.fn_Diagnostic_WritePerf "TDB.Reset.Total | Success=" & VBA.CStr(fn_ResetForm), startedAt
    On Error GoTo 0
    Exit Function
Failed:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "TDB_RESET_FAILED | " & diagnostic
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
    Resume Cleanup
End Function

Public Function fn_ApplyResetValues(ByVal bindingContext As obj_UiBindingContext) As Boolean
    Dim field As Variant
    Dim mapping As Variant
    Dim pair As Variant
    Dim handlers As Object
    Dim command As Object
    Dim value As Variant
    Dim isObject As Boolean
    Dim startedAt As Double

    fn_ApplyResetValues = True
    If m_resetValues Is Nothing Then Exit Function
    startedAt = VBA.Timer
    On Error GoTo Failed
    Set handlers = VBA.CreateObject("Scripting.Dictionary")
    For Each mapping In VBA.Split(fn_Config("forms.reset.restoreHandlers"), "|")
        pair = VBA.Split(VBA.CStr(mapping), "=")
        If UBound(pair) <> 1 Then VBA.Err.Raise 5, , "Invalid forms.reset.restoreHandlers mapping."
        handlers.Add pair(0), pair(1)
    Next mapping
    ' Restore selectors before the first render, using the normal selection commands.
    If Not bindingContext.SetValue("Form", "ResetInProgress", True) Then VBA.Err.Raise 5, , fn_Config("text.ResetFailed")
    For Each field In VBA.Split(fn_Config("forms.reset.deepPreserve"), "|")
        ex_Core.fn_Diagnostic_WriteLog "TDB_RESET_RESTORE_STARTED | Field=" & VBA.CStr(field)
        If Not m_resetValues.Exists(VBA.CStr(field)) Or Not handlers.Exists(VBA.CStr(field)) Then
            VBA.Err.Raise 5, , "Missing reset value or handler: " & VBA.CStr(field)
        End If
        If Not bindingContext.SetValue("Form", VBA.CStr(field), m_resetValues(VBA.CStr(field))) Then
            VBA.Err.Raise 5, , fn_Config("text.ResetFailed")
        End If
        If VBA.Len(VBA.CStr(handlers(VBA.CStr(field)))) > 0 Then
            If Not bindingContext.TryGetValue("Commands", VBA.CStr(handlers(VBA.CStr(field))), value, command, isObject) Then
                VBA.Err.Raise 5, , fn_Config("text.ResetFailed")
            End If
            If Not isObject Or command Is Nothing Then VBA.Err.Raise 5, , fn_Config("text.ResetFailed")
            If Not command.Execute() Then VBA.Err.Raise 5, , fn_Config("text.ResetFailed")
        End If
        ex_Core.fn_Diagnostic_WriteLog "TDB_RESET_RESTORE_COMPLETED | Field=" & VBA.CStr(field)
    Next field
    If Not bindingContext.SetValue("Form", "ResetInProgress", False) Then VBA.Err.Raise 5, , fn_Config("text.ResetFailed")
    ex_Core.fn_Diagnostic_WritePerf "TDB.Reset.RestoreSelectors", startedAt
    Exit Function
Failed:
    VBA.Err.Raise VBA.Err.Number, "TDB.Reset.RestoreSelectors", VBA.Err.Description
End Function

Public Sub fn_FocusBasisCell(ByVal target As Range)
    Dim context As obj_UiRenderContext
    Dim value As Variant
    Dim command As Object
    Dim isObject As Boolean

    If target Is Nothing Then Exit Sub
    If Not ex_UiRuntime.fn_TryGetContext(target.Worksheet, context) Then Exit Sub
    If Not context.BindingContext.TryGetValue("Commands", "FocusBasisCell", value, command, isObject) Then Exit Sub
    If Not isObject Or command Is Nothing Then Exit Sub
    command.ExecuteWithPayload target
End Sub

' } // namespace API
' --------------------------------------