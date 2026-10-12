Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_UiPageManager.fn_Module_Dispose
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
' } // namespace API
' --------------------------------------