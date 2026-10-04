Attribute VB_Name = "ex_UiPageFactory"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Create(ByVal pageId As String) As obj_IPage
    pageId = VBA.LCase$(VBA.Trim$(pageId))
    Select Case pageId
        Case "personaleventbuilder"
            Set fn_Create = New obj_PEB_PgMain
        Case "personnelatdisposaldayscalculation"
            Set fn_Create = New obj_PADC_PgMain
        Case Else
            ex_WindowsUi.fn_ShowMessage "The page is not registered: " & pageId, _
                VBA.vbExclamation, "PersonalEventBuilder"
    End Select
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------