Attribute VB_Name = "ex_PEB_PersonNameTests"
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
Public Function fn_Run(ByRef diagnostic As String) As Boolean
    Dim values As Object

    On Error GoTo EH
    If ex_PEB_Export.fn_CleanPersonName("SURNAME First Patronymic (3561012017)") <> "SURNAME First Patronymic" Then
        Err.Raise 5, , "A trailing IPN must be removed from the name."
    End If
    If ex_PEB_Export.fn_CleanPersonName("SURNAME (note) First ( 3561012017 ) ") <> "SURNAME (note) First" Then
        Err.Raise 5, , "Non-IPN parentheses must be preserved."
    End If
    If ex_PEB_Export.fn_CleanPersonName("SURNAME First (123)") <> "SURNAME First (123)" Then
        Err.Raise 5, , "A shorter numeric note must be preserved."
    End If
    If ex_PEB_Export.fn_CleanPersonName("SURNAME First (12345678901)") <> "SURNAME First (12345678901)" Then
        Err.Raise 5, , "A longer numeric note must be preserved."
    End If
    If ex_PEB_Export.fn_CleanPersonName("SURNAME First") <> "SURNAME First" Then
        Err.Raise 5, , "An unchanged name must stay unchanged."
    End If
    Set values = CreateObject("Scripting.Dictionary")
    values("PersonName") = "SURNAME First (3561012017)"
    values("ReportName") = "REPORTER First (1234567890)"
    values("PersonId") = "3561012017"
    ex_PEB_Export.fn_NormalizeNames values
    ex_PEB_Export.fn_NormalizeNames values
    If values("PersonName") <> "SURNAME First" Or values("ReportName") <> "REPORTER First" Or values("PersonId") <> "3561012017" Then
        Err.Raise 5, , "Name normalization must be idempotent and preserve the separate tax ID."
    End If
    fn_Run = True
    Exit Function
EH:
    diagnostic = Err.Description
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------