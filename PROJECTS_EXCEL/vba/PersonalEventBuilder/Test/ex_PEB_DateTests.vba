Attribute VB_Name = "ex_PEB_DateTests"
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
    Dim referenceDate As Date

    On Error GoTo EH
    referenceDate = DateSerial(2026, 10, 31)
    If ex_PEB_Export.fn_ParseDate("30", "DocumentDate", referenceDate) <> DateSerial(2026, 10, 30) Then
        Err.Raise 5, , "Day shorthand must use the order month and year."
    End If
    If ex_PEB_Export.fn_ParseDate("30.9", "DocumentDate", referenceDate) <> DateSerial(2026, 9, 30) Then
        Err.Raise 5, , "Day/month shorthand must use the order year."
    End If
    If ex_PEB_Export.fn_ParseDate("30.9.2024", "DocumentDate", referenceDate) <> DateSerial(2024, 9, 30) Then
        Err.Raise 5, , "A complete date must retain its explicit year."
    End If
    referenceDate = DateSerial(2025, 12, 31)
    If ex_PEB_Export.fn_ParseDate(1, "DocumentDate", referenceDate) <> DateSerial(2025, 12, 1) Then
        Err.Raise 5, , "Numeric day shorthand must be independent of the system calendar."
    End If
    If ex_PEB_Export.fn_ExtensionValue("30", "date", "DocumentDate", DateSerial(2026, 10, 31)) <> DateSerial(2026, 10, 30) Then
        Err.Raise 5, , "Child document dates must use the same reference as primary dates."
    End If
    fn_Run = True
    Exit Function
EH:
    diagnostic = Err.Description
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------