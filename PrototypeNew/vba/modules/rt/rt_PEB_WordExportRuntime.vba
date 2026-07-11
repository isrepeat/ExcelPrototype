Attribute VB_Name = "rt_PEB_WordExportRuntime"
Option Explicit

' Word.Application is deliberately cached at module scope. Exporter class
' instances are short-lived, while this runtime survives between exports.
Private g_WordApp As Object

Public Function fn_GetOrCreateWordApp(ByRef outWordApp As Object) As Boolean
    Set outWordApp = Nothing

    If private_IsWordAppAlive(g_WordApp) Then
        Set outWordApp = g_WordApp
        fn_GetOrCreateWordApp = True
        Exit Function
    End If

    Set g_WordApp = Nothing
    On Error Resume Next
    Set g_WordApp = VBA.CreateObject("Word.Application")
    On Error GoTo 0
    If g_WordApp Is Nothing Then
        VBA.MsgBox "PrototypeNew: Microsoft Word could not be started.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    g_WordApp.Visible = False
    g_WordApp.DisplayAlerts = 0
    Set outWordApp = g_WordApp
    fn_GetOrCreateWordApp = True
End Function

Public Sub fn_Dispose(Optional ByVal quitWord As Boolean = True)
    On Error Resume Next
    If quitWord And private_IsWordAppAlive(g_WordApp) Then g_WordApp.Quit
    Set g_WordApp = Nothing
    On Error GoTo 0
End Sub

Private Function private_IsWordAppAlive(ByVal wordApp As Object) As Boolean
    Dim documentCount As Long
    If wordApp Is Nothing Then Exit Function
    On Error GoTo NotAlive
    documentCount = wordApp.Documents.Count
    private_IsWordAppAlive = True
    Exit Function
NotAlive:
    private_IsWordAppAlive = False
End Function
