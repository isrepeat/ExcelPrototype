Attribute VB_Name = "rt_PEB_WordExportRuntime"
Option Explicit

' Word.Application is deliberately cached at module scope. Exporter class
' instances are short-lived, while this runtime survives between exports.
Private g_WordApp As Object
Private g_OwnsWordApp As Boolean

Public Function fn_GetOrCreateWordApp(ByRef outWordApp As Object) As Boolean
    Set outWordApp = Nothing

    If private_IsWordAppAlive(g_WordApp) Then
        Set outWordApp = g_WordApp
        fn_GetOrCreateWordApp = True
        Exit Function
    End If

    Set g_WordApp = Nothing
    g_OwnsWordApp = False
    On Error Resume Next
    Set g_WordApp = VBA.CreateObject("Word.Application")
    On Error GoTo 0
    g_OwnsWordApp = Not g_WordApp Is Nothing
    If g_WordApp Is Nothing Then
        VBA.MsgBox "PrototypeNew: Microsoft Word could not be started.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    If g_OwnsWordApp Then
        g_WordApp.Visible = False
        g_WordApp.DisplayAlerts = 0
    End If
    Set outWordApp = g_WordApp
    fn_GetOrCreateWordApp = True
End Function

' Возвращает уже открытый документ по точному FullName либо открывает его
' через runtime Word.Application. outOpenedHere определяет lifecycle:
' пользовательский документ после операции сохраняется, но не закрывается.
Public Function fn_TryAcquireWordDocument( _
    ByVal documentPath As String, _
    ByRef outWordApp As Object, _
    ByRef outWordDoc As Object, _
    ByRef outOpenedHere As Boolean _
) As Boolean
    Dim wordDoc As Object
    Dim runningWordApp As Object
    Dim errorDescription As String

    Set outWordApp = Nothing
    Set outWordDoc = Nothing
    outOpenedHere = False
    documentPath = VBA.Trim$(documentPath)
    If VBA.Len(documentPath) = 0 Then Exit Function
    On Error GoTo EH
    If private_IsWordAppAlive(g_WordApp) Then
        For Each wordDoc In g_WordApp.Documents
            If VBA.StrComp( _
                VBA.Trim$(VBA.CStr(wordDoc.FullName)), _
                documentPath, _
                VBA.vbTextCompare) = 0 Then
                If wordDoc.ReadOnly Then
                    VBA.MsgBox _
                        "WORD-документ открыт только для чтения:" & _
                        VBA.vbCrLf & documentPath, _
                        VBA.vbExclamation, _
                        "PrsnlEventBuilder / WORD export"
                    Exit Function
                End If
                Set outWordApp = g_WordApp
                Set outWordDoc = wordDoc
                fn_TryAcquireWordDocument = True
                Exit Function
            End If
        Next wordDoc
    End If

    ' К пользовательскому Word подключаемся только ради уже открытого
    ' целевого документа. Закрытые файлы по-прежнему обрабатываются через
    ' собственный скрытый runtime, чтобы не переключать окна пользователя.
    On Error Resume Next
    Set runningWordApp = VBA.GetObject(, "Word.Application")
    On Error GoTo EH
    If Not runningWordApp Is Nothing Then
        For Each wordDoc In runningWordApp.Documents
            If VBA.StrComp( _
                VBA.Trim$(VBA.CStr(wordDoc.FullName)), _
                documentPath, _
                VBA.vbTextCompare) = 0 Then
                If wordDoc.ReadOnly Then
                    VBA.MsgBox _
                        "WORD-документ открыт только для чтения:" & _
                        VBA.vbCrLf & documentPath, _
                        VBA.vbExclamation, _
                        "PrsnlEventBuilder / WORD export"
                    Exit Function
                End If
                Set outWordApp = runningWordApp
                Set outWordDoc = wordDoc
                fn_TryAcquireWordDocument = True
                Exit Function
            End If
        Next wordDoc
    End If

    If Not fn_GetOrCreateWordApp(outWordApp) Then Exit Function
    Set outWordDoc = outWordApp.Documents.Open( _
        documentPath, False, False, False)
    outOpenedHere = Not outWordDoc Is Nothing
    fn_TryAcquireWordDocument = Not outWordDoc Is Nothing
    Exit Function

EH:
    errorDescription = Err.Description
    Set outWordDoc = Nothing
    outOpenedHere = False
    VBA.MsgBox _
        "Не вдалося підключитися до WORD-документа:" & _
        VBA.vbCrLf & documentPath & VBA.vbCrLf & _
        "Помилка: " & errorDescription, _
        VBA.vbExclamation, _
        "PrsnlEventBuilder / WORD export"
End Function

Public Sub fn_Dispose(Optional ByVal quitWord As Boolean = True)
    On Error Resume Next
    If quitWord And g_OwnsWordApp And private_IsWordAppAlive(g_WordApp) Then
        g_WordApp.Quit
    End If
    Set g_WordApp = Nothing
    g_OwnsWordApp = False
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
