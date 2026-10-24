Attribute VB_Name = "ex_PEB_WordCleanupTests"
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
    Dim wordApp As Object
    Dim document As Object
    Dim anchors As obj_PEB_WordAnchors
    Dim cases As New Collection
    Dim scenario As Variant
    Dim body As String
    Dim expected As String
    Dim target As Object
    Dim field As Object
    Dim count As Long

    On Error GoTo EH
    body = "date" & VBA.vbCr & VBA.vbCr & "person"
    cases.Add VBA.Array(private_Block("Empty", VBA.vbCr), VBA.vbCr, False)
    cases.Add VBA.Array("Lead" & VBA.vbCr & private_Block("Filled", body) & "Tail" & VBA.vbCr, _
        "Lead" & VBA.vbCr & private_Visible("Filled", body) & "Tail" & VBA.vbCr, False)
    cases.Add VBA.Array(private_Block("A", body) & private_Block("B", VBA.vbCr) & private_Block("C", body), _
        private_Visible("A", body) & private_Visible("C", body), False)
    cases.Add VBA.Array(private_Block("A", VBA.vbCr) & private_Block("B", VBA.vbCr), VBA.vbCr, False)
    cases.Add VBA.Array("Lead" & VBA.vbCr & "{\export:Legacy_Begin}" & VBA.vbCr & _
        "{\export:Legacy_End}" & VBA.vbCr & "Tail" & VBA.vbCr, _
        "Lead" & VBA.vbCr & VBA.vbCr & VBA.vbCr & "Tail" & VBA.vbCr, False)
    cases.Add VBA.Array(private_Block("A", "FieldPlaceholder") & private_Block("B", body), _
        private_Visible("A", "FieldPlaceholder") & private_Visible("B", body), True)
    cases.Add VBA.Array("{\block:Standalone_Begin}" & VBA.vbCr & VBA.vbCr & "Heading" & VBA.vbCr & _
        VBA.vbCr & "{\export:Standalone_Begin}" & VBA.vbCr & "date" & VBA.vbCr & _
        "{\export:Standalone_End}" & VBA.vbCr & "{\block:Standalone_End}" & VBA.vbCr, _
        VBA.vbCr & "Heading" & VBA.vbCr & VBA.vbCr & "date" & VBA.vbCr, False)
    Set wordApp = VBA.CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    For Each scenario In cases
        count = count + 1
        Set document = wordApp.Documents.Add
        document.Content.Text = scenario(0)
        expected = scenario(1)
        If scenario(2) Then
            Set target = document.Content.Duplicate
            If Not target.Find.Execute("FieldPlaceholder") Then
                Err.Raise 5, , "Cleanup field fixture was not found."
            End If
            Set field = document.Fields.Add(target, -1, "QUOTE ""FieldValue""", False)
            field.Update
            expected = VBA.Replace$(expected, "FieldPlaceholder", field.Result.Text)
        End If
        Set anchors = New obj_PEB_WordAnchors
        If Not anchors.Initialize(document, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If Not anchors.TryRemoveMarkers(document, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If VBA.StrComp(document.Content.Text, expected, VBA.vbBinaryCompare) <> 0 Then
            Err.Raise 5, , "Unexpected cleanup result in case " & VBA.CStr(count)
        End If
        anchors.Dispose
        Set anchors = Nothing
        document.Close 0
        Set document = Nothing
    Next scenario
    fn_Run = True
Cleanup:
    On Error Resume Next
    If Not anchors Is Nothing Then
        anchors.Dispose
    End If
    If Not document Is Nothing Then
        document.Close 0
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit 0
    End If
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = "Word cleanup case " & VBA.CStr(count) & ": " & VBA.Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Block( _
    ByVal identifier As String, _
    ByVal body As String _
) As String
    private_Block = "{\block:" & identifier & "_Begin}" & VBA.vbCr & VBA.vbCr & _
        identifier & " heading" & VBA.vbCr & VBA.vbCr & "{\export:" & identifier & "_Begin}" & _
        body & "{\export:" & identifier & "_End}" & VBA.vbCr & "{\block:" & identifier & "_End}" & VBA.vbCr
End Function

Private Function private_Visible( _
    ByVal identifier As String, _
    ByVal body As String _
) As String
    private_Visible = VBA.vbCr & identifier & " heading" & VBA.vbCr & VBA.vbCr & body & VBA.vbCr
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------