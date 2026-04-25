Attribute VB_Name = "ex_ControlPartsRuntime"
Option Explicit

Private g_ControlParts As Collection

' //
' // API
' //
Public Sub m_ResetControlParts()
    Set g_ControlParts = Nothing
End Sub


Public Function m_RegisterControlPart( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal partName As String, _
    ByVal partRange As Range _
) As Boolean
    Dim entry As Object

    If ws Is Nothing Then
        VBA.MsgBox "PrototypeNew: worksheet is not specified for control part registration.", VBA.vbExclamation
        Exit Function
    End If
    If partRange Is Nothing Then
        VBA.MsgBox "PrototypeNew: range is not specified for control part registration.", VBA.vbExclamation
        Exit Function
    End If

    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    partName = VBA.LCase$(VBA.Trim$(partName))

    If VBA.Len(controlType) = 0 Then
        VBA.MsgBox "PrototypeNew: control part registration requires non-empty control type.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(partName) = 0 Then
        VBA.MsgBox "PrototypeNew: control part registration requires non-empty part name.", VBA.vbExclamation
        Exit Function
    End If

    private_EnsureControlPartsStorage

    Set entry = CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("SheetName") = VBA.LCase$(ws.Name)
    entry("ControlType") = controlType
    entry("ControlName") = controlName
    entry("PartName") = partName
    Set entry("Range") = partRange

    g_ControlParts.Add entry
    m_RegisterControlPart = True
End Function


Public Function m_TryResolveControlPartScope( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal partName As String, _
    ByRef outScope As Range, _
    ByRef outColumnScope As Range _
) As Boolean
    Dim entry As Object
    Dim partRange As Range
    Dim wsKey As String

    If ws Is Nothing Then
        VBA.MsgBox "PrototypeNew: worksheet is not specified for control part selector.", VBA.vbExclamation
        Exit Function
    End If

    wsKey = VBA.LCase$(ws.Name)
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    partName = VBA.LCase$(VBA.Trim$(partName))

    If VBA.Len(controlType) = 0 Then
        VBA.MsgBox "PrototypeNew: control part selector requires non-empty type.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(partName) = 0 Then
        VBA.MsgBox "PrototypeNew: control part selector requires non-empty part.", VBA.vbExclamation
        Exit Function
    End If

    If g_ControlParts Is Nothing Then
        m_TryResolveControlPartScope = True
        Exit Function
    End If

    For Each entry In g_ControlParts
        If VBA.LCase$(VBA.CStr(entry("SheetName"))) <> wsKey Then GoTo ContinueEntry
        If VBA.LCase$(VBA.CStr(entry("ControlType"))) <> controlType Then GoTo ContinueEntry
        If VBA.Len(controlName) > 0 Then
            If VBA.LCase$(VBA.CStr(entry("ControlName"))) <> controlName Then GoTo ContinueEntry
        End If
        If VBA.LCase$(VBA.CStr(entry("PartName"))) <> partName Then GoTo ContinueEntry

        Set partRange = Nothing
        On Error Resume Next
        Set partRange = entry("Range")
        On Error GoTo 0
        If partRange Is Nothing Then GoTo ContinueEntry

        If outScope Is Nothing Then
            Set outScope = partRange
        Else
            Set outScope = Application.Union(outScope, partRange)
        End If

ContinueEntry:
    Next entry

    If Not outScope Is Nothing Then
        Set outColumnScope = outScope.EntireColumn
    End If

    m_TryResolveControlPartScope = True
End Function

' //
' // Internal
' //

Private Sub private_EnsureControlPartsStorage()
    If Not g_ControlParts Is Nothing Then Exit Sub
    Set g_ControlParts = New Collection
End Sub
