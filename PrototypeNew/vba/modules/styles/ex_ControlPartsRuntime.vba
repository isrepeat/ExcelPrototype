Attribute VB_Name = "ex_ControlPartsRuntime"
Option Explicit

Private g_ControlParts As Collection

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
        MsgBox "PrototypeNew: worksheet is not specified for control part registration.", vbExclamation
        Exit Function
    End If
    If partRange Is Nothing Then
        MsgBox "PrototypeNew: range is not specified for control part registration.", vbExclamation
        Exit Function
    End If

    controlType = LCase$(Trim$(controlType))
    controlName = LCase$(Trim$(controlName))
    partName = LCase$(Trim$(partName))

    If Len(controlType) = 0 Then
        MsgBox "PrototypeNew: control part registration requires non-empty control type.", vbExclamation
        Exit Function
    End If
    If Len(partName) = 0 Then
        MsgBox "PrototypeNew: control part registration requires non-empty part name.", vbExclamation
        Exit Function
    End If

    mp_EnsureControlPartsStorage

    Set entry = CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("SheetName") = LCase$(ws.Name)
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
        MsgBox "PrototypeNew: worksheet is not specified for control part selector.", vbExclamation
        Exit Function
    End If

    wsKey = LCase$(ws.Name)
    controlType = LCase$(Trim$(controlType))
    controlName = LCase$(Trim$(controlName))
    partName = LCase$(Trim$(partName))

    If Len(controlType) = 0 Then
        MsgBox "PrototypeNew: control part selector requires non-empty type.", vbExclamation
        Exit Function
    End If
    If Len(partName) = 0 Then
        MsgBox "PrototypeNew: control part selector requires non-empty part.", vbExclamation
        Exit Function
    End If

    If g_ControlParts Is Nothing Then
        m_TryResolveControlPartScope = True
        Exit Function
    End If

    For Each entry In g_ControlParts
        If LCase$(CStr(entry("SheetName"))) <> wsKey Then GoTo ContinueEntry
        If LCase$(CStr(entry("ControlType"))) <> controlType Then GoTo ContinueEntry
        If Len(controlName) > 0 Then
            If LCase$(CStr(entry("ControlName"))) <> controlName Then GoTo ContinueEntry
        End If
        If LCase$(CStr(entry("PartName"))) <> partName Then GoTo ContinueEntry

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

Private Sub mp_EnsureControlPartsStorage()
    If Not g_ControlParts Is Nothing Then Exit Sub
    Set g_ControlParts = New Collection
End Sub

