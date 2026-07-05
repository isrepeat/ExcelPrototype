Attribute VB_Name = "ex_ControlPartsRuntime"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private g_ControlParts As Collection
Private g_ControlColumnAliases As Collection

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_ControlPartsRuntime.fn_Module_Dispose"
#End If
    On Error Resume Next
    Set g_ControlParts = Nothing
    Set g_ControlColumnAliases = Nothing
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Sub fn_ResetControlParts()
    Set g_ControlParts = Nothing
    Set g_ControlColumnAliases = Nothing
End Sub

Public Function fn_RemoveControlPartsByWorksheetName(ByVal worksheetName As String) As Boolean
    Dim worksheetKey As String

    worksheetKey = VBA.LCase$(VBA.Trim$(worksheetName))
    If VBA.Len(worksheetKey) = 0 Then
        fn_RemoveControlPartsByWorksheetName = True
        Exit Function
    End If

    Call private_RemoveEntriesByWorksheetKey(g_ControlParts, worksheetKey)
    Call private_RemoveEntriesByWorksheetKey(g_ControlColumnAliases, worksheetKey)

    fn_RemoveControlPartsByWorksheetName = True
End Function


Public Function fn_RegisterControlPart( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal partName As String, _
    ByVal partRange As Range _
) As Boolean
    Dim entry As Object

    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for control part registration."
#End If
        Exit Function
    End If
    If partRange Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: range is not specified for control part registration."
#End If
        Exit Function
    End If

    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    partName = VBA.LCase$(VBA.Trim$(partName))

    If VBA.Len(controlType) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: control part registration requires non-empty control type."
#End If
        Exit Function
    End If
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: control part registration requires non-empty part name."
#End If
        Exit Function
    End If

    private_EnsureControlPartsStorage

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("SheetName") = VBA.LCase$(ws.Name)
    entry("ControlType") = controlType
    entry("ControlName") = controlName
    entry("PartName") = partName
    Set entry("Range") = partRange

    g_ControlParts.Add entry
    fn_RegisterControlPart = True
End Function

Public Function fn_RegisterControlColumnAlias( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal columnAlias As String, _
    ByVal columnRange As Range _
) As Boolean
    Dim entry As Object

    If ws Is Nothing Then Exit Function
    If columnRange Is Nothing Then Exit Function

    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    columnAlias = VBA.LCase$(VBA.Trim$(columnAlias))

    If VBA.Len(controlType) = 0 Then Exit Function
    If VBA.Len(columnAlias) = 0 Then Exit Function

    private_EnsureControlColumnAliasesStorage

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("SheetName") = VBA.LCase$(ws.Name)
    entry("ControlType") = controlType
    entry("ControlName") = controlName
    entry("ColumnAlias") = columnAlias
    Set entry("Range") = columnRange

    g_ControlColumnAliases.Add entry
    fn_RegisterControlColumnAlias = True
End Function

Public Function fn_TryResolveControlColumnAliasScope( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal columnAlias As String, _
    ByRef outColumnScope As Range _
) As Boolean
    Dim entry As Object
    Dim aliasRange As Range
    Dim wsKey As String

    If ws Is Nothing Then Exit Function

    wsKey = VBA.LCase$(ws.Name)
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    columnAlias = VBA.LCase$(VBA.Trim$(columnAlias))

    If VBA.Len(controlType) = 0 Then Exit Function
    If VBA.Len(columnAlias) = 0 Then Exit Function

    If g_ControlColumnAliases Is Nothing Then
        fn_TryResolveControlColumnAliasScope = True
        Exit Function
    End If

    For Each entry In g_ControlColumnAliases
        If VBA.LCase$(VBA.CStr(entry("SheetName"))) <> wsKey Then GoTo ContinueEntry
        If VBA.LCase$(VBA.CStr(entry("ControlType"))) <> controlType Then GoTo ContinueEntry
        If VBA.Len(controlName) > 0 Then
            If VBA.LCase$(VBA.CStr(entry("ControlName"))) <> controlName Then GoTo ContinueEntry
        End If
        If VBA.LCase$(VBA.CStr(entry("ColumnAlias"))) <> columnAlias Then GoTo ContinueEntry

        Set aliasRange = Nothing
        On Error Resume Next
        Set aliasRange = entry("Range")
        On Error GoTo 0
        If aliasRange Is Nothing Then GoTo ContinueEntry

        If outColumnScope Is Nothing Then
            Set outColumnScope = aliasRange
        Else
            Set outColumnScope = Application.Union(outColumnScope, aliasRange)
        End If

ContinueEntry:
    Next entry

    fn_TryResolveControlColumnAliasScope = True
End Function


Public Function fn_TryResolveControlPartScope( _
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
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for control part selector."
#End If
        Exit Function
    End If

    wsKey = VBA.LCase$(ws.Name)
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    partName = VBA.LCase$(VBA.Trim$(partName))

    If VBA.Len(controlType) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: control part selector requires non-empty type."
#End If
        Exit Function
    End If
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: control part selector requires non-empty part."
#End If
        Exit Function
    End If

    If g_ControlParts Is Nothing Then
        fn_TryResolveControlPartScope = True
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

    fn_TryResolveControlPartScope = True
End Function

' //
' // Internal
' //

Private Function private_RemoveEntriesByWorksheetKey( _
    ByRef entries As Collection, _
    ByVal worksheetKey As String _
) As Long
    Dim entryIndex As Long
    Dim entry As Object
    Dim entrySheetName As String

    worksheetKey = VBA.LCase$(VBA.Trim$(worksheetKey))
    If VBA.Len(worksheetKey) = 0 Then Exit Function
    If entries Is Nothing Then Exit Function

    For entryIndex = entries.Count To 1 Step -1
        Set entry = Nothing
        entrySheetName = VBA.vbNullString

        On Error Resume Next
        Set entry = entries.Item(entryIndex)
        If Not entry Is Nothing Then entrySheetName = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("SheetName"))))
        On Error GoTo 0

        If VBA.StrComp(entrySheetName, worksheetKey, VBA.vbTextCompare) = 0 Then
            On Error Resume Next
            If Not entry Is Nothing Then Set entry("Range") = Nothing
            entries.Remove entryIndex
            On Error GoTo 0
            private_RemoveEntriesByWorksheetKey = private_RemoveEntriesByWorksheetKey + 1
        End If
    Next entryIndex

    If entries.Count = 0 Then Set entries = Nothing
End Function

Private Sub private_EnsureControlPartsStorage()
    If Not g_ControlParts Is Nothing Then Exit Sub
    Set g_ControlParts = New Collection
End Sub

Private Sub private_EnsureControlColumnAliasesStorage()
    If Not g_ControlColumnAliases Is Nothing Then Exit Sub
    Set g_ControlColumnAliases = New Collection
End Sub
