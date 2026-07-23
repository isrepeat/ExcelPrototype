Attribute VB_Name = "ex_ControlPartsRuntime"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Runtime-индекс визуальных частей control используется selector-движком как
' связь "control/part/source alias -> Range". При partial render весь индекс не
' сбрасывается: записи target удаляются и создаются его renderer-ом заново, а
' ranges перенесённых siblings транслируются вместе с worksheet subtree.
'
' Это не только оптимизация. Если metadata отстанет от листа, последующий
' локальный style pass применит rule к старым координатам другого контрола.
Private g_ControlParts As Collection
Private g_ControlPartShapes As Object
Private g_ControlColumnAliases As Collection
Private g_ControlSourceAliases As Collection

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_ControlPartsRuntime.fn_Module_Dispose"
#End If
    On Error Resume Next
    Set g_ControlParts = Nothing
    Set g_ControlPartShapes = Nothing
    Set g_ControlColumnAliases = Nothing
    Set g_ControlSourceAliases = Nothing
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Sub fn_ResetControlParts()
    Set g_ControlParts = Nothing
    Set g_ControlPartShapes = Nothing
    Set g_ControlColumnAliases = Nothing
    Set g_ControlSourceAliases = Nothing
End Sub

Public Function fn_RemoveControlPartsByWorksheetName(ByVal worksheetName As String) As Boolean
    Dim worksheetKey As String

    worksheetKey = VBA.LCase$(VBA.Trim$(worksheetName))
    If VBA.Len(worksheetKey) = 0 Then
        fn_RemoveControlPartsByWorksheetName = True
        Exit Function
    End If

    Call private_RemoveEntriesByWorksheetKey(g_ControlParts, worksheetKey)
    private_RemoveControlPartShapeBuckets worksheetKey
    Call private_RemoveEntriesByWorksheetKey(g_ControlColumnAliases, worksheetKey)
    Call private_RemoveEntriesByWorksheetKey(g_ControlSourceAliases, worksheetKey)

    fn_RemoveControlPartsByWorksheetName = True
End Function


Public Function fn_RemoveControlPartsByControl( _
    ByVal worksheetName As String, _
    ByVal controlName As String _
) As Boolean
    Dim worksheetKey As String
    Dim controlKey As String

    worksheetKey = VBA.LCase$(VBA.Trim$(worksheetName))
    controlKey = VBA.LCase$(VBA.Trim$(controlName))
    If VBA.Len(worksheetKey) = 0 Or VBA.Len(controlKey) = 0 Then Exit Function

    ' Локальный render заменяет только parts изменившегося контрола.
    ' Parts остальных контролов нужны retained style pipeline и не сбрасываются.
    Call private_RemoveEntriesByControlKey(g_ControlParts, worksheetKey, controlKey)
    private_RemoveControlPartShapeBuckets worksheetKey, controlKey
    Call private_RemoveEntriesByControlKey(g_ControlColumnAliases, worksheetKey, controlKey)
    Call private_RemoveEntriesByControlKey(g_ControlSourceAliases, worksheetKey, controlKey)

    fn_RemoveControlPartsByControl = True
End Function


Public Function fn_TryGetControlVisualScope( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByRef outScope As Range _
) As Boolean
    Dim entry As Variant
    Dim entryRange As Range

    Set outScope = Nothing
    If ws Is Nothing Then Exit Function
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    If VBA.Len(controlName) = 0 Then Exit Function
    If g_ControlParts Is Nothing Then
        fn_TryGetControlVisualScope = True
        Exit Function
    End If

    For Each entry In g_ControlParts
        If VBA.LCase$(VBA.CStr(entry("SheetName"))) <> VBA.LCase$(ws.Name) Then GoTo ContinueEntry
        If VBA.LCase$(VBA.CStr(entry("ControlName"))) <> controlName Then GoTo ContinueEntry
        Set entryRange = Nothing
        On Error Resume Next
        Set entryRange = entry("Range")
        On Error GoTo 0
        If entryRange Is Nothing Then GoTo ContinueEntry
        If outScope Is Nothing Then
            Set outScope = entryRange
        Else
            Set outScope = Application.Union(outScope, entryRange)
        End If
ContinueEntry:
    Next entry
    fn_TryGetControlVisualScope = True
End Function


Public Function fn_TranslateControlPartsBelow( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal rowDelta As Long _
) As Boolean
    If ws Is Nothing Or firstRow <= 0 Then Exit Function
    If rowDelta = 0 Then
        fn_TranslateControlPartsBelow = True
        Exit Function
    End If

    If Not private_TranslateEntryRangesBelow(g_ControlParts, ws, firstRow, rowDelta) Then Exit Function
    If Not private_TranslateEntryRangesBelow(g_ControlColumnAliases, ws, firstRow, rowDelta) Then Exit Function
    If Not private_TranslateEntryRangesBelow(g_ControlSourceAliases, ws, firstRow, rowDelta) Then Exit Function

    fn_TranslateControlPartsBelow = True
End Function


Public Function fn_TranslateControlPartsInRegion( _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
) As Boolean
    If ws Is Nothing Then Exit Function

    ' Быстрый reflow переносит generated UI через Copy, поэтому сохранённые
    ' Excel Range остаются привязаны к исходным адресам. Явно переводим metadata
    ' всех частей, полностью входящих в перемещаемый layout patch.
    If Not private_TranslateEntryRangesInRegion( _
        g_ControlParts, ws, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If Not private_TranslateEntryRangesInRegion( _
        g_ControlColumnAliases, ws, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If Not private_TranslateEntryRangesInRegion( _
        g_ControlSourceAliases, ws, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function

    fn_TranslateControlPartsInRegion = True
End Function

Public Function fn_RegisterControlSourceAlias( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal sourceAlias As String, _
    ByVal sourceAliasTemplate As String, _
    ByVal sourceRange As Range _
) As Boolean
    Dim entry As Object

    If ws Is Nothing Or sourceRange Is Nothing Then Exit Function
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    sourceAlias = VBA.LCase$(VBA.Trim$(sourceAlias))
    sourceAliasTemplate = VBA.LCase$(VBA.Trim$(sourceAliasTemplate))
    If VBA.Len(controlType) = 0 Then Exit Function
    If VBA.Len(sourceAlias) = 0 And VBA.Len(sourceAliasTemplate) = 0 Then Exit Function

    private_EnsureControlSourceAliasesStorage
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("SheetName") = VBA.LCase$(ws.Name)
    entry("ControlType") = controlType
    entry("ControlName") = controlName
    entry("SourceAlias") = sourceAlias
    entry("SourceAliasTemplate") = sourceAliasTemplate
    Set entry("Range") = sourceRange
    g_ControlSourceAliases.Add entry
    fn_RegisterControlSourceAlias = True
End Function

Public Function fn_TryResolveControlSourceAliasScope( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal sourceAlias As String, _
    ByVal sourceAliasTemplate As String, _
    ByRef outScope As Range _
) As Boolean
    Dim entry As Object
    Dim aliasRange As Range
    Dim wsKey As String

    If ws Is Nothing Then Exit Function
    wsKey = VBA.LCase$(ws.Name)
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    sourceAlias = VBA.LCase$(VBA.Trim$(sourceAlias))
    sourceAliasTemplate = VBA.LCase$(VBA.Trim$(sourceAliasTemplate))
    If VBA.Len(controlType) = 0 Then Exit Function

    If g_ControlSourceAliases Is Nothing Then
        fn_TryResolveControlSourceAliasScope = True
        Exit Function
    End If

    For Each entry In g_ControlSourceAliases
        If VBA.LCase$(VBA.CStr(entry("SheetName"))) <> wsKey Then GoTo ContinueEntry
        If VBA.LCase$(VBA.CStr(entry("ControlType"))) <> controlType Then GoTo ContinueEntry
        If VBA.Len(controlName) > 0 Then
            If VBA.LCase$(VBA.CStr(entry("ControlName"))) <> controlName Then GoTo ContinueEntry
        End If
        If VBA.Len(sourceAlias) > 0 Then
            If VBA.LCase$(VBA.CStr(entry("SourceAlias"))) <> sourceAlias Then GoTo ContinueEntry
        End If
        If VBA.Len(sourceAliasTemplate) > 0 Then
            If VBA.LCase$(VBA.CStr(entry("SourceAliasTemplate"))) <> sourceAliasTemplate Then GoTo ContinueEntry
        End If

        Set aliasRange = Nothing
        On Error Resume Next
        Set aliasRange = entry("Range")
        On Error GoTo 0
        If aliasRange Is Nothing Then GoTo ContinueEntry
        If outScope Is Nothing Then
            Set outScope = aliasRange
        Else
            Set outScope = Application.Union(outScope, aliasRange)
        End If
ContinueEntry:
    Next entry

    fn_TryResolveControlSourceAliasScope = True
End Function


Public Function fn_RegisterControlPart( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal partName As String, _
    ByVal partRange As Range, _
    Optional ByVal partShape As Variant _
) As Boolean
    Dim entry As Object
    Dim registeredShape As Shape
    Dim shapeBucket As Collection
    Dim shapeBucketKey As String

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

    ' Shape индексируется тем же semantic part, что и Range. Style pipeline
    ' получает точные targets без повторного перебора всей ws.Shapes.
    If VBA.IsObject(partShape) Then Set registeredShape = partShape
    If Not registeredShape Is Nothing Then
        private_EnsureControlPartShapesStorage
        shapeBucketKey = private_ControlPartShapeBucketKey( _
            ws.Name, controlType, controlName, partName)
        If g_ControlPartShapes.Exists(shapeBucketKey) Then
            Set shapeBucket = g_ControlPartShapes(shapeBucketKey)
        Else
            Set shapeBucket = New Collection
            Set g_ControlPartShapes(shapeBucketKey) = shapeBucket
        End If
        shapeBucket.Add registeredShape
    End If
    fn_RegisterControlPart = True
End Function

Public Function fn_TryResolveControlPartShapes( _
    ByVal ws As Worksheet, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal partName As String, _
    ByRef outShapes As Collection _
) As Boolean
    Dim bucketKey As Variant
    Dim bucket As Collection
    Dim shapeItem As Variant
    Dim exactKey As String
    Dim keyPrefix As String
    Dim keySuffix As String

    Set outShapes = New Collection
    If ws Is Nothing Then Exit Function
    controlType = VBA.LCase$(VBA.Trim$(controlType))
    controlName = VBA.LCase$(VBA.Trim$(controlName))
    partName = VBA.LCase$(VBA.Trim$(partName))
    If VBA.Len(controlType) = 0 Or VBA.Len(partName) = 0 Then Exit Function
    If g_ControlPartShapes Is Nothing Then
        fn_TryResolveControlPartShapes = True
        Exit Function
    End If

    If VBA.Len(controlName) > 0 Then
        exactKey = private_ControlPartShapeBucketKey(ws.Name, controlType, controlName, partName)
        If g_ControlPartShapes.Exists(exactKey) Then
            Set bucket = g_ControlPartShapes(exactKey)
            For Each shapeItem In bucket
                outShapes.Add shapeItem
            Next shapeItem
        End If
        fn_TryResolveControlPartShapes = True
        Exit Function
    End If

    keyPrefix = VBA.LCase$(ws.Name) & "|" & controlType & "|"
    keySuffix = "|" & partName
    For Each bucketKey In g_ControlPartShapes.Keys
        If VBA.Left$(VBA.CStr(bucketKey), VBA.Len(keyPrefix)) <> keyPrefix Then GoTo ContinueBucket
        If VBA.Right$(VBA.CStr(bucketKey), VBA.Len(keySuffix)) <> keySuffix Then GoTo ContinueBucket
        Set bucket = g_ControlPartShapes(bucketKey)
        For Each shapeItem In bucket
            outShapes.Add shapeItem
        Next shapeItem
ContinueBucket:
    Next bucketKey

    fn_TryResolveControlPartShapes = True
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

Private Function private_RemoveEntriesByControlKey( _
    ByRef entries As Collection, _
    ByVal worksheetKey As String, _
    ByVal controlKey As String _
) As Long
    Dim entryIndex As Long
    Dim entry As Object
    Dim entrySheetName As String
    Dim entryControlName As String

    If entries Is Nothing Then Exit Function

    For entryIndex = entries.Count To 1 Step -1
        Set entry = Nothing
        entrySheetName = VBA.vbNullString
        entryControlName = VBA.vbNullString
        On Error Resume Next
        Set entry = entries.Item(entryIndex)
        If Not entry Is Nothing Then
            entrySheetName = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("SheetName"))))
            entryControlName = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("ControlName"))))
        End If
        On Error GoTo 0

        If entrySheetName = worksheetKey And entryControlName = controlKey Then
            On Error Resume Next
            Set entry("Range") = Nothing
            entries.Remove entryIndex
            On Error GoTo 0
            private_RemoveEntriesByControlKey = private_RemoveEntriesByControlKey + 1
        End If
    Next entryIndex

    If entries.Count = 0 Then Set entries = Nothing
End Function

Private Function private_TranslateEntryRangesBelow( _
    ByRef entries As Collection, _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim entry As Variant
    Dim entryRange As Range
    Dim translatedRange As Range
    Dim newRow As Long

    If entries Is Nothing Then
        private_TranslateEntryRangesBelow = True
        Exit Function
    End If

    For Each entry In entries
        If VBA.LCase$(VBA.Trim$(VBA.CStr(entry("SheetName")))) <> VBA.LCase$(ws.Name) Then GoTo ContinueEntry
        Set entryRange = Nothing
        On Error Resume Next
        Set entryRange = entry("Range")
        On Error GoTo 0
        If entryRange Is Nothing Then GoTo ContinueEntry
        If entryRange.Row < firstRow Then GoTo ContinueEntry

        newRow = entryRange.Row + rowDelta
        If newRow <= 0 Then Exit Function
        Set translatedRange = ws.Range( _
            ws.Cells(newRow, entryRange.Column), _
            ws.Cells(newRow + entryRange.Rows.Count - 1, entryRange.Column + entryRange.Columns.Count - 1))
        Set entry("Range") = translatedRange
ContinueEntry:
    Next entry

    private_TranslateEntryRangesBelow = True
End Function

Private Function private_TranslateEntryRangesInRegion( _
    ByRef entries As Collection, _
    ByVal ws As Worksheet, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim entry As Variant
    Dim entryRange As Range
    Dim translatedRange As Range
    Dim entryRowEnd As Long
    Dim entryColEnd As Long
    Dim newRow As Long

    If entries Is Nothing Or rowDelta = 0 Then
        private_TranslateEntryRangesInRegion = True
        Exit Function
    End If

    For Each entry In entries
        If VBA.LCase$(VBA.Trim$(VBA.CStr(entry("SheetName")))) <> _
            VBA.LCase$(ws.Name) Then GoTo ContinueEntry

        Set entryRange = Nothing
        On Error Resume Next
        Set entryRange = entry("Range")
        On Error GoTo 0
        If entryRange Is Nothing Then GoTo ContinueEntry

        entryRowEnd = entryRange.Row + entryRange.Rows.Count - 1
        entryColEnd = entryRange.Column + entryRange.Columns.Count - 1
        If entryRange.Row < rowStart Or entryRowEnd > rowEnd Then GoTo ContinueEntry
        If entryRange.Column < colStart Or entryColEnd > colEnd Then GoTo ContinueEntry

        newRow = entryRange.Row + rowDelta
        If newRow <= 0 Then Exit Function
        Set translatedRange = ws.Range( _
            ws.Cells(newRow, entryRange.Column), _
            ws.Cells(newRow + entryRange.Rows.Count - 1, entryColEnd))
        Set entry("Range") = translatedRange
ContinueEntry:
    Next entry

    private_TranslateEntryRangesInRegion = True
End Function

Private Sub private_EnsureControlPartsStorage()
    If Not g_ControlParts Is Nothing Then Exit Sub
    Set g_ControlParts = New Collection
End Sub

Private Sub private_EnsureControlPartShapesStorage()
    If Not g_ControlPartShapes Is Nothing Then Exit Sub
    Set g_ControlPartShapes = VBA.CreateObject("Scripting.Dictionary")
    g_ControlPartShapes.CompareMode = 1
End Sub

Private Function private_ControlPartShapeBucketKey( _
    ByVal sheetName As String, _
    ByVal controlType As String, _
    ByVal controlName As String, _
    ByVal partName As String _
) As String
    private_ControlPartShapeBucketKey = VBA.LCase$(VBA.Trim$(sheetName)) & "|" & _
        VBA.LCase$(VBA.Trim$(controlType)) & "|" & _
        VBA.LCase$(VBA.Trim$(controlName)) & "|" & _
        VBA.LCase$(VBA.Trim$(partName))
End Function

Private Sub private_RemoveControlPartShapeBuckets( _
    ByVal worksheetKey As String, _
    Optional ByVal controlKey As String = VBA.vbNullString _
)
    Dim bucketKey As Variant
    Dim keysToRemove As Collection
    Dim removeKey As Variant
    Dim keyPrefix As String
    Dim controlMarker As String

    If g_ControlPartShapes Is Nothing Then Exit Sub
    worksheetKey = VBA.LCase$(VBA.Trim$(worksheetKey))
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    keyPrefix = worksheetKey & "|"
    controlMarker = "|" & controlKey & "|"
    Set keysToRemove = New Collection

    For Each bucketKey In g_ControlPartShapes.Keys
        If VBA.Left$(VBA.CStr(bucketKey), VBA.Len(keyPrefix)) <> keyPrefix Then GoTo ContinueKey
        If VBA.Len(controlKey) > 0 Then
            If VBA.InStr(1, VBA.CStr(bucketKey), controlMarker, VBA.vbBinaryCompare) = 0 Then GoTo ContinueKey
        End If
        keysToRemove.Add VBA.CStr(bucketKey)
ContinueKey:
    Next bucketKey
    For Each removeKey In keysToRemove
        g_ControlPartShapes.Remove VBA.CStr(removeKey)
    Next removeKey
End Sub

Private Sub private_EnsureControlColumnAliasesStorage()
    If Not g_ControlColumnAliases Is Nothing Then Exit Sub
    Set g_ControlColumnAliases = New Collection
End Sub

Private Sub private_EnsureControlSourceAliasesStorage()
    If Not g_ControlSourceAliases Is Nothing Then Exit Sub
    Set g_ControlSourceAliases = New Collection
End Sub
