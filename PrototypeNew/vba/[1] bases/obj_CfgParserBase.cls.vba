VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_CfgParserBase"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_IsDisposed As Boolean
Private m_ConfigTable As obj_ConfigTable

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    ' База только сохраняет ссылку на модель.
    ' Конкретный парсер уже решает, какие ключи обязательные/опциональные.
    m_IsDisposed = False
    Set m_ConfigTable = configTable
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_ConfigTable = Nothing
    On Error GoTo 0
End Sub

Public Property Get ConfigTable() As obj_ConfigTable
    Set ConfigTable = m_ConfigTable
End Property

Public Function TryGetConfigEntries(ByRef outConfigEntries As Collection) As Boolean
    Dim configEntries As list__obj_ConfigEntry

    ' Выгружаем строки из obj_ConfigTable в обычную коллекцию
    ' для быстрого обхода ключ/значение.
    Set outConfigEntries = Nothing
    If m_ConfigTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CfgParserBase: config table is not assigned."
#End If
        Exit Function
    End If

    Set configEntries = m_ConfigTable.Items
    If configEntries Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CfgParserBase: config entries are not initialized."
#End If
        Exit Function
    End If
    If configEntries.Count <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CfgParserBase: config entries are empty."
#End If
        Exit Function
    End If

    Set outConfigEntries = configEntries.AsCollection
    If outConfigEntries Is Nothing Then Exit Function
    If outConfigEntries.Count <= 0 Then Exit Function

    TryGetConfigEntries = True
End Function

Public Function BuildConfigDictionary( _
    ByVal configEntries As Collection, _
    ByRef outCfgMap As Object _
) As Boolean
    Dim entryObj As Variant
    Dim entry As obj_ConfigEntry
    Dim keyText As String

    ' Нормализуем строки таблицы в словарь:
    ' key   -> в нижний регистр + trim
    ' value -> trim
    ' При дубликатах побеждает последнее значение (поведение для override-конфигов).
    Set outCfgMap = ex_Helpers.fn_CreateDictionaryTextCompare()
    If outCfgMap Is Nothing Then Exit Function
    If configEntries Is Nothing Then
        BuildConfigDictionary = True
        Exit Function
    End If

    For Each entryObj In configEntries
        If Not VBA.IsObject(entryObj) Then GoTo ContinueEntry
        If VBA.StrComp(VBA.TypeName(entryObj), "obj_ConfigEntry", VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        Set entry = entryObj

        keyText = VBA.Trim$(entry.Key)
        If VBA.Len(keyText) = 0 Then GoTo ContinueEntry

        outCfgMap(VBA.LCase$(keyText)) = VBA.Trim$(entry.Value)
ContinueEntry:
    Next entryObj

    BuildConfigDictionary = True
End Function

Public Function TryGetRequiredConfigValue( _
    ByVal cfgMap As Object, _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    Dim normalizedKey As String

    ' Обязательный ключ: должен существовать и быть непустым.
    outValue = VBA.vbNullString
    If cfgMap Is Nothing Then Exit Function

    normalizedKey = VBA.LCase$(VBA.Trim$(keyName))
    If VBA.Len(normalizedKey) = 0 Then Exit Function
    If Not cfgMap.Exists(normalizedKey) Then Exit Function

    outValue = VBA.Trim$(VBA.CStr(cfgMap(normalizedKey)))
    If VBA.Len(outValue) = 0 Then Exit Function

    TryGetRequiredConfigValue = True
End Function

Public Function GetOptionalConfigValue( _
    ByVal cfgMap As Object, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = VBA.vbNullString _
) As String
    Dim normalizedKey As String

    ' Опциональный ключ: если есть - вернуть значение, иначе вернуть значение по умолчанию.
    GetOptionalConfigValue = defaultValue
    If cfgMap Is Nothing Then Exit Function

    normalizedKey = VBA.LCase$(VBA.Trim$(keyName))
    If VBA.Len(normalizedKey) = 0 Then Exit Function
    If Not cfgMap.Exists(normalizedKey) Then Exit Function

    GetOptionalConfigValue = VBA.Trim$(VBA.CStr(cfgMap(normalizedKey)))
End Function

Public Function TryGetRequiredConfigList( _
    ByVal cfgMap As Object, _
    ByVal keyName As String, _
    ByRef outItems As Collection _
) As Boolean
    Dim rawText As String

    ' Обязательный список: ключ обязателен, после разбиения должен быть хотя бы 1 элемент.
    Set outItems = Nothing
    If Not TryGetRequiredConfigValue(cfgMap, keyName, rawText) Then Exit Function

    Set outItems = SplitListToCollection(rawText)
    If outItems Is Nothing Then Exit Function
    If outItems.Count <= 0 Then Exit Function

    TryGetRequiredConfigList = True
End Function

Public Function SplitListToCollection(ByVal rawText As String) As Collection
    Dim result As Collection
    Dim parts As Variant
    Dim token As String
    Dim i As Long

    Set result = New Collection
    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then
        Set SplitListToCollection = result
        Exit Function
    End If

    ' Поддерживаем оба разделителя, чтобы конфиг было проще редактировать вручную.
    rawText = VBA.Replace$(rawText, ",", ";")
    parts = VBA.Split(rawText, ";")

    For i = LBound(parts) To UBound(parts)
        token = VBA.Trim$(VBA.CStr(parts(i)))
        If VBA.Len(token) > 0 Then result.Add token
    Next i

    Set SplitListToCollection = result
End Function

Public Function TryResolveWithOptionalResolver( _
    ByVal rawValue As String, _
    ByVal resolverName As String, _
    ByVal resolverArgs As String, _
    ByRef outResolvedValue As String _
) As Boolean
    Dim callName As String
    Dim resolvedValue As Variant

    outResolvedValue = VBA.Trim$(rawValue)
    resolverName = VBA.Trim$(resolverName)
    resolverArgs = VBA.Trim$(resolverArgs)

    ' Resolver опционален:
    ' - пустой resolverName => значение используется как есть
    ' - непустой resolverName => вызываем Application.Run(resolver, rawValue, resolverArgs)
    If VBA.Len(resolverName) = 0 Then
        TryResolveWithOptionalResolver = True
        Exit Function
    End If

    If VBA.InStr(1, resolverName, "!", VBA.vbBinaryCompare) > 0 Then
        callName = resolverName
    Else
        callName = "'" & ThisWorkbook.Name & "'!" & resolverName
    End If

    On Error GoTo EH_RESOLVE
    resolvedValue = Application.Run(callName, rawValue, resolverArgs)
    On Error GoTo 0

    outResolvedValue = VBA.Trim$(VBA.CStr(resolvedValue))
    If VBA.Len(outResolvedValue) = 0 Then Exit Function

    TryResolveWithOptionalResolver = True
    Exit Function

EH_RESOLVE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "CfgParserBase: resolver failed '" & resolverName & "': " & Err.Description
#End If
End Function

Public Function ResolvePathLocal(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = VBA.Trim$(inputPath)
    If VBA.Len(inputPath) = 0 Then Exit Function

    ' Абсолютный путь возвращаем как есть.
    ' Относительный путь резолвим относительно ThisWorkbook.Path.
    If VBA.Left$(inputPath, 2) = "\\" Or VBA.InStr(1, inputPath, ":\", VBA.vbTextCompare) > 0 Then
        ResolvePathLocal = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If VBA.Len(basePath) = 0 Then basePath = CurDir$
    If VBA.Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    ResolvePathLocal = basePath & inputPath
End Function
