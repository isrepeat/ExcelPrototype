Attribute VB_Name = "rt_HotkeyRuntime"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const MAX_HOTKEY_SLOTS As Long = 20

' Application.OnKey глобален на уровне Excel/книги; Excel не умеет ограничивать
' такую привязку конкретным листом.
' Поэтому этот модуль отвечает только за физическую глобальную связку:
'   OnKey token -> фиксированный slot macro -> rt_Bridge.fn_OnHotkey(hotkeyKey)
' Ограничение по странице происходит позже: rt_Bridge/PageBase смотрят активный лист.
' Словарь Pages под каждым хоткеем держит глобальную привязку живой, пока хотя бы
' одна страница продолжает иметь route для этого hotkey.
Private g_EntryByHotkey As Object
Private g_HotkeyBySlot As Object

Public Sub fn_Module_Dispose()
    fn_UnregisterAllHotkeys
End Sub

Public Function fn_TryNormalizeHotkey(ByVal hotkeyText As String, ByRef outHotkeyKey As String) As Boolean
    Dim parts() As String
    Dim part As Variant
    Dim token As String
    Dim keyPart As String
    Dim hasCtrl As Boolean
    Dim hasShift As Boolean
    Dim hasAlt As Boolean

    outHotkeyKey = VBA.vbNullString
    hotkeyText = VBA.UCase$(VBA.Trim$(hotkeyText))
    hotkeyText = VBA.Replace$(hotkeyText, " ", VBA.vbNullString)
    If VBA.Len(hotkeyText) = 0 Then Exit Function

    parts = VBA.Split(hotkeyText, "+")
    For Each part In parts
        token = VBA.UCase$(VBA.Trim$(VBA.CStr(part)))
        If VBA.Len(token) = 0 Then GoTo ContinuePart

        Select Case token
            Case "CTRL", "CONTROL", "^"
                hasCtrl = True
            Case "SHIFT", "+"
                hasShift = True
            Case "ALT", "OPTION", "%"
                hasAlt = True
            Case Else
                If VBA.Len(keyPart) > 0 Then Exit Function
                keyPart = private_NormalizeKeyPart(token)
                If VBA.Len(keyPart) = 0 Then Exit Function
        End Select

ContinuePart:
    Next part

    If VBA.Len(keyPart) = 0 Then Exit Function

    If hasCtrl Then outHotkeyKey = outHotkeyKey & "^"
    If hasShift Then outHotkeyKey = outHotkeyKey & "+"
    If hasAlt Then outHotkeyKey = outHotkeyKey & "%"
    outHotkeyKey = outHotkeyKey & keyPart
    fn_TryNormalizeHotkey = True
End Function

Public Function fn_RegisterPageHotkey(ByVal pageId As String, ByVal hotkeyKey As String) As Boolean
    Dim entry As Object
    Dim pages As Object
    Dim slotIndex As Long

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then Exit Function
    If VBA.Len(hotkeyKey) = 0 Then Exit Function

    private_EnsureStorage

    ' Если другая страница уже использует этот физический хоткей, оставляем одну
    ' OnKey-привязку и просто добавляем pageId. Dispatch выберет нужный route
    ' по активному листу.
    If g_EntryByHotkey.Exists(hotkeyKey) Then
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        pages(pageId) = True
        fn_RegisterPageHotkey = True
        Exit Function
    End If

    slotIndex = private_AllocateSlotIndex()
    If slotIndex <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "hotkey-runtime: no free OnKey slot for '" & private_EscapeForLog(hotkeyKey) & "'."
#End If
        VBA.MsgBox "PrototypeNew: no free hotkey slots. Increase MAX_HOTKEY_SLOTS in rt_HotkeyRuntime.", VBA.vbExclamation, "PrototypeNew / Hotkeys"
        Exit Function
    End If

    Set pages = VBA.CreateObject("Scripting.Dictionary")
    pages.CompareMode = 1
    pages(pageId) = True

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("SlotIndex") = slotIndex
    Set entry("Pages") = pages

    Set g_EntryByHotkey(hotkeyKey) = entry
    g_HotkeyBySlot(VBA.CStr(slotIndex)) = hotkeyKey

    ' OnKey не умеет надежно передавать произвольные аргументы в один общий макрос,
    ' поэтому каждый уникальный hotkey ведет в фиксированный slot macro, а token
    ' хоткея достается из g_HotkeyBySlot.
    If Not private_AssignOnKey(hotkeyKey, private_BuildSlotMacroRef(slotIndex)) Then
        g_HotkeyBySlot.Remove VBA.CStr(slotIndex)
        g_EntryByHotkey.Remove hotkeyKey
        Exit Function
    End If

    fn_RegisterPageHotkey = True
End Function

Public Sub fn_UnregisterPageHotkeys(ByVal pageId As String)
    Dim hotkeyKey As Variant
    Dim keysToRemove As Collection
    Dim entry As Object
    Dim pages As Object
    Dim removeKey As Variant

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then Exit Sub
    If g_EntryByHotkey Is Nothing Then Exit Sub

    ' Убираем только эту страницу из общих hotkeys. Физическая OnKey-привязка
    ' освобождается только когда ни одна страница больше не ссылается на token.
    Set keysToRemove = New Collection
    For Each hotkeyKey In g_EntryByHotkey.Keys
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        If Not pages Is Nothing Then
            If pages.Exists(pageId) Then pages.Remove pageId
            If pages.Count = 0 Then keysToRemove.Add VBA.CStr(hotkeyKey)
        End If
    Next hotkeyKey

    For Each removeKey In keysToRemove
        private_UnregisterHotkey VBA.CStr(removeKey)
    Next removeKey
End Sub

Public Sub fn_UnregisterAllHotkeys()
    Dim hotkeyKey As Variant
    Dim keysToRemove As Collection
    Dim removeKey As Variant

    If g_EntryByHotkey Is Nothing Then Exit Sub

    Set keysToRemove = New Collection
    For Each hotkeyKey In g_EntryByHotkey.Keys
        keysToRemove.Add VBA.CStr(hotkeyKey)
    Next hotkeyKey

    For Each removeKey In keysToRemove
        private_UnregisterHotkey VBA.CStr(removeKey)
    Next removeKey

    Set g_EntryByHotkey = Nothing
    Set g_HotkeyBySlot = Nothing
End Sub

Public Sub fn_DispatchSlot(ByVal slotIndex As Long)
    Dim hotkeyKey As String

    private_EnsureStorage
    hotkeyKey = VBA.vbNullString
    If g_HotkeyBySlot.Exists(VBA.CStr(slotIndex)) Then
        hotkeyKey = VBA.CStr(g_HotkeyBySlot(VBA.CStr(slotIndex)))
    End If
    If VBA.Len(hotkeyKey) = 0 Then Exit Sub

    ' Bridge определяет страницу по активному листу и вызывает локальный route страницы.
    rt_Bridge.fn_OnHotkey hotkeyKey
End Sub

Public Sub fn_OnHotkeySlot01()
    fn_DispatchSlot 1
End Sub

Public Sub fn_OnHotkeySlot02()
    fn_DispatchSlot 2
End Sub

Public Sub fn_OnHotkeySlot03()
    fn_DispatchSlot 3
End Sub

Public Sub fn_OnHotkeySlot04()
    fn_DispatchSlot 4
End Sub

Public Sub fn_OnHotkeySlot05()
    fn_DispatchSlot 5
End Sub

Public Sub fn_OnHotkeySlot06()
    fn_DispatchSlot 6
End Sub

Public Sub fn_OnHotkeySlot07()
    fn_DispatchSlot 7
End Sub

Public Sub fn_OnHotkeySlot08()
    fn_DispatchSlot 8
End Sub

Public Sub fn_OnHotkeySlot09()
    fn_DispatchSlot 9
End Sub

Public Sub fn_OnHotkeySlot10()
    fn_DispatchSlot 10
End Sub

Public Sub fn_OnHotkeySlot11()
    fn_DispatchSlot 11
End Sub

Public Sub fn_OnHotkeySlot12()
    fn_DispatchSlot 12
End Sub

Public Sub fn_OnHotkeySlot13()
    fn_DispatchSlot 13
End Sub

Public Sub fn_OnHotkeySlot14()
    fn_DispatchSlot 14
End Sub

Public Sub fn_OnHotkeySlot15()
    fn_DispatchSlot 15
End Sub

Public Sub fn_OnHotkeySlot16()
    fn_DispatchSlot 16
End Sub

Public Sub fn_OnHotkeySlot17()
    fn_DispatchSlot 17
End Sub

Public Sub fn_OnHotkeySlot18()
    fn_DispatchSlot 18
End Sub

Public Sub fn_OnHotkeySlot19()
    fn_DispatchSlot 19
End Sub

Public Sub fn_OnHotkeySlot20()
    fn_DispatchSlot 20
End Sub

Private Sub private_EnsureStorage()
    If g_EntryByHotkey Is Nothing Then
        Set g_EntryByHotkey = VBA.CreateObject("Scripting.Dictionary")
        g_EntryByHotkey.CompareMode = 1
    End If

    If g_HotkeyBySlot Is Nothing Then
        Set g_HotkeyBySlot = VBA.CreateObject("Scripting.Dictionary")
        g_HotkeyBySlot.CompareMode = 1
    End If
End Sub

Private Function private_AllocateSlotIndex() As Long
    Dim idx As Long

    private_EnsureStorage
    For idx = 1 To MAX_HOTKEY_SLOTS
        If Not g_HotkeyBySlot.Exists(VBA.CStr(idx)) Then
            private_AllocateSlotIndex = idx
            Exit Function
        End If
    Next idx
End Function

Private Function private_AssignOnKey(ByVal hotkeyKey As String, ByVal macroRef As String) As Boolean
    On Error GoTo EH_ASSIGN
    Application.OnKey hotkeyKey, macroRef
    private_AssignOnKey = True
    Exit Function

EH_ASSIGN:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "hotkey-runtime: failed to assign OnKey '" & private_EscapeForLog(hotkeyKey) & "': " & Err.Description
#End If
End Function

Private Sub private_UnregisterHotkey(ByVal hotkeyKey As String)
    Dim entry As Object
    Dim slotIndex As Long

    If g_EntryByHotkey Is Nothing Then Exit Sub
    If VBA.Len(hotkeyKey) = 0 Then Exit Sub

    On Error Resume Next
    Application.OnKey hotkeyKey
    On Error GoTo 0

    If g_EntryByHotkey.Exists(hotkeyKey) Then
        Set entry = g_EntryByHotkey(hotkeyKey)
        slotIndex = VBA.CLng(entry("SlotIndex"))
        If Not g_HotkeyBySlot Is Nothing Then
            If g_HotkeyBySlot.Exists(VBA.CStr(slotIndex)) Then g_HotkeyBySlot.Remove VBA.CStr(slotIndex)
        End If
        g_EntryByHotkey.Remove hotkeyKey
    End If
End Sub

Private Function private_BuildSlotMacroRef(ByVal slotIndex As Long) As String
    Dim wbName As String

    wbName = VBA.Replace$(ThisWorkbook.Name, "'", "''")
    private_BuildSlotMacroRef = "'" & wbName & "'!rt_HotkeyRuntime.fn_OnHotkeySlot" & VBA.Format$(slotIndex, "00")
End Function

Private Function private_NormalizeKeyPart(ByVal token As String) As String
    token = VBA.UCase$(VBA.Trim$(token))
    Select Case token
        Case "ENTER", "RETURN"
            ' Основной Enter/Return в Application.OnKey — это "~".
            ' "{ENTER}" означает Enter на цифровой клавиатуре.
            private_NormalizeKeyPart = "~"
        Case "NUMENTER", "NUMPADENTER"
            private_NormalizeKeyPart = "{ENTER}"
        Case "ESC", "ESCAPE"
            private_NormalizeKeyPart = "{ESC}"
        Case "TAB"
            private_NormalizeKeyPart = "{TAB}"
        Case "BACKSPACE", "BKSP"
            private_NormalizeKeyPart = "{BACKSPACE}"
        Case "DELETE", "DEL"
            private_NormalizeKeyPart = "{DELETE}"
        Case "SPACE"
            private_NormalizeKeyPart = " "
        Case "UP", "DOWN", "LEFT", "RIGHT", "HOME", "END", "PGUP", "PGDN"
            private_NormalizeKeyPart = "{" & token & "}"
        Case Else
            If VBA.Len(token) = 1 Then
                ' Заглавные буквы в OnKey фактически подразумевают Shift.
                ' Поэтому буквы храним в lowercase, а Shift выражаем только через "+".
                If token Like "[A-Z]" Then
                    private_NormalizeKeyPart = VBA.LCase$(token)
                Else
                    private_NormalizeKeyPart = token
                End If
            ElseIf VBA.Left$(token, 1) = "F" And VBA.Len(token) <= 3 Then
                private_NormalizeKeyPart = "{" & token & "}"
            End If
    End Select
End Function

Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace$(VBA.CStr(valueText), "'", "''")
End Function
