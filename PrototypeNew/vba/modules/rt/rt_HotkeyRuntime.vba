Attribute VB_Name = "rt_HotkeyRuntime"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const MAX_HOTKEY_SLOTS As Long = 20

' Application.OnKey глобален на уровне Excel/книги; Excel не умеет ограничивать
' такую привязку конкретным листом. Поэтому runtime хранит routes всех страниц,
' но физически включает OnKey только для активной страницы.
Private g_EntryByHotkey As Object
Private g_HotkeyBySlot As Object
Private g_ActiveHotkeyByKey As Object
Private g_ActivePageId As String
Private g_IsShuttingDown As Boolean

Public Sub fn_Module_Dispose()
    fn_UnregisterAllHotkeys
End Sub

Public Sub fn_BeginSession()
    g_IsShuttingDown = False
End Sub

Public Sub fn_BeginShutdown()
    ' Сначала запрещаем любые поздние регистрации из Workbook_Deactivate,
    ' render/restore callbacks и только затем снимаем глобальные OnKey routes.
    g_IsShuttingDown = True
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

    If g_IsShuttingDown Then
        fn_RegisterPageHotkey = True
        Exit Function
    End If

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then Exit Function
    If VBA.Len(hotkeyKey) = 0 Then Exit Function

    private_EnsureStorage

    ' Если другая страница уже использует этот token, оставляем один slot macro
    ' и просто добавляем pageId. Физически OnKey включается только для active page.
    If g_EntryByHotkey.Exists(hotkeyKey) Then
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        pages(pageId) = True
        If VBA.StrComp(g_ActivePageId, pageId, VBA.vbTextCompare) = 0 Then
            If Not private_AssignPhysicalHotkey(hotkeyKey) Then Exit Function
        End If
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

    If VBA.StrComp(g_ActivePageId, pageId, VBA.vbTextCompare) = 0 Then
        If Not private_AssignPhysicalHotkey(hotkeyKey) Then Exit Function
    End If

    fn_RegisterPageHotkey = True
End Function

Public Function fn_ActivatePageHotkeys(ByVal pageId As String) As Boolean
    Dim hotkeyKey As Variant
    Dim entry As Object
    Dim pages As Object
    Dim hasFailure As Boolean

    If g_IsShuttingDown Then
        fn_ActivatePageHotkeys = True
        Exit Function
    End If

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    private_EnsureStorage

    ' Повторный render активной страницы часто заново строит те же page-local routes.
    ' Если физические Application.OnKey уже соответствуют этой странице, не снимаем
    ' и не назначаем их повторно: OnKey — одна из самых дорогих частей rerender-а.
    If private_AreActivePageHotkeysCurrent(pageId) Then
        fn_ActivatePageHotkeys = True
        Exit Function
    End If

    private_UnassignActivePhysicalHotkeys
    g_ActivePageId = pageId

    If VBA.Len(pageId) = 0 Then
        fn_ActivatePageHotkeys = True
        Exit Function
    End If

    For Each hotkeyKey In g_EntryByHotkey.Keys
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        If Not pages Is Nothing Then
            If pages.Exists(pageId) Then
                If Not private_AssignPhysicalHotkey(VBA.CStr(hotkeyKey)) Then hasFailure = True
            End If
        End If
    Next hotkeyKey

    private_RemoveEmptyHotkeyEntries
    fn_ActivatePageHotkeys = Not hasFailure
End Function

Private Function private_AreActivePageHotkeysCurrent(ByVal pageId As String) As Boolean
    Dim expectedKeys As Object
    Dim hotkeyKey As Variant
    Dim entry As Object
    Dim pages As Object

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then
        private_AreActivePageHotkeysCurrent = (VBA.Len(VBA.Trim$(g_ActivePageId)) = 0)
        Exit Function
    End If
    If VBA.StrComp(g_ActivePageId, pageId, VBA.vbTextCompare) <> 0 Then Exit Function
    If g_EntryByHotkey Is Nothing Then Exit Function
    If g_ActiveHotkeyByKey Is Nothing Then Exit Function

    Set expectedKeys = VBA.CreateObject("Scripting.Dictionary")
    expectedKeys.CompareMode = 1

    For Each hotkeyKey In g_EntryByHotkey.Keys
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        If Not pages Is Nothing Then
            If pages.Exists(pageId) Then expectedKeys(VBA.CStr(hotkeyKey)) = True
        End If
    Next hotkeyKey

    If expectedKeys.Count <> g_ActiveHotkeyByKey.Count Then Exit Function
    For Each hotkeyKey In expectedKeys.Keys
        If Not g_ActiveHotkeyByKey.Exists(VBA.CStr(hotkeyKey)) Then Exit Function
    Next hotkeyKey

    private_AreActivePageHotkeysCurrent = True
End Function

Public Sub fn_UnregisterPageHotkeys(ByVal pageId As String, Optional ByVal keepActivePhysicalHotkeys As Boolean = False)
    Dim hotkeyKey As Variant
    Dim keysToRemove As Collection
    Dim entry As Object
    Dim pages As Object
    Dim removeKey As Variant
    Dim keepPhysicalForActivePage As Boolean

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then Exit Sub
    If g_EntryByHotkey Is Nothing Then Exit Sub

    keepPhysicalForActivePage = _
        (keepActivePhysicalHotkeys And VBA.StrComp(g_ActivePageId, pageId, VBA.vbTextCompare) = 0)

    ' Убираем только эту страницу из общих hotkeys. Физическая OnKey-привязка
    ' обычно освобождается только когда ни одна страница больше не ссылается на token.
    ' Во время rerender активной страницы можно временно оставить физические OnKey:
    ' если после render будут зарегистрированы те же hotkeys, Excel не придется
    ' снимать и назначать их заново.
    Set keysToRemove = New Collection
    For Each hotkeyKey In g_EntryByHotkey.Keys
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        If Not pages Is Nothing Then
            If pages.Exists(pageId) Then pages.Remove pageId
            If pages.Count = 0 And Not keepPhysicalForActivePage Then keysToRemove.Add VBA.CStr(hotkeyKey)
        End If
    Next hotkeyKey

    For Each removeKey In keysToRemove
        private_UnregisterHotkey VBA.CStr(removeKey)
    Next removeKey

    If VBA.StrComp(g_ActivePageId, pageId, VBA.vbTextCompare) = 0 And Not keepPhysicalForActivePage Then
        private_UnassignActivePhysicalHotkeys
        g_ActivePageId = VBA.vbNullString
    End If
End Sub

Public Sub fn_UnregisterAllHotkeys()
    Dim hotkeyKey As Variant
    Dim keysToRemove As Collection
    Dim removeKey As Variant

    private_UnassignActivePhysicalHotkeys
    g_ActivePageId = VBA.vbNullString

    If g_EntryByHotkey Is Nothing Then
        Set g_HotkeyBySlot = Nothing
        Set g_ActiveHotkeyByKey = Nothing
        Exit Sub
    End If

    Set keysToRemove = New Collection
    For Each hotkeyKey In g_EntryByHotkey.Keys
        keysToRemove.Add VBA.CStr(hotkeyKey)
    Next hotkeyKey

    For Each removeKey In keysToRemove
        private_UnregisterHotkey VBA.CStr(removeKey)
    Next removeKey

    Set g_EntryByHotkey = Nothing
    Set g_HotkeyBySlot = Nothing
    Set g_ActiveHotkeyByKey = Nothing
End Sub

Public Sub fn_DispatchSlot(ByVal slotIndex As Long)
    Dim hotkeyKey As String

    If g_IsShuttingDown Then Exit Sub
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

    If g_ActiveHotkeyByKey Is Nothing Then
        Set g_ActiveHotkeyByKey = VBA.CreateObject("Scripting.Dictionary")
        g_ActiveHotkeyByKey.CompareMode = 1
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

Private Function private_AssignPhysicalHotkey(ByVal hotkeyKey As String) As Boolean
    Dim entry As Object
    Dim slotIndex As Long

    private_EnsureStorage
    If VBA.Len(hotkeyKey) = 0 Then Exit Function
    If g_ActiveHotkeyByKey.Exists(hotkeyKey) Then
        private_AssignPhysicalHotkey = True
        Exit Function
    End If
    If Not g_EntryByHotkey.Exists(hotkeyKey) Then Exit Function

    Set entry = g_EntryByHotkey(hotkeyKey)
    slotIndex = VBA.CLng(entry("SlotIndex"))
    If Not private_AssignOnKey(hotkeyKey, private_BuildSlotMacroRef(slotIndex)) Then Exit Function

    g_ActiveHotkeyByKey(hotkeyKey) = True
    private_AssignPhysicalHotkey = True
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

Private Sub private_UnassignActivePhysicalHotkeys()
    Dim hotkeyKey As Variant

    If g_ActiveHotkeyByKey Is Nothing Then Exit Sub

    For Each hotkeyKey In g_ActiveHotkeyByKey.Keys
        On Error Resume Next
        Application.OnKey VBA.CStr(hotkeyKey)
        On Error GoTo 0
    Next hotkeyKey

    Set g_ActiveHotkeyByKey = VBA.CreateObject("Scripting.Dictionary")
    g_ActiveHotkeyByKey.CompareMode = 1
End Sub

Private Sub private_UnregisterHotkey(ByVal hotkeyKey As String)
    Dim entry As Object
    Dim slotIndex As Long

    If g_EntryByHotkey Is Nothing Then Exit Sub
    If VBA.Len(hotkeyKey) = 0 Then Exit Sub

    If Not g_ActiveHotkeyByKey Is Nothing Then
        If g_ActiveHotkeyByKey.Exists(hotkeyKey) Then
            On Error Resume Next
            Application.OnKey hotkeyKey
            On Error GoTo 0
            g_ActiveHotkeyByKey.Remove hotkeyKey
        End If
    End If

    If g_EntryByHotkey.Exists(hotkeyKey) Then
        Set entry = g_EntryByHotkey(hotkeyKey)
        slotIndex = VBA.CLng(entry("SlotIndex"))
        If Not g_HotkeyBySlot Is Nothing Then
            If g_HotkeyBySlot.Exists(VBA.CStr(slotIndex)) Then g_HotkeyBySlot.Remove VBA.CStr(slotIndex)
        End If
        g_EntryByHotkey.Remove hotkeyKey
    End If
End Sub

Private Sub private_RemoveEmptyHotkeyEntries()
    Dim hotkeyKey As Variant
    Dim keysToRemove As Collection
    Dim entry As Object
    Dim pages As Object
    Dim removeKey As Variant

    If g_EntryByHotkey Is Nothing Then Exit Sub

    Set keysToRemove = New Collection
    For Each hotkeyKey In g_EntryByHotkey.Keys
        Set entry = g_EntryByHotkey(hotkeyKey)
        Set pages = entry("Pages")
        If pages Is Nothing Then
            keysToRemove.Add VBA.CStr(hotkeyKey)
        ElseIf pages.Count = 0 Then
            keysToRemove.Add VBA.CStr(hotkeyKey)
        End If
    Next hotkeyKey

    For Each removeKey In keysToRemove
        private_UnregisterHotkey VBA.CStr(removeKey)
    Next removeKey
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
