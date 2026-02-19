Attribute VB_Name = "ex_Settings"
Option Explicit

' =============================================================================
' Enum для режимов вывода данных
' =============================================================================

Public Enum OutputMode
    PersonTimeline = 1     ' Персональная карта с временной шкалой
    StateTableOnly = 2     ' Только таблица состояния
    EventsTableOnly = 3    ' Только таблица событий
End Enum

' =============================================================================
' Константы флагов
' =============================================================================

Private Const FLAG_OUTPUT_MODE As String = "Settings.OutputMode"

' =============================================================================
' Public API: Булевы флаги
' =============================================================================

Public Function m_GetBoolFlag(ByVal flagName As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo NoProp
    m_GetBoolFlag = CBool(ThisWorkbook.CustomDocumentProperties(flagName).Value)
    Exit Function
NoProp:
    m_GetBoolFlag = defaultValue
End Function

Public Sub m_SetBoolFlag(ByVal flagName As String, ByVal valueBool As Boolean)
    On Error GoTo AddProp
    ThisWorkbook.CustomDocumentProperties(flagName).Value = valueBool
    Exit Sub
AddProp:
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=flagName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=valueBool
End Sub

' =============================================================================
' Public API: Enum флаги (режим вывода)
' =============================================================================

Public Function m_GetOutputMode() As OutputMode
    Dim defaultModeValue As Long

    On Error GoTo NoProp
    m_GetOutputMode = CLng(ThisWorkbook.CustomDocumentProperties(FLAG_OUTPUT_MODE).Value)
    Exit Function
NoProp:
    defaultModeValue = mp_GetDefaultModeValue()
    If defaultModeValue <= 0 Then
        MsgBox "Mode variants for control 'btnMode' are not configured in UI.xml.", vbExclamation
        m_GetOutputMode = 0
        Exit Function
    End If

    m_SetOutputMode CLng(defaultModeValue)
    m_GetOutputMode = CLng(defaultModeValue)
End Function

Public Sub m_SetOutputMode(ByVal mode As OutputMode)
    On Error GoTo AddProp
    ThisWorkbook.CustomDocumentProperties(FLAG_OUTPUT_MODE).Value = CLng(mode)
    Exit Sub
AddProp:
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=FLAG_OUTPUT_MODE, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeNumber, _
        Value:=CLng(mode)
End Sub

Public Function m_GetOutputModeString() As String
    Dim mode As OutputMode
    Dim variants As Variant
    Dim variantRow As Long
    Dim displayText As String

    mode = m_GetOutputMode()

    variants = mp_GetModeVariants()
    If Not mp_ArrayHasItems(variants) Then
        m_GetOutputModeString = "Unknown"
        Exit Function
    End If

    variantRow = mp_FindVariantRowByValue(variants, CLng(mode))
    If variantRow > 0 Then
        displayText = CStr(variants(variantRow, 4))
        If Len(displayText) > 0 Then
            m_GetOutputModeString = displayText
            Exit Function
        End If
    End If

    m_GetOutputModeString = "Mode" & CStr(CLng(mode))
End Function

Public Function m_GetOutputModeDisplay() As String
    Dim mode As OutputMode
    Dim variants As Variant
    Dim variantRow As Long
    Dim displayText As String

    mode = m_GetOutputMode()
    variants = mp_GetModeVariants()
    If Not mp_ArrayHasItems(variants) Then Exit Function

    variantRow = mp_FindVariantRowByValue(variants, CLng(mode))
    If variantRow <= 0 Then Exit Function

    displayText = CStr(variants(variantRow, 4))
    If Len(displayText) > 0 Then
        m_GetOutputModeDisplay = displayText
    Else
        m_GetOutputModeDisplay = CStr(variants(variantRow, 2))
    End If
End Function

' =============================================================================
' Утилиты для переключения режимов
' =============================================================================

Public Sub m_CycleOutputMode()
    Dim currentModeValue As Long
    Dim nextModeValue As Long
    Dim variants As Variant

    variants = mp_GetModeVariants()
    If Not mp_ArrayHasItems(variants) Then
        MsgBox "Mode variants for control 'btnMode' are not configured in UI.xml.", vbExclamation
        Exit Sub
    End If

    currentModeValue = CLng(m_GetOutputMode())
    nextModeValue = mp_GetNextVariantValue(variants, currentModeValue)
    If nextModeValue <= 0 Then
        MsgBox "Failed to resolve next mode value from UI.xml modeVariants.", vbExclamation
        Exit Sub
    End If

    m_SetOutputMode CLng(nextModeValue)
    Call ex_Messaging.m_ShowNotice("Mode changed to: " & m_GetOutputModeDisplay(), 3)
End Sub

Public Sub m_SetOutputModeByString(ByVal modeStr As String)
    Dim mode As OutputMode
    
    Select Case LCase(Trim(modeStr))
        Case "persontimeline"
            mode = PersonTimeline
        Case "statetableonly"
            mode = StateTableOnly
        Case "eventstableonly"
            mode = EventsTableOnly
        Case Else
            Call ex_Messaging.m_ShowNotice("Unknown mode: " & modeStr, 3)
            Exit Sub
    End Select
    
    m_SetOutputMode mode
End Sub

' =============================================================================
' Макрос для одной кнопки переключения режимов
' =============================================================================

Public Sub m_SwitchMode_OnClick()
    ' Одна кнопка для переключения между тремя режимами
    ' При каждом клике - переходит на следующий режим
    m_CycleOutputMode
    m_UpdateModeButton
End Sub

' =============================================================================
' Обновление визуального состояния кнопки
' =============================================================================

Public Sub m_UpdateModeButton()
    On Error GoTo EH

    Dim currentMode As OutputMode
    Dim ws As Worksheet
    Dim btn As Shape
    Dim variants As Variant
    Dim variantRow As Long
    Dim captionText As String
    Dim styleName As String

    currentMode = m_GetOutputMode()
    variants = mp_GetModeVariants()
    If Not mp_ArrayHasItems(variants) Then
        MsgBox "Mode variants for control 'btnMode' are not configured in UI.xml.", vbExclamation
        Exit Sub
    End If

    Set ws = ws_Dev
    Set btn = ex_ConfigProfilesManager.m_GetShapeByName(ws, "btnMode")
    If btn Is Nothing Then
        Call ex_Messaging.m_ShowNotice("Button 'btnMode' not found on Dev sheet", 3)
        Exit Sub
    End If

    variantRow = mp_FindVariantRowByValue(variants, CLng(currentMode))
    If variantRow <= 0 Then
        MsgBox "Current mode value '" & CStr(CLng(currentMode)) & "' is not present in UI.xml modeVariants.", vbExclamation
        Exit Sub
    End If

    captionText = CStr(variants(variantRow, 2))
    If Len(captionText) = 0 Then
        captionText = ex_UiXmlProvider.m_GetControlAttribute("btnMode", "caption", ThisWorkbook)
    End If

    If Len(captionText) > 0 Then
        btn.TextFrame.Characters.Text = captionText
    End If

    styleName = CStr(variants(variantRow, 3))
    If Len(styleName) = 0 Then
        styleName = ex_UiXmlProvider.m_GetControlAttribute("btnMode", "style", ThisWorkbook)
    End If
    If Len(styleName) > 0 Then
        If Not ex_UiXmlProvider.m_ApplyControlStyleByName(ws, "btnMode", styleName, ThisWorkbook) Then
            Exit Sub
        End If
    End If

    Exit Sub
EH:
    Call ex_Messaging.m_ShowNotice("Error updating mode button: " & Err.Description, 3)
End Sub

Private Function mp_GetModeVariants() As Variant
    mp_GetModeVariants = ex_UiXmlProvider.m_GetModeVariantsByControl("btnMode", ThisWorkbook)
End Function

Private Function mp_GetDefaultModeValue() As Long
    Dim variants As Variant

    variants = mp_GetModeVariants()
    If Not mp_ArrayHasItems(variants) Then Exit Function
    mp_GetDefaultModeValue = CLng(variants(LBound(variants, 1), 1))
End Function

Private Function mp_ArrayHasItems(ByVal values As Variant) As Boolean
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayHasItems = (UBound(values, 1) >= LBound(values, 1))
    End If
    Exit Function
EH:
    mp_ArrayHasItems = False
End Function

Private Function mp_FindVariantRowByValue(ByVal variants As Variant, ByVal modeValue As Long) As Long
    Dim i As Long

    If Not mp_ArrayHasItems(variants) Then Exit Function

    For i = LBound(variants, 1) To UBound(variants, 1)
        If CLng(variants(i, 1)) = modeValue Then
            mp_FindVariantRowByValue = i
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetNextVariantValue(ByVal variants As Variant, ByVal currentModeValue As Long) As Long
    Dim currentRow As Long

    If Not mp_ArrayHasItems(variants) Then Exit Function

    currentRow = mp_FindVariantRowByValue(variants, currentModeValue)
    If currentRow = 0 Then
        Exit Function
    End If

    If currentRow >= UBound(variants, 1) Then
        mp_GetNextVariantValue = CLng(variants(LBound(variants, 1), 1))
    Else
        mp_GetNextVariantValue = CLng(variants(currentRow + 1, 1))
    End If
End Function
