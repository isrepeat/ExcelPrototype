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

Private Const FLAG_OUTPUT_MODE As String = "Display.OutputMode"

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
    On Error GoTo NoProp
    m_GetOutputMode = CLng(ThisWorkbook.CustomDocumentProperties(FLAG_OUTPUT_MODE).Value)
    Exit Function
NoProp:
    ' По умолчанию - Timeline
    m_SetOutputMode PersonTimeline
    m_GetOutputMode = PersonTimeline
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
    mode = m_GetOutputMode()
    
    Select Case mode
        Case PersonTimeline
            m_GetOutputModeString = "PersonTimeline"
        Case StateTableOnly
            m_GetOutputModeString = "StateTableOnly"
        Case EventsTableOnly
            m_GetOutputModeString = "EventsTableOnly"
        Case Else
            m_GetOutputModeString = "Unknown"
    End Select
End Function

Public Function m_GetOutputModeDisplay() As String
    Dim mode As OutputMode
    mode = m_GetOutputMode()
    
    Select Case mode
        Case PersonTimeline
            m_GetOutputModeDisplay = "Timeline"
        Case StateTableOnly
            m_GetOutputModeDisplay = "State"
        Case EventsTableOnly
            m_GetOutputModeDisplay = "Events"
        Case Else
            m_GetOutputModeDisplay = ""
    End Select
End Function

' =============================================================================
' Утилиты для переключения режимов
' =============================================================================

Public Sub m_CycleOutputMode()
    ' Циклически переключает режимы: Timeline -> State -> Events -> Timeline
    Dim currentMode As OutputMode
    Dim nextMode As OutputMode
    
    currentMode = m_GetOutputMode()
    
    Select Case currentMode
        Case PersonTimeline
            nextMode = StateTableOnly
        Case StateTableOnly
            nextMode = EventsTableOnly
        Case EventsTableOnly
            nextMode = PersonTimeline
        Case Else
            nextMode = PersonTimeline
    End Select
    
    m_SetOutputMode nextMode
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

    ' Обновляет текст и цвет кнопки с начальным текстом "ModeButton"
    ' Текст обновляется на "ModeButton [Timeline]", "ModeButton [State]" или "ModeButton [Events]"
    
    On Error GoTo EH
    
    Dim currentMode As OutputMode
    currentMode = m_GetOutputMode()
    
    ' Получаем фигуру с листа Dev
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dev")
    
    Dim btn As Shape
    Dim i As Integer
    Dim btnFound As Boolean
    btnFound = False
    
    ' Ищем фигуру, текст которой начинается с "ModeButton"
    For i = 1 To ws.Shapes.Count
        On Error Resume Next
        If ws.Shapes(i).HasTextFrame Then
            If InStr(1, ws.Shapes(i).TextFrame.Characters.Text, "ModeButton", vbTextCompare) = 1 Then
                Set btn = ws.Shapes(i)
                btnFound = True
                On Error GoTo EH
                Exit For
            End If
        End If
        On Error GoTo EH
    Next i
    
    If Not btnFound Then
        Call ex_Messaging.m_ShowNotice("Button with 'ModeButton' text not found on Dev sheet", 3)
        Exit Sub
    End If
    
    ' Константы цветов
    Const COLOR_TIMELINE As Long = &H0070C0     ' Синий
    Const COLOR_STATE As Long = &H70AD47        ' Зелёный
    Const COLOR_EVENTS As Long = &HC55A11       ' Оранжевый
    Const COLOR_TEXT As Long = &HFFFFFF         ' Белый текст
    
    Dim btnText As String
    Dim btnColor As Long
    
    ' Определяем текст и цвет в зависимости от режима
    Select Case currentMode
        Case PersonTimeline
            btnText = "ModeButton [Timeline]"
            btnColor = COLOR_TIMELINE
            
        Case StateTableOnly
            btnText = "ModeButton [State]"
            btnColor = COLOR_STATE
            
        Case EventsTableOnly
            btnText = "ModeButton [Events]"
            btnColor = COLOR_EVENTS
            
        Case Else
            btnText = "ModeButton [Unknown]"
            btnColor = &H808080
    End Select
    
    ' Обновляем текст кнопки - просто присваиваем текст
    btn.TextFrame.Characters.Text = btnText
    
    ' Форматируем текст
    With btn.TextFrame.Characters.Font
        .Size = 14
        .Bold = True
        .Color = COLOR_TEXT
    End With
    
    ' Центрируем текст
    btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    btn.TextFrame.VerticalAlignment = xlVAlignCenter
    
    ' Обновляем цвет кнопки
    With btn.Fill
        .Solid
        .ForeColor.RGB = btnColor
    End With
    
    Exit Sub
    
EH:
    Call ex_Messaging.m_ShowNotice("Error updating button: " & Err.Description, 3)
End Sub
