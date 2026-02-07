Attribute VB_Name = "ex_Run"
Option Explicit

Public Sub m_Run()
    Dim resultTable As Variant
    Dim keyColumnName As String

    On Error GoTo EH

    keyColumnName = mp_GetKeyColumnFromConfig()
    If Len(keyColumnName) = 0 Then
        keyColumnName = "Id"
    End If

    ' 1) Загружаем внешние таблицы в g_Old / g_New
    ex_SourceLoader.m_LoadOldNewFromConfigToInternalSheets

    ' 2) Сравниваем как раньше (движок тот же)
    resultTable = ex_TableComparer.m_CompareSheets( _
        "Old", _
        "New", _
        keyColumnName _
    )

    ' 3) Пишем результат как раньше
    ex_ResultWriter.m_WriteTableToResultSheet resultTable
    Exit Sub

EH:
    MsgBox "Run error: " & Err.Description, vbExclamation
End Sub

Public Sub m_Clear()
    ex_SheetManager.m_DeleteResultSheets
End Sub

' -----------------------------------------------------------------------------
' СТАРАЯ ТЕСТОВАЯ ЛОГИКА (закомментирована по твоей просьбе)
' -----------------------------------------------------------------------------
'Public Sub TestCompareOldNew()
'    Dim resultTable As Variant
'
'    ex_TestData.GenerateTestTables
'
'    resultTable = ex_TableComparer.CompareSheets( _
'        "Old", _
'        "New", _
'        "Id" _
'    )
'
'    ex_ResultWriter.WriteTableToResultSheet resultTable
'End Sub

Private Function mp_GetKeyColumnFromConfig() As String
    On Error GoTo Fallback

    mp_GetKeyColumnFromConfig = ex_Config.m_GetConfigValue("KeyColumnName", "Id")
    Exit Function

Fallback:
    mp_GetKeyColumnFromConfig = "Id"
End Function
