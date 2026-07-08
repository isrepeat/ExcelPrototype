VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PersonalCardSqlRowPcsr"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_ISqlRowProcessor

Private Const OWNER_ALIAS_FIO As String = "FIO"
Private Const DOC_NOTE_ALIAS As String = "DocNote"
Private Const TVO_MARKER As String = "ТВО"

Private Const IDENTITY_KEY_POSITION As String = "POSITION"
Private Const IDENTITY_KEY_ROW_KIND As String = "ROW_KIND"
Private Const IDENTITY_KEY_IS_META As String = "IS_META_ROW"
Private Const IDENTITY_KEY_IS_VIRTUAL As String = "IS_VIRTUAL_ROW"
Private Const IDENTITY_KEY_OWNER_ACTIVE As String = "OWNER_ACTIVE"

Private Const IDENTITY_VALUE_POS_BEFORE As String = "Before"
Private Const IDENTITY_VALUE_POS_AFTER As String = "After"
Private Const IDENTITY_VALUE_ROW_KIND_OWNER As String = "Owner"
Private Const IDENTITY_VALUE_ROW_KIND_META As String = "Meta"
Private Const IDENTITY_VALUE_ROW_KIND_COMMON As String = "Common"

Private m_InputTable As obj_TableDynamic
Private m_SqlMetaTableScheme As obj_SqlMetaTableScheme
Private m_OwnerColIndex As Long
Private m_DocNoteColIndex As Long
Private m_CommonKey As String
Private m_HasCommonKey As Boolean
Private m_IsOwnerActive As Boolean
Private m_IsInitialized As Boolean
Private m_IsDisposed As Boolean

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
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_ISqlRowProcessor_Initialize( _
    ByVal inputTable As obj_TableDynamic, _
    ByVal sqlParams As obj_SqlParams _
) As Boolean
    obj_ISqlRowProcessor_Initialize = Me.Initialize(inputTable, sqlParams)
End Function

Private Function obj_ISqlRowProcessor_HandleRow( _
    ByVal row As obj_Row _
) As Boolean
    obj_ISqlRowProcessor_HandleRow = Me.HandleRow(row)
End Function

Private Function obj_ISqlRowProcessor_BuildResult() As obj_TableDynamic
    Set obj_ISqlRowProcessor_BuildResult = Me.BuildResult()
End Function

Private Sub obj_ISqlRowProcessor_Dispose()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal inputTable As obj_TableDynamic, _
    ByVal sqlParams As obj_SqlParams _
) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Call private_ResetProcessingState
    If inputTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.Initialize: input table is Nothing."
#End If
        Exit Function
    End If
    If sqlParams Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.Initialize: sql params is Nothing."
#End If
        Exit Function
    End If

    Set m_InputTable = inputTable
    Set m_SqlMetaTableScheme = New obj_SqlMetaTableScheme
    If m_SqlMetaTableScheme Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.Initialize: failed to create SqlMetaTableScheme."
#End If
        Exit Function
    End If

    Call m_InputTable.TryGetColumnIndexByAlias(OWNER_ALIAS_FIO, m_OwnerColIndex)
    Call m_InputTable.TryGetColumnIndexByAlias(DOC_NOTE_ALIAS, m_DocNoteColIndex)
    If m_OwnerColIndex <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.Initialize: owner alias '" & OWNER_ALIAS_FIO & "' is not available in input table."
#End If
        Exit Function
    End If
    If Not ex_HelpersSql.fn_TryExtractWhereEqualsValue(sqlParams.WhereConditions, m_CommonKey) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.Initialize: failed to extract CommonKey from WhereConditions='" & sqlParams.WhereConditions & "'."
#End If
        Exit Function
    End If
    m_CommonKey = VBA.Trim$(m_CommonKey)
    m_HasCommonKey = (VBA.Len(m_CommonKey) > 0)
    If Not m_HasCommonKey Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.Initialize: extracted CommonKey is empty. WhereConditions='" & sqlParams.WhereConditions & "'."
#End If
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardSqlRowPcsr.Initialize: columns=" & VBA.CStr(m_InputTable.ColumnCount) & "; ownerAlias='" & OWNER_ALIAS_FIO & "' index=" & VBA.CStr(m_OwnerColIndex) & "; docNoteAlias='" & DOC_NOTE_ALIAS & "' index=" & VBA.CStr(m_DocNoteColIndex) & "; commonKey='" & m_CommonKey & "'"
#End If

    m_IsInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    Call private_ResetProcessingState
End Sub

Public Function HandleRow( _
    ByVal row As obj_Row _
) As Boolean
    Dim ownerValue As String
    Dim docNoteValue As String
    Dim isOwnerFilled As Boolean
    Dim isTargetOwner As Boolean
    Dim wasOwnerActive As Boolean
    Dim isMetaRow As Boolean
    Dim rowKind As String
    Dim positionValue As String
    Dim rowIdentity As obj_ItemIdentity
    Dim rowIndexForLog As Long

    On Error GoTo EH_HANDLE_ROW

    If Not m_IsInitialized Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: processor is not initialized."
#End If
        Exit Function
    End If
    If row Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: row is Nothing."
#End If
        Exit Function
    End If
    If m_InputTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: input table is Nothing."
#End If
        Exit Function
    End If
    If m_SqlMetaTableScheme Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: SqlMetaTableScheme is Nothing."
#End If
        Exit Function
    End If
    If m_OwnerColIndex <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr: owner alias '" & OWNER_ALIAS_FIO & "' is not available in table."
#End If
        Exit Function
    End If
    If Not m_HasCommonKey Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: CommonKey is not initialized."
#End If
        Exit Function
    End If

    ownerValue = VBA.Trim$(row.GetCellValue(m_OwnerColIndex))
    isOwnerFilled = (VBA.Len(ownerValue) > 0)
    wasOwnerActive = m_IsOwnerActive
    If isOwnerFilled Then
        isTargetOwner = ex_Helpers.fn_IsStringEquals(ownerValue, m_CommonKey, VBA.vbTextCompare)
    End If

    ' META-строка определяется по маркеру в DocNote.
    If m_DocNoteColIndex > 0 Then
        docNoteValue = row.GetCellValue(m_DocNoteColIndex)
        isMetaRow = (VBA.InStr(1, docNoteValue, TVO_MARKER, VBA.vbTextCompare) > 0)
    End If

    If isOwnerFilled Then
        If isTargetOwner Then
            ' При кастомном processor SQL читает полный диапазон без WHERE.
            ' Поэтому здесь вручную восстанавливаем фильтр CommonKey:
            ' активной становится только owner-строка с нужным ФИО.
            m_IsOwnerActive = True
        ElseIf isMetaRow And wasOwnerActive Then
            ' В DailyEvents meta-строка может иметь свой заполненный ПІБ
            ' другого человека. Если перед ней уже найден target owner,
            ' не считаем такую строку новым owner, а сохраняем как meta.
            m_IsOwnerActive = True
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "PersonalCardSqlRowPcsr.HandleRow: keeping non-target owner row as meta rowIndex=" & VBA.CStr(row.Index) & "; owner='" & ownerValue & "'; marker='" & TVO_MARKER & "'"
#End If
        Else
            m_IsOwnerActive = False
            HandleRow = True
            Exit Function
        End If
    End If

    If isMetaRow Then
        ' TODO: здесь останется только определение группы/связи meta-строки.
        If m_IsOwnerActive Then
            positionValue = IDENTITY_VALUE_POS_AFTER
        Else
            positionValue = IDENTITY_VALUE_POS_BEFORE
        End If
    End If

    If Not isOwnerFilled Then
        ' Пока полноценная pending/before-группировка не переписана, строки без
        ' owner выводим только если это meta-строки активного target owner.
        If Not (isMetaRow And m_IsOwnerActive) Then
            HandleRow = True
            Exit Function
        End If
    End If

    If isTargetOwner And isMetaRow Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogWarning "PersonalCardSqlRowPcsr.HandleRow: target owner row contains meta marker; row will be kept only as owner row. rowIndex=" & VBA.CStr(row.Index) & "; owner='" & ownerValue & "'"
#End If
        isMetaRow = False
    End If

    rowKind = IDENTITY_VALUE_ROW_KIND_COMMON
    If isTargetOwner Then
        rowKind = IDENTITY_VALUE_ROW_KIND_OWNER
    ElseIf isMetaRow Then
        rowKind = IDENTITY_VALUE_ROW_KIND_META
    End If
    If VBA.Len(positionValue) = 0 Then positionValue = IDENTITY_VALUE_POS_AFTER

    Set rowIdentity = New obj_ItemIdentity
    If rowIdentity Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: failed to create ItemIdentity rowIndex=" & VBA.CStr(row.Index)
#End If
        Exit Function
    End If
    If Not rowIdentity.Add(IDENTITY_KEY_POSITION, positionValue) Then GoTo IdentityFail
    If Not rowIdentity.Add(IDENTITY_KEY_ROW_KIND, rowKind) Then GoTo IdentityFail
    If Not rowIdentity.Add(IDENTITY_KEY_IS_META, isMetaRow) Then GoTo IdentityFail
    If Not rowIdentity.Add(IDENTITY_KEY_IS_VIRTUAL, False) Then GoTo IdentityFail
    If Not rowIdentity.Add(IDENTITY_KEY_OWNER_ACTIVE, m_IsOwnerActive) Then GoTo IdentityFail

    If Not m_SqlMetaTableScheme.AddRow(row, rowIdentity) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: failed to add row to SqlMetaTableScheme rowIndex=" & VBA.CStr(row.Index) & "; rowKind='" & rowKind & "'; position='" & positionValue & "'; isMeta=" & VBA.CStr(isMetaRow)
#End If
        Exit Function
    End If

    HandleRow = True
    Exit Function

IdentityFail:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.HandleRow: failed to add identity rowIndex=" & VBA.CStr(row.Index) & "; rowKind='" & rowKind & "'; position='" & positionValue & "'; isMeta=" & VBA.CStr(isMetaRow) & "; ownerActive=" & VBA.CStr(m_IsOwnerActive)
#End If
    Exit Function

EH_HANDLE_ROW:
#If LOGGING_DEBUG_ENABLED Then
    rowIndexForLog = 0
    If Not row Is Nothing Then rowIndexForLog = row.Index
    ex_Core.fn_Diagnostic_LogError _
        "PersonalCardSqlRowPcsr.HandleRow: exception rowIndex=" & _
        VBA.CStr(rowIndexForLog) & " err=[" & _
        VBA.CStr(Err.Number) & "] " & Err.Description
#End If
End Function

Public Function BuildResult() As obj_TableDynamic
    Dim resultTable As obj_TableDynamic
    Dim sqlMetaRowItems As list__obj_SqlMetaRowItem
    Dim sqlMetaRowItem As obj_SqlMetaRowItem
    Dim sourceColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim copiedRow As obj_Row
    Dim rowIdentity As obj_ItemIdentity
    Dim isVirtualRow As Boolean
    Dim i As Long

    On Error GoTo EH_BUILD_RESULT

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardSqlRowPcsr.BuildResult: start"
#End If

    If Not m_IsInitialized Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: processor is not initialized."
#End If
        Exit Function
    End If
    If m_InputTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: input table is Nothing."
#End If
        Exit Function
    End If
    If m_SqlMetaTableScheme Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: SqlMetaTableScheme is Nothing."
#End If
        Exit Function
    End If

    Set resultTable = New obj_TableDynamic
    resultTable.SectionTitle = m_InputTable.SectionTitle

    For i = 1 To m_InputTable.Columns.Count
        Set sourceColumn = m_InputTable.Columns.Item(i)
        If sourceColumn Is Nothing Then GoTo ContinueColumn
        If Not resultTable.PushColumn(sourceColumn) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: failed to push column index=" & VBA.CStr(i) & " name='" & sourceColumn.Name & "'"
#End If
            Exit Function
        End If
ContinueColumn:
    Next i

    Set sqlMetaRowItems = m_SqlMetaTableScheme.SqlMetaRowItems
    If Not sqlMetaRowItems Is Nothing Then
        For i = 1 To sqlMetaRowItems.Count
            Set sqlMetaRowItem = sqlMetaRowItems.Item(i)
            If sqlMetaRowItem Is Nothing Then GoTo ContinueRow

            Set sourceRow = sqlMetaRowItem.SourceRow
            If sourceRow Is Nothing Then
                Set rowIdentity = sqlMetaRowItem.ItemIdentity
                isVirtualRow = False
                If Not rowIdentity Is Nothing Then Call rowIdentity.TryGetBoolean(IDENTITY_KEY_IS_VIRTUAL, isVirtualRow)
                If isVirtualRow Then
                    Set copiedRow = private_CreateVirtualRow(resultTable.ColumnCount)
                    If copiedRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
                        ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: failed to create virtual row at item index=" & VBA.CStr(i)
#End If
                        Exit Function
                    End If
                Else
                    GoTo ContinueRow
                End If
            Else
                Set copiedRow = sourceRow.Clone(resultTable.ColumnCount)
                If copiedRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
                    ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: failed to clone source row index=" & VBA.CStr(sourceRow.Index)
#End If
                    Exit Function
                End If
            End If

            If Not resultTable.PushRow(copiedRow) Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: failed to push result row sourceIndex=" & VBA.CStr(copiedRow.Index)
#End If
                Exit Function
            End If
ContinueRow:
        Next i
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardSqlRowPcsr.BuildResult: result rows=" & VBA.CStr(resultTable.RowCount) & "; columns=" & VBA.CStr(resultTable.ColumnCount) & "; metaItems=" & VBA.CStr(m_SqlMetaTableScheme.Count)
#End If

    Set BuildResult = resultTable
    Exit Function

EH_BUILD_RESULT:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PersonalCardSqlRowPcsr.BuildResult: exception itemIndex=" & VBA.CStr(i) & " err=[" & VBA.CStr(Err.Number) & "] " & Err.Description
#End If
End Function

' //
' // Internal
' //
Private Sub private_ResetProcessingState()
    m_IsInitialized = False
    Set m_InputTable = Nothing
    Set m_SqlMetaTableScheme = Nothing
    m_OwnerColIndex = 0
    m_DocNoteColIndex = 0
    m_CommonKey = VBA.vbNullString
    m_HasCommonKey = False
    m_IsOwnerActive = False
End Sub

Private Function private_CreateVirtualRow(ByVal columnCount As Long) As obj_Row
    Dim result As obj_Row
    Dim i As Long

    If columnCount <= 0 Then Exit Function

    Set result = New obj_Row
    If result Is Nothing Then Exit Function
    result.Desc = "__virtual_row"

    For i = 1 To columnCount
        result.PushCellRaw VBA.vbNullString
    Next i

    Set private_CreateVirtualRow = result
End Function
