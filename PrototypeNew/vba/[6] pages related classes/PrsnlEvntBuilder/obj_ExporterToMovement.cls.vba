VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ExporterToMovement"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IDataExporter

Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_Data As obj_PrsnlEvntBuilderData

Private Const SAVE_ALREADY_OPEN_WORKBOOK As Boolean = False
Private Const MOVEMENT_TARGET_COLUMN_COUNT As Long = 6
Private Const MOVEMENT_TRAILING_EMPTY_LOOKBACK_ROWS As Long = 20
Private Const MOVEMENT_SOURCE_INCOMING_NO As String = "Вх. №"
Private Const MOVEMENT_CONTEXT_MANUAL_ORDER_NO As String = "ManualOrderNo"
Private Const MOVEMENT_CONTEXT_SECTION_TYPE As String = "SectionType"
Private Const MOVEMENT_SOURCE_EVENT As String = "Подія"
Private Const MOVEMENT_SOURCE_INCOMING_DATE As String = "Вх. дата"
Private Const MOVEMENT_SOURCE_DEPARTURE_DATE As String = "З"
Private Const MOVEMENT_SOURCE_DURATION_TERM As String = "Термін вибуття"
Private Const MOVEMENT_SOURCE_DURATION_DAYS As String = "На скільки"
Private Const MOVEMENT_SOURCE_ESCORT_DOC As String = "Супровідний документ"
Private Const MOVEMENT_SOURCE_VK_NO As String = "В/к №"
Private Const MOVEMENT_TARGET_ORDER_NO As String = "Наказ вибуття"
Private Const MOVEMENT_TARGET_FOOD_FROM As String = "З продовольчого"
Private Const MOVEMENT_TARGET_DEPARTURE As String = "Вибуття"
Private Const MOVEMENT_TARGET_ARRIVAL_ORDER_NO As String = "Наказ прибуття"
Private Const MOVEMENT_TARGET_ON_FOOD As String = "На продовольче"
Private Const MOVEMENT_TARGET_ARRIVAL As String = "Прибуття"
Private Const MOVEMENT_TARGET_IPN As String = "ІПН"
Private Const MOVEMENT_TARGET_DURATION_DAYS As String = "На скільки"
Private Const MOVEMENT_TARGET_VK_NO As String = "В/к №"
Private Const MOVEMENT_TARGET_EVENT As String = "Подія"
Private Const MOVEMENT_ORDERS_SHEET_NAME As String = "Накази"
Private Const MOVEMENT_ORDER_DATE_TABLE_NAME As String = "tbMapNumberOrderToDate"
Private Const MOVEMENT_ORDER_DATE_COLUMN_NAME As String = "Дата наказу"
Private Const MOVEMENT_ORDER_NO_COLUMN_NAME As String = "Номер наказу"

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IDataExporter_Export( _
    ByVal sourceTables As Collection, _
    Optional ByVal context As Object = Nothing _
) As Boolean
    obj_IDataExporter_Export = Me.Export(sourceTables, context)
End Function

' //
' // API
' //
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    private_LogMethodEntry "Initialize"

    m_IsDisposed = False
    Set m_Base = New obj_DataExporterBase
    Set m_Data = New obj_PrsnlEvntBuilderData

    If Not m_Base.Initialize(configTable, "Movement", "PrototypeNew / Movement export") Then Exit Function
    private_LogInfo "movement:init workbook='" & private_EscapeForLog(m_Base.TargetWorkbookPath) & _
        "' sheet='" & private_EscapeForLog(m_Base.TargetSheetName) & _
        "' start='" & private_EscapeForLog(m_Base.TargetRangeStartMarker) & _
        "' end='" & private_EscapeForLog(m_Base.TargetRangeEndMarker) & "'"

    Initialize = True
End Function

Public Function TryGetSectionTypeOptions(ByRef outSectionTypeOptions As Collection) As Boolean
    private_LogMethodEntry "TryGetSectionTypeOptions"
    Set outSectionTypeOptions = m_Data.SectionTypeNames
    If outSectionTypeOptions Is Nothing Then Exit Function
    TryGetSectionTypeOptions = (outSectionTypeOptions.Count > 0)
End Function

Public Sub Dispose()
    private_LogMethodEntry "Dispose"
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Base Is Nothing Then m_Base.Dispose
    Set m_Base = Nothing
    Set m_Data = Nothing
    On Error GoTo 0
End Sub

Public Function Export( _
    ByVal sourceTables As Collection, _
    Optional ByVal context As Object = Nothing _
) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetTable As ListObject
    Dim insertedRow As ListRow
    Dim targetRowRange As Range
    Dim targetValues As Variant
    Dim openedByExporter As Boolean
    Dim fastModeStarted As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim outgoingOrderNo As Variant
    Dim outgoingFoodFromDate As Variant
    Dim outgoingDepartureDate As Variant
    Dim closingOrderNo As Variant
    Dim closingOnFoodDate As Variant
    Dim closingArrivalDate As Variant
    Dim mirrorOpeningOrderNo As Variant
    Dim mirrorOpeningFoodFromDate As Variant
    Dim mirrorOpeningDepartureDate As Variant
    Dim writeSpecialOpeningFields As Boolean
    Dim specialDurationValue As Variant
    Dim specialVkNoValue As Variant
    Dim sectionTypeRaw As String
    Dim sectionTypeNormalized As String
    Dim mappedEventText As String
    Dim shouldWriteMappedEvent As Boolean
    Dim isClosingEvent As Boolean
    Dim isMirrorTransferEvent As Boolean
    Dim closingTargetIpn As String

    On Error GoTo EH
    private_LogMethodEntry "Export"
    private_LogInfo "movement:export start"

    If m_IsDisposed Then
        private_LogError "Movement exporter is disposed."
        VBA.MsgBox "PrototypeNew: Movement exporter is disposed.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function
    sectionTypeRaw = private_GetSectionTypeTextFromContext(context, sourceTable)
    sectionTypeNormalized = private_NormalizeText(sectionTypeRaw)
    isClosingEvent = private_IsClosingSectionType(sectionTypeNormalized)
    isMirrorTransferEvent = private_IsMirrorTransferSectionType(sectionTypeRaw)
    If isMirrorTransferEvent Then isClosingEvent = False
    private_LogInfo "movement:source rows=" & VBA.CStr(sourceTable.RowCount) & " cols=" & VBA.CStr(sourceTable.ColumnCount)
    private_LogInfo "movement:event-section raw='" & private_EscapeForLog(sectionTypeRaw) & "' normalized='" & private_EscapeForLog(sectionTypeNormalized) & "'"
    private_LogInfo "movement:event-mode closing=" & private_BoolText(isClosingEvent)
    private_LogInfo "movement:event-mode mirror-transfer=" & private_BoolText(isMirrorTransferEvent)

    If Not private_TryBuildSpecialOpeningValues(sourceTable, sectionTypeRaw, writeSpecialOpeningFields, specialDurationValue, specialVkNoValue) Then Exit Function
    private_LogInfo "movement:special-open-fields enabled=" & private_BoolText(writeSpecialOpeningFields) & _
        " duration='" & private_EscapeForLog(VBA.CStr(specialDurationValue)) & _
        "' vk='" & private_EscapeForLog(VBA.CStr(specialVkNoValue)) & "'"

    shouldWriteMappedEvent = private_TryMapSectionTypeToEventText(sectionTypeRaw, mappedEventText)
    private_LogInfo "movement:mapped-event enabled=" & private_BoolText(shouldWriteMappedEvent) & " value='" & private_EscapeForLog(mappedEventText) & "'"

    If Not isClosingEvent Then
        If Not private_TryBuildMovementRowValues(sourceTable, sectionTypeRaw, targetValues) Then Exit Function
        private_LogInfo "movement:row-values rank='" & private_EscapeForLog(VBA.CStr(targetValues(1, 1))) & _
            "' fio='" & private_EscapeForLog(VBA.CStr(targetValues(1, 2))) & _
            "' ipn='" & private_EscapeForLog(VBA.CStr(targetValues(1, 3))) & _
            "' position='" & private_EscapeForLog(VBA.CStr(targetValues(1, 4))) & _
            "' event='" & private_EscapeForLog(VBA.CStr(targetValues(1, 5))) & _
            "' destination='" & private_EscapeForLog(VBA.CStr(targetValues(1, 6))) & "'"
    End If

    m_Base.BeginFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    fastModeStarted = True
    private_LogInfo "movement:fast-mode enabled"

    If Not m_Base.TryOpenTargetWorkbook(targetWb, openedByExporter) Then GoTo CleanFail
    private_LogInfo "movement:workbook ready name='" & private_EscapeForLog(targetWb.Name) & "' openedByExporter=" & private_BoolText(openedByExporter)

    If Not m_Base.TryGetWorksheet(targetWb, m_Base.ResolveTargetWorksheetName(), targetWs) Then GoTo CleanFail
    private_LogInfo "movement:worksheet ready name='" & private_EscapeForLog(targetWs.Name) & "'"
    If Not m_Base.TryFindConfiguredTargetTable(targetWs, targetTable) Then GoTo CleanFail
    private_LogInfo "movement:table ready name='" & private_EscapeForLog(targetTable.Name) & "' rows=" & VBA.CStr(targetTable.ListRows.Count) & " cols=" & VBA.CStr(targetTable.ListColumns.Count)

    If isClosingEvent Then
        private_LogInfo "movement:event-branch selected='closing-only'"
        If Not private_TryBuildMovementClosingValues(sourceTable, targetWb, context, closingOrderNo, closingOnFoodDate, closingArrivalDate) Then GoTo CleanFail
        private_LogInfo "movement:closing-values orderNo='" & private_EscapeForLog(VBA.CStr(closingOrderNo)) & _
            "' arrival='" & private_EscapeForLog(private_FormatLogDateValue(closingArrivalDate)) & _
            "' onFood='" & private_EscapeForLog(private_FormatLogDateValue(closingOnFoodDate)) & "'"

        If Not private_TryGetRequiredSourceText(sourceTable, sourceTable.Rows.Item(1), MOVEMENT_TARGET_IPN, closingTargetIpn) Then GoTo CleanFail
        If Not private_TryFindLastRowByIpn(targetTable, closingTargetIpn, targetRowRange) Then GoTo CleanFail
        private_LogInfo "movement:closing-row address='" & private_EscapeForLog(targetRowRange.Address(False, False)) & "'"
        If Not private_TryWriteMovementClosingRow(targetTable, targetRowRange, closingOrderNo, closingOnFoodDate, closingArrivalDate) Then GoTo CleanFail
        private_LogInfo "movement:write-closing-row done"
    ElseIf isMirrorTransferEvent Then
        private_LogInfo "movement:event-branch selected='mirror-close-then-open'"
        If Not private_TryBuildMovementClosingValues(sourceTable, targetWb, context, closingOrderNo, closingOnFoodDate, closingArrivalDate) Then GoTo CleanFail
        private_LogInfo "movement:mirror-close-values orderNo='" & private_EscapeForLog(VBA.CStr(closingOrderNo)) & _
            "' arrival='" & private_EscapeForLog(private_FormatLogDateValue(closingArrivalDate)) & _
            "' onFood='" & private_EscapeForLog(private_FormatLogDateValue(closingOnFoodDate)) & "'"

        If Not private_TryGetRequiredSourceText(sourceTable, sourceTable.Rows.Item(1), MOVEMENT_TARGET_IPN, closingTargetIpn) Then GoTo CleanFail
        If Not private_TryFindLastRowByIpn(targetTable, closingTargetIpn, targetRowRange) Then GoTo CleanFail
        private_LogInfo "movement:mirror-close-row address='" & private_EscapeForLog(targetRowRange.Address(False, False)) & "'"
        If Not private_TryWriteMovementClosingRow(targetTable, targetRowRange, closingOrderNo, closingOnFoodDate, closingArrivalDate) Then GoTo CleanFail
        private_LogInfo "movement:mirror-write-closing-row done"

        mirrorOpeningOrderNo = closingOrderNo
        mirrorOpeningDepartureDate = closingArrivalDate
        mirrorOpeningFoodFromDate = closingOnFoodDate
        private_LogInfo "movement:mirror-open-values orderNo='" & private_EscapeForLog(VBA.CStr(mirrorOpeningOrderNo)) & _
            "' departure='" & private_EscapeForLog(private_FormatLogDateValue(mirrorOpeningDepartureDate)) & _
            "' foodFrom='" & private_EscapeForLog(private_FormatLogDateValue(mirrorOpeningFoodFromDate)) & "'"

        If Not private_TryGetAppendRowRange(targetTable, targetRowRange, insertedRow) Then GoTo CleanFail
        private_LogInfo "movement:mirror-append-row address='" & private_EscapeForLog(targetRowRange.Address(False, False)) & "'"
        If Not private_TryWriteMovementRow(targetTable, targetRowRange, targetValues, mirrorOpeningOrderNo, mirrorOpeningFoodFromDate, mirrorOpeningDepartureDate, writeSpecialOpeningFields, specialDurationValue, specialVkNoValue, shouldWriteMappedEvent, mappedEventText) Then GoTo CleanFail
        private_LogInfo "movement:mirror-write-open-row done"
    Else
        private_LogInfo "movement:event-branch selected='opening-only'"
        If Not private_TryBuildMovementOutgoingValues(sourceTable, targetWb, context, outgoingOrderNo, outgoingFoodFromDate, outgoingDepartureDate) Then GoTo CleanFail
        private_LogInfo "movement:outgoing orderNo='" & private_EscapeForLog(VBA.CStr(outgoingOrderNo)) & _
            "' departure='" & private_EscapeForLog(private_FormatLogDateValue(outgoingDepartureDate)) & _
            "' foodFrom='" & private_EscapeForLog(private_FormatLogDateValue(outgoingFoodFromDate)) & "'"
        If Not private_TryGetAppendRowRange(targetTable, targetRowRange, insertedRow) Then GoTo CleanFail
        private_LogInfo "movement:append-row address='" & private_EscapeForLog(targetRowRange.Address(False, False)) & "'"
        If Not private_TryWriteMovementRow(targetTable, targetRowRange, targetValues, outgoingOrderNo, outgoingFoodFromDate, outgoingDepartureDate, writeSpecialOpeningFields, specialDurationValue, specialVkNoValue, shouldWriteMappedEvent, mappedEventText) Then GoTo CleanFail
        private_LogInfo "movement:write-row done"
    End If

    If Not openedByExporter And SAVE_ALREADY_OPEN_WORKBOOK Then targetWb.Save
    private_LogInfo "movement:export success"
    Export = True
    GoTo CleanExit

CleanFail:
    private_LogError "Movement export failed before completion."
    Export = False
    If Not insertedRow Is Nothing Then
        On Error Resume Next
        insertedRow.Delete
        On Error GoTo 0
    End If

CleanExit:
    If openedByExporter Then
        On Error Resume Next
        targetWb.Close SaveChanges:=Export
        On Error GoTo 0
    End If
    If fastModeStarted Then m_Base.RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    Exit Function

EH:
    private_LogError "Movement export exception: [" & VBA.CStr(Err.Number) & "] " & Err.Description
    VBA.MsgBox "PrototypeNew: Movement export failed. " & Err.Description, VBA.vbExclamation, "PrototypeNew / Movement export"
    On Error Resume Next
    If openedByExporter Then targetWb.Close SaveChanges:=False
    If fastModeStarted Then m_Base.RestoreFastExcelMode prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation
    On Error GoTo 0
End Function

Private Function private_IsMirrorTransferSectionType(ByVal sectionTypeText As String) As Boolean
    private_IsMirrorTransferSectionType = m_Data.IsMovementMirrorTransferSectionType(sectionTypeText)
End Function

Private Function private_GetSectionTypeTextFromContext( _
    ByVal context As Object, _
    ByVal sourceTable As obj_TableDynamic _
) As String
    Dim sectionTypeText As String

    sectionTypeText = private_GetContextText(context, MOVEMENT_CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) > 0 Then
        private_GetSectionTypeTextFromContext = sectionTypeText
        Exit Function
    End If
    If sourceTable Is Nothing Then Exit Function
    sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionTypeText) > 0 Then
        private_GetSectionTypeTextFromContext = sectionTypeText
        Exit Function
    End If
End Function

Private Function private_GetContextText(ByVal context As Object, ByVal keyText As String) As String
    If context Is Nothing Then Exit Function
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function

    On Error Resume Next
    If context.Exists(keyText) Then private_GetContextText = VBA.Trim$(VBA.CStr(context(keyText)))
    If Err.Number <> 0 Then
        Err.Clear
        private_GetContextText = VBA.Trim$(VBA.CStr(VBA.CallByName(context, keyText, VbGet)))
    End If
    On Error GoTo 0
End Function

Private Function private_TryBuildSpecialOpeningValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sectionTypeText As String, _
    ByRef outShouldWrite As Boolean, _
    ByRef outDurationValue As Variant, _
    ByRef outVkNoValue As Variant _
) As Boolean
    Dim sourceRow As obj_Row

    outShouldWrite = False
    outDurationValue = VBA.vbNullString
    outVkNoValue = VBA.vbNullString

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    If Not private_ShouldWriteSpecialOpeningFieldsForSectionType(sectionTypeText) Then
        private_TryBuildSpecialOpeningValues = True
        Exit Function
    End If

    outShouldWrite = True
    outDurationValue = private_GetOptionalSourceTextByAnyColumn(sourceTable, sourceRow, MOVEMENT_SOURCE_DURATION_TERM, MOVEMENT_SOURCE_DURATION_DAYS)
    outVkNoValue = private_GetOptionalSourceTextByAnyColumn(sourceTable, sourceRow, MOVEMENT_SOURCE_ESCORT_DOC, MOVEMENT_SOURCE_VK_NO, "Док. №")
    private_TryBuildSpecialOpeningValues = True
End Function

Private Function private_ShouldWriteSpecialOpeningFieldsForSectionType(ByVal rawSectionType As String) As Boolean
    private_ShouldWriteSpecialOpeningFieldsForSectionType = m_Data.ShouldWriteMovementSpecialOpeningFields(rawSectionType)
End Function

Private Function private_GetOptionalSourceTextByAnyColumn( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ParamArray columnNames() As Variant _
) As Variant
    Dim columnName As Variant
    Dim valueObj As Variant

    If sourceTable Is Nothing Then Exit Function
    If sourceRow Is Nothing Then Exit Function

    For Each columnName In columnNames
        valueObj = private_GetOptionalSourceText(sourceTable, sourceRow, VBA.CStr(columnName))
        If VBA.Len(VBA.Trim$(VBA.CStr(valueObj))) > 0 Then
            private_GetOptionalSourceTextByAnyColumn = valueObj
            Exit Function
        End If
    Next columnName
End Function

Private Function private_IsClosingSectionType(ByVal normalizedSectionType As String) As Boolean
    private_IsClosingSectionType = m_Data.IsMovementClosingSectionType(normalizedSectionType)
End Function

Private Function private_TryBuildMovementClosingValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal targetWorkbook As Workbook, _
    ByVal context As Object, _
    ByRef outOrderNo As Variant, _
    ByRef outOnFoodDate As Variant, _
    ByRef outArrivalDate As Variant _
) As Boolean
    Dim sourceRow As obj_Row
    Dim manualOrderNo As Variant
    Dim arrivalDateRaw As Variant
    Dim arrivalDateFromForm As Date
    Dim effectiveArrivalDate As Date
    Dim orderDateByNo As Date
    Dim hasArrivalDateFromForm As Boolean
    Dim hasOrderDateByNo As Boolean

    outOrderNo = VBA.vbNullString
    outOnFoodDate = VBA.vbNullString
    outArrivalDate = VBA.vbNullString

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    manualOrderNo = private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO)
    outOrderNo = manualOrderNo
    If VBA.Len(VBA.Trim$(VBA.CStr(outOrderNo))) = 0 Then
        outOrderNo = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_INCOMING_NO)
    End If

    arrivalDateRaw = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_DEPARTURE_DATE)
    hasArrivalDateFromForm = private_TryParseIncomingDateWithOrderContext(arrivalDateRaw, outOrderNo, targetWorkbook, arrivalDateFromForm)
    If VBA.Len(VBA.Trim$(VBA.CStr(arrivalDateRaw))) > 0 And Not hasArrivalDateFromForm Then
        private_LogError "Movement failed to resolve arrival date raw='" & private_EscapeForLog(VBA.CStr(arrivalDateRaw)) & "' by orderNo='" & private_EscapeForLog(VBA.CStr(outOrderNo)) & "'."
        VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & MOVEMENT_SOURCE_DEPARTURE_DATE & "' from value '" & VBA.CStr(arrivalDateRaw) & "' and order number '" & VBA.CStr(outOrderNo) & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    hasOrderDateByNo = private_TryResolveOrderDateByNumber(targetWorkbook, outOrderNo, orderDateByNo)

    If hasArrivalDateFromForm Then
        outArrivalDate = arrivalDateFromForm
    ElseIf hasOrderDateByNo Then
        outArrivalDate = orderDateByNo
    End If

    If hasOrderDateByNo Then
        If hasArrivalDateFromForm Then
            effectiveArrivalDate = arrivalDateFromForm
        Else
            effectiveArrivalDate = orderDateByNo
        End If

        If effectiveArrivalDate > orderDateByNo Then
            outOnFoodDate = effectiveArrivalDate
        Else
            outOnFoodDate = VBA.DateAdd("d", 1, orderDateByNo)
        End If
    ElseIf hasArrivalDateFromForm Then
        outOnFoodDate = arrivalDateFromForm
    End If

    private_TryBuildMovementClosingValues = True
End Function

Private Function private_TryFindLastRowByIpn( _
    ByVal targetTable As ListObject, _
    ByVal ipnValue As String, _
    ByRef outRowRange As Range _
) As Boolean
    Dim ipnColumnIndex As Long
    Dim rowIndex As Long
    Dim candidateValue As String
    Dim expectedValue As String

    Set outRowRange = Nothing
    If targetTable Is Nothing Then Exit Function

    expectedValue = private_NormalizeComparableToken(ipnValue)
    If VBA.Len(expectedValue) = 0 Then
        private_LogError "Movement closing failed because source IПН is empty."
        VBA.MsgBox "PrototypeNew: Movement closing requires source value '" & MOVEMENT_TARGET_IPN & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    ipnColumnIndex = private_FindTargetColumnIndex(targetTable, MOVEMENT_TARGET_IPN)
    If ipnColumnIndex <= 0 Then
        private_LogError "Movement closing failed because target column '" & private_EscapeForLog(MOVEMENT_TARGET_IPN) & "' was not found."
        VBA.MsgBox "PrototypeNew: Movement target column '" & MOVEMENT_TARGET_IPN & "' was not found.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    For rowIndex = targetTable.ListRows.Count To 1 Step -1
        candidateValue = private_NormalizeComparableToken(targetTable.ListRows.Item(rowIndex).Range.Cells(1, ipnColumnIndex).Value2)
        If VBA.StrComp(candidateValue, expectedValue, VBA.vbTextCompare) = 0 Then
            Set outRowRange = targetTable.ListRows.Item(rowIndex).Range
            private_TryFindLastRowByIpn = True
            Exit Function
        End If
    Next rowIndex

    private_LogError "Movement closing row was not found by ІПН='" & private_EscapeForLog(expectedValue) & "'."
    VBA.MsgBox "PrototypeNew: failed to find latest Movement row by '" & MOVEMENT_TARGET_IPN & "' = '" & ipnValue & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
End Function

Private Function private_NormalizeComparableToken(ByVal rawValue As Variant) As String
    Dim valueText As String

    valueText = VBA.Trim$(VBA.CStr(rawValue))
    valueText = VBA.Replace(valueText, " ", VBA.vbNullString)
    private_NormalizeComparableToken = valueText
End Function

Private Function private_TryWriteMovementClosingRow( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal arrivalOrderNo As Variant, _
    ByVal onFoodDate As Variant, _
    ByVal arrivalDate As Variant _
) As Boolean
    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ARRIVAL_ORDER_NO, arrivalOrderNo) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ON_FOOD, onFoodDate) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ARRIVAL, arrivalDate) Then Exit Function

    private_TryWriteMovementClosingRow = True
End Function

' //
' // Internal
' //
Private Function private_TryBuildMovementRowValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sectionTypeText As String, _
    ByRef outValues As Variant _
) As Boolean
    Dim sourceRow As obj_Row
    Dim requiredValue As Variant
    Dim destinationValue As Variant
    Dim mappedEventText As String

    Set sourceRow = Nothing
    If VBA.IsArray(outValues) Then Erase outValues
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    ReDim outValues(1 To 1, 1 To MOVEMENT_TARGET_COLUMN_COUNT)

    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "Звання", requiredValue) Then Exit Function
    outValues(1, 1) = requiredValue
    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "ПІБ", requiredValue) Then Exit Function
    outValues(1, 2) = requiredValue
    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "ІПН", requiredValue) Then Exit Function
    outValues(1, 3) = requiredValue
    If Not private_TryGetRequiredSourceText(sourceTable, sourceRow, "Код посади", requiredValue) Then Exit Function
    outValues(1, 4) = requiredValue

    If private_TryMapSectionTypeToEventText(sectionTypeText, mappedEventText) Then
        outValues(1, 5) = mappedEventText
    Else
        outValues(1, 5) = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_EVENT)
    End If

    destinationValue = private_ResolveMovementDestinationValue(sourceTable, sourceRow, sectionTypeText)
    outValues(1, 6) = destinationValue
    private_LogInfo "movement:event resolved value='" & private_EscapeForLog(VBA.CStr(outValues(1, 5))) & "'"
    private_LogInfo "movement:destination resolved value='" & private_EscapeForLog(VBA.CStr(destinationValue)) & "'"

    private_TryBuildMovementRowValues = True
End Function

Private Function private_ResolveMovementDestinationValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal sectionTypeText As String _
) As Variant
    If m_Data.UsesMovementVacationDestination(sectionTypeText) Then
        private_ResolveMovementDestinationValue = private_GetOptionalSourceText(sourceTable, sourceRow, "Відпустка")
        Exit Function
    End If

    private_ResolveMovementDestinationValue = private_GetOptionalSourceText(sourceTable, sourceRow, "Лікарня скорочена назва")
End Function

Private Function private_TryGetRequiredSourceText( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal columnName As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim columnIndex As Long

    outValue = VBA.vbNullString
    columnIndex = private_GetSourceColumnIndex(sourceTable, columnName)
    If columnIndex <= 0 Then
        private_LogError "Movement source column is missing: '" & private_EscapeForLog(columnName) & "'."
        VBA.MsgBox "PrototypeNew: Movement export requires source column '" & columnName & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    outValue = sourceRow.GetCellValue(columnIndex)
    private_LogInfo "movement:required-source column='" & private_EscapeForLog(columnName) & "' index=" & VBA.CStr(columnIndex) & " value='" & private_EscapeForLog(VBA.CStr(outValue)) & "'"
    private_TryGetRequiredSourceText = True
End Function

Private Function private_GetOptionalSourceText( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceRow As obj_Row, _
    ByVal columnName As String _
) As Variant
    Dim columnIndex As Long

    columnIndex = private_GetSourceColumnIndex(sourceTable, columnName)
    If columnIndex <= 0 Then
        private_LogInfo "movement:optional-source column missing='" & private_EscapeForLog(columnName) & "'"
        private_GetOptionalSourceText = VBA.vbNullString
        Exit Function
    End If

    private_GetOptionalSourceText = sourceRow.GetCellValue(columnIndex)
    private_LogInfo "movement:optional-source column='" & private_EscapeForLog(columnName) & "' index=" & VBA.CStr(columnIndex) & " value='" & private_EscapeForLog(VBA.CStr(private_GetOptionalSourceText)) & "'"
End Function

Private Function private_GetSourceColumnIndex(ByVal sourceTable As obj_TableDynamic, ByVal columnName As String) As Long
    Dim sourceColIndex As Long
    Dim sourceColumn As obj_Column
    Dim expectedName As String

    If sourceTable Is Nothing Then Exit Function
    expectedName = private_NormalizeText(columnName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For sourceColIndex = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(sourceColIndex)
        If sourceColumn Is Nothing Then GoTo ContinueColumn
        If VBA.StrComp(private_NormalizeText(sourceColumn.Name), expectedName, VBA.vbTextCompare) = 0 Then
            private_GetSourceColumnIndex = sourceColIndex
            Exit Function
        End If

ContinueColumn:
    Next sourceColIndex
End Function

Private Function private_TryGetAppendRowRange( _
    ByVal targetTable As ListObject, _
    ByRef outRowRange As Range, _
    ByRef outInsertedRow As ListRow _
) As Boolean
    Dim reusableRowIndex As Long

    Set outRowRange = Nothing
    Set outInsertedRow = Nothing
    If targetTable Is Nothing Then Exit Function

    If targetTable.ListColumns.Count < MOVEMENT_TARGET_COLUMN_COUNT Then
        private_LogError "Movement target table has fewer than 6 columns. Actual=" & VBA.CStr(targetTable.ListColumns.Count)
        VBA.MsgBox "PrototypeNew: Movement target table has fewer than 6 columns.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    reusableRowIndex = private_FindTrailingEmptyReuseRowIndex(targetTable)
    If reusableRowIndex > 0 Then
        Set outRowRange = targetTable.ListRows.Item(reusableRowIndex).Range
        private_LogInfo "movement:reuse-empty-row index=" & VBA.CStr(reusableRowIndex) & " address='" & private_EscapeForLog(outRowRange.Address(False, False)) & "'"
        private_TryGetAppendRowRange = Not outRowRange Is Nothing
        Exit Function
    End If

    Set outInsertedRow = targetTable.ListRows.Add
    If outInsertedRow Is Nothing Then
        private_LogError "Movement append row returned Nothing."
        VBA.MsgBox "PrototypeNew: failed to append Movement row to target table.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    Set outRowRange = outInsertedRow.Range
    private_TryGetAppendRowRange = Not outRowRange Is Nothing
End Function

Private Function private_FindTrailingEmptyReuseRowIndex(ByVal targetTable As ListObject) As Long
    Dim firstScanRowIndex As Long
    Dim rowIndex As Long
    Dim firstTrailingEmptyRowIndex As Long

    If targetTable Is Nothing Then Exit Function
    If targetTable.ListRows.Count <= 0 Then Exit Function

    firstScanRowIndex = targetTable.ListRows.Count - MOVEMENT_TRAILING_EMPTY_LOOKBACK_ROWS + 1
    If firstScanRowIndex < 1 Then firstScanRowIndex = 1

    For rowIndex = targetTable.ListRows.Count To firstScanRowIndex Step -1
        If private_IsRowEmpty(targetTable.ListRows.Item(rowIndex).Range) Then
            firstTrailingEmptyRowIndex = rowIndex
        Else
            Exit For
        End If
    Next rowIndex

    private_FindTrailingEmptyReuseRowIndex = firstTrailingEmptyRowIndex
End Function

Private Function private_IsRowEmpty(ByVal rowRange As Range) As Boolean
    Dim cellObj As Range
    Dim valueText As String

    If rowRange Is Nothing Then Exit Function

    For Each cellObj In rowRange.Cells
        If cellObj.HasFormula Then GoTo ContinueCell
        valueText = VBA.Trim$(VBA.CStr(cellObj.Value2))
        If VBA.Len(valueText) > 0 Then Exit Function

ContinueCell:
    Next cellObj

    private_IsRowEmpty = True
End Function

Private Function private_TryWriteMovementRow( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByRef targetValues As Variant, _
    ByVal outgoingOrderNo As Variant, _
    ByVal outgoingFoodFromDate As Variant, _
    ByVal outgoingDepartureDate As Variant, _
    ByVal writeSpecialFields As Boolean, _
    ByVal specialDurationValue As Variant, _
    ByVal specialVkNoValue As Variant, _
    ByVal writeMappedEvent As Boolean, _
    ByVal mappedEventText As String _
) As Boolean
    Dim columnIndex As Long

    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function
    If rowRange.Columns.Count < MOVEMENT_TARGET_COLUMN_COUNT Then
        private_LogError "Movement target row has fewer than 6 cells. Actual=" & VBA.CStr(rowRange.Columns.Count)
        VBA.MsgBox "PrototypeNew: Movement target row has fewer than 6 cells.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    For columnIndex = 1 To MOVEMENT_TARGET_COLUMN_COUNT
        If Not private_TryWriteCellValueWithFormulaPolicy(rowRange.Cells(1, columnIndex), targetValues(1, columnIndex), "movement:write-cell col=" & VBA.CStr(columnIndex)) Then Exit Function
    Next columnIndex

    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_ORDER_NO, outgoingOrderNo) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_FOOD_FROM, outgoingFoodFromDate) Then Exit Function
    If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_DEPARTURE, outgoingDepartureDate) Then Exit Function

    If writeSpecialFields Then
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_DURATION_DAYS, specialDurationValue) Then Exit Function
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_VK_NO, specialVkNoValue) Then Exit Function
    End If

    If writeMappedEvent Then
        If Not private_TryWriteNamedColumnValue(targetTable, rowRange, MOVEMENT_TARGET_EVENT, mappedEventText) Then Exit Function
    End If

    private_TryWriteMovementRow = True
End Function

Private Function private_TryMapSectionTypeToEventText( _
    ByVal rawSectionType As String, _
    ByRef outEventText As String _
) As Boolean
    private_TryMapSectionTypeToEventText = m_Data.TryMapMovementSectionTypeToEventText(rawSectionType, outEventText)
End Function

Private Function private_TryBuildMovementOutgoingValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal targetWorkbook As Workbook, _
    ByVal context As Object, _
    ByRef outOrderNo As Variant, _
    ByRef outFoodFromDate As Variant, _
    ByRef outDepartureDate As Variant _
) As Boolean
    Dim sourceRow As obj_Row
    Dim incomingDate As Date
    Dim departureDateFromForm As Date
    Dim orderDateByNo As Date
    Dim manualOrderNo As Variant
    Dim incomingDateRaw As Variant
    Dim departureDateRaw As Variant
    Dim hasIncomingDate As Boolean
    Dim hasDepartureDateFromForm As Boolean
    Dim hasOrderDateByNo As Boolean

    outOrderNo = VBA.vbNullString
    outFoodFromDate = VBA.vbNullString
    outDepartureDate = VBA.vbNullString

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    manualOrderNo = private_GetContextText(context, MOVEMENT_CONTEXT_MANUAL_ORDER_NO)
    outOrderNo = manualOrderNo
    If VBA.Len(VBA.Trim$(VBA.CStr(outOrderNo))) = 0 Then
        outOrderNo = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_INCOMING_NO)
    End If
    incomingDateRaw = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_INCOMING_DATE)
    departureDateRaw = private_GetOptionalSourceText(sourceTable, sourceRow, MOVEMENT_SOURCE_DEPARTURE_DATE)

    hasIncomingDate = private_TryParseIncomingDateWithOrderContext(incomingDateRaw, outOrderNo, targetWorkbook, incomingDate)
    If VBA.Len(VBA.Trim$(VBA.CStr(incomingDateRaw))) > 0 And Not hasIncomingDate Then
        private_LogError "Movement failed to resolve incoming date raw='" & private_EscapeForLog(VBA.CStr(incomingDateRaw)) & "' by orderNo='" & private_EscapeForLog(VBA.CStr(outOrderNo)) & "'."
        VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & MOVEMENT_SOURCE_INCOMING_DATE & "' from value '" & VBA.CStr(incomingDateRaw) & "' and order number '" & VBA.CStr(outOrderNo) & "'.", VBA.vbExclamation, "PrototypeNew / Movement export"
        Exit Function
    End If

    hasDepartureDateFromForm = private_TryParseDateValue(departureDateRaw, departureDateFromForm)
    hasOrderDateByNo = private_TryResolveOrderDateByNumber(targetWorkbook, outOrderNo, orderDateByNo)

    If hasDepartureDateFromForm Then
        outDepartureDate = departureDateFromForm
    ElseIf hasIncomingDate Then
        outDepartureDate = incomingDate
    End If

    If hasOrderDateByNo Then
        If hasDepartureDateFromForm And departureDateFromForm > orderDateByNo Then
            outFoodFromDate = departureDateFromForm
        Else
            outFoodFromDate = VBA.DateAdd("d", 1, orderDateByNo)
        End If
    ElseIf hasDepartureDateFromForm Then
        outFoodFromDate = departureDateFromForm
    End If

    private_TryBuildMovementOutgoingValues = True
End Function

Private Function private_TryParseIncomingDateWithOrderContext( _
    ByVal rawIncomingDate As Variant, _
    ByVal rawOrderNo As Variant, _
    ByVal targetWorkbook As Workbook, _
    ByRef outDateValue As Date _
) As Boolean
    Dim valueText As String
    Dim normalizedText As String
    Dim parts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim orderDate As Date
    Dim hasOrderDate As Boolean

    valueText = VBA.Trim$(VBA.CStr(rawIncomingDate))
    If VBA.Len(valueText) = 0 Then Exit Function

    normalizedText = VBA.Replace(valueText, "-", ".")
    normalizedText = VBA.Replace(normalizedText, "/", ".")
    normalizedText = VBA.Replace(normalizedText, " ", VBA.vbNullString)
    If VBA.Len(normalizedText) = 0 Then Exit Function

    parts = VBA.Split(normalizedText, ".")

    If UBound(parts) = 2 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.CLng(VBA.Val(parts(1)))
        yearValue = VBA.CLng(VBA.Val(parts(2)))
        private_TryParseIncomingDateWithOrderContext = private_TryBuildDate(dayValue, monthValue, yearValue, outDateValue)
        Exit Function
    End If

    hasOrderDate = private_TryResolveOrderDateByNumber(targetWorkbook, rawOrderNo, orderDate)

    If UBound(parts) = 1 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.CLng(VBA.Val(parts(1)))
        If hasOrderDate Then
            yearValue = VBA.Year(orderDate)
            private_TryParseIncomingDateWithOrderContext = private_TryBuildDate(dayValue, monthValue, yearValue, outDateValue)
            Exit Function
        End If
    End If

    If UBound(parts) = 0 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        If hasOrderDate Then
            monthValue = VBA.Month(orderDate)
            yearValue = VBA.Year(orderDate)
            private_TryParseIncomingDateWithOrderContext = private_TryBuildDate(dayValue, monthValue, yearValue, outDateValue)
            Exit Function
        End If
    End If

    private_TryParseIncomingDateWithOrderContext = private_TryParseDateValue(rawIncomingDate, outDateValue)
End Function

Private Function private_TryResolveOrderDateByNumber( _
    ByVal targetWorkbook As Workbook, _
    ByVal rawOrderNo As Variant, _
    ByRef outOrderDate As Date _
) As Boolean
    Dim ws As Worksheet
    Dim ordersTable As ListObject
    Dim orderDateColumnIndex As Long
    Dim orderNoColumnIndex As Long
    Dim orderNoToken As String
    Dim rowIndex As Long
    Dim candidateOrderToken As String
    Dim candidateDate As Date

    If targetWorkbook Is Nothing Then Exit Function

    orderNoToken = private_NormalizeOrderNumberToken(rawOrderNo)
    If VBA.Len(orderNoToken) = 0 Then Exit Function

    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(MOVEMENT_ORDERS_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        private_LogInfo "movement:orders-sheet missing name='" & private_EscapeForLog(MOVEMENT_ORDERS_SHEET_NAME) & "'"
        Exit Function
    End If

    Set ordersTable = private_FindListObjectByName(ws, MOVEMENT_ORDER_DATE_TABLE_NAME)
    If ordersTable Is Nothing Then
        private_LogInfo "movement:order-date table missing name='" & private_EscapeForLog(MOVEMENT_ORDER_DATE_TABLE_NAME) & "' sheet='" & private_EscapeForLog(ws.Name) & "'"
        Exit Function
    End If
    If ordersTable.DataBodyRange Is Nothing Then Exit Function

    orderDateColumnIndex = private_FindTargetColumnIndex(ordersTable, MOVEMENT_ORDER_DATE_COLUMN_NAME)
    orderNoColumnIndex = private_FindTargetColumnIndex(ordersTable, MOVEMENT_ORDER_NO_COLUMN_NAME)
    If orderDateColumnIndex <= 0 Or orderNoColumnIndex <= 0 Then
        private_LogInfo "movement:order-date table columns missing date='" & private_EscapeForLog(MOVEMENT_ORDER_DATE_COLUMN_NAME) & "' no='" & private_EscapeForLog(MOVEMENT_ORDER_NO_COLUMN_NAME) & "'"
        Exit Function
    End If

    For rowIndex = 1 To ordersTable.ListRows.Count
        candidateOrderToken = private_NormalizeOrderNumberToken(ordersTable.DataBodyRange.Cells(rowIndex, orderNoColumnIndex).Value2)
        If VBA.Len(candidateOrderToken) = 0 Then GoTo ContinueRow
        If VBA.StrComp(candidateOrderToken, orderNoToken, VBA.vbTextCompare) <> 0 Then GoTo ContinueRow

        If private_TryParseDateValue(ordersTable.DataBodyRange.Cells(rowIndex, orderDateColumnIndex).Value2, candidateDate) Then
            outOrderDate = candidateDate
            private_LogInfo "movement:order-date resolved table='" & private_EscapeForLog(ordersTable.Name) & "' orderNo='" & private_EscapeForLog(orderNoToken) & "' date='" & private_EscapeForLog(private_FormatLogDateValue(outOrderDate)) & "'"
            private_TryResolveOrderDateByNumber = True
            Exit Function
        End If

ContinueRow:
    Next rowIndex
End Function

Private Function private_FindListObjectByName( _
    ByVal ws As Worksheet, _
    ByVal tableName As String _
) As ListObject
    Dim tableObj As ListObject
    Dim expectedName As String

    If ws Is Nothing Then Exit Function
    expectedName = VBA.Trim$(tableName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For Each tableObj In ws.ListObjects
        If VBA.StrComp(tableObj.Name, expectedName, VBA.vbTextCompare) = 0 Then
            Set private_FindListObjectByName = tableObj
            Exit Function
        End If
    Next tableObj
End Function

Private Function private_NormalizeOrderNumberToken(ByVal rawValue As Variant) As String
    Dim valueText As String
    Dim numericValue As Double

    valueText = VBA.Trim$(VBA.CStr(rawValue))
    valueText = VBA.Replace(valueText, " ", VBA.vbNullString)
    If VBA.Len(valueText) = 0 Then Exit Function

    If VBA.IsNumeric(valueText) Then
        numericValue = VBA.CDbl(valueText)
        If numericValue >= 0 Then
            private_NormalizeOrderNumberToken = VBA.CStr(VBA.CLng(numericValue))
            Exit Function
        End If
    End If

    private_NormalizeOrderNumberToken = valueText
End Function

Private Function private_TryBuildDate( _
    ByVal dayValue As Long, _
    ByVal monthValue As Long, _
    ByVal yearValue As Long, _
    ByRef outDateValue As Date _
) As Boolean
    Dim candidate As Date

    If yearValue < 1900 Then Exit Function
    If monthValue < 1 Or monthValue > 12 Then Exit Function
    If dayValue < 1 Or dayValue > 31 Then Exit Function

    On Error Resume Next
    candidate = VBA.DateSerial(yearValue, monthValue, dayValue)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If VBA.Day(candidate) <> dayValue Then Exit Function
    If VBA.Month(candidate) <> monthValue Then Exit Function
    If VBA.Year(candidate) <> yearValue Then Exit Function

    outDateValue = candidate
    private_TryBuildDate = True
End Function

Private Function private_TryWriteNamedColumnValue( _
    ByVal targetTable As ListObject, _
    ByVal rowRange As Range, _
    ByVal targetColumnName As String, _
    ByVal incomingValue As Variant _
) As Boolean
    Dim targetColumnIndex As Long

    If targetTable Is Nothing Then Exit Function
    If rowRange Is Nothing Then Exit Function

    targetColumnIndex = private_FindTargetColumnIndex(targetTable, targetColumnName)
    If targetColumnIndex <= 0 Then
        private_LogInfo "movement:target-column missing='" & private_EscapeForLog(targetColumnName) & "'"
        private_TryWriteNamedColumnValue = True
        Exit Function
    End If

    private_TryWriteNamedColumnValue = private_TryWriteCellValueWithFormulaPolicy( _
        rowRange.Cells(1, targetColumnIndex), _
        incomingValue, _
        "movement:write-named col='" & private_EscapeForLog(targetColumnName) & "'")
End Function

Private Function private_FindTargetColumnIndex(ByVal targetTable As ListObject, ByVal targetColumnName As String) As Long
    Dim columnObj As ListColumn
    Dim expectedName As String
    Dim candidateName As String

    If targetTable Is Nothing Then Exit Function

    expectedName = private_NormalizeText(targetColumnName)
    If VBA.Len(expectedName) = 0 Then Exit Function

    For Each columnObj In targetTable.ListColumns
        candidateName = private_NormalizeText(VBA.CStr(columnObj.Name))
        If VBA.StrComp(candidateName, expectedName, VBA.vbTextCompare) = 0 Then
            private_FindTargetColumnIndex = columnObj.Index
            Exit Function
        End If
    Next columnObj
End Function

Private Function private_TryWriteCellValueWithFormulaPolicy( _
    ByVal targetCell As Range, _
    ByVal incomingValue As Variant, _
    ByVal logPrefix As String _
) As Boolean
    Dim incomingText As String

    If targetCell Is Nothing Then Exit Function

    incomingText = VBA.Trim$(VBA.CStr(incomingValue))
    If VBA.Len(incomingText) = 0 And targetCell.HasFormula Then
        private_LogInfo logPrefix & " skip-formula-preserve address='" & private_EscapeForLog(targetCell.Address(False, False)) & "'"
        private_TryWriteCellValueWithFormulaPolicy = True
        Exit Function
    End If

    private_LogInfo logPrefix & " value='" & private_EscapeForLog(VBA.CStr(incomingValue)) & "'"
    targetCell.Value2 = incomingValue
    private_TryWriteCellValueWithFormulaPolicy = True
End Function

Private Function private_TryParseDateValue(ByVal rawValue As Variant, ByRef outDateValue As Date) As Boolean
    Dim valueText As String
    Dim normalizedText As String
    Dim parts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim serialValue As Double

    valueText = VBA.Trim$(VBA.CStr(rawValue))
    If VBA.Len(valueText) = 0 Then Exit Function

    If VBA.IsNumeric(valueText) Then
        serialValue = VBA.CDbl(valueText)
        If serialValue > 0 Then
            outDateValue = VBA.DateSerial(1899, 12, 30) + serialValue
            private_TryParseDateValue = True
            Exit Function
        End If
    End If

    normalizedText = VBA.Replace(valueText, "-", ".")
    parts = VBA.Split(normalizedText, ".")
    If UBound(parts) = 2 Then
        dayValue = VBA.CLng(VBA.Val(parts(0)))
        monthValue = VBA.CLng(VBA.Val(parts(1)))
        yearValue = VBA.CLng(VBA.Val(parts(2)))
        If yearValue >= 1900 And monthValue >= 1 And monthValue <= 12 And dayValue >= 1 And dayValue <= 31 Then
            On Error Resume Next
            outDateValue = VBA.DateSerial(yearValue, monthValue, dayValue)
            If Err.Number = 0 Then
                private_TryParseDateValue = True
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        End If
    End If

    If VBA.IsDate(valueText) Then
        outDateValue = VBA.CDate(valueText)
        private_TryParseDateValue = True
    End If
End Function

Private Function private_FormatLogDateValue(ByVal valueObj As Variant) As String
    If VBA.IsDate(valueObj) Then
        private_FormatLogDateValue = VBA.Format$(VBA.CDate(valueObj), "dd.mm.yyyy")
    Else
        private_FormatLogDateValue = VBA.CStr(valueObj)
    End If
End Function

Private Function private_NormalizeText(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueText)))
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ".", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "/", " ")
    valueText = VBA.Replace(valueText, "(", " ")
    valueText = VBA.Replace(valueText, ")", " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_NormalizeText = VBA.Trim$(valueText)
End Function

Private Sub private_LogMethodEntry(ByVal methodName As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "exporter:enter " & VBA.Trim$(methodName)
#End If
End Sub

Private Sub private_LogInfo(ByVal message As String)
    ex_Core.fn_Diagnostic_LogInfo "movement-export: " & message
End Sub

Private Sub private_LogError(ByVal message As String)
    ex_Core.fn_Diagnostic_LogError "Movement export: " & message
End Sub

Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace(VBA.CStr(valueText), "'", "''")
End Function

Private Function private_BoolText(ByVal value As Boolean) As String
    If value Then
        private_BoolText = "True"
    Else
        private_BoolText = "False"
    End If
End Function
