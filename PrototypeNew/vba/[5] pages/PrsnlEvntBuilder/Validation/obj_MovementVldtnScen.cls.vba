VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_MovementVldtnScen"
Option Explicit

Implements obj_IMultiSourcesScenario

Private Const TABLES_RUNTIME_KEY As String = "RuntimeItems.MultiSourcesView.Tables"
Private Const PREVIOUS_ORDERS_RUNTIME_KEY As String = "RuntimeItems.MovementValidation.PreviousOrders"
Private Const ORDER_NO_INPUT_NAME As String = "OrderNoInput"
Private Const FIO_ALIAS As String = "FIO"
Private Const EVENT_ALIAS As String = "Event"
Private Const ORDER_ALIAS As String = "OutOrder"
Private Const DATE_ALIAS As String = "HospitalizationDate"
Private Const OUT_DATE_ALIAS As String = "OutDate"
Private Const PLANNED_RETURN_ALIAS As String = "PlannedReturnDate"
Private Const RETURN_DATE_ALIAS As String = "ReturnDate"
Private Const NOTES_ALIAS As String = "Примітка"
Private Const STATUS_OK_TAG As String = "validation-ok"
Private Const STATUS_CONFLICT_TAG As String = "validation-conflict"
Private Const ERROR_TITLE As String = "PrototypeNew / MovementValidation"

Private m_Page As obj_IPage
Private m_ConfigTable As obj_ConfigTable
Private m_CfgParser As obj_MultiSourcesViewCfgParser
Private m_MovementTableRef As String
Private m_EjosTableRef As String
Private m_MedicalTableRef As String
Private m_EjosEvents As Collection
Private m_MedicalEvents As Collection
Private m_OrderDateProviderClass As String
Private m_OrderNoText As String
Private m_EmbeddedOrderDate As Date
Private m_EmbeddedResultRuntimeKey As String
Private m_IsEmbedded As Boolean
Private m_IsDisposed As Boolean

Public Function InitializeEmbedded( _
    ByVal page As obj_IPage, _
    ByVal configTable As obj_ConfigTable, _
    ByVal orderNo As String, _
    ByVal orderDate As Date, _
    ByVal resultRuntimeKey As String _
) As Boolean
    Dim multiSourcesViewCfgParser As obj_MultiSourcesViewCfgParser

    If page Is Nothing Or configTable Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(orderNo)) = 0 Or orderDate = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(resultRuntimeKey)) = 0 Then Exit Function
    m_IsDisposed = False
    m_IsEmbedded = True
    Set m_Page = page
    Set m_ConfigTable = configTable
    m_OrderNoText = VBA.Trim$(orderNo)
    m_EmbeddedOrderDate = VBA.DateValue(orderDate)
    m_EmbeddedResultRuntimeKey = VBA.Trim$(resultRuntimeKey)
    Set multiSourcesViewCfgParser = New obj_MultiSourcesViewCfgParser
    If Not multiSourcesViewCfgParser.Initialize(configTable) Then Exit Function
    If Not private_TryReadSettings(configTable) Then Exit Function
    Set m_CfgParser = multiSourcesViewCfgParser
    InitializeEmbedded = True
End Function

Public Function RunEmbedded( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    If Not m_IsEmbedded Then Exit Function
    RunEmbedded = private_RunPipeline(notifyChange)
End Function

Private Function obj_IMultiSourcesScenario_Initialize( _
    ByVal page As obj_IPage, _
    ByVal configTable As obj_ConfigTable _
) As Boolean
    Dim multiSourcesViewCfgParser As obj_MultiSourcesViewCfgParser
    Dim pageBase As obj_PageBase
    Dim emptyItems As Collection

    If page Is Nothing Or configTable Is Nothing Then Exit Function
    m_IsDisposed = False
    m_IsEmbedded = False
    Set m_Page = page
    Set m_ConfigTable = configTable
    Set multiSourcesViewCfgParser = New obj_MultiSourcesViewCfgParser
    If Not multiSourcesViewCfgParser.Initialize(configTable) Then Exit Function
    If Not private_TryReadSettings(configTable) Then Exit Function
    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set emptyItems = New Collection
    If Not pageBase.RuntimeSources.SetItemsSource( _
        PREVIOUS_ORDERS_RUNTIME_KEY, emptyItems, False) Then Exit Function
    Set m_CfgParser = multiSourcesViewCfgParser
    obj_IMultiSourcesScenario_Initialize = True
End Function

Private Sub obj_IMultiSourcesScenario_Dispose()
    private_Dispose
End Sub

Private Function obj_IMultiSourcesScenario_RunPipeline( _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    obj_IMultiSourcesScenario_RunPipeline = private_RunPipeline(notifyChange)
End Function

Private Property Get obj_IMultiSourcesScenario_OrderNoText() As String
    obj_IMultiSourcesScenario_OrderNoText = m_OrderNoText
End Property

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    private_Dispose
    On Error GoTo 0
End Sub

Private Sub private_Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    Set m_CfgParser = Nothing
    Set m_ConfigTable = Nothing
    Set m_EjosEvents = Nothing
    Set m_MedicalEvents = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Private Function private_TryReadSettings( _
    ByVal configTable As obj_ConfigTable _
) As Boolean
    Dim cfgTableParser As obj_CfgTableParser
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim rawEjosEvents As String
    Dim rawMedicalEvents As String

    Set cfgTableParser = New obj_CfgTableParser
    If Not cfgTableParser.Initialize(configTable, Me) Then Exit Function
    If Not cfgTableParser.CfgParserBase.TryGetConfigEntries( _
        configEntries) Then Exit Function
    If Not cfgTableParser.CfgParserBase.BuildConfigDictionary( _
        configEntries, cfgMap) Then Exit Function

    If Not private_TryGetRequiredValue(cfgTableParser, cfgMap, _
        "MovementValidation.MovementTable", m_MovementTableRef) Then Exit Function
    If Not private_TryGetRequiredValue(cfgTableParser, cfgMap, _
        "MovementValidation.EjosTable", m_EjosTableRef) Then Exit Function
    If Not private_TryGetRequiredValue(cfgTableParser, cfgMap, _
        "MovementValidation.MedicalTable", m_MedicalTableRef) Then Exit Function
    If Not private_TryGetRequiredValue(cfgTableParser, cfgMap, _
        "MovementValidation.EjosEvents", rawEjosEvents) Then Exit Function
    If Not private_TryGetRequiredValue(cfgTableParser, cfgMap, _
        "MovementValidation.MedicalEvents", rawMedicalEvents) Then Exit Function
    If Not private_TryGetRequiredValue(cfgTableParser, cfgMap, _
        "MovementValidation.OrderDateProviderClass", _
        m_OrderDateProviderClass) Then Exit Function

    Set m_EjosEvents = cfgTableParser.CfgParserBase.SplitListToCollection( _
        rawEjosEvents)
    Set m_MedicalEvents = cfgTableParser.CfgParserBase.SplitListToCollection( _
        rawMedicalEvents)
    If m_EjosEvents Is Nothing Or m_EjosEvents.Count = 0 Then
        private_ShowError "MovementValidation.EjosEvents is empty."
        Exit Function
    End If
    If m_MedicalEvents Is Nothing Or m_MedicalEvents.Count = 0 Then
        private_ShowError "MovementValidation.MedicalEvents is empty."
        Exit Function
    End If

    cfgTableParser.Dispose
    private_TryReadSettings = True
End Function

Private Function private_TryGetRequiredValue( _
    ByVal cfgTableParser As obj_CfgTableParser, _
    ByVal cfgMap As Object, _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    If Not cfgTableParser.CfgParserBase.TryGetRequiredConfigValue( _
        cfgMap, keyName, outValue) Then
        private_ShowError "Required config key is missing or empty: " & keyName
        Exit Function
    End If
    private_TryGetRequiredValue = True
End Function

Private Function private_RunPipeline( _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    Dim orderNo As String
    Dim orderDate As Date
    Dim orderDateProvider As Object
    Dim movementColumns As Collection
    Dim ejosColumns As Collection
    Dim medicalColumns As Collection
    Dim movementParams As obj_SqlParams
    Dim ejosParams As obj_SqlParams
    Dim medicalParamsList As Collection
    Dim medicalParams As obj_SqlParams
    Dim movementData As obj_TableData
    Dim ejosData As obj_TableData
    Dim medicalData As obj_TableData
    Dim ejosMovementNames As Object
    Dim medicalMovementNames As Object
    Dim latestMovementRowByFio As Object
    Dim resultTables As Collection
    Dim previousOrderTables As Collection
    Dim ejosResultTable As obj_TableDynamic
    Dim medicalResultTable As obj_TableDynamic
    Dim pageBase As obj_PageBase

    On Error GoTo EH
    If m_IsDisposed Or m_CfgParser Is Nothing Then
        private_ShowError "MovementValidation scenario is not initialized."
        Exit Function
    End If
    If m_IsEmbedded Then
        orderNo = m_OrderNoText
        orderDate = m_EmbeddedOrderDate
    Else
        If Not private_TryReadOrderNo(orderNo) Then Exit Function
        If Not private_TryCreateOrderDateProvider( _
            m_OrderDateProviderClass, orderDateProvider) Then Exit Function
        If Not orderDateProvider.Initialize(m_ConfigTable) Then Exit Function
        If Not orderDateProvider.TryResolveOrderDateByNumber(orderNo, orderDate) Then
            private_ShowError "Наказ № " & orderNo & _
                " не знайдено у довіднику «Накази»."
            GoTo CleanFail
        End If
    End If

    Set movementColumns = private_CreateStringCollection( _
        FIO_ALIAS, EVENT_ALIAS, ORDER_ALIAS, OUT_DATE_ALIAS, _
        PLANNED_RETURN_ALIAS, RETURN_DATE_ALIAS)
    Set ejosColumns = private_CreateStringCollection( _
        "Rank", FIO_ALIAS, "IPN", EVENT_ALIAS, "Destination", _
        "DepartureDate", "OrderDate", ORDER_ALIAS, "Duration", _
        PLANNED_RETURN_ALIAS, "Field74", "Field75A", "Field75B", _
        "Basis", "SequenceNo", "Field2", "Field3")
    Set medicalColumns = private_CreateStringCollection( _
        FIO_ALIAS, "IPN", DATE_ALIAS, "BasisDocName", "BasisDocDate", _
        "BasisDocNumber", "FullHospitalName", "TreatmentStatus")
    If Not m_CfgParser.TryBuildTableSqlParams(m_MovementTableRef, _
        movementColumns, movementParams) Then GoTo CleanFail
    If Not m_CfgParser.TryBuildTableSqlParams(m_EjosTableRef, _
        ejosColumns, ejosParams) Then GoTo CleanFail
    If Not m_CfgParser.TryBuildTableSqlParamsList(m_MedicalTableRef, _
        medicalColumns, medicalParamsList) Then GoTo CleanFail
    If medicalParamsList.Count <> 1 Then
        private_ShowError _
            "Resolver строевой записки должен вернуть ровно один последний файл."
        GoTo CleanFail
    End If
    Set medicalParams = medicalParamsList.Item(1)

    movementParams.MaxRows = 20000
    ejosParams.MaxRows = 20000
    medicalParams.MaxRows = 20000
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData( _
        movementParams, movementData) Then GoTo CleanFail
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData( _
        ejosParams, ejosData) Then GoTo CleanFail
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData( _
        medicalParams, medicalData) Then GoTo CleanFail

    Set ejosMovementNames = private_BuildMovementNameSet( _
        movementData, orderNo, m_EjosEvents)
    Set medicalMovementNames = private_BuildMovementNameSet( _
        movementData, orderNo, m_MedicalEvents)
    Set latestMovementRowByFio = private_BuildLatestMovementRowMap(movementData)

    Set resultTables = New Collection
    If Not private_TryBuildEjosTable( _
        "ЄЖОС_last: перевірка Movement", ejosData, orderDate, _
        ejosMovementNames, ejosResultTable) Then GoTo CleanFail
    resultTables.Add ejosResultTable
    If Not private_TryBuildMedicalTable( _
        "Стройова записка: перевірка Movement", medicalData, orderDate, _
        medicalMovementNames, latestMovementRowByFio, movementData, _
        medicalResultTable) Then GoTo CleanFail
    resultTables.Add medicalResultTable

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then GoTo CleanFail
    If m_IsEmbedded Then
        If Not pageBase.RuntimeSources.RemoveItemsSource( _
            VBA.LCase$(m_EmbeddedResultRuntimeKey)) Then GoTo CleanFail
        If Not pageBase.RuntimeSources.SetItemsSource( _
            VBA.LCase$(m_EmbeddedResultRuntimeKey), resultTables, False) Then GoTo CleanFail
    Else
        If Not private_TryBuildPreviousOrderTables( _
            orderNo, orderDate, orderDateProvider, _
            previousOrderTables) Then GoTo CleanFail
        If Not pageBase.RuntimeSources.SetItemsSource( _
            PREVIOUS_ORDERS_RUNTIME_KEY, previousOrderTables, False) Then GoTo CleanFail
        If Not pageBase.RuntimeSources.SetItemsSource( _
            TABLES_RUNTIME_KEY, resultTables, False) Then GoTo CleanFail
    End If
    If notifyChange Then
        If Not rt_PageManager.fn_RenderPage( _
            m_Page, "movementvalidation:run") Then GoTo CleanFail
    End If

    rt_Messaging.fn_ShowStatusBarSuccess _
        "MovementValidation: ЄЖОС і стройова записка перевірені.", 4
    If Not orderDateProvider Is Nothing Then orderDateProvider.Dispose
    private_RunPipeline = True
    Exit Function

CleanFail:
    If Not orderDateProvider Is Nothing Then orderDateProvider.Dispose
    Exit Function
EH:
    If Not orderDateProvider Is Nothing Then orderDateProvider.Dispose
    private_ShowError "Помилка MovementValidation: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description
    Err.Clear
End Function

Private Function private_TryBuildPreviousOrderTables( _
    ByVal currentOrderNo As String, _
    ByVal currentOrderDate As Date, _
    ByVal orderDateProvider As Object, _
    ByRef outTables As Collection _
) As Boolean
    Dim normalizedOrderNo As String
    Dim currentOrderNumber As Long
    Dim listedOrderNo As String
    Dim listedOrderDate As Date
    Dim tableObj As obj_TableDynamic
    Dim offset As Long

    Set outTables = Nothing
    normalizedOrderNo = private_NormalizeOrderNo(currentOrderNo)
    If Not VBA.IsNumeric(normalizedOrderNo) Then
        private_ShowError "Для списку попередніх наказів потрібен " & _
            "числовий номер наказу: '" & currentOrderNo & "'."
        Exit Function
    End If
    currentOrderNumber = VBA.CLng(normalizedOrderNo)
    If currentOrderNumber <= 4 Then
        private_ShowError "Номер наказу має бути більшим за 4."
        Exit Function
    End If

    Set tableObj = New obj_TableDynamic
    If Not private_AddTableColumn(tableObj, "Дата наказу", 1) Then Exit Function
    If Not private_AddTableColumn(tableObj, "Номер наказу", 2) Then Exit Function
    For offset = 0 To 4
        listedOrderNo = VBA.CStr(currentOrderNumber - offset)
        If offset = 0 Then
            listedOrderDate = currentOrderDate
        Else
            If Not orderDateProvider.TryResolveOrderDateByNumber( _
                listedOrderNo, listedOrderDate) Then
                private_ShowError "Не знайдено дату наказу № " & _
                    listedOrderNo & "."
                Exit Function
            End If
        End If
        If Not private_AddPreviousOrderRow(tableObj, _
            listedOrderDate, listedOrderNo) Then Exit Function
    Next offset

    Set outTables = New Collection
    outTables.Add tableObj
    private_TryBuildPreviousOrderTables = True
End Function

Private Function private_AddTableColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal columnName As String, _
    ByVal position As Long _
) As Boolean
    Dim columnObj As obj_Column

    Set columnObj = New obj_Column
    columnObj.Name = columnName
    columnObj.Position = position
    If Not columnObj.AddAlias(columnName) Then Exit Function
    private_AddTableColumn = tableObj.PushColumn(columnObj)
End Function

Private Function private_AddPreviousOrderRow( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal orderDate As Date, _
    ByVal orderNo As String _
) As Boolean
    Dim rowObj As obj_Row
    Dim cellObj As obj_Cell

    Set rowObj = New obj_Row
    Set cellObj = New obj_Cell
    cellObj.Value = VBA.Format$(orderDate, "dd.mm.yyyy")
    If Not rowObj.PushCell(cellObj) Then Exit Function
    Set cellObj = New obj_Cell
    cellObj.Value = orderNo
    If Not rowObj.PushCell(cellObj) Then Exit Function
    private_AddPreviousOrderRow = tableObj.PushRow(rowObj)
End Function

Private Function private_TryCreateOrderDateProvider( _
    ByVal className As String, _
    ByRef outProvider As Object _
) As Boolean
    Set outProvider = Nothing
    className = VBA.Trim$(className)
    If VBA.Len(className) = 0 Then
        private_ShowError _
            "MovementValidation.OrderDateProviderClass is empty."
        Exit Function
    End If

    Select Case VBA.LCase$(className)
        Case VBA.LCase$("obj_PEB_ExptrCommonDataPrvdr")
            Set outProvider = New obj_PEB_ExptrCommonDataPrvdr
        Case Else
            private_ShowError "Unsupported order date provider class: " & _
                className
            Exit Function
    End Select
    private_TryCreateOrderDateProvider = Not outProvider Is Nothing
End Function

Private Function private_TryReadOrderNo(ByRef outOrderNo As String) As Boolean
    Dim rawControl As Object
    Dim orderNoInput As obj_InputControlVM
    Dim rawValue As String

    outOrderNo = VBA.vbNullString
    If Not m_Page.TryGetRegisteredControlByName( _
        ORDER_NO_INPUT_NAME, rawControl) Then
        private_ShowError "Не знайдено поле номера наказу '" & _
            ORDER_NO_INPUT_NAME & "'."
        Exit Function
    End If
    If rawControl Is Nothing Or Not TypeOf rawControl Is obj_InputControlVM Then
        private_ShowError "Контрол '" & ORDER_NO_INPUT_NAME & _
            "' має неправильний тип."
        Exit Function
    End If
    Set orderNoInput = rawControl
    If Not orderNoInput.TryGetValue(rawValue) Then Exit Function
    outOrderNo = VBA.Trim$(rawValue)
    If VBA.Len(outOrderNo) = 0 Then
        private_ShowError "Введіть номер поточного наказу."
        Exit Function
    End If
    m_OrderNoText = outOrderNo
    private_TryReadOrderNo = True
End Function

Private Function private_CreateStringCollection( _
    ParamArray values() As Variant _
) As Collection
    Dim result As Collection
    Dim item As Variant

    Set result = New Collection
    For Each item In values
        result.Add VBA.CStr(item)
    Next item
    Set private_CreateStringCollection = result
End Function

Private Function private_BuildMovementNameSet( _
    ByVal movementData As obj_TableData, _
    ByVal orderNo As String, _
    ByVal allowedEvents As Collection _
) As Object
    Dim result As Object
    Dim rowIndex As Long
    Dim fioText As String
    Dim normalizedFio As String

    Set result = ex_Helpers.fn_CreateDictionaryTextCompare()
    For rowIndex = 1 To movementData.RowCount
        If VBA.StrComp(private_NormalizeOrderNo( _
            movementData.ValueAt(rowIndex, 3)), _
            private_NormalizeOrderNo(orderNo), VBA.vbTextCompare) = 0 Then
            If private_TextIsInList( _
                movementData.ValueAt(rowIndex, 2), allowedEvents) Then
                fioText = VBA.Trim$(movementData.ValueAt(rowIndex, 1))
                normalizedFio = private_NormalizeFio(fioText)
                If VBA.Len(normalizedFio) > 0 Then
                    result(normalizedFio) = fioText
                End If
            End If
        End If
    Next rowIndex
    Set private_BuildMovementNameSet = result
End Function

Private Function private_BuildLatestMovementRowMap( _
    ByVal movementData As obj_TableData _
) As Object
    Dim result As Object
    Dim rowIndex As Long
    Dim normalizedFio As String

    Set result = ex_Helpers.fn_CreateDictionaryTextCompare()
    ' Movement хранится в хронологическом порядке. Последняя физическая
    ' строка человека заменяет предыдущую и становится текущим состоянием.
    For rowIndex = 1 To movementData.RowCount
        normalizedFio = private_NormalizeFio( _
            movementData.ValueAt(rowIndex, 1))
        If VBA.Len(normalizedFio) > 0 Then
            result(normalizedFio) = rowIndex
        End If
    Next rowIndex
    Set private_BuildLatestMovementRowMap = result
End Function

Private Function private_TryBuildEjosTable( _
    ByVal titleText As String, _
    ByVal ejosData As obj_TableData, _
    ByVal orderDate As Date, _
    ByVal currentOrderNames As Object, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim columnNames As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim normalizedFio As String
    Dim statusTag As String

    Set outTable = Nothing
    columnNames = Array( _
        "65", "66", "67", "68", "69", "70", "71 (А)", "71 (Б)", _
        "72", "73", "74", "75 (А)", "75 (Б)", "76", _
        "Стовпець1", "Стовпець2", "Стовпець3")
    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = titleText
    For columnIndex = LBound(columnNames) To UBound(columnNames)
        If Not private_AddTableColumn(tableObj, _
            VBA.CStr(columnNames(columnIndex)), columnIndex + 1) Then Exit Function
    Next columnIndex

    For rowIndex = 1 To ejosData.RowCount
        If private_DateMatchesOrder( _
            ejosData.ValueAt(rowIndex, 7), orderDate) Then
            normalizedFio = private_NormalizeFio( _
                ejosData.ValueAt(rowIndex, 2))
            statusTag = VBA.vbNullString
            If currentOrderNames.Exists(normalizedFio) Then
                statusTag = STATUS_OK_TAG
            End If

            Set rowObj = New obj_Row
            For columnIndex = 1 To ejosData.ColumnCount
                If Not private_PushTaggedCell(rowObj, _
                    ejosData.ValueAt(rowIndex, columnIndex), _
                    statusTag) Then Exit Function
            Next columnIndex
            If Not tableObj.PushRow(rowObj) Then Exit Function
        End If
    Next rowIndex

    Set outTable = tableObj
    private_TryBuildEjosTable = True
End Function

Private Function private_TryBuildMedicalTable( _
    ByVal titleText As String, _
    ByVal medicalData As obj_TableData, _
    ByVal orderDate As Date, _
    ByVal currentOrderNames As Object, _
    ByVal latestMovementRowByFio As Object, _
    ByVal movementData As obj_TableData, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim columnNames As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim normalizedFio As String
    Dim statusTag As String
    Dim noteText As String
    Dim latestMovementRow As Long

    Set outTable = Nothing
    columnNames = Array( _
        "FIO", "IPN", "HospitalizationDate", "BasisDocName", _
        "BasisDocDate", "BasisDocNumber", "FullHospitalName", _
        "TreatmentStatus", NOTES_ALIAS)
    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = titleText
    For columnIndex = LBound(columnNames) To UBound(columnNames)
        If Not private_AddTableColumn(tableObj, _
            VBA.CStr(columnNames(columnIndex)), columnIndex + 1) Then Exit Function
    Next columnIndex

    For rowIndex = 1 To medicalData.RowCount
        If private_DateMatchesOrder( _
            medicalData.ValueAt(rowIndex, 3), orderDate) Then
            normalizedFio = private_NormalizeFio( _
                medicalData.ValueAt(rowIndex, 1))
            statusTag = VBA.vbNullString
            noteText = VBA.vbNullString
            If currentOrderNames.Exists(normalizedFio) Then
                statusTag = STATUS_OK_TAG
            ElseIf latestMovementRowByFio.Exists(normalizedFio) Then
                latestMovementRow = VBA.CLng( _
                    latestMovementRowByFio(normalizedFio))
                If private_TryBuildConflictNote( _
                    movementData, latestMovementRow, noteText) Then
                    statusTag = STATUS_CONFLICT_TAG
                End If
            End If

            Set rowObj = New obj_Row
            For columnIndex = 1 To medicalData.ColumnCount
                If Not private_PushTaggedCell(rowObj, _
                    medicalData.ValueAt(rowIndex, columnIndex), _
                    statusTag) Then Exit Function
            Next columnIndex
            If Not private_PushTaggedCell(rowObj, noteText, statusTag) Then Exit Function
            If Not tableObj.PushRow(rowObj) Then Exit Function
        End If
    Next rowIndex

    Set outTable = tableObj
    private_TryBuildMedicalTable = True
End Function

Private Function private_TryBuildConflictNote( _
    ByVal movementData As obj_TableData, _
    ByVal rowIndex As Long, _
    ByRef outNoteText As String _
) As Boolean
    Dim eventText As String
    Dim outDateText As String
    Dim plannedReturnText As String
    Dim returnDateText As String

    outNoteText = VBA.vbNullString
    If movementData Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > movementData.RowCount Then Exit Function
    eventText = VBA.Trim$(movementData.ValueAt(rowIndex, 2))
    outDateText = private_FormatMovementDateText( _
        movementData.ValueAt(rowIndex, 4))
    plannedReturnText = private_FormatMovementDateText( _
        movementData.ValueAt(rowIndex, 5))
    returnDateText = private_FormatMovementDateText( _
        movementData.ValueAt(rowIndex, 6))
    If VBA.Len(returnDateText) > 0 Then Exit Function

    Select Case VBA.LCase$(eventText)
        Case VBA.LCase$("Стаціонарне лікування")
            outNoteText = "Уже лечится с " & outDateText
        Case VBA.LCase$("Відпустка для лікування")
            outNoteText = "В отпуске для лечения с " & outDateText
            If VBA.Len(plannedReturnText) > 0 Then
                outNoteText = outNoteText & " по " & plannedReturnText
            End If
        Case Else
            Exit Function
    End Select
    private_TryBuildConflictNote = True
End Function

Private Function private_FormatMovementDateText( _
    ByVal rawDateValue As Variant _
) As String
    Dim rawDateText As String
    Dim serialDate As Double

    If VBA.IsError(rawDateValue) Or VBA.IsNull(rawDateValue) Or _
        VBA.IsEmpty(rawDateValue) Then Exit Function

    rawDateText = VBA.Trim$(VBA.CStr(rawDateValue))
    If VBA.Len(rawDateText) = 0 Then Exit Function

    ' Формульная ячейка Movement через ADO может вернуть не отображаемую
    ' дату, а её числовой Excel serial, например 46236.
    If VBA.IsNumeric(rawDateValue) Then
        serialDate = VBA.CDbl(rawDateValue)
        If serialDate >= 1 And serialDate <= 2958465 Then
            private_FormatMovementDateText = VBA.Format$( _
                VBA.CDate(serialDate), "dd.mm.yyyy")
            Exit Function
        End If
    End If

    private_FormatMovementDateText = rawDateText
End Function

Private Function private_PushTaggedCell( _
    ByVal rowObj As obj_Row, _
    ByVal valueText As String, _
    ByVal tagName As String _
) As Boolean
    Dim cellObj As obj_Cell

    Set cellObj = New obj_Cell
    cellObj.Value = valueText
    If VBA.Len(VBA.Trim$(tagName)) > 0 Then
        If Not cellObj.AddTag(tagName) Then Exit Function
    End If
    private_PushTaggedCell = rowObj.PushCell(cellObj)
End Function

Private Function private_DateMatchesOrder( _
    ByVal rawDateText As String, _
    ByVal orderDate As Date _
) As Boolean
    Dim actualDate As Date
    Dim expectedPrefix As String

    rawDateText = VBA.Trim$(rawDateText)
    expectedPrefix = VBA.Format$(orderDate, "dd.mm.yyyy")
    If VBA.StrComp(VBA.Left$(rawDateText, VBA.Len(expectedPrefix)), _
        expectedPrefix, VBA.vbTextCompare) = 0 Then
        private_DateMatchesOrder = True
        Exit Function
    End If
    If ex_Helpers.fn_TryResolveDateWithContext( _
        rawDateText, orderDate, actualDate) Then
        private_DateMatchesOrder = _
            (VBA.DateValue(actualDate) = VBA.DateValue(orderDate))
    End If
End Function

Private Function private_TextIsInList( _
    ByVal valueText As String, _
    ByVal allowedValues As Collection _
) As Boolean
    Dim item As Variant

    For Each item In allowedValues
        If VBA.StrComp(VBA.Trim$(valueText), _
            VBA.Trim$(VBA.CStr(item)), VBA.vbTextCompare) = 0 Then
            private_TextIsInList = True
            Exit Function
        End If
    Next item
End Function

Private Function private_NormalizeFio(ByVal fioText As String) As String
    fioText = VBA.Replace$(fioText, VBA.ChrW$(160), " ")
    fioText = VBA.Replace$(fioText, "’", "'")
    fioText = VBA.Replace$(fioText, "ʼ", "'")
    fioText = VBA.Trim$(fioText)
    Do While VBA.InStr(1, fioText, "  ", VBA.vbBinaryCompare) > 0
        fioText = VBA.Replace$(fioText, "  ", " ")
    Loop
    private_NormalizeFio = VBA.LCase$(fioText)
End Function

Private Function private_NormalizeOrderNo(ByVal orderNoText As String) As String
    orderNoText = VBA.Replace$(orderNoText, "№", "")
    orderNoText = VBA.Replace$(orderNoText, " ", "")
    private_NormalizeOrderNo = VBA.LCase$(VBA.Trim$(orderNoText))
End Function

Private Sub private_ShowError(ByVal messageText As String)
    VBA.MsgBox messageText, VBA.vbExclamation, ERROR_TITLE
End Sub
