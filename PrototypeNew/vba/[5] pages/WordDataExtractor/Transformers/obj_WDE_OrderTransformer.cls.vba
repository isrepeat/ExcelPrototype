VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WDE_OrderTransformer"
Option Explicit

Implements obj_ITableTransformer

Private Const FIO_COLUMN_ALIAS As String = "fio"
Private Const IPN_COLUMN_ALIAS As String = "ipn"
Private Const UNRESOLVED_FIO_TAG As String = "fio-not-normalized"
Private Const CONFIG_STATE_PATH As String = "Source.Personnel.FilePath"
Private Const CONFIG_STATE_RANGE As String = "Personnel.Sheet[StateMain].SheetName"
Private Const CONFIG_STATE_FIO_HEADER As String = "Personnel.Sheet[StateMain].Map[FIO]"
Private Const CONFIG_STATE_IPN_HEADER As String = "Personnel.Sheet[StateMain].Map[IPN]"

Private m_IsDisposed As Boolean

Private Function obj_ITableTransformer_Transform( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal configTable As obj_ConfigTable, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim targetTable As obj_TableDynamic
    Dim fioColumnIndex As Long, ipnColumnIndex As Long
    Dim hasIpnColumn As Boolean
    Dim statePath As String, stateTableRef As String
    Dim stateFioHeader As String, stateIpnHeader As String
    Dim stateEngine As obj_ExtWorkbookQueryEngine
    Dim commonData As obj_PEB_ExptrCommonDataPrvdr

    Set outTable = Nothing
    If sourceTable Is Nothing Or configTable Is Nothing Then Exit Function
    If Not sourceTable.TryGetColumnIndexByAlias( _
        FIO_COLUMN_ALIAS, fioColumnIndex) Then
        private_ShowError "В таблице не найдена колонка с алиасом '" & _
            FIO_COLUMN_ALIAS & "'."
        Exit Function
    End If
    hasIpnColumn = sourceTable.TryGetColumnIndexByAlias( _
        IPN_COLUMN_ALIAS, ipnColumnIndex)

    ' State-конфигурация загружается один раз на таблицу, а query engine
    ' переиспользует соединение для всех её строк. Это существенно дешевле,
    ' чем открывать большой файл особового состава для каждого человека.
    If Not private_TryLoadStateConfig(configTable, statePath, _
        stateTableRef, stateFioHeader, stateIpnHeader) Then Exit Function
    Set stateEngine = New obj_ExtWorkbookQueryEngine
    If Not stateEngine.Initialize Then Exit Function
    Set commonData = New obj_PEB_ExptrCommonDataPrvdr
    If Not commonData.Initialize() Then
        stateEngine.Dispose
        Exit Function
    End If

    If Not private_TryCloneTableStructure(sourceTable, targetTable) Then GoTo CleanFail
    If Not private_TryTransformRows(sourceTable, targetTable, _
        fioColumnIndex, ipnColumnIndex, hasIpnColumn, stateEngine, _
        statePath, stateTableRef, stateIpnHeader, stateFioHeader, _
        commonData) Then GoTo CleanFail

    stateEngine.Dispose
    commonData.Dispose
    Set outTable = targetTable
    obj_ITableTransformer_Transform = True
    Exit Function

CleanFail:
    On Error Resume Next
    stateEngine.Dispose
    commonData.Dispose
    On Error GoTo 0
End Function

Private Function private_TryCloneTableStructure( _
    ByVal sourceTable As obj_TableDynamic, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim sourceColumn As obj_Column, targetColumn As obj_Column
    Dim aliasItem As Variant
    Dim i As Long

    Set outTable = New obj_TableDynamic
    outTable.SectionTitle = sourceTable.SectionTitle
    outTable.SourceAlias = sourceTable.SourceAlias
    outTable.SourceAliasTemplate = sourceTable.SourceAliasTemplate
    For i = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(i)
        If sourceColumn Is Nothing Then Exit Function
        Set targetColumn = New obj_Column
        targetColumn.Name = sourceColumn.Name
        targetColumn.Position = i
        targetColumn.FormatKind = sourceColumn.FormatKind
        For Each aliasItem In sourceColumn.Aliases
            If Not targetColumn.AddAlias(VBA.CStr(aliasItem)) Then Exit Function
        Next aliasItem
        If Not outTable.PushColumn(targetColumn) Then Exit Function
    Next i
    private_TryCloneTableStructure = True
End Function

Private Function private_TryTransformRows( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal targetTable As obj_TableDynamic, _
    ByVal fioColumnIndex As Long, _
    ByVal ipnColumnIndex As Long, _
    ByVal hasIpnColumn As Boolean, _
    ByVal stateEngine As obj_ExtWorkbookQueryEngine, _
    ByVal statePath As String, _
    ByVal stateTableRef As String, _
    ByVal stateIpnHeader As String, _
    ByVal stateFioHeader As String, _
    ByVal commonData As obj_PEB_ExptrCommonDataPrvdr _
) As Boolean
    Dim sourceRow As obj_Row, targetRow As obj_Row
    Dim fioDeclined As String, fioDefault As String, ipnText As String
    Dim personFound As Boolean
    Dim i As Long

    For i = 1 To sourceTable.RowCount
        Set sourceRow = sourceTable.Rows.Item(i)
        If sourceRow Is Nothing Then Exit Function
        fioDeclined = VBA.Trim$(VBA.CStr( _
            sourceRow.GetCellValue(fioColumnIndex)))
        ipnText = VBA.vbNullString
        If hasIpnColumn Then ipnText = VBA.Trim$(VBA.CStr( _
            sourceRow.GetCellValue(ipnColumnIndex)))

        personFound = False
        fioDefault = VBA.vbNullString
        ' ИПН — устойчивый уникальный ключ, поэтому при его наличии ШПО не
        ' используется. Без ИПН остаётся обратный поиск склонённой формы в АЛФ.
        If VBA.Len(ipnText) > 0 Then
            If Not private_TryFindStateFioByIpn(stateEngine, statePath, _
                stateTableRef, stateIpnHeader, stateFioHeader, ipnText, _
                personFound, fioDefault) Then Exit Function
        Else
            If Not commonData.TryFindFioDefaultByDeclinedForm( _
                fioDeclined, personFound, fioDefault) Then Exit Function
        End If
        ' Отсутствие человека не должно отменять extraction или уничтожать
        ' исходные данные. Оставляем склонённое ФИО и ставим семантический тег;
        ' конкретное оформление задаётся снаружи selector-ом TableList.
        If Not personFound Or VBA.Len(VBA.Trim$(fioDefault)) = 0 Then
            fioDefault = fioDeclined
            ex_Core.fn_Diagnostic_LogError _
                "WordDataExtractor/OrderTransformer: person-not-found ipn='" & _
                ipnText & "' fio='" & fioDeclined & "'"
        End If

        Set targetRow = sourceRow.Clone(sourceTable.ColumnCount)
        If targetRow Is Nothing Then Exit Function
        If personFound And VBA.Len(VBA.Trim$(fioDefault)) > 0 Then
            If Not targetRow.SetCellRaw( _
                fioColumnIndex, fioDefault) Then Exit Function
        Else
            If Not targetRow.SetCellRaw(fioColumnIndex, fioDefault) Then Exit Function
            If Not targetRow.AddCellTag( _
                fioColumnIndex, UNRESOLVED_FIO_TAG) Then Exit Function
        End If
        If Not targetTable.PushRow(targetRow) Then Exit Function
    Next i
    private_TryTransformRows = True
End Function

Private Function private_TryFindStateFioByIpn( _
    ByVal queryEngine As obj_ExtWorkbookQueryEngine, _
    ByVal statePath As String, _
    ByVal stateTableRef As String, _
    ByVal ipnHeader As String, _
    ByVal fioHeader As String, _
    ByVal ipnText As String, _
    ByRef outFound As Boolean, _
    ByRef outFio As String _
) As Boolean
    Dim query As obj_ExtWorkbookQuery
    Dim resultTable As obj_TableDynamic
    Dim resultRow As obj_Row

    outFound = False
    outFio = VBA.vbNullString
    Set query = New obj_ExtWorkbookQuery
    query.SourcePath = statePath
    query.TableRef = stateTableRef
    query.MaxRows = 2
    If Not query.AddSelectColumn(fioHeader) Then Exit Function
    If Not query.AddCondition(ipnHeader, _
        en_ExtWorkbookQueryOp.ExtQueryOpEquals, ipnText, True) Then Exit Function
    If Not queryEngine.TryExecute(query, resultTable) Then Exit Function
    If resultTable Is Nothing Then Exit Function
    If resultTable.RowCount = 0 Then
        private_TryFindStateFioByIpn = True
        Exit Function
    End If
    If resultTable.RowCount > 1 Then
        ' ИПН обязан быть уникальным. Выбор первой из нескольких строк скрыл бы
        ' повреждение State, поэтому такое ФИО считается ненормализованным.
        ex_Core.fn_Diagnostic_LogError _
            "WordDataExtractor/OrderTransformer: duplicate State rows for IPN '" & _
            ipnText & "'."
        private_TryFindStateFioByIpn = True
        Exit Function
    End If
    Set resultRow = resultTable.Rows.Item(1)
    If resultRow Is Nothing Then Exit Function
    If Not resultRow.TryGetCellValueByColumn(fioHeader, outFio) Then Exit Function
    outFio = VBA.Trim$(outFio)
    outFound = (VBA.Len(outFio) > 0)
    private_TryFindStateFioByIpn = True
End Function

Private Function private_TryLoadStateConfig( _
    ByVal configTable As obj_ConfigTable, _
    ByRef outStatePath As String, _
    ByRef outStateTableRef As String, _
    ByRef outFioHeader As String, _
    ByRef outIpnHeader As String _
) As Boolean
    Dim parser As obj_CfgParserBase
    Dim entries As Collection, cfgMap As Object
    Dim rawTableRef As String

    Set parser = New obj_CfgParserBase
    If Not parser.Initialize(configTable) Then Exit Function
    If Not parser.TryGetConfigEntries(entries) Then Exit Function
    If Not parser.BuildConfigDictionary(entries, cfgMap) Then Exit Function
    If Not parser.TryGetRequiredConfigValue( _
        cfgMap, CONFIG_STATE_PATH, outStatePath) Then Exit Function
    If Not parser.TryGetRequiredConfigValue( _
        cfgMap, CONFIG_STATE_RANGE, rawTableRef) Then Exit Function
    If Not parser.TryGetRequiredConfigValue( _
        cfgMap, CONFIG_STATE_FIO_HEADER, outFioHeader) Then Exit Function
    If Not parser.TryGetRequiredConfigValue( _
        cfgMap, CONFIG_STATE_IPN_HEADER, outIpnHeader) Then Exit Function

    outStatePath = private_ResolveWorkbookPath(outStatePath)
    If VBA.Len(outStatePath) = 0 Or VBA.Len(VBA.Dir$(outStatePath)) = 0 Then
        private_ShowError "Не найден файл State: " & outStatePath
        Exit Function
    End If
    ' QueryEngine принимает одну ADO-совместимую ссылку и одинаково разбирает
    ' её как для закрытой книги (ACE), так и для уже открытого Worksheet.
    rawTableRef = VBA.Replace(VBA.Trim$(rawTableRef), "]", "]]")
    outStateTableRef = "[" & rawTableRef & "]"
    private_TryLoadStateConfig = True
End Function

Private Function private_ResolveWorkbookPath( _
    ByVal workbookPath As String _
) As String
    workbookPath = VBA.Trim$(workbookPath)
    If VBA.InStr(1, workbookPath, ":", VBA.vbBinaryCompare) > 0 Or _
        VBA.Left$(workbookPath, 2) = "\\" Then
        private_ResolveWorkbookPath = workbookPath
    Else
        private_ResolveWorkbookPath = _
            ex_XmlCore.fn_CombineBasePath(ThisWorkbook, workbookPath)
    End If
End Function

Private Sub obj_ITableTransformer_Dispose()
    m_IsDisposed = True
End Sub

Private Sub private_ShowError(ByVal messageText As String)
    ex_Core.fn_Diagnostic_LogError _
        "WordDataExtractor/OrderTransformer: " & messageText
    VBA.MsgBox messageText, VBA.vbExclamation, _
        "PrototypeNew / WordDataExtractor transformer"
End Sub
