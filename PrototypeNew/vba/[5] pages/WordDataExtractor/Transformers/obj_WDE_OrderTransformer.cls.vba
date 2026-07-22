VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WDE_OrderTransformer"
Option Explicit

Implements obj_ITableTransformer

Private Const FIO_COLUMN_ALIAS As String = "fio"
Private Const IPN_COLUMN_ALIAS As String = "ipn"
Private Const RANK_COLUMN_ALIAS As String = "rank"
Private Const UNRESOLVED_FIO_TAG As String = "fio-not-normalized"
Private Const CONFIG_STATE_PATH As String = "Source.Personnel.FilePath"
Private Const CONFIG_STATE_RANGE As String = "Personnel.Sheet[StateMain].SheetName"
Private Const CONFIG_STATE_FIO_HEADER As String = "Personnel.Sheet[StateMain].Map[FIO]"
Private Const CONFIG_STATE_IPN_HEADER As String = "Personnel.Sheet[StateMain].Map[IPN]"

Private m_IsDisposed As Boolean
Private m_ResourcesReady As Boolean
Private m_StatePath As String
Private m_StateTableRef As String
Private m_StateFioHeader As String
Private m_StateIpnHeader As String
Private m_StateEngine As obj_ExtWorkbookQueryEngine
Private m_CommonData As obj_PEB_ExptrCommonDataPrvdr

Private Function obj_ITableTransformer_Transform( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal configTable As obj_ConfigTable, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim targetTable As obj_TableDynamic
    Dim fioColumnIndex As Long, ipnColumnIndex As Long, rankColumnIndex As Long
    Dim hasIpnColumn As Boolean, hasRankColumn As Boolean

    Set outTable = Nothing
    If m_IsDisposed Then Exit Function
    If sourceTable Is Nothing Or configTable Is Nothing Then Exit Function
    If Not sourceTable.TryGetColumnIndexByAlias( _
        FIO_COLUMN_ALIAS, fioColumnIndex) Then
        ' Transformer нормализует только кадровые таблицы. Служебные таблицы
        ' pipeline (например, "Дати") не содержат ФИО и проходят без изменений.
        Set outTable = sourceTable
        obj_ITableTransformer_Transform = True
        Exit Function
    End If
    ' Optional datasets без строк не требуют ни клонирования, ни открытия
    ' State/АЛФ. Исходная модель неизменяема для последующих стадий pipeline.
    If sourceTable.RowCount = 0 Then
        Set outTable = sourceTable
        obj_ITableTransformer_Transform = True
        Exit Function
    End If
    hasIpnColumn = sourceTable.TryGetColumnIndexByAlias( _
        IPN_COLUMN_ALIAS, ipnColumnIndex)
    hasRankColumn = sourceTable.TryGetColumnIndexByAlias( _
        RANK_COLUMN_ALIAS, rankColumnIndex)

    ' Один экземпляр transformer обслуживает все таблицы extraction, поэтому
    ' открытые подключения переиспользуются до единственного Dispose.
    If Not private_TryEnsureResources(configTable) Then Exit Function

    If Not private_TryCloneTableStructure(sourceTable, targetTable) Then Exit Function
    If Not private_TryTransformRows(sourceTable, targetTable, _
        fioColumnIndex, ipnColumnIndex, hasIpnColumn, _
        rankColumnIndex, hasRankColumn) Then Exit Function

    Set outTable = targetTable
    obj_ITableTransformer_Transform = True
End Function

Private Function private_TryEnsureResources( _
    ByVal configTable As obj_ConfigTable _
) As Boolean
    If m_ResourcesReady Then
        private_TryEnsureResources = True
        Exit Function
    End If

    If Not private_TryLoadStateConfig(configTable, m_StatePath, _
        m_StateTableRef, m_StateFioHeader, m_StateIpnHeader) Then Exit Function
    m_ResourcesReady = True
    private_TryEnsureResources = True
End Function

Private Function private_TryEnsureStateEngine() As Boolean
    If Not m_StateEngine Is Nothing Then
        private_TryEnsureStateEngine = True
        Exit Function
    End If
    Set m_StateEngine = New obj_ExtWorkbookQueryEngine
    If Not m_StateEngine.Initialize Then
        Set m_StateEngine = Nothing
        Exit Function
    End If
    private_TryEnsureStateEngine = True
End Function

Private Function private_TryEnsureCommonData() As Boolean
    If Not m_CommonData Is Nothing Then
        private_TryEnsureCommonData = True
        Exit Function
    End If
    Set m_CommonData = New obj_PEB_ExptrCommonDataPrvdr
    If Not m_CommonData.Initialize() Then
        Set m_CommonData = Nothing
        Exit Function
    End If
    private_TryEnsureCommonData = True
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
    ByVal rankColumnIndex As Long, _
    ByVal hasRankColumn As Boolean _
) As Boolean
    Dim sourceRow As obj_Row, targetRow As obj_Row
    Dim fioDeclined As String, fioDefault As String, ipnText As String
    Dim rankDeclined As String, rankDefault As String
    Dim personFound As Boolean, rankFound As Boolean
    Dim i As Long

    For i = 1 To sourceTable.RowCount
        Set sourceRow = sourceTable.Rows.Item(i)
        If sourceRow Is Nothing Then Exit Function
        fioDeclined = VBA.Trim$(VBA.CStr( _
            sourceRow.GetCellValue(fioColumnIndex)))
        fioDeclined = private_RemoveRankServiceSuffixFromFio(fioDeclined)
        ipnText = VBA.vbNullString
        If hasIpnColumn Then ipnText = VBA.Trim$(VBA.CStr( _
            sourceRow.GetCellValue(ipnColumnIndex)))

        personFound = False
        fioDefault = VBA.vbNullString
        ' ИПН — устойчивый уникальный ключ, поэтому при его наличии ШПО не
        ' используется. Без ИПН остаётся обратный поиск склонённой формы в АЛФ.
        If VBA.Len(ipnText) > 0 Then
            If Not private_TryEnsureStateEngine() Then Exit Function
            If Not private_TryFindStateFioByIpn(m_StateEngine, m_StatePath, _
                m_StateTableRef, m_StateIpnHeader, m_StateFioHeader, ipnText, _
                personFound, fioDefault) Then Exit Function
        Else
            If Not private_TryEnsureCommonData() Then Exit Function
            If Not m_CommonData.TryFindFioDefaultByDeclinedForm( _
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
        If hasRankColumn Then
            rankDeclined = VBA.Trim$(VBA.CStr( _
                sourceRow.GetCellValue(rankColumnIndex)))
            rankDefault = VBA.vbNullString
            rankFound = False
            If Not private_TryEnsureCommonData() Then Exit Function
            If Not m_CommonData.TryFindRankDefaultByDeclinedForm( _
                rankDeclined, rankFound, rankDefault) Then Exit Function
            If rankFound And VBA.Len(VBA.Trim$(rankDefault)) > 0 Then
                If Not targetRow.SetCellRaw( _
                    rankColumnIndex, rankDefault) Then Exit Function
            End If
        End If
        If Not targetTable.PushRow(targetRow) Then Exit Function
    Next i
    private_TryTransformRows = True
End Function

Private Function private_RemoveRankServiceSuffixFromFio( _
    ByVal fioText As String _
) As String
    Const MEDICAL_SERVICE_PREFIX As String = "медичної служби "

    fioText = VBA.Trim$(fioText)
    If VBA.StrComp(VBA.Left$(fioText, VBA.Len(MEDICAL_SERVICE_PREFIX)), _
        MEDICAL_SERVICE_PREFIX, vbTextCompare) = 0 Then
        fioText = VBA.Trim$(VBA.Mid$(fioText, _
            VBA.Len(MEDICAL_SERVICE_PREFIX) + 1))
    End If
    private_RemoveRankServiceSuffixFromFio = fioText
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
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_StateEngine Is Nothing Then m_StateEngine.Dispose
    If Not m_CommonData Is Nothing Then m_CommonData.Dispose
    Set m_StateEngine = Nothing
    Set m_CommonData = Nothing
    m_ResourcesReady = False
    On Error GoTo 0
End Sub

Private Sub private_ShowError(ByVal messageText As String)
    ex_Core.fn_Diagnostic_LogError _
        "WordDataExtractor/OrderTransformer: " & messageText
    VBA.MsgBox messageText, VBA.vbExclamation, _
        "PrototypeNew / WordDataExtractor transformer"
End Sub
