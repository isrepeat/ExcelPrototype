Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const OPENNING_EXTERNAL_SOURCES_TYPE = 2

Private Const DIAGNOSTIC_LOG_FILE_REL_PATH As String = "Logs\\templateprototype.log"
Private Const NO_QUERY_RESULTS_TEXT As String = "<No query results>"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const LISTS_SHEET As String = "Переліки"

' ModePayments: зашитые ключи key/value-конфига для источников.
Private Const cnst_CfgKey_StateBookPath As String = "Стан (шлях)"
Private Const cnst_CfgKey_StateTableName As String = "Стан (таблиця)"
Private Const cnst_CfgKey_PaymentsBookPath As String = "Виплати (шлях)"
Private Const cnst_CfgKey_PaymentsTableName As String = "Виплати (таблиця)"
Private Const cnst_CfgKey_RankBookPath As String = "Звання (шлях)"
Private Const cnst_CfgKey_RankTableName As String = "Звання (таблиця)"
Private Const cnst_CfgKey_PositionBookPath As String = "Посади (шлях)"
Private Const cnst_CfgKey_PositionTableName As String = "Посади (таблиця)"
Private Const cnst_ModePayents_CfgKey_ResultItemTemplate As String = "Результат (шаблон рядка)"

' ModePayments: локальные таблицы текущей книги.
Private Const cnst_ModePayents_AdditionalInfoTableName As String = "Виплати_додаткова_інформація"

' ModePayments: ссылки на параметры режима в ячейках.
Private Const cnst_ModePayents_InParam_PaymentTypeCell As String = "Виплати!C2"
Private Const cnst_ModePayents_InParam_OrderNumberCell As String = "Виплати!E2"
Private Const cnst_ModePayents_InParam_YearCell As String = "Виплати!G2"
Private Const cnst_ModePayents_ResultCell As String = "Виплати!C4"

' Дескриптор внешней книги для безопасного открытия/закрытия.
Private Type private_WorkbookHandle
    Workbook As Workbook
    OpenedByCode As Boolean
    WasWindowVisible As Boolean
End Type

' --------------------------------------
'  Public API
' --------------------------------------

' --------------------------------------
'  namespace ModePayments {
' --------------------------------------
' Ручная точка входа режима: назначается кнопке напрямую в Excel.
Public Sub fn_ModePayments_Run()
    Dim cfg As Object
    Dim modeParams As Object
    Dim stateSrc As private_WorkbookHandle
    Dim paymentsSrc As private_WorkbookHandle
    Dim rankSrc As private_WorkbookHandle
    Dim positionSrc As private_WorkbookHandle
    Dim stateLo As ListObject
    Dim paymentsLo As ListObject
    Dim rankLo As ListObject
    Dim positionLo As ListObject
    Dim additionalInfoLo As ListObject
    Dim resultText As String
    Dim paymentTypeText As String
    Dim orderNumberText As String
    Dim yearText As String
    Dim resultItemTemplate As String
    Dim stateBookPath As String
    Dim stateTableName As String
    Dim paymentsBookPath As String
    Dim paymentsTableName As String
    Dim rankBookPath As String
    Dim rankTableName As String
    Dim positionBookPath As String
    Dim positionTableName As String

    On Error GoTo ErrHandler

    ' Читаем параметры режима с листа: вид выплаты и номер приказа.
    Set modeParams = private_ModePayments_ParamsParser()
    paymentTypeText = private_Dict_RequireText( _
        modeParams, _
        "PaymentTypeText", _
        "Не найден параметр выплаты" _
    )
    orderNumberText = private_Dict_RequireText( _
        modeParams, _
        "OrderNumberText", _
        "Не найден параметр номера приказа" _
    )
    yearText = private_Dict_RequireText( _
        modeParams, _
        "YearText", _
        "Не найден параметр года" _
    )

    ' Читаем Settings: пути, имена ListObject-таблиц и шаблон строки результата.
    Set cfg = private_Config_LoadSettingsDict()
    stateBookPath = private_Config_RequireText(cfg, cnst_CfgKey_StateBookPath)
    stateTableName = private_Config_RequireText(cfg, cnst_CfgKey_StateTableName)
    paymentsBookPath = private_Config_RequireText(cfg, cnst_CfgKey_PaymentsBookPath)
    paymentsTableName = private_Config_RequireText(cfg, cnst_CfgKey_PaymentsTableName)
    rankBookPath = private_Config_RequireText(cfg, cnst_CfgKey_RankBookPath)
    rankTableName = private_Config_RequireText(cfg, cnst_CfgKey_RankTableName)
    positionBookPath = private_Config_RequireText(cfg, cnst_CfgKey_PositionBookPath)
    positionTableName = private_Config_RequireText(cfg, cnst_CfgKey_PositionTableName)
    resultItemTemplate = private_Config_RequireText(cfg, cnst_ModePayents_CfgKey_ResultItemTemplate)

#If OPENNING_EXTERNAL_SOURCES_TYPE = 1 Or OPENNING_EXTERNAL_SOURCES_TYPE = 2 Then
    ' Открываем внешние книги по выбранному препроцессорному режиму.
    stateSrc = private_Workbook_GetSafe(stateBookPath, True, True)
    paymentsSrc = private_Workbook_GetSafe(paymentsBookPath, True, True)
    rankSrc = private_Workbook_GetSafe(rankBookPath, True, True)
    positionSrc = private_Workbook_GetSafe(positionBookPath, True, True)

    ' После открытия работаем только с ListObject, без SQL/ADODB.
    Set stateLo = private_ListObject_FindByName(stateSrc.Workbook, stateTableName)
    Set paymentsLo = private_ListObject_FindByName(paymentsSrc.Workbook, paymentsTableName)
    Set rankLo = private_ListObject_FindByName(rankSrc.Workbook, rankTableName)
    Set positionLo = private_ListObject_FindByName(positionSrc.Workbook, positionTableName)
#Else
    VBA.Err.Raise VBA.vbObjectError + 4900, , "Неподдерживаемый OPENNING_EXTERNAL_SOURCES_TYPE."
#End If

    Set additionalInfoLo = private_ListObject_FindByNameOnSheet( _
        LISTS_SHEET, _
        cnst_ModePayents_AdditionalInfoTableName _
    )

    ' Собираем итоговый текст и пишем его в объединенную область результата.
    resultText = private_ModePayments_BuildResultText( _
        paymentsLo, _
        stateLo, _
        rankLo, _
        positionLo, _
        additionalInfoLo, _
        paymentTypeText, _
        orderNumberText, _
        yearText, _
        resultItemTemplate _
    )

    private_CellRef_WriteText cnst_ModePayents_ResultCell, resultText

CleanExit:
    ' Тип 1 закрывает источники; тип 2 оставляет их в Excel-сессии.
#If OPENNING_EXTERNAL_SOURCES_TYPE = 1 Then
    private_Workbook_CloseSafe stateSrc, False
    private_Workbook_CloseSafe paymentsSrc, False
    private_Workbook_CloseSafe rankSrc, False
    private_Workbook_CloseSafe positionSrc, False
#ElseIf OPENNING_EXTERNAL_SOURCES_TYPE = 2 Then
    private_Workbook_ReleaseHandle stateSrc
    private_Workbook_ReleaseHandle paymentsSrc
    private_Workbook_ReleaseHandle rankSrc
    private_Workbook_ReleaseHandle positionSrc
#End If
    Exit Sub

ErrHandler:
    fn_Diagnostic_LogError "fn_ModePayments_Run: " & VBA.Err.Description
    VBA.MsgBox _
        "An error occurred while running ModePayments." & VBA.vbCrLf & _
        "See diagnostic log for details." & VBA.vbCrLf & _
        "Error code: " & VBA.CStr(VBA.Err.Number), _
        VBA.vbCritical
    Resume CleanExit
End Sub
' --------------------------------------
'  } // namespace ModePayments
' --------------------------------------


' --------------------------------------
'  namespace Diagnostic {
' --------------------------------------
Public Sub fn_Diagnostic_LogInfo(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_Write messageText
#End If
End Sub


Public Sub fn_Diagnostic_LogError(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_Write "error: " & messageText
#End If
End Sub


Public Sub fn_Diagnostic_LogWarning(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_Write "warning: " & messageText
#End If
End Sub


Public Sub fn_Diagnostic_ClearLog()
    private_Diagnostic_ClearFile
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------


' //
' // Internal
' //
' --------------------------------------
'  namespace Config {
' --------------------------------------
Private Function private_Config_LoadSettingsDict() As Object
    Dim cfg As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim k As String
    Dim v As String

    Set cfg = private_Dict_CreateTextMap()
    Set ws = private_Worksheet_GetByNameOrFail(SETTINGS_SHEET, "лист конфигурации")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        k = VBA.Trim$(VBA.CStr(ws.Cells(r, "A").Value2))
        v = VBA.Trim$(VBA.CStr(ws.Cells(r, "B").Value2))
        If VBA.Len(k) = 0 Then GoTo ContinueRow
        If cfg.Exists(k) Then
            VBA.Err.Raise VBA.vbObjectError + 4100, , "Дублирующийся ключ конфига: " & k
        End If
        cfg.Add k, v
ContinueRow:
    Next r

    Set private_Config_LoadSettingsDict = cfg
End Function


Private Function private_Config_RequireText(ByVal cfg As Object, ByVal keyName As String) As String
    If cfg Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 4101, , "Конфиг не загружен."
    End If

    If Not cfg.Exists(keyName) Then
        VBA.Err.Raise VBA.vbObjectError + 4102, , "Не найден ключ конфига: " & keyName
    End If

    private_Config_RequireText = VBA.Trim$(VBA.CStr(cfg(keyName)))
    If VBA.Len(private_Config_RequireText) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 4103, , "Пустое значение ключа конфига: " & keyName
    End If
End Function
' --------------------------------------
'  } // namespace Config
' --------------------------------------


' --------------------------------------
'  namespace ModePayments {
' --------------------------------------
Private Function private_ModePayments_ParamsParser() As Object
    Dim paramsDict As Object

    ' UI-параметры режима живут на листе Виплати, а не в Settings.
    Set paramsDict = private_Dict_CreateTextMap()
    paramsDict("PaymentTypeText") = private_CellRef_ReadRequiredText(cnst_ModePayents_InParam_PaymentTypeCell, "Виплата")
    paramsDict("OrderNumberText") = private_CellRef_ReadRequiredText(cnst_ModePayents_InParam_OrderNumberCell, "Номер наказу")
    paramsDict("YearText") = private_CellRef_ReadRequiredText(cnst_ModePayents_InParam_YearCell, "Рік")
    Set private_ModePayments_ParamsParser = paramsDict
End Function


Private Function private_ModePayments_BuildResultText( _
    ByVal paymentsLo As ListObject, _
    ByVal stateLo As ListObject, _
    ByVal rankLo As ListObject, _
    ByVal positionLo As ListObject, _
    ByVal additionalInfoLo As ListObject, _
    ByVal paymentTypeText As String, _
    ByVal orderNumberText As String, _
    ByVal yearText As String, _
    ByVal resultItemTemplate As String _
) As String
    Dim stateDict As Object
    Dim rankDict As Object
    Dim positionDict As Object
    Dim stateRowDict As Object
    Dim rankRowDict As Object
    Dim paymentValues As Variant
    Dim r As Long
    Dim taxId As String
    Dim rankText As String
    Dim positionCode As String
    Dim rowPaymentType As String
    Dim rowOrderNumber As String
    Dim headerText As String
    Dim footerText As String
    Dim resultText As String

    ' Проверяем, что источники вообще загружены как ListObject.
    private_ListObject_RequireLoaded paymentsLo, "Payments"
    private_ListObject_RequireLoaded stateLo, "State"
    private_ListObject_RequireLoaded rankLo, "Rank"
    private_ListObject_RequireLoaded positionLo, "Position"
    private_ListObject_RequireLoaded additionalInfoLo, "AdditionalInfo"

    If paymentsLo.DataBodyRange Is Nothing Then
        private_ModePayments_BuildResultText = NO_QUERY_RESULTS_TEXT
        Exit Function
    End If

    ' Справочники разворачиваем в словари для lookup по ключам.
    Set stateDict = private_ExternalSources_BuildStateDict(stateLo)
    Set rankDict = private_ExternalSources_BuildRankDict(rankLo)
    Set positionDict = private_ExternalSources_BuildPositionDict(positionLo)

    headerText = private_ModePayments_BuildResultHeaderText( _
        additionalInfoLo, _
        paymentTypeText, _
        yearText _
    )

    For r = 1 To paymentsLo.DataBodyRange.Rows.Count
        paymentValues = paymentsLo.DataBodyRange.Rows(r).Value2

        ' Зашитые поля Payments для фильтра: Вид, Наказ.
        rowPaymentType = private_Text_NormalizeLookupToken( _
            private_ListRow_FieldText( _
                paymentsLo, _
                paymentValues, _
                "Вид" _
            ) _
        )
        rowOrderNumber = private_Text_NormalizeLookupToken( _
            private_ListRow_FieldText( _
                paymentsLo, _
                paymentValues, _
                "Наказ" _
            ) _
        )

        If VBA.StrComp(rowPaymentType, private_Text_NormalizeLookupToken(paymentTypeText), VBA.vbTextCompare) = 0 _
           And VBA.StrComp(rowOrderNumber, private_Text_NormalizeLookupToken(orderNumberText), VBA.vbTextCompare) = 0 Then

            ' Зашитые поля Payments для результата: ІПН, Звання, Посада.
            taxId = private_Text_Require( _
                private_ListRow_FieldText( _
                    paymentsLo, _
                    paymentValues, _
                    "ІПН" _
                ), _
                "Payments row has empty TaxId." _
            )
            rankText = private_Text_NormalizeLookupToken( _
                private_Text_Require( _
                    private_ListRow_FieldText( _
                        paymentsLo, _
                        paymentValues, _
                        "Звання" _
                    ), _
                    "Payments row has empty Rank." _
                ) _
            )
            positionCode = private_Text_Require( _
                private_ListRow_FieldText( _
                    paymentsLo, _
                    paymentValues, _
                    "Посада" _
                ), _
                "Payments row has empty PositionCode." _
            )

            If Not stateDict.Exists(taxId) Then
                VBA.Err.Raise VBA.vbObjectError + 5101, , "State row not found for TaxId: " & taxId
            End If

            If Not rankDict.Exists(rankText) Then
                VBA.Err.Raise VBA.vbObjectError + 5102, , "Rank declension row not found for Rank: " & rankText
            End If

            If Not positionDict.Exists(positionCode) Then
                VBA.Err.Raise VBA.vbObjectError + 5103, , "Position declension row not found for PositionCode: " & positionCode
            End If

            Set stateRowDict = stateDict(taxId)
            Set rankRowDict = rankDict(rankText)
            If VBA.Len(resultText) > 0 Then resultText = resultText & VBA.vbCrLf & VBA.vbCrLf

            resultText = resultText & private_ModePayments_BuildResultItemText( _
                paymentsLo, _
                paymentValues, _
                stateRowDict, _
                rankRowDict, _
                positionDict(positionCode), _
                taxId, _
                resultItemTemplate _
            )
        End If
    Next r

    footerText = private_ModePayments_BuildResultFooterText()

    If VBA.Len(resultText) = 0 Then
        private_ModePayments_BuildResultText = NO_QUERY_RESULTS_TEXT
    Else
        If VBA.Len(headerText) > 0 Then resultText = headerText & VBA.vbCrLf & VBA.vbCrLf & resultText
        If VBA.Len(footerText) > 0 Then resultText = resultText & VBA.vbCrLf & VBA.vbCrLf & footerText
        private_ModePayments_BuildResultText = resultText
    End If
End Function


Private Function private_ModePayments_BuildResultHeaderText( _
    ByVal additionalInfoLo As ListObject, _
    ByVal paymentTypeText As String, _
    ByVal yearText As String _
) As String
    Dim values As Variant
    Dim r As Long
    Dim rowPaymentType As String
    Dim headerTemplate As String

    If additionalInfoLo.DataBodyRange Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 5150, , "AdditionalInfo table returned no rows."
    End If

    For r = 1 To additionalInfoLo.DataBodyRange.Rows.Count
        values = additionalInfoLo.DataBodyRange.Rows(r).Value2
        rowPaymentType = private_Text_NormalizeLookupToken( _
            private_ListRow_FieldText( _
                additionalInfoLo, _
                values, _
                "Виплата" _
            ) _
        )

        If VBA.StrComp(rowPaymentType, private_Text_NormalizeLookupToken(paymentTypeText), VBA.vbTextCompare) = 0 Then
            headerTemplate = private_Text_Require( _
                private_ListRow_FieldText( _
                    additionalInfoLo, _
                    values, _
                    "Заголовок" _
                ), _
                "AdditionalInfo row has empty Header." _
            )
            private_ModePayments_BuildResultHeaderText = private_ModePayments_ResolveHeaderText( _
                headerTemplate, _
                yearText _
            )
            Exit Function
        End If
    Next r

    VBA.Err.Raise VBA.vbObjectError + 5151, , "Header row not found for PaymentType: " & paymentTypeText
End Function


Private Function private_ModePayments_BuildResultFooterText() As String
    private_ModePayments_BuildResultFooterText = ""
End Function


Private Function private_ModePayments_ResolveHeaderText( _
    ByVal headerTemplate As String, _
    ByVal yearText As String _
) As String
    headerTemplate = VBA.Replace(headerTemplate, "{YEAR}", yearText)
    private_ModePayments_ResolveHeaderText = headerTemplate
End Function


Private Function private_ModePayments_BuildResultItemText( _
    ByVal paymentsLo As ListObject, _
    ByVal paymentValues As Variant, _
    ByVal stateRowDict As Object, _
    ByVal rankRowDict As Object, _
    ByVal positionDative As String, _
    ByVal taxId As String, _
    ByVal resultItemTemplate As String _
) As String
    Dim reportText As String
    Dim reportDateText As String
    Dim orderText As String

    reportText = private_Text_Require( _
        private_ListRow_FieldText( _
            paymentsLo, _
            paymentValues, _
            "Рапорт" _
        ), _
        "Payments row has empty Report." _
    )
    reportDateText = private_Text_Require( _
        private_Text_FormatDisplayDate( _
            private_ListRow_FieldValue( _
                paymentsLo, _
                paymentValues, _
                "Дата" _
            ) _
        ), _
        "Payments row has empty ReportDate." _
    )

    ' Первая строка результата настраивается через Settings-шаблон.
    orderText = private_ModePayments_BuildPersonalTextWithTemplate( _
        resultItemTemplate, _
        private_Text_CapitalizeFirst(rankRowDict("RankDative")), _
        stateRowDict("DativeName"), _
        taxId, _
        private_Text_LowerFirst(positionDative) _
    )

    ' Основание пока остается фиксированным.
    orderText = orderText & _
                VBA.vbCrLf & _
                private_ModePayments_BuildBasisText( _
                    rankRowDict("RankGenitive"), _
                    stateRowDict("GenitiveNameShortText"), _
                    reportText, _
                    reportDateText _
                )

    private_ModePayments_BuildResultItemText = orderText
End Function


Private Function private_ModePayments_BuildPersonalTextWithTemplate( _
    ByVal resultItemTemplate As String, _
    ByVal rankText As String, _
    ByVal fioText As String, _
    ByVal taxId As String, _
    ByVal positionText As String _
) As String
    ' Поддерживаемые плейсхолдеры: {Rank}, {FIO}, {IPN}, {Position}.
    resultItemTemplate = private_Text_Require( _
        resultItemTemplate, _
        "Result item template is empty." _
    )
    resultItemTemplate = VBA.Replace(resultItemTemplate, "{Rank}", rankText)
    resultItemTemplate = VBA.Replace(resultItemTemplate, "{FIO}", fioText)
    resultItemTemplate = VBA.Replace(resultItemTemplate, "{IPN}", taxId)
    resultItemTemplate = VBA.Replace(resultItemTemplate, "{Position}", positionText)

    private_ModePayments_BuildPersonalTextWithTemplate = resultItemTemplate
End Function


Private Function private_ModePayments_BuildBasisText( _
    ByVal rankGenitive As String, _
    ByVal genitiveNameShortText As String, _
    ByVal reportText As String, _
    ByVal reportDateText As String _
) As String
    ' Формат основания намеренно оставлен фиксированным: "Підстава: рапорт ...".
    private_ModePayments_BuildBasisText = private_Text_UaPidstavaRaport() & _
                                          private_Text_LowerFirst(rankGenitive) & " " & _
                                          genitiveNameShortText & " (" & _
                                          private_Text_UaIncomingNo() & reportText & _
                                          private_Text_UaFrom() & reportDateText & ")."
End Function
' --------------------------------------
'  } // namespace ModePayments
' --------------------------------------


' --------------------------------------
'  namespace ExternalSources {
' --------------------------------------
Private Function private_ExternalSources_BuildStateDict(ByVal stateLo As ListObject) As Object
    Dim stateDict As Object
    Dim rowDict As Object
    Dim values As Variant
    Dim r As Long
    Dim taxId As String
    Dim dativeName As String
    Dim genitiveName As String
    Dim genitiveNameShortText As String

    If stateLo.DataBodyRange Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 5110, , "State query returned no rows."
    End If

    Set stateDict = private_Dict_CreateTextMap()

    For r = 1 To stateLo.DataBodyRange.Rows.Count
        values = stateLo.DataBodyRange.Rows(r).Value2
        taxId = private_ListRow_FieldText( _
            stateLo, _
            values, _
            "ІПН" _
        )
        dativeName = private_ListRow_FieldText( _
            stateLo, _
            values, _
            "Давальний" _
        )
        genitiveName = private_ListRow_FieldText( _
            stateLo, _
            values, _
            "Родовий" _
        )
        genitiveNameShortText = private_Text_ToShortName(genitiveName)

        If VBA.Len(taxId) > 0 And VBA.Len(dativeName) > 0 And VBA.Len(genitiveNameShortText) > 0 Then
            Set rowDict = private_Dict_CreateTextMap()
            rowDict("DativeName") = dativeName
            rowDict("GenitiveNameShortText") = genitiveNameShortText
            If stateDict.Exists(taxId) Then stateDict.Remove taxId
            stateDict.Add taxId, rowDict
        End If
    Next r

    If stateDict.Count = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 5111, , "State query returned no usable TaxId/DativeName/GenitiveName rows."
    End If

    Set private_ExternalSources_BuildStateDict = stateDict
End Function


Private Function private_ExternalSources_BuildRankDict(ByVal rankLo As ListObject) As Object
    Dim rankDict As Object
    Dim rowDict As Object
    Dim values As Variant
    Dim r As Long
    Dim rankText As String
    Dim rankGenitive As String
    Dim rankDative As String

    If rankLo.DataBodyRange Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 5120, , "RankQuery returned no rows."
    End If

    Set rankDict = private_Dict_CreateTextMap()

    For r = 1 To rankLo.DataBodyRange.Rows.Count
        values = rankLo.DataBodyRange.Rows(r).Value2
        rankText = private_Text_NormalizeLookupToken( _
            private_ListRow_FieldText( _
                rankLo, _
                values, _
                "Звання" _
            ) _
        )
        rankGenitive = private_ListRow_FieldText( _
            rankLo, _
            values, _
            "Родовий" _
        )
        rankDative = private_ListRow_FieldText( _
            rankLo, _
            values, _
            "Давальний" _
        )
        If VBA.Len(rankText) > 0 And VBA.Len(rankGenitive) > 0 And VBA.Len(rankDative) > 0 Then
            Set rowDict = private_Dict_CreateTextMap()
            rowDict("RankGenitive") = rankGenitive
            rowDict("RankDative") = rankDative
            If rankDict.Exists(rankText) Then rankDict.Remove rankText
            rankDict.Add rankText, rowDict
        End If
    Next r

    If rankDict.Count = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 5121, , "RankQuery returned no usable Rank/RankGenitive/RankDative rows."
    End If

    Set private_ExternalSources_BuildRankDict = rankDict
End Function


Private Function private_ExternalSources_BuildPositionDict(ByVal positionLo As ListObject) As Object
    Dim positionDict As Object
    Dim values As Variant
    Dim r As Long
    Dim positionCode As String
    Dim positionDative As String

    If positionLo.DataBodyRange Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 5130, , "PositionQuery returned no rows."
    End If

    Set positionDict = private_Dict_CreateTextMap()

    For r = 1 To positionLo.DataBodyRange.Rows.Count
        values = positionLo.DataBodyRange.Rows(r).Value2
        positionCode = private_ListRow_FieldText( _
            positionLo, _
            values, _
            "Код" _
        )
        positionDative = private_ListRow_FieldText( _
            positionLo, _
            values, _
            "Давальний" _
        )
        If VBA.Len(positionCode) > 0 And VBA.Len(positionDative) > 0 Then
            positionDict(positionCode) = positionDative
        End If
    Next r

    If positionDict.Count = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 5131, , "PositionQuery returned no usable PositionCode/PositionDative rows."
    End If

    Set private_ExternalSources_BuildPositionDict = positionDict
End Function
' --------------------------------------
'  } // namespace ExternalSources
' --------------------------------------


' --------------------------------------
'  namespace ListObject {
' --------------------------------------
Private Function private_ListObject_FindByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    tableName = VBA.Trim$(tableName)
    If VBA.Len(tableName) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 4800, , "Имя таблицы пустое."
    End If

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If VBA.StrComp(lo.Name, tableName, VBA.vbTextCompare) = 0 Then
                Set private_ListObject_FindByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    VBA.Err.Raise VBA.vbObjectError + 4801, , "ListObject не найден: " & tableName
End Function


Private Function private_ListObject_FindByNameOnSheet( _
    ByVal sheetName As String, _
    ByVal tableName As String _
) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    tableName = VBA.Trim$(tableName)
    If VBA.Len(tableName) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 4806, , "ListObject name is empty."
    End If

    Set ws = private_Worksheet_GetByNameOrFail(sheetName, "ListObject host sheet")

    For Each lo In ws.ListObjects
        If VBA.StrComp(lo.Name, tableName, VBA.vbTextCompare) = 0 Then
            Set private_ListObject_FindByNameOnSheet = lo
            Exit Function
        End If
    Next lo

    VBA.Err.Raise VBA.vbObjectError + 4807, , "ListObject not found on sheet '" & ws.Name & "': " & tableName
End Function


Private Sub private_ListObject_RequireLoaded(ByVal lo As ListObject, ByVal tableAlias As String)
    If lo Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 4802, , "ListObject не загружен: " & tableAlias
    End If
End Sub


Private Function private_ListRow_FieldText( _
    ByVal lo As ListObject, _
    ByVal rowValues As Variant, _
    ByVal fieldName As String _
) As String
    Dim columnIndex As Long

    columnIndex = private_ListObject_ColumnIndex(lo, fieldName)
    If columnIndex <= 0 Then
        fn_Diagnostic_LogError "Не найдена колонка в ListObject '" & lo.Name & "': " & fieldName
        VBA.Err.Raise VBA.vbObjectError + 4804, , "Обязательная колонка не найдена в ListObject '" & lo.Name & "': " & fieldName
    End If

    private_ListRow_FieldText = VBA.Trim$(private_Text_NullToEmptyString(rowValues(1, columnIndex)))
End Function


Private Function private_ListRow_FieldValue( _
    ByVal lo As ListObject, _
    ByVal rowValues As Variant, _
    ByVal fieldName As String _
) As Variant
    Dim columnIndex As Long

    columnIndex = private_ListObject_ColumnIndex(lo, fieldName)
    If columnIndex <= 0 Then
        fn_Diagnostic_LogError "Не найдена колонка в ListObject '" & lo.Name & "': " & fieldName
        VBA.Err.Raise VBA.vbObjectError + 4805, , "Обязательная колонка не найдена в ListObject '" & lo.Name & "': " & fieldName
    End If

    private_ListRow_FieldValue = rowValues(1, columnIndex)
End Function


Private Function private_ListObject_ColumnIndex(ByVal lo As ListObject, ByVal fieldName As String) As Long
    Dim c As Long

    fieldName = private_Text_NormalizeLookupToken(fieldName)
    For c = 1 To lo.ListColumns.Count
        If VBA.StrComp(private_Text_NormalizeLookupToken(lo.ListColumns(c).Name), fieldName, VBA.vbTextCompare) = 0 Then
            private_ListObject_ColumnIndex = c
            Exit Function
        End If
    Next c
End Function
' --------------------------------------
'  } // namespace ListObject
' --------------------------------------


' --------------------------------------
'  namespace CellRef {
' --------------------------------------
Private Function private_CellRef_ReadRequiredText(ByVal cellRef As String, ByVal paramTitle As String) As String
    Dim p As Long
    Dim ws As Worksheet
    Dim sheetName As String
    Dim cellAddress As String
    Dim valueText As String

    p = VBA.InStr(1, cellRef, "!")
    If p <= 1 Or p = VBA.Len(cellRef) Then
        VBA.Err.Raise VBA.vbObjectError + 4020, , "Неверный формат ссылки параметра: " & cellRef
    End If

    sheetName = VBA.Left$(cellRef, p - 1)
    cellAddress = VBA.Mid$(cellRef, p + 1)
    Set ws = private_Worksheet_GetByNameOrFail(sheetName, "лист параметра '" & paramTitle & "'")
    valueText = VBA.Trim$(VBA.CStr(ws.Range(cellAddress).Value2))

    If VBA.Len(valueText) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 4021, , "Параметр '" & paramTitle & "' пустой: " & cellRef
    End If

    private_CellRef_ReadRequiredText = valueText
End Function


Private Sub private_CellRef_WriteText(ByVal cellRef As String, ByVal valueText As String)
    Dim p As Long
    Dim ws As Worksheet
    Dim sheetName As String
    Dim cellAddress As String

    p = VBA.InStr(1, cellRef, "!")
    If p <= 1 Or p = VBA.Len(cellRef) Then
        VBA.Err.Raise VBA.vbObjectError + 4022, , "Неверный формат ссылки вывода: " & cellRef
    End If

    sheetName = VBA.Left$(cellRef, p - 1)
    cellAddress = VBA.Mid$(cellRef, p + 1)
    Set ws = private_Worksheet_GetByNameOrFail(sheetName, "лист вывода")

    With ws.Range(cellAddress).MergeArea
        .ClearContents
        .Cells(1, 1).Value = valueText
        .WrapText = True
    End With
End Sub


Private Function private_Worksheet_GetByNameOrFail( _
    ByVal sheetName As String, _
    ByVal sheetRole As String _
) As Worksheet
    Dim ws As Worksheet

    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 4023, , "Пустое имя листа (" & sheetRole & ")."
    End If

    For Each ws In Application.ThisWorkbook.Worksheets
        If VBA.StrComp(ws.Name, sheetName, VBA.vbTextCompare) = 0 Then
            Set private_Worksheet_GetByNameOrFail = ws
            Exit Function
        End If
    Next ws

    VBA.Err.Raise VBA.vbObjectError + 4024, , "Лист не найден (" & sheetRole & "): " & sheetName
End Function
' --------------------------------------
'  } // namespace CellRef
' --------------------------------------


' --------------------------------------
'  namespace Workbook {
' --------------------------------------
Private Function private_Workbook_GetSafe( _
    ByVal filePath As String, _
    Optional ByVal openHidden As Boolean = True, _
    Optional ByVal readOnly As Boolean = True _
) As private_WorkbookHandle
#If OPENNING_EXTERNAL_SOURCES_TYPE = 1 Or OPENNING_EXTERNAL_SOURCES_TYPE = 2 Then
    Dim wb As Workbook
    Dim h As private_WorkbookHandle

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 4700, , "private_Workbook_GetSafe: путь к файлу пустой."
    End If
    filePath = private_Workbook_ResolvePath(filePath)

    For Each wb In Application.Workbooks
        If VBA.StrComp(wb.FullName, filePath, VBA.vbTextCompare) = 0 Then
            Set h.Workbook = wb
            h.OpenedByCode = False
            h.WasWindowVisible = VBA.IIf(wb.Windows.Count > 0, wb.Windows(1).Visible, False)
            private_Workbook_GetSafe = h
            Exit Function
        End If
    Next wb

    Set wb = Application.Workbooks.Open( _
        Filename:=filePath, _
        ReadOnly:=readOnly, _
        UpdateLinks:=False, _
        AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True _
    )

    Set h.Workbook = wb
    h.OpenedByCode = True
    h.WasWindowVisible = VBA.IIf(wb.Windows.Count > 0, wb.Windows(1).Visible, False)

    If openHidden And wb.Windows.Count > 0 Then
        wb.Windows(1).Visible = False
    End If

    private_Workbook_GetSafe = h
#Else
    VBA.Err.Raise VBA.vbObjectError + 4901, , "private_Workbook_GetSafe недоступен при текущем OPENNING_EXTERNAL_SOURCES_TYPE."
#End If
End Function


Private Function private_Workbook_ResolvePath(ByVal filePath As String) As String
    Dim basePath As String

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function
    If VBA.Mid$(filePath, 2, 2) = ":\" Or VBA.Left$(filePath, 2) = "\\" Or VBA.Left$(filePath, 1) = "/" Then
        private_Workbook_ResolvePath = filePath
        Exit Function
    End If

    basePath = VBA.Trim$(Application.ThisWorkbook.Path)
    If VBA.Len(basePath) = 0 Then
        private_Workbook_ResolvePath = filePath
        Exit Function
    End If

    Do While VBA.Left$(filePath, 2) = ".\" Or VBA.Left$(filePath, 2) = "./"
        filePath = VBA.Mid$(filePath, 3)
    Loop

    Do While VBA.Left$(filePath, 1) = "\" Or VBA.Left$(filePath, 1) = "/"
        filePath = VBA.Mid$(filePath, 2)
    Loop

    private_Workbook_ResolvePath = basePath & Application.PathSeparator & filePath
End Function


' Главное правило: закрываем только если книгу открыл код.
Private Sub private_Workbook_CloseSafe( _
    ByRef h As private_WorkbookHandle, _
    Optional ByVal saveChanges As Boolean = False _
)
#If OPENNING_EXTERNAL_SOURCES_TYPE = 1 Then
    On Error Resume Next
    If h.Workbook Is Nothing Then Exit Sub
    If h.OpenedByCode Then h.Workbook.Close SaveChanges:=saveChanges
    Set h.Workbook = Nothing
#Else
    VBA.Err.Raise VBA.vbObjectError + 4902, , "private_Workbook_CloseSafe недоступен при текущем OPENNING_EXTERNAL_SOURCES_TYPE."
#End If
End Sub


Private Sub private_Workbook_ReleaseHandle(ByRef h As private_WorkbookHandle)
    Set h.Workbook = Nothing
End Sub
' --------------------------------------
'  } // namespace Workbook
' --------------------------------------


' --------------------------------------
'  namespace Diagnostic {
' --------------------------------------
Private Sub private_Diagnostic_Write(ByVal messageText As String)
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object
    Dim lineText As String

    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
    If VBA.Len(VBA.Trim$(Application.ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = Application.ThisWorkbook.Path & "\\" & DIAGNOSTIC_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If VBA.Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
        lineText = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
        Set stream = fso.OpenTextFile(logPath, 8, True)
        If Not stream Is Nothing Then
            stream.WriteLine lineText
            stream.Close
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub private_Diagnostic_ClearFile()
    Dim logPath As String
    Dim fso As Object
    Dim stream As Object

    If VBA.Len(VBA.Trim$(Application.ThisWorkbook.Path)) = 0 Then Exit Sub
    logPath = Application.ThisWorkbook.Path & "\\" & DIAGNOSTIC_LOG_FILE_REL_PATH

    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        Set stream = fso.OpenTextFile(logPath, 2, True)
        If Not stream Is Nothing Then stream.Close
    End If
    On Error GoTo 0
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------


' --------------------------------------
'  namespace Dict {
' --------------------------------------
Private Function private_Dict_CreateTextMap() As Object
    Set private_Dict_CreateTextMap = VBA.CreateObject("Scripting.Dictionary")
    private_Dict_CreateTextMap.CompareMode = VBA.vbTextCompare
End Function


Private Function private_Dict_RequireText( _
    ByVal targetDict As Object, _
    ByVal key As String, _
    ByVal errPrefix As String _
) As String
    If targetDict Is Nothing Then
        VBA.Err.Raise VBA.vbObjectError + 3000, , errPrefix
    End If

    If Not targetDict.Exists(key) Then
        VBA.Err.Raise VBA.vbObjectError + 3001, , errPrefix & ": " & key
    End If

    private_Dict_RequireText = VBA.Trim$(VBA.CStr(targetDict(key)))
    If VBA.Len(private_Dict_RequireText) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 3002, , errPrefix & ": " & key
    End If
End Function
' --------------------------------------
'  } // namespace Dict
' --------------------------------------


' --------------------------------------
'  namespace Text {
' --------------------------------------
Private Function private_Text_Require(ByVal value As String, ByVal errDescription As String) As String
    private_Text_Require = VBA.Trim$(value)
    If VBA.Len(private_Text_Require) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 5140, , errDescription
    End If
End Function


Private Function private_Text_NullToEmptyString(ByVal value As Variant) As String
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        private_Text_NullToEmptyString = ""
    Else
        private_Text_NullToEmptyString = VBA.CStr(value)
    End If
End Function


Private Function private_Text_FormatDisplayDate(ByVal value As Variant) As String
    Dim textValue As String

    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function

    If VBA.IsDate(value) Then
        private_Text_FormatDisplayDate = VBA.Format$(VBA.CDate(value), "dd.mm.yyyy")
        Exit Function
    End If

    textValue = VBA.Trim$(VBA.CStr(value))
    If VBA.Len(textValue) = 0 Then Exit Function

    If VBA.IsNumeric(textValue) And VBA.CDbl(textValue) >= 20000 And VBA.CDbl(textValue) <= 60000 Then
        private_Text_FormatDisplayDate = VBA.Format$(VBA.CDate(VBA.CDbl(textValue)), "dd.mm.yyyy")
    Else
        private_Text_FormatDisplayDate = textValue
    End If
End Function


Private Function private_Text_CapitalizeFirst(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then Exit Function

    private_Text_CapitalizeFirst = VBA.UCase$(VBA.Left$(text, 1)) & VBA.Mid$(text, 2)
End Function


Private Function private_Text_LowerFirst(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then Exit Function

    private_Text_LowerFirst = VBA.LCase$(VBA.Left$(text, 1)) & VBA.Mid$(text, 2)
End Function


Private Function private_Text_ToShortName(ByVal fullName As String) As String
    Dim parts As Variant
    Dim firstInitial As String
    Dim patronymicInitial As String

    fullName = private_Text_NormalizeSpaces(fullName)
    If VBA.Len(fullName) = 0 Then Exit Function

    parts = VBA.Split(fullName, " ")
    If UBound(parts) < 2 Then
        private_Text_ToShortName = fullName
        Exit Function
    End If

    firstInitial = VBA.UCase$(VBA.Left$(VBA.CStr(parts(1)), 1))
    patronymicInitial = VBA.UCase$(VBA.Left$(VBA.CStr(parts(2)), 1))
    private_Text_ToShortName = private_Text_CapitalizeWord(VBA.CStr(parts(0))) & " " & firstInitial & "." & patronymicInitial & "."
End Function


Private Function private_Text_NormalizeSpaces(ByVal text As String) As String
    text = VBA.Trim$(text)
    Do While VBA.InStr(1, text, "  ") > 0
        text = VBA.Replace(text, "  ", " ")
    Loop
    private_Text_NormalizeSpaces = text
End Function


Private Function private_Text_NormalizeLookupToken(ByVal text As String) As String
    text = private_Text_NullToEmptyString(text)
    text = VBA.Trim$(text)
    text = VBA.Replace(text, VBA.ChrW$(160), " ")
    text = VBA.Replace(text, VBA.ChrW$(8239), " ")
    text = VBA.Replace(text, VBA.ChrW$(173), "")
    text = VBA.Replace(text, VBA.ChrW$(8208), "-")
    text = VBA.Replace(text, VBA.ChrW$(8209), "-")
    text = VBA.Replace(text, VBA.ChrW$(8210), "-")
    text = VBA.Replace(text, VBA.ChrW$(8211), "-")
    text = VBA.Replace(text, VBA.ChrW$(8212), "-")
    text = VBA.Replace(text, VBA.ChrW$(8722), "-")
    text = VBA.Replace(text, VBA.ChrW$(8217), "'")
    text = VBA.Replace(text, VBA.ChrW$(96), "'")
    text = private_Text_NormalizeSpaces(text)
    private_Text_NormalizeLookupToken = text
End Function


Private Function private_Text_CapitalizeWord(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then Exit Function

    private_Text_CapitalizeWord = VBA.UCase$(VBA.Left$(text, 1)) & VBA.LCase$(VBA.Mid$(text, 2))
End Function


Private Function private_Text_UaPidstavaRaport() As String
    private_Text_UaPidstavaRaport = VBA.ChrW$(1055) & VBA.ChrW$(1110) & VBA.ChrW$(1076) & VBA.ChrW$(1089) & VBA.ChrW$(1090) & VBA.ChrW$(1072) & VBA.ChrW$(1074) & VBA.ChrW$(1072) & ": " & VBA.ChrW$(1088) & VBA.ChrW$(1072) & VBA.ChrW$(1087) & VBA.ChrW$(1086) & VBA.ChrW$(1088) & VBA.ChrW$(1090) & " "
End Function


Private Function private_Text_UaIncomingNo() As String
    private_Text_UaIncomingNo = VBA.ChrW$(1074) & VBA.ChrW$(1093) & ". " & VBA.ChrW$(8470) & " "
End Function


Private Function private_Text_UaFrom() As String
    private_Text_UaFrom = " " & VBA.ChrW$(1074) & VBA.ChrW$(1110) & VBA.ChrW$(1076) & " "
End Function
' --------------------------------------
'  } // namespace Text
' --------------------------------------
