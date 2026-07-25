Attribute VB_Name = "ex_ExtWorkbookQuery"
Option Explicit

' Общие операции фильтрации для обоих режимов внешнего запроса.
' Смысл операторов не зависит от backend: SQL и открытый Worksheet обязаны
' отбирать один и тот же набор строк.
Public Enum en_ExtWorkbookQueryOp
    ExtQueryOpEquals = 1
    ExtQueryOpNotEquals = 2
    ExtQueryOpContains = 3
    ExtQueryOpStartsWith = 4
    ExtQueryOpEndsWith = 5
    ExtQueryOpIsEmpty = 6
    ExtQueryOpIsNotEmpty = 7
    ' Сравнение выполняется над текстовым представлением значения. Для дат и
    ' чисел вызывающий код должен передавать формат с корректным лексическим
    ' порядком либо использовать отдельный типизированный API в будущем.
    ExtQueryOpGreaterThan = 8
    ExtQueryOpLessThan = 9
End Enum

' Политика обработки строковых значений, которые ACE возвращает ровно
' 255 символами. Query хранит снимок режима конкретного запуска.
Public Enum en_AdoLongValuesMode
    AdoLongValuesMarkCandidates = 0
    AdoLongValuesHydrate = 1
End Enum
