Attribute VB_Name = "ex_QueryRunner"
Option Explicit

' ============================================================================
' ex_QueryRunner
' ----------------------------------------------------------------------------
' Модуль-источник данных.
'
' Отвечает ТОЛЬКО за получение данных.
' Не пишет в Excel, не форматирует, не подсвечивает.
'
' В реальном проекте здесь будет:
' - чтение из других файлов
' - Power Query
' - сравнение таблиц
' - любые вычисления
'
' Возвращает данные в виде 2D-массива Variant:
'   (строки, колонки)
' ============================================================================

Public Function GetResultTable() As Variant
    ' Тестовые данные (HelloWorld),
    ' чтобы проверить весь pipeline
    
    Dim data(1 To 6, 1 To 4) As Variant
    
    ' Заголовки таблицы
    data(1, 1) = "ID"
    data(1, 2) = "Name"
    data(1, 3) = "Status"
    data(1, 4) = "Value"
    
    ' Строки данных
    data(2, 1) = 1
    data(2, 2) = "Alpha"
    data(2, 3) = "Added"
    data(2, 4) = 10
    
    data(3, 1) = 2
    data(3, 2) = "Beta"
    data(3, 3) = "Changed"
    data(3, 4) = 20
    
    data(4, 1) = 3
    data(4, 2) = "Gamma"
    data(4, 3) = "Removed"
    data(4, 4) = 30
    
    data(5, 1) = 4
    data(5, 2) = "Delta"
    data(5, 3) = "Error"
    data(5, 4) = "Bad row"
    
    data(6, 1) = 5
    data(6, 2) = "Epsilon"
    data(6, 3) = "OK"
    data(6, 4) = 50
    
    ' Возвращаем таблицу вызывающему коду
    GetResultTable = data
End Function