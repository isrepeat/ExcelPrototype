VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IUaInflector"
Option Explicit

' Общий контракт для взаимозаменяемых реализаций украинского склонения.
' targetCase передаётся стабильным английским идентификатором:
' "genitive", "accusative" или "dative".
'
' True означает, что реализация обработала запрос. Результат при этом может
' совпасть с sourceText: неизменяемое слово не является ошибкой склонения.
Public Function TryInflect( _
    ByVal sourceText As String, _
    ByVal targetCase As String, _
    ByRef outText As String _
) As Boolean
End Function
