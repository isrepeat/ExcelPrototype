VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WordDataExtractorCfgParser"
Option Explicit

Private m_CfgParserBase As obj_CfgParserBase
Private m_ConfigEntries As Collection
Private m_CfgMap As Object
Private m_IsDisposed As Boolean

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    m_IsDisposed = False
    Set m_CfgParserBase = New obj_CfgParserBase
    Set m_ConfigEntries = Nothing
    Set m_CfgMap = Nothing

    If configTable Is Nothing Then Exit Function
    ' Сам parser является resolverDataContext режима. Благодаря этому
    ' transformer получает уже разрешённый путь и не знает ни о binding,
    ' ни о правилах выбора последней даты в имени файла.
    If Not m_CfgParserBase.Initialize(configTable, Me) Then Exit Function
    If Not m_CfgParserBase.TryGetConfigEntries(m_ConfigEntries) Then Exit Function
    If Not m_CfgParserBase.BuildConfigDictionary( _
        m_ConfigEntries, m_CfgMap) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CfgParserBase Is Nothing Then m_CfgParserBase.Dispose
    Set m_CfgParserBase = Nothing
    Set m_ConfigEntries = Nothing
    Set m_CfgMap = Nothing
    On Error GoTo 0
End Sub

Public Function ResolveLatestByDmyPattern(ByVal rawValue As String) As String
    ResolveLatestByDmyPattern = _
        ex_SourceResolver.fn_ResolveLatestByDmyPattern(rawValue)
End Function

Public Function TryGetRequiredValue( _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    outValue = VBA.vbNullString
    If m_CfgParserBase Is Nothing Then Exit Function
    If m_CfgMap Is Nothing Then Exit Function

    TryGetRequiredValue = m_CfgParserBase.TryGetRequiredConfigValue( _
        m_CfgMap, keyName, outValue)
End Function

Public Function GetOptionalValue( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = VBA.vbNullString _
) As String
    GetOptionalValue = defaultValue
    If m_CfgParserBase Is Nothing Then Exit Function
    If m_CfgMap Is Nothing Then Exit Function

    GetOptionalValue = m_CfgParserBase.GetOptionalConfigValue( _
        m_CfgMap, keyName, defaultValue)
End Function
