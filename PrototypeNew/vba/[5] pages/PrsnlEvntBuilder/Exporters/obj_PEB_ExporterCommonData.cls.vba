VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExporterCommonData"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Private m_IsDisposed As Boolean

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
' // API
' //
Public Function Initialize() As Boolean
    m_IsDisposed = False
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
End Sub

' Общий data-provider для экспортеров PrsnlEvntBuilder.
' Здесь будут жить запросы к таблицам, которые нужны нескольким экспортерам:
' даты приказов, общие справочники, cross-export lookup и т.п.
Public Function TryResolveOrderDateByNumber( _
    ByVal orderNo As Variant, _
    ByRef outOrderDate As Date _
) As Boolean
    If m_IsDisposed Then Exit Function

    ' Каркас: реализация будет добавлена, когда общий источник данных
    ' и контракт lookup-таблиц будут окончательно зафиксированы.
    TryResolveOrderDateByNumber = False
End Function
