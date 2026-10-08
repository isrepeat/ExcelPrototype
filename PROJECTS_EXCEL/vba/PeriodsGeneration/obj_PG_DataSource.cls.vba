VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PG_Configuration
Private m_runContext As obj_PG_RunContext

' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //

Public Function Initialize( _
    ByVal configuration As obj_PG_Configuration, _
    ByVal runContext As obj_PG_RunContext _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Or runContext Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    Set m_runContext = runContext
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_configuration = Nothing
    Set m_runContext = Nothing
End Sub

Public Function FindSourceTable() As ListObject
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    Dim table As ListObject
    Dim found As ListObject
    Dim column As ListColumn
    Dim key As Variant
    Dim columnName As String

    private_EnsureReady
    For Each workbook In Application.Workbooks
        m_runContext.CheckCancel 0
        If Not workbook Is ThisWorkbook Then
            Set worksheet = Nothing
            On Error Resume Next
            Set worksheet = workbook.Worksheets(m_configuration.GetText("legacy.SOURCE_SHEET_NAME"))
            On Error GoTo 0
            If Not worksheet Is Nothing Then
                For Each table In worksheet.ListObjects
                    If VBA.StrComp(table.Name, m_configuration.GetText("legacy.SOURCE_TABLE_NAME"), vbBinaryCompare) = 0 Then
                        If Not found Is Nothing Then
                            m_configuration.Fail m_configuration.GetText("message.MultipleSources")
                        End If
                        Set found = table
                    End If
                Next table
            End If
        End If
    Next workbook
    If found Is Nothing Then
        m_configuration.Fail m_configuration.GetText("message.SourceMissing")
    End If
    For Each key In VBA.Array( _
        "legacy.SOURCE_COL_RANK", _
        "legacy.SOURCE_COL_NAME", _
        "legacy.SOURCE_COL_TAX_ID", _
        "legacy.SOURCE_COL_POSITION", _
        "legacy.SOURCE_COL_EVENT", _
        "legacy.SOURCE_COL_PERIOD_FROM", _
        "legacy.SOURCE_COL_PERIOD_TO", _
        "legacy.SOURCE_COL_DEPARTURE_ORDER", _
        "legacy.SOURCE_COL_ARRIVAL_ORDER" _
    )
        columnName = m_configuration.GetText(VBA.CStr(key))
        Set column = Nothing
        On Error Resume Next
        Set column = found.ListColumns(columnName)
        On Error GoTo 0
        If column Is Nothing Then
            m_configuration.Fail VBA.Replace(m_configuration.GetText("message.ColumnMissing"), "{column}", columnName)
        End If
    Next key
    Set FindSourceTable = found
End Function

' //
' // Вспомогательные методы
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_DataSource", "Service is not initialized."
    End If
End Sub