VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_ResultWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_configuration As obj_PADC_Configuration
Private Const FORMAT_DATE As String = "dd.mm.yyyy"

' //
' // Lifecycle
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
    ByVal pageBase As obj_PageBase, _
    ByVal configuration As obj_PADC_Configuration _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If pageBase Is Nothing Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    If configuration Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_pageBase = Nothing
    Set m_configuration = Nothing
End Sub

Public Sub Publish( _
    ByRef names() As Variant, _
    ByRef ids() As Variant, _
    ByRef starts() As Variant, _
    ByRef periods() As Variant, _
    ByRef days() As Variant, _
    ByRef thresholds() As Variant, _
    ByVal count As Long, _
    ByVal failures As Collection, _
    ByVal lastDay As Long _
)
    Dim values() As Variant
    Dim headers As Variant
    Dim errorValues() As Variant
    Dim errorHeaders As Variant
    Dim item As Variant
    Dim i As Long
    Dim resultTable As obj_UiRawTable
    Dim errorTable As obj_UiRawTable
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim through As Date
    Dim threshold As Date
    Dim include As Boolean
    Dim separator As String

    private_EnsureReady
    headers = VBA.Array(m_configuration.GetText("header.FullName"), m_configuration.GetText("header.TaxId"), m_configuration.GetText("header.StartDate"), _
        m_configuration.GetText("header.Periods"), m_configuration.GetText("header.CountedDays"), m_configuration.GetText("header.ThresholdDate"), _
        m_configuration.GetText("export.header.From"), m_configuration.GetText("export.header.Through"), m_configuration.GetText("export.header.Period"), _
        m_configuration.GetText("export.header.Include"))
    Set resultTable = New obj_UiRawTable
    If count > 0 Then
        monthStart = VBA.DateSerial(VBA.Year(VBA.CDate(lastDay)), VBA.Month(VBA.CDate(lastDay)) - 1, 1)
        monthEnd = VBA.DateSerial(VBA.Year(VBA.CDate(lastDay)), VBA.Month(VBA.CDate(lastDay)), 0)
        separator = " " & m_configuration.GetText("format.periodRangeSeparator") & " "
        ReDim values(1 To count, 1 To 10)
        For i = 1 To count
            values(i, 1) = names(i, 1)
            values(i, 2) = ids(i, 1)
            values(i, 3) = VBA.Format$(starts(i, 1), FORMAT_DATE)
            values(i, 4) = periods(i, 1)
            values(i, 5) = days(i, 1)
            include = True
            through = monthEnd
            If Not VBA.IsEmpty(thresholds(i, 1)) Then
                threshold = VBA.CDate(thresholds(i, 1))
                values(i, 6) = VBA.Format$(threshold, FORMAT_DATE)
                include = threshold > monthStart
                If threshold <= monthEnd Then
                    through = threshold - 1
                End If
            End If
            If include Then
                values(i, 7) = VBA.Format$(monthStart, FORMAT_DATE)
                values(i, 8) = VBA.Format$(through, FORMAT_DATE)
                values(i, 9) = values(i, 7) & separator & values(i, 8)
            End If
            values(i, 10) = include
        Next i
        If Not resultTable.Initialize(values, headers) Then
            m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
        End If
    Else
        If Not resultTable.InitializeEmpty(headers) Then
            m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
        End If
    End If
    errorHeaders = VBA.Array(m_configuration.GetText("header.errors.FullName"), m_configuration.GetText("header.errors.TaxId"), m_configuration.GetText("header.errors.Description"))
    Set errorTable = New obj_UiRawTable
    If failures.Count > 0 Then
        ReDim errorValues(1 To failures.Count, 1 To 3)
        For i = 1 To failures.Count
            item = failures(i)
            errorValues(i, 1) = item(0)
            errorValues(i, 2) = item(1)
            errorValues(i, 3) = item(2)
        Next i
        If Not errorTable.Initialize(errorValues, errorHeaders) Then
            m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
        End If
    Else
        If Not errorTable.InitializeEmpty(errorHeaders) Then
            m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
        End If
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "CalculationRows", resultTable) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "Errors", errorTable) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    If Not m_pageBase.Render() Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
End Sub

' //
' // Private
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_ResultWriter", "Service is not initialized."
    End If
End Sub