VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_configuration As obj_PADC_Configuration
Private m_validation As obj_PADC_Validation
Private m_runContext As obj_PADC_RunContext
Private m_inputPath As String
Private m_sheetName As String
Private m_rangeAddress As String
Private m_endDay As Long
Private Const INPUT_REFERENCE_SEPARATOR As String = "|"
Private Const INPUT_SHEET_SEPARATOR As String = "!"
Private Const PARAM_INPUT_PATH_CELL As String = "I2"
Private Const FORMAT_DATE As String = "dd.mm.yyyy"
Private Const FILE_SYSTEM_PROG_ID As String = "Scripting.FileSystemObject"

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get InputPath() As String
    private_EnsureReady
    InputPath = m_inputPath
End Property

Public Property Get SheetName() As String
    private_EnsureReady
    SheetName = m_sheetName
End Property

Public Property Get RangeAddress() As String
    private_EnsureReady
    RangeAddress = m_rangeAddress
End Property

Public Property Get EndDay() As Long
    private_EnsureReady
    EndDay = m_endDay
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal pageBase As obj_PageBase, _
    ByVal configuration As obj_PADC_Configuration, _
    ByVal validation As obj_PADC_Validation, _
    ByVal runContext As obj_PADC_RunContext, _
    ByVal parameterSheet As Worksheet _
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
    If validation Is Nothing Then
        Exit Function
    End If
    Set m_validation = validation
    If runContext Is Nothing Then
        Exit Function
    End If
    Set m_runContext = runContext
    If parameterSheet Is Nothing Then
        Exit Function
    End If
    private_ReadCalculationParameters parameterSheet
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
    Set m_validation = Nothing
    Set m_runContext = Nothing
End Sub

' //
' // Private
' //
Private Sub private_ReadCalculationParameters(ByVal ws As Worksheet)
    Dim fileSystem As Object

    m_inputPath = m_validation.RequiredText(private_ReadFormValue("InputReference"), "Form.InputReference")
    m_endDay = m_validation.ReadDay(private_ReadFormValue("EndDate"), ws.Parent.Date1904, "Form.EndDate")
    private_ParseInputReference
    m_inputPath = private_ResolveInputPath(m_inputPath)
    m_runContext.LogDebug m_configuration.GetText("text.InputFile") & m_inputPath
    m_runContext.LogDebug m_configuration.GetText("text.EndDate") & VBA.Format$(VBA.CDate(m_endDay), FORMAT_DATE)
    m_runContext.CheckCancel 0
    Set fileSystem = VBA.CreateObject(FILE_SYSTEM_PROG_ID)
    If Not fileSystem.FileExists(m_inputPath) Then
        m_configuration.Fail ws.Name & "!" & PARAM_INPUT_PATH_CELL & m_configuration.GetText("text.FileNotFoundOrInaccessible") & m_inputPath
    End If
End Sub

Private Sub private_ParseInputReference()
    Dim parts As Variant
    Dim location As String
    Dim separatorPosition As Long
    Dim address As String
    Dim bounds As Variant
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long

    parts = VBA.Split(m_inputPath, INPUT_REFERENCE_SEPARATOR)
    If UBound(parts) <> 1 Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
    m_inputPath = VBA.Trim$(parts(0))
    location = VBA.Trim$(parts(1))
    separatorPosition = VBA.InStrRev(location, INPUT_SHEET_SEPARATOR)
    If separatorPosition <= 1 Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
    m_sheetName = VBA.Trim$(VBA.Left$(location, separatorPosition - 1))
    If VBA.Left$(m_sheetName, 1) = "'" And VBA.Right$(m_sheetName, 1) = "'" Then
        m_sheetName = VBA.Replace(VBA.Mid$(m_sheetName, 2, _
            VBA.Len(m_sheetName) - 2), "''", "'")
    End If
    address = VBA.UCase$(VBA.Replace(VBA.Trim$(VBA.Mid$(location, separatorPosition + 1)), "$", ""))
    If VBA.Len(m_inputPath) = 0 Or VBA.Len(m_sheetName) = 0 Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
    bounds = VBA.Split(address, ":")
    If UBound(bounds) <> 1 Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
    private_ParseCell VBA.CStr(bounds(0)), firstRow, firstColumn
    private_ParseCell VBA.CStr(bounds(1)), lastRow, lastColumn
    If lastRow <= firstRow Or lastColumn < firstColumn Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
    m_rangeAddress = address
End Sub

Private Sub private_ParseCell( _
    ByVal address As String, _
    ByRef rowNumber As Long, _
    ByRef columnNumber As Long _
)
    Dim i As Long
    Dim character As String
    Dim digits As String

    For i = 1 To VBA.Len(address)
        character = VBA.Mid$(address, i, 1)
        If character Like "[A-Z]" And VBA.Len(digits) = 0 Then
            columnNumber = columnNumber * 26 + VBA.AscW(character) - VBA.AscW("A") + 1
            If columnNumber > 16384 Then
                m_configuration.Fail m_configuration.GetText("text.InvalidReference")
            End If
        ElseIf character Like "[0-9]" And columnNumber > 0 Then
            digits = digits & character
            If VBA.Len(digits) > 7 Then
                m_configuration.Fail m_configuration.GetText("text.InvalidReference")
            End If
        Else
            m_configuration.Fail m_configuration.GetText("text.InvalidReference")
        End If
    Next i
    If VBA.Len(digits) = 0 Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
    rowNumber = VBA.CLng(digits)
    If rowNumber < 1 Or rowNumber > 1048576 Or VBA.Left$(digits, 1) = "0" Then
        m_configuration.Fail m_configuration.GetText("text.InvalidReference")
    End If
End Sub

Private Function private_ResolveInputPath(ByVal inputPath As String) As String
    Dim fileSystem As Object
    Dim basePath As String

    Set fileSystem = VBA.CreateObject(FILE_SYSTEM_PROG_ID)
    inputPath = VBA.Replace(inputPath, "/", "\")
    If VBA.Left$(inputPath, 2) = "\\" Or inputPath Like "[A-Za-z]:\*" Then
        private_ResolveInputPath = fileSystem.GetAbsolutePathName(inputPath)
        Exit Function
    End If
    If VBA.Left$(inputPath, 1) = "\" Or VBA.InStr(1, inputPath, ":", vbBinaryCompare) > 0 Then
        m_configuration.Fail m_configuration.GetText("text.AmbiguousPath") & inputPath
    End If
    basePath = ThisWorkbook.Path
    If VBA.Len(basePath) = 0 Or VBA.InStr(1, basePath, "://", vbBinaryCompare) > 0 Then
        m_configuration.Fail m_configuration.GetText("text.RelativePathBase")
    End If
    private_ResolveInputPath = fileSystem.GetAbsolutePathName(fileSystem.BuildPath(basePath, inputPath))
End Function

Private Function private_ReadFormValue(ByVal fieldName As String) As Variant
    Dim value As Variant
    Dim valueObject As Object
    Dim isObject As Boolean

    If Not m_pageBase.BindingContext.TryGetValue("Form", fieldName, value, valueObject, isObject) Then
        m_configuration.Fail "Form." & fieldName
    End If
    private_ReadFormValue = value
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_Parameters", "Service is not initialized."
    End If
End Sub