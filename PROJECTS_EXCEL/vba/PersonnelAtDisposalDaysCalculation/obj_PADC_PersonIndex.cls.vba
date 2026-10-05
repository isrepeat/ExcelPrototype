VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_PersonIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PADC_Configuration
Private m_validation As obj_PADC_Validation
Private m_runContext As obj_PADC_RunContext
Private m_taxIndex As Object
Private m_nameIndex As Object
Private m_nameIds As Object
Private Const DICTIONARY_PROG_ID As String = "Scripting.Dictionary"

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
    ByRef data As Variant, _
    ByVal count As Long, _
    ByVal taxColumn As Long, _
    ByVal nameColumn As Long, _
    ByVal configuration As obj_PADC_Configuration, _
    ByVal validation As obj_PADC_Validation, _
    ByVal runContext As obj_PADC_RunContext _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Or validation Is Nothing Or runContext Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    Set m_validation = validation
    Set m_runContext = runContext
    m_runContext.LogStage "PersonIndex.Started", "Rows=" & count
    private_BuildPersonIndex data, count, taxColumn, nameColumn
    m_runContext.LogStage "PersonIndex.Completed", "TaxIds=" & m_taxIndex.Count & " | Names=" & m_nameIndex.Count
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_taxIndex = Nothing
    Set m_nameIndex = Nothing
    Set m_nameIds = Nothing
    Set m_configuration = Nothing
    Set m_validation = Nothing
    Set m_runContext = Nothing
End Sub

Public Function Find( _
    ByVal taxId As String, _
    ByVal fullName As String, _
    ByRef matchError As String _
) As Collection
    Dim ids As Object

    private_EnsureReady
    matchError = vbNullString
    If m_taxIndex.Exists(taxId) Then
        Set Find = m_taxIndex(taxId)
    ElseIf m_nameIndex.Exists(fullName) Then
        Set Find = m_nameIndex(fullName)
        Set ids = m_nameIds(fullName)
        If ids.Count > 1 Then
            matchError = m_configuration.GetText("legacy.MSG_AMBIGUOUS_NAME") & fullName
        End If
    Else
        Set Find = New Collection
    End If
End Function

' //
' // Private
' //
Private Sub private_BuildPersonIndex( _
    ByRef data As Variant, _
    ByVal count As Long, _
    ByVal taxColumn As Long, _
    ByVal nameColumn As Long _
)
    Dim i As Long
    Dim taxId As String
    Dim fullName As String
    Dim rows As Collection
    Dim ids As Object

    Set m_taxIndex = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set m_nameIndex = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set m_nameIds = VBA.CreateObject(DICTIONARY_PROG_ID)
    m_nameIndex.CompareMode = vbTextCompare
    m_nameIds.CompareMode = vbTextCompare
    For i = 1 To count
        m_runContext.CheckCancel i
        taxId = m_validation.MatchText(data(i, taxColumn))
        fullName = m_validation.MatchText(data(i, nameColumn))
        If VBA.Len(taxId) > 0 Then
            If Not m_taxIndex.Exists(taxId) Then
                Set rows = New Collection
                m_taxIndex.Add taxId, rows
            End If
            Set rows = m_taxIndex(taxId)
            rows.Add i
        End If
        If VBA.Len(fullName) > 0 Then
            If Not m_nameIndex.Exists(fullName) Then
                Set rows = New Collection
                m_nameIndex.Add fullName, rows
                Set ids = VBA.CreateObject(DICTIONARY_PROG_ID)
                m_nameIds.Add fullName, ids
            End If
            Set rows = m_nameIndex(fullName)
            rows.Add i
            Set ids = m_nameIds(fullName)
            If VBA.Len(taxId) > 0 Then
                ids(taxId) = True
            End If
        End If
    Next i
End Sub

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_PersonIndex", "Service is not initialized."
    End If
End Sub