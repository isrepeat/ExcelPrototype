VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageEntityLookupCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PageEntityLookup.Controller"
Private Const CANDIDATE_TABLES_RUNTIME_KEY As String = "RuntimeItems.EntityLookup.CandidateTables"

Private m_LookupFeature As obj_EntityLookupFeature
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get LookupFeature() As obj_EntityLookupFeature
    Set LookupFeature = m_LookupFeature
End Property

' //
' // API
' //
Public Function Initialize(ByVal page As Object) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageEntityLookupCtrl.Initialize"
#End If
    Dim pageBase As obj_PageBase
    Dim pageInterface As obj_IPage

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageEntityLookupCtrl initialization failed because page is not specified."
#End If
        Exit Function
    End If
    On Error Resume Next
    Set pageInterface = page
    On Error GoTo 0
    If pageInterface Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageEntityLookupCtrl initialization failed because page does not implement obj_IPage."
#End If
        Exit Function
    End If

    m_IsDisposed = False
    Set pageBase = pageInterface.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.RuntimeSources Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_LookupFeature = New obj_EntityLookupFeature
    If Not m_LookupFeature.Initialize( _
        pageInterface, _
        CANDIDATE_TABLES_RUNTIME_KEY, _
        "entitylookup") Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageEntityLookupCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_LookupFeature Is Nothing Then m_LookupFeature.Dispose
    Set m_LookupFeature = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    UpdateData = m_LookupFeature.UpdateData(configControl)
End Function

Public Function PrepareLookupRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    PrepareLookupRuntime = m_LookupFeature.PrepareLookupRuntime(notifyChange)
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    ClearLookupCandidates = m_LookupFeature.ClearLookupCandidates(renderNow)
End Function

Public Function SearchCandidates( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByRef outCandidateCount As Long, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    If m_LookupFeature Is Nothing Then Exit Function
    SearchCandidates = m_LookupFeature.SearchCandidates(lookupKey, queryText, outCandidateCount, notifyChange)
End Function

Public Function TryGetLookupKeys(ByRef outLookupKeys As Collection) As Boolean
    Set outLookupKeys = Nothing
    If m_LookupFeature Is Nothing Then Exit Function
    TryGetLookupKeys = m_LookupFeature.TryGetLookupKeys(outLookupKeys)
End Function
