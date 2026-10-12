VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TDB_PgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IPage

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean
Private m_pageBase As obj_PageBase
Private m_controller As obj_TDB_PgMainController

Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    If m_isDisposed Or m_isInitialized Then Exit Function

    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, VBA.Trim$(profileId)) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_TDB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    obj_IPage_Initialize = True
    m_isInitialized = True
End Function

Private Function obj_IPage_Render() As Boolean
    If m_pageBase Is Nothing Then Exit Function
    obj_IPage_Render = m_pageBase.Render()
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If target Is Nothing Then Exit Function
    obj_IPage_HandleCellChange = True
End Function

Private Sub obj_IPage_Dispose()
    Me.Dispose
End Sub

Public Sub Dispose()
    If m_isDisposed Then Exit Sub

    m_isDisposed = True
    m_isInitialized = False
    If Not m_controller Is Nothing Then m_controller.Dispose
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then m_pageBase.Dispose
    Set m_pageBase = Nothing
End Sub

Private Function private_TryRegisterBindings() As Boolean
    Dim bindingContext As obj_UiBindingContext
    Dim key As Variant
    Dim valueText As String

    If m_pageBase Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function

    If Not bindingContext.SetValue("Text", "Title", "TravelDocumentsBuilder") Then Exit Function
    If Not bindingContext.SetValue("Form", "DocumentType", "Відпускний квиток") Then Exit Function
    If Not bindingContext.SetValue( _
            "Data", "DocumentTypes", "Відпускний квиток,Посвідчення про відрядження") Then Exit Function
    For Each key In VBA.Array("PersonName", "PersonId", "PersonRank", "PersonPosition", _
            "PersonUnit", "DocumentTarget", "DepartureDate", "ReturnDate", "Purpose")
        If Not bindingContext.SetValue("Form", VBA.CStr(key), VBA.vbNullString) Then Exit Function
    Next key
    If Not bindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not bindingContext.SetValue("Form", "AutoFillPersonnel", True) Then Exit Function
    For Each key In VBA.Array("DocumentTypeLabel", "PersonnelPrompt", "PersonNameLabel", _
            "PersonIdLabel", "PersonRankLabel", "PersonPositionLabel", "PersonUnitLabel", _
            "RefreshPersonnel", "AutoFillPersonnelLabel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & VBA.CStr(key), valueText) Then Exit Function
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    If Not ex_Core.fn_TryGetWorkbookConfigValue( _
            "TravelDocumentsBuilder::text.PersonnelPrompt", valueText) Then Exit Function
    If Not bindingContext.SetValue("Text", "PersonnelStatus", valueText) Then Exit Function
    private_TryRegisterBindings = private_ApplyDocumentMode("Відпускний квиток")
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim bindingContext As obj_UiBindingContext

    If m_pageBase Is Nothing Or m_controller Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "SearchPersonnel", "SearchPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "SelectPersonnel", "SelectPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "RefreshPersonnel", "RefreshPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "DocumentTypeChanged", "DocumentTypeChangedHandler") Then Exit Function
    private_TryRegisterCommands = True
End Function

Private Function private_TryRegisterCommand( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal commandKey As String, _
    ByVal methodName As String _
) As Boolean
    Dim command As obj_UiCommand

    If bindingContext Is Nothing Or m_controller Is Nothing Then Exit Function
    Set command = New obj_UiCommand
    If Not command.Initialize(m_controller, methodName) Then Exit Function
    private_TryRegisterCommand = bindingContext.SetObject("Commands", commandKey, command)
End Function

Private Function private_ApplyDocumentMode(ByVal documentType As String) As Boolean
    Dim bindingContext As obj_UiBindingContext
    Dim prefix As String
    Dim key As Variant
    Dim valueText As String

    If m_pageBase Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function

    Select Case VBA.Trim$(documentType)
        Case "Відпускний квиток"
            prefix = "VacationTicket"
        Case "Посвідчення про відрядження"
            prefix = "BusinessTripCertificate"
        Case Else
            Exit Function
    End Select
    For Each key In VBA.Array("DocumentFormTitle", "DocumentTargetLabel", _
            "DepartureDateLabel", "ReturnDateLabel", "PurposeLabel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & prefix & "." & VBA.CStr(key), valueText) Then Exit Function
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    private_ApplyDocumentMode = True
End Function