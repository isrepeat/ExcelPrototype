VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiCellBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiBindingListener

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_worksheetName As String
Private m_cellAddress As String
Private m_bindingContext As obj_UiBindingContext
Private m_sourceName As String
Private m_bindingPath As String
Private m_command As obj_UiCommand
Private m_twoWay As Boolean
Private m_refreshListeners As Collection

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
Public Property Get WorksheetName() As String
    WorksheetName = m_worksheetName
End Property

Public Property Get CellAddress() As String
    CellAddress = m_cellAddress
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal worksheetName As String, _
    ByVal cellAddress As String, _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal sourceName As String, _
    ByVal bindingPath As String, _
    ByVal command As obj_UiCommand _
) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(worksheetName)) = 0 Or _
       VBA.Len(VBA.Trim$(cellAddress)) = 0 Then Exit Function
    If bindingContext Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(sourceName)) = 0 Or _
       VBA.Len(VBA.Trim$(bindingPath)) = 0 Then Exit Function
    m_worksheetName = worksheetName
    m_cellAddress = cellAddress
    Set m_bindingContext = bindingContext
    m_sourceName = sourceName
    m_bindingPath = bindingPath
    Set m_command = command
    m_twoWay = True
    Set m_refreshListeners = New Collection
    Initialize = True
    m_isInitialized = Initialize
    m_bindingContext.Subscribe Me
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_refreshListeners = Nothing
    Set m_command = Nothing
    If Not m_bindingContext Is Nothing Then m_bindingContext.Unsubscribe Me
    Set m_bindingContext = Nothing
    m_worksheetName = VBA.vbNullString
    m_cellAddress = VBA.vbNullString
    m_sourceName = VBA.vbNullString
    m_bindingPath = VBA.vbNullString
End Sub

Public Sub SetTwoWay(ByVal enabled As Boolean)
    m_twoWay = enabled
End Sub

Public Sub SubscribeRefresh(ByVal listener As obj_IUiCellRefreshListener)
    Dim registered As obj_IUiCellRefreshListener

    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Cell binding is not initialized."
    If listener Is Nothing Then Err.Raise 5, , "Cell refresh listener is required."
    For Each registered In m_refreshListeners
        If registered Is listener Then Exit Sub
    Next registered
    m_refreshListeners.Add listener
End Sub

Public Sub UnsubscribeRefresh(ByVal listener As obj_IUiCellRefreshListener)
    Dim index As Long
    Dim registered As obj_IUiCellRefreshListener

    If m_refreshListeners Is Nothing Then Exit Sub
    For index = m_refreshListeners.Count To 1 Step -1
        Set registered = m_refreshListeners(index)
        If registered Is listener Then m_refreshListeners.Remove index
    Next index
End Sub

Public Function HandleCellChange(ByVal target As Range) As Boolean
    If m_isDisposed Or target Is Nothing Then Exit Function
    If VBA.StrComp(target.Parent.Name, m_worksheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.StrComp(target.Address(False, False), m_cellAddress, VBA.vbTextCompare) <> 0 Then Exit Function
    If Not m_twoWay Then
        private_bindingContext_ValueChanged m_sourceName, m_bindingPath
        HandleCellChange = True
        Exit Function
    End If
    If Not m_bindingContext.TrySetPathValue( _
            m_sourceName, m_bindingPath, target.Value2) Then Exit Function
    If Not m_command Is Nothing Then
        If Not m_command.Execute() Then Exit Function
    End If
    HandleCellChange = True
End Function

' //
' // Private
' //
Private Sub obj_IUiBindingListener_OnValueChanged( _
    ByVal sourceName As String, _
    ByVal bindingPath As String _
)
    private_bindingContext_ValueChanged sourceName, bindingPath
End Sub

Private Sub private_bindingContext_ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim previousEnableEvents As Boolean
    Dim target As Range

    If m_isDisposed Then Exit Sub
    If VBA.StrComp(sourceName, m_sourceName, VBA.vbTextCompare) <> 0 Then Exit Sub
    If VBA.StrComp(bindingPath, m_bindingPath, VBA.vbTextCompare) <> 0 Then
        If VBA.StrComp(VBA.Left$(m_bindingPath, VBA.Len(bindingPath) + 1), _
                bindingPath & ".", VBA.vbTextCompare) <> 0 Then Exit Sub
    End If
    previousEnableEvents = Application.EnableEvents
    On Error GoTo EH
    If Not m_bindingContext.TryGetValue(m_sourceName, m_bindingPath, _
            value, sourceObject, isObject) Then
        VBA.Err.Raise VBA.vbObjectError + 2101, "obj_UiCellBinding", _
            "Binding value was not found: " & m_sourceName & "." & m_bindingPath
    End If
    If isObject Then
        VBA.Err.Raise VBA.vbObjectError + 2102, "obj_UiCellBinding", _
            "A cell binding requires a scalar value: " & m_sourceName & "." & m_bindingPath
    End If
    Set target = ThisWorkbook.Worksheets(m_worksheetName).Range(m_cellAddress)
    Application.EnableEvents = False
    target.Value2 = value
    Application.EnableEvents = previousEnableEvents
    private_NotifyValueRefreshed
    Exit Sub
EH:
    Application.EnableEvents = previousEnableEvents
    ex_WindowsUi.fn_ShowMessage "Cannot refresh bound cell " & m_worksheetName & "!" & m_cellAddress & _
        ": " & VBA.Err.Description, VBA.vbExclamation, "Binding"
End Sub

Private Sub private_NotifyValueRefreshed()
    Dim snapshot As New Collection
    Dim listener As obj_IUiCellRefreshListener

    If m_refreshListeners Is Nothing Then Exit Sub
    For Each listener In m_refreshListeners
        snapshot.Add listener
    Next listener
    For Each listener In snapshot
        If m_isDisposed Then Exit Sub
        listener.OnValueRefreshed
    Next listener
End Sub