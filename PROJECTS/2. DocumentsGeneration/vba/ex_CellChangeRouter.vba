Option Explicit

Private changeRoutes As Collection
Private selectionRoutes As Collection
Private isDispatchingChange As Boolean
Private isDispatchingSelection As Boolean

' --------------------------------------
' namespace API {
' --------------------------------------
' Очищает runtime-маршруты перед их регистрацией функциями формы.
Public Sub fn_Reset()
    Set changeRoutes = New Collection
    Set selectionRoutes = New Collection
    isDispatchingChange = False
    isDispatchingSelection = False
    ex_Helpers.LogDebug "Cell route registry reset"
End Sub

Public Function fn_RegisterChangeRoute( _
    ByVal worksheetName As String, _
    ByVal rangeAddress As String, _
    ByVal callbackName As String _
) As Boolean
    fn_RegisterChangeRoute = private_Routes_TryRegisterRoute( _
        changeRoutes, worksheetName, rangeAddress, callbackName, "Change")
End Function

Public Function fn_RegisterSelectionRoute( _
    ByVal worksheetName As String, _
    ByVal rangeAddress As String, _
    ByVal callbackName As String _
) As Boolean
    fn_RegisterSelectionRoute = private_Routes_TryRegisterRoute( _
        selectionRoutes, worksheetName, rangeAddress, callbackName, _
        "SelectionChange")
End Function

' Маршрутизирует изменение ячейки в зарегистрированный callback.
Public Sub fn_OnSheetChange( _
    ByVal changedSheet As Object, _
    ByVal target As Range _
)
    Dim eventsWereEnabled As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    If isDispatchingChange Then Exit Sub

    On Error GoTo EH
    eventsWereEnabled = Application.EnableEvents
    isDispatchingChange = True
    Application.EnableEvents = False
    private_Routes_DispatchRoutes changeRoutes, changedSheet, target
CleanExit:
    Application.EnableEvents = eventsWereEnabled
    isDispatchingChange = False
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Application.EnableEvents = eventsWereEnabled
    isDispatchingChange = False
    VBA.Err.Raise errorNumber, "ex_CellChangeRouter.fn_OnSheetChange", _
        errorDescription
End Sub

' Маршрутизирует выбор ячейки в зарегистрированный callback.
Public Sub fn_OnSheetSelectionChange( _
    ByVal changedSheet As Object, _
    ByVal target As Range _
)
    Dim errorNumber As Long
    Dim errorDescription As String

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    If isDispatchingSelection Then Exit Sub

    On Error GoTo EH
    isDispatchingSelection = True
    private_Routes_DispatchRoutes selectionRoutes, changedSheet, target
CleanExit:
    isDispatchingSelection = False
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    isDispatchingSelection = False
    VBA.Err.Raise errorNumber, "ex_CellChangeRouter.fn_OnSheetSelectionChange", _
        errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Routes {
' --------------------------------------
Private Function private_Routes_TryRegisterRoute( _
    ByRef routes As Collection, _
    ByVal worksheetName As String, _
    ByVal rangeAddress As String, _
    ByVal callbackName As String, _
    ByVal eventName As String _
) As Boolean
    Dim routeItem As Object

    worksheetName = VBA.Trim$(worksheetName)
    rangeAddress = VBA.Trim$(rangeAddress)
    callbackName = VBA.Trim$(callbackName)
    If VBA.Len(worksheetName) = 0 Or VBA.Len(rangeAddress) = 0 Or _
       VBA.Len(callbackName) = 0 Then
        VBA.MsgBox "Invalid " & eventName & " route. Sheet, range, and callback " & _
            "must be specified.", VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If routes Is Nothing Then Set routes = New Collection

    Set routeItem = VBA.CreateObject("Scripting.Dictionary")
    routeItem.CompareMode = VBA.vbTextCompare
    routeItem("WorksheetName") = worksheetName
    routeItem("RangeAddress") = rangeAddress
    routeItem("CallbackName") = callbackName
    routes.Add routeItem
    ex_Helpers.LogDebug "Cell route registered | Event=" & eventName & _
        " | Sheet=" & worksheetName & " | Range=" & rangeAddress & _
        " | Callback=" & callbackName
    private_Routes_TryRegisterRoute = True
End Function

Private Sub private_Routes_DispatchRoutes( _
    ByVal routes As Collection, _
    ByVal changedSheet As Object, _
    ByVal target As Range _
)
    Dim routeItem As Object
    Dim routeRange As Range
    Dim matchedRange As Range
    Dim callbackName As String

    If routes Is Nothing Then Exit Sub
    For Each routeItem In routes
        If VBA.StrComp(VBA.CStr(routeItem("WorksheetName")), _
                VBA.CStr(changedSheet.Name), VBA.vbTextCompare) <> 0 Then
            GoTo ContinueRoute
        End If
        Set routeRange = changedSheet.Range( _
            VBA.CStr(routeItem("RangeAddress")))
        Set matchedRange = Application.Intersect(target, routeRange)
        If matchedRange Is Nothing Then GoTo ContinueRoute
        callbackName = VBA.CStr(routeItem("CallbackName"))
        ex_Helpers.LogDebug "Cell route matched | Sheet=" & _
            VBA.CStr(changedSheet.Name) & " | Range=" & _
            routeRange.Address(False, False) & " | Callback=" & callbackName
        private_Routes_InvokeRouteCallback callbackName, changedSheet, matchedRange
ContinueRoute:
    Next routeItem
End Sub

Private Sub private_Routes_InvokeRouteCallback( _
    ByVal callbackName As String, _
    ByVal changedSheet As Object, _
    ByVal target As Range _
)
    Dim macroReference As String

    macroReference = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & callbackName
    Application.Run macroReference, changedSheet, target
End Sub
' --------------------------------------
' } // namespace Routes
' --------------------------------------