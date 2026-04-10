Attribute VB_Name = "ex_ShapeClickDispatcher"
Option Explicit

' Универсальный runtime-диспетчер кликов по shape.
' Причина существования модуля: Shape.OnAction может вызвать только public macro
' из стандартного модуля, но не метод конкретного экземпляра класса-контрола.
'
' Модуль хранит:
' 1) registry контролов (controlKey -> VM/object)
' 2) registry маршрутов shape (shapeName -> controlKey + method + arg)
' и делегирует клики в нужный метод нужного объекта.
'
' controlKey -> Object (обычно VM-контрол)
Private g_ControlByKey As Object
' shapeName -> { ControlKey, MethodName, HasArg, ArgValue }
Private g_RouteByShape As Object
' Последний факт диспатчинга (универсальный контекст)
Private g_LastDispatchContext As Object
' Последний контекст выбранного элемента select (совместимость)
Private g_LastSelectContext As Object

' === Универсальный API диспетчера ===
Public Sub m_ResetDispatcher()
    Set g_ControlByKey = Nothing
    Set g_RouteByShape = Nothing
    Set g_LastDispatchContext = Nothing
    Set g_LastSelectContext = Nothing
End Sub

Public Function m_RegisterControl(ByVal controlKey As String, ByVal controlVm As Object) As Boolean
    controlKey = LCase$(Trim$(controlKey))
    If Len(controlKey) = 0 Then
        MsgBox "ShapeClickDispatcher: control key is empty.", vbExclamation
        Exit Function
    End If
    If controlVm Is Nothing Then
        MsgBox "ShapeClickDispatcher: control VM is not specified for key '" & controlKey & "'.", vbExclamation
        Exit Function
    End If

    mp_EnsureStorage
    Set g_ControlByKey(controlKey) = controlVm
    m_RegisterControl = True
End Function

Public Function m_RegisterShapeRoute( _
    ByVal shapeName As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    Dim shapeKey As String
    Dim entry As Object

    shapeKey = LCase$(Trim$(shapeName))
    controlKey = LCase$(Trim$(controlKey))
    methodName = Trim$(methodName)

    If Len(shapeKey) = 0 Then
        MsgBox "ShapeClickDispatcher: shape name is empty.", vbExclamation
        Exit Function
    End If
    If Len(controlKey) = 0 Then
        MsgBox "ShapeClickDispatcher: control key is empty for shape '" & shapeName & "'.", vbExclamation
        Exit Function
    End If
    If Len(methodName) = 0 Then
        MsgBox "ShapeClickDispatcher: method name is empty for shape '" & shapeName & "'.", vbExclamation
        Exit Function
    End If

    mp_EnsureStorage
    If Not g_ControlByKey.Exists(controlKey) Then
        MsgBox "ShapeClickDispatcher: control '" & controlKey & "' is not registered for shape '" & shapeName & "'.", vbExclamation
        Exit Function
    End If

    Set entry = CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("ControlKey") = controlKey
    entry("MethodName") = methodName
    entry("HasArg") = CBool(hasArg)
    If hasArg Then
        entry("ArgValue") = argValue
    Else
        entry("ArgValue") = Empty
    End If

    Set g_RouteByShape(shapeKey) = entry
    m_RegisterShapeRoute = True
End Function

Public Function m_UnregisterControl(ByVal controlKey As String) As Boolean
    Dim routeKey As Variant
    Dim routeEntry As Object
    Dim controlKeyNorm As String
    Dim routeKeysToRemove As Collection
    Dim removeKey As Variant

    controlKeyNorm = LCase$(Trim$(controlKey))
    If Len(controlKeyNorm) = 0 Then
        MsgBox "ShapeClickDispatcher: control key is empty.", vbExclamation
        Exit Function
    End If

    mp_EnsureStorage

    If g_ControlByKey.Exists(controlKeyNorm) Then
        g_ControlByKey.Remove controlKeyNorm
    End If

    Set routeKeysToRemove = New Collection
    For Each routeKey In g_RouteByShape.Keys
        Set routeEntry = g_RouteByShape(routeKey)
        If LCase$(Trim$(CStr(routeEntry("ControlKey")))) = controlKeyNorm Then
            routeKeysToRemove.Add CStr(routeKey)
        End If
    Next routeKey

    For Each removeKey In routeKeysToRemove
        g_RouteByShape.Remove CStr(removeKey)
    Next removeKey

    m_UnregisterControl = True
End Function

' Универсальная точка входа для Shape.OnAction.
Public Sub m_OnShapeClick()
    Dim callerShapeName As String
    Dim routeEntry As Object
    Dim controlKey As String
    Dim methodName As String
    Dim hasArg As Boolean
    Dim argValue As Variant
    Dim controlVm As Object
    Dim actionOk As Boolean
    Dim hasSelectContext As Boolean

    On Error Resume Next
    callerShapeName = CStr(Application.Caller)
    On Error GoTo 0

    If Len(Trim$(callerShapeName)) = 0 Then Exit Sub
    If Not mp_TryGetShapeRoute(callerShapeName, routeEntry) Then
        Call mp_TryCloseOpenSelectPopups(vbNullString)
        Exit Sub
    End If

    controlKey = LCase$(Trim$(CStr(routeEntry("ControlKey"))))
    methodName = Trim$(CStr(routeEntry("MethodName")))
    hasArg = CBool(routeEntry("HasArg"))
    If hasArg Then
        argValue = routeEntry("ArgValue")
    Else
        argValue = Empty
    End If

    If Not mp_TryGetControl(controlKey, controlVm) Then
        ' Контрол недоступен: чистим висячий route и выходим.
        mp_RemoveShapeRoute callerShapeName
        Exit Sub
    End If

    ' Перед выполнением действия закрываем dropdown у других select-контролов.
    ' Для select, по которому кликнули, закрытие не делаем:
    ' header должен корректно toggle-иться, item должен обработать выбор.
    If LCase$(TypeName(controlVm)) = "obj_selectcontrolvm" Then
        If Not mp_TryCloseOpenSelectPopups(controlKey) Then Exit Sub
    Else
        If Not mp_TryCloseOpenSelectPopups(vbNullString) Then Exit Sub
    End If

    If Not mp_TryInvokeControlAction(controlVm, methodName, hasArg, argValue, actionOk) Then Exit Sub
    If Not actionOk Then Exit Sub

    hasSelectContext = mp_TryUpdateSelectContext(controlVm, methodName, controlKey)
    mp_SetLastDispatchContext callerShapeName, controlKey, methodName, hasArg, argValue, hasSelectContext
End Sub

Public Sub m_SetSelectContextFromVm(ByVal selectId As String, ByVal controlVm As obj_SelectControlVM)
    selectId = LCase$(Trim$(selectId))
    If Len(selectId) = 0 Then Exit Sub
    If controlVm Is Nothing Then Exit Sub
    mp_SetLastSelectContextFromVm controlVm, selectId
End Sub

Public Function m_GetLastSelectedCaption(Optional ByVal defaultValue As String = vbNullString) As String
    If g_LastSelectContext Is Nothing Then
        m_GetLastSelectedCaption = defaultValue
        Exit Function
    End If
    If Not g_LastSelectContext.Exists("SelectedCaption") Then
        m_GetLastSelectedCaption = defaultValue
        Exit Function
    End If
    m_GetLastSelectedCaption = CStr(g_LastSelectContext("SelectedCaption"))
End Function

' === Internal ===
Private Sub mp_EnsureStorage()
    If g_ControlByKey Is Nothing Then
        Set g_ControlByKey = CreateObject("Scripting.Dictionary")
        g_ControlByKey.CompareMode = 1
    End If

    If g_RouteByShape Is Nothing Then
        Set g_RouteByShape = CreateObject("Scripting.Dictionary")
        g_RouteByShape.CompareMode = 1
    End If
End Sub

Private Function mp_TryGetShapeRoute(ByVal shapeName As String, ByRef outEntry As Object) As Boolean
    Dim shapeKey As String

    If g_RouteByShape Is Nothing Then Exit Function

    shapeKey = LCase$(Trim$(shapeName))
    If Len(shapeKey) = 0 Then Exit Function
    If Not g_RouteByShape.Exists(shapeKey) Then Exit Function

    Set outEntry = g_RouteByShape(shapeKey)
    mp_TryGetShapeRoute = True
End Function

Private Sub mp_RemoveShapeRoute(ByVal shapeName As String)
    Dim shapeKey As String

    If g_RouteByShape Is Nothing Then Exit Sub
    shapeKey = LCase$(Trim$(shapeName))
    If Len(shapeKey) = 0 Then Exit Sub
    If g_RouteByShape.Exists(shapeKey) Then
        g_RouteByShape.Remove shapeKey
    End If
End Sub

Private Function mp_TryGetControl(ByVal controlKey As String, ByRef outControl As Object) As Boolean
    If g_ControlByKey Is Nothing Then Exit Function
    If Not g_ControlByKey.Exists(controlKey) Then Exit Function

    Set outControl = g_ControlByKey(controlKey)
    If outControl Is Nothing Then Exit Function

    mp_TryGetControl = True
End Function

Private Function mp_TryCloseOpenSelectPopups(Optional ByVal excludeControlKey As String = vbNullString) As Boolean
    Dim routeControlKey As Variant
    Dim routeControlKeyNorm As String
    Dim routeControlVm As Object
    Dim closeResult As Variant

    routeControlKeyNorm = LCase$(Trim$(excludeControlKey))
    If g_ControlByKey Is Nothing Then
        mp_TryCloseOpenSelectPopups = True
        Exit Function
    End If

    On Error GoTo EH_CLOSE
    For Each routeControlKey In g_ControlByKey.Keys
        If Len(routeControlKeyNorm) > 0 Then
            If LCase$(Trim$(CStr(routeControlKey))) = routeControlKeyNorm Then GoTo ContinueControl
        End If

        Set routeControlVm = g_ControlByKey(routeControlKey)
        If routeControlVm Is Nothing Then GoTo ContinueControl
        If LCase$(TypeName(routeControlVm)) <> "obj_selectcontrolvm" Then GoTo ContinueControl

        closeResult = CallByName(routeControlVm, "m_RuntimeCloseDropdown", VbMethod)
        If VarType(closeResult) = vbBoolean Then
            If Not CBool(closeResult) Then Exit Function
        End If

ContinueControl:
    Next routeControlKey
    On Error GoTo 0

    mp_TryCloseOpenSelectPopups = True
    Exit Function

EH_CLOSE:
    MsgBox "ShapeClickDispatcher: failed to close select dropdowns: " & Err.Description, vbExclamation
End Function

Private Function mp_TryInvokeControlAction( _
    ByVal controlVm As Object, _
    ByVal methodName As String, _
    ByVal hasArg As Boolean, _
    ByVal argValue As Variant, _
    ByRef outActionOk As Boolean _
) As Boolean
    Dim resultValue As Variant

    If controlVm Is Nothing Then Exit Function
    methodName = Trim$(methodName)
    If Len(methodName) = 0 Then
        MsgBox "ShapeClickDispatcher: method name is empty.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_INVOKE
    If hasArg Then
        resultValue = CallByName(controlVm, methodName, VbMethod, argValue)
    Else
        resultValue = CallByName(controlVm, methodName, VbMethod)
    End If
    On Error GoTo 0

    ' Если обработчик вернул Boolean=False, считаем действие неуспешным.
    ' Если вернул Sub/другой тип — трактуем как успешный вызов.
    If VarType(resultValue) = vbBoolean Then
        outActionOk = CBool(resultValue)
    Else
        outActionOk = True
    End If

    mp_TryInvokeControlAction = True
    Exit Function

EH_INVOKE:
    MsgBox "ShapeClickDispatcher: failed to invoke method '" & methodName & "' on '" & TypeName(controlVm) & "': " & Err.Description, vbExclamation
End Function

Private Function mp_TryUpdateSelectContext(ByVal controlVm As Object, ByVal methodName As String, ByVal controlKey As String) As Boolean
    If controlVm Is Nothing Then Exit Function
    If LCase$(TypeName(controlVm)) <> "obj_selectcontrolvm" Then Exit Function

    methodName = LCase$(Trim$(methodName))
    Select Case methodName
        Case "m_runtimehandleitemclick", "m_runtimehandleheaderclick"
            If Not controlVm.m_RuntimeHasSelectedItem() Then Exit Function
            mp_SetLastSelectContextFromVm controlVm, controlKey
            mp_TryUpdateSelectContext = True
    End Select
End Function

Private Sub mp_SetLastDispatchContext( _
    ByVal shapeName As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    ByVal hasArg As Boolean, _
    ByVal argValue As Variant, _
    ByVal hasSelectContext As Boolean _
)
    Set g_LastDispatchContext = CreateObject("Scripting.Dictionary")
    g_LastDispatchContext.CompareMode = 1

    g_LastDispatchContext("ShapeName") = CStr(shapeName)
    g_LastDispatchContext("ControlKey") = CStr(controlKey)
    g_LastDispatchContext("MethodName") = CStr(methodName)
    g_LastDispatchContext("HasArg") = CBool(hasArg)
    g_LastDispatchContext("HasSelectContext") = CBool(hasSelectContext)

    If hasArg Then
        If IsObject(argValue) Then
            g_LastDispatchContext("ArgType") = TypeName(argValue)
        Else
            g_LastDispatchContext("ArgValue") = CStr(argValue)
        End If
    End If
End Sub

Private Sub mp_SetLastSelectContextFromVm(ByVal controlVm As obj_SelectControlVM, ByVal selectId As String)
    If controlVm Is Nothing Then Exit Sub
    If Not controlVm.m_RuntimeHasSelectedItem() Then Exit Sub

    Set g_LastSelectContext = CreateObject("Scripting.Dictionary")
    g_LastSelectContext.CompareMode = 1
    g_LastSelectContext("SelectId") = LCase$(Trim$(selectId))
    g_LastSelectContext("ControlName") = CStr(controlVm.m_GetControlKey())
    g_LastSelectContext("SelectedIndex") = CLng(controlVm.m_RuntimeGetSelectedIndex())
    g_LastSelectContext("SelectedCaption") = CStr(controlVm.m_RuntimeGetSelectedCaption())
    g_LastSelectContext("SelectedId") = CStr(controlVm.m_RuntimeGetSelectedId())
End Sub
