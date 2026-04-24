Attribute VB_Name = "ex_RuntimeSourceResolver"
Option Explicit

Private Const PAGE_RUNTIME_SOURCE_ARG As String = "PageRuntimeSource"
Private Const GLOBAL_RUNTIME_SOURCE_ARG As String = "GlobalRuntimeSource"
Private Const SETTINGS_RUNTIME_SOURCE_KEY As String = "Settings"
Private Const RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_NONE As Long = 0
Private Const RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_PAGE As Long = 1
Private Const RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_GLOBAL As Long = 2

' //
' // API
' //
' Callstack[1]: ex_LayoutListRenderer.private_RenderLayoutControl -> ex_RuntimeSourceResolver.m_TryResolveItemsSource
' Callstack[2]: obj_SelectControlVM.private_ResolveItems -> ex_RuntimeSourceResolver.m_TryResolveItemsSource
' Callstack[3]: obj_ConfigControlVM.private_ResolveItems -> ex_RuntimeSourceResolver.m_TryResolveItemsSource
Public Function m_TryResolveItemsSource( _
    ByVal runtimeSources As obj_PageRuntimeSources, _
    ByVal rawSource As String, _
    ByRef outItems As Collection _
) As Boolean
    Dim sourceMap As Object
    Dim resolvedValue As Variant
    Dim sourceKey As String
    Dim runtimeSourceBindingType As Long
    Dim runtimeSourceKey As String

    Set outItems = Nothing

    If runtimeSources Is Nothing Then
        VBA.MsgBox "PrototypeNew: runtime sources are not specified for itemsSource resolve.", VBA.vbExclamation
        Exit Function
    End If

    rawSource = VBA.Trim$(rawSource)
    If VBA.Len(rawSource) = 0 Then
        VBA.MsgBox "PrototypeNew: list itemsSource is required.", VBA.vbExclamation
        Exit Function
    End If

    If Not private_TryExtractRuntimeSourceBinding(rawSource, runtimeSourceBindingType, runtimeSourceKey) Then Exit Function

    If runtimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_NONE Then
        ' itemsSource должен быть runtime source expression
        ' или Binding, который резолвится сразу в Collection.
        If Not private_IsBindingExpression(rawSource) Then
            VBA.MsgBox "PrototypeNew: list itemsSource must use runtime source expression ({PageRuntimeSource='...'} / {GlobalRuntimeSource='...'}) or Binding that resolves to Collection.", VBA.vbExclamation
            Exit Function
        End If

        Set sourceMap = runtimeSources.ItemsSourceMap
        If sourceMap Is Nothing Then
            VBA.MsgBox "PrototypeNew: page itemsSource map is not initialized.", VBA.vbExclamation
            Exit Function
        End If

        If Not ex_BindingRuntime.m_TryResolveValueBinding(rawSource, sourceMap, resolvedValue) Then Exit Function
        If VBA.IsObject(resolvedValue) Then
            If VBA.TypeName(resolvedValue) <> "Collection" Then
                VBA.MsgBox "PrototypeNew: list itemsSource must resolve to Collection.", VBA.vbExclamation
                Exit Function
            End If

            Set outItems = resolvedValue
            m_TryResolveItemsSource = True
            Exit Function
        End If

        VBA.MsgBox "PrototypeNew: list itemsSource Binding must resolve to Collection object.", VBA.vbExclamation
        Exit Function
    End If

    sourceKey = VBA.LCase$(runtimeSourceKey)

    Select Case runtimeSourceBindingType
        Case RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_PAGE
            m_TryResolveItemsSource = runtimeSources.TryGetItemsSourceByKey(sourceKey, outItems, False)
            Exit Function

        Case RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_GLOBAL
            m_TryResolveItemsSource = ex_Core.m_RuntimeSource_TryGetGlobalItemsSourceByKey(sourceKey, outItems, False)
            Exit Function
    End Select

    VBA.MsgBox "PrototypeNew: unsupported itemsSource runtime source type.", VBA.vbExclamation
End Function


' Callstack[1]: obj_ControlBase.TryResolveDataContext -> ex_RuntimeSourceResolver.m_TryResolveObjectSource
' Callstack[2]: ex_LayoutItemControlRenderer.private_TryResolveObjectSourceByText -> ex_RuntimeSourceResolver.m_TryResolveObjectSource
Public Function m_TryResolveObjectSource( _
    ByVal runtimeSources As obj_PageRuntimeSources, _
    ByVal rawSource As String, _
    ByRef outObject As Object, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim sourceMap As Object
    Dim resolvedValue As Variant
    Dim sourceKey As String
    Dim runtimeSourceBindingType As Long
    Dim runtimeSourceKey As String

    Set outObject = Nothing

    If runtimeSources Is Nothing Then
        VBA.MsgBox "PrototypeNew: runtime sources are not specified for objectSource resolve.", VBA.vbExclamation
        Exit Function
    End If

    rawSource = VBA.Trim$(rawSource)
    If VBA.Len(rawSource) = 0 Then
        If allowMissing Then
            m_TryResolveObjectSource = True
            Exit Function
        End If
        VBA.MsgBox "PrototypeNew: itemControl objectSource is required.", VBA.vbExclamation
        Exit Function
    End If

    If Not private_TryExtractRuntimeSourceBinding(rawSource, runtimeSourceBindingType, runtimeSourceKey) Then Exit Function

    If runtimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_NONE Then
        ' objectSource/dataContext должен быть runtime source expression
        ' или Binding, который резолвится сразу в Object.
        If Not private_IsBindingExpression(rawSource) Then
            VBA.MsgBox "PrototypeNew: objectSource must use runtime source expression ({PageRuntimeSource='...'} / {GlobalRuntimeSource='...'}) or Binding that resolves to object.", VBA.vbExclamation
            Exit Function
        End If

        Set sourceMap = runtimeSources.ObjectSourceMap
        If sourceMap Is Nothing Then
            VBA.MsgBox "PrototypeNew: page objectSource map is not initialized.", VBA.vbExclamation
            Exit Function
        End If

        If Not ex_BindingRuntime.m_TryResolveValueBinding(rawSource, sourceMap, resolvedValue) Then Exit Function
        If VBA.IsObject(resolvedValue) Then
            Set outObject = resolvedValue
            m_TryResolveObjectSource = True
            Exit Function
        End If

        VBA.MsgBox "PrototypeNew: objectSource Binding must resolve to object.", VBA.vbExclamation
        Exit Function
    End If

    sourceKey = VBA.LCase$(runtimeSourceKey)

    Select Case runtimeSourceBindingType
        Case RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_PAGE
            m_TryResolveObjectSource = runtimeSources.TryGetObjectSourceByKey(sourceKey, outObject, allowMissing)
            Exit Function

        Case RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_GLOBAL
            ' settings — special-case:
            ' возвращаем snapshot из Settings.xml (через ex_Core.m_Settings_TryGetObjectSource),
            ' а не объект из глобальной map, чтобы видеть изменения файла без ручной ре-регистрации.
            If VBA.StrComp(sourceKey, SETTINGS_RUNTIME_SOURCE_KEY, VBA.vbTextCompare) = 0 Then
                If ex_Core.m_Settings_TryGetObjectSource(outObject, Not allowMissing) Then
                    m_TryResolveObjectSource = True
                ElseIf allowMissing Then
                    m_TryResolveObjectSource = True
                End If
                Exit Function
            End If

            m_TryResolveObjectSource = ex_Core.m_RuntimeSource_TryGetGlobalObjectSourceByKey(sourceKey, outObject, allowMissing)
            Exit Function
    End Select

    VBA.MsgBox "PrototypeNew: unsupported objectSource runtime source type.", VBA.vbExclamation
End Function


' //
' // Internal
' //
Private Function private_TryExtractRuntimeSourceBinding( _
    ByVal rawSource As String, _
    ByRef outRuntimeSourceBindingType As Long, _
    ByRef outRuntimeSourceKey As String _
) As Boolean
    Dim normalizedText As String
    Dim expressionBody As String
    Dim eqPos As Long
    Dim argName As String
    Dim argValue As String
    Dim quoteChar As String

    outRuntimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_NONE
    outRuntimeSourceKey = VBA.vbNullString

    normalizedText = VBA.Trim$(rawSource)
    If VBA.Len(normalizedText) < 3 Then
        private_TryExtractRuntimeSourceBinding = True
        Exit Function
    End If
    If VBA.Left$(normalizedText, 1) <> "{" Then
        private_TryExtractRuntimeSourceBinding = True
        Exit Function
    End If
    If VBA.Right$(normalizedText, 1) <> "}" Then
        private_TryExtractRuntimeSourceBinding = True
        Exit Function
    End If

    expressionBody = VBA.Trim$(VBA.Mid$(normalizedText, 2, VBA.Len(normalizedText) - 2))
    If VBA.Len(expressionBody) = 0 Then
        private_TryExtractRuntimeSourceBinding = True
        Exit Function
    End If

    eqPos = VBA.InStr(1, expressionBody, "=", VBA.vbBinaryCompare)
    If eqPos <= 1 Then
        private_TryExtractRuntimeSourceBinding = True
        Exit Function
    End If

    argName = VBA.Trim$(VBA.Left$(expressionBody, eqPos - 1))
    Select Case True
        Case VBA.StrComp(argName, PAGE_RUNTIME_SOURCE_ARG, VBA.vbTextCompare) = 0
            outRuntimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_PAGE

        Case VBA.StrComp(argName, GLOBAL_RUNTIME_SOURCE_ARG, VBA.vbTextCompare) = 0
            outRuntimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_GLOBAL

        Case Else
            private_TryExtractRuntimeSourceBinding = True
            Exit Function
    End Select

    argValue = VBA.Trim$(VBA.Mid$(expressionBody, eqPos + 1))
    If VBA.Len(argValue) = 0 Then
        If outRuntimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_PAGE Then
            VBA.MsgBox "PrototypeNew: PageRuntimeSource key is empty.", VBA.vbExclamation
        Else
            VBA.MsgBox "PrototypeNew: GlobalRuntimeSource key is empty.", VBA.vbExclamation
        End If
        Exit Function
    End If

    quoteChar = VBA.Left$(argValue, 1)
    If VBA.Len(argValue) >= 2 And (quoteChar = """" Or quoteChar = "'") And VBA.Right$(argValue, 1) = quoteChar Then
        argValue = VBA.Mid$(argValue, 2, VBA.Len(argValue) - 2)
    End If

    outRuntimeSourceKey = VBA.Trim$(argValue)
    If VBA.Len(outRuntimeSourceKey) = 0 Then
        If outRuntimeSourceBindingType = RUNTIME_SOURCE_BINDING_EXPRESSION_TYPE_PAGE Then
            VBA.MsgBox "PrototypeNew: PageRuntimeSource key is empty.", VBA.vbExclamation
        Else
            VBA.MsgBox "PrototypeNew: GlobalRuntimeSource key is empty.", VBA.vbExclamation
        End If
        Exit Function
    End If

    private_TryExtractRuntimeSourceBinding = True
End Function


Private Function private_IsBindingExpression(ByVal rawText As String) As Boolean
    Dim normalized As String

    normalized = VBA.Trim$(rawText)
    If VBA.Len(normalized) = 0 Then Exit Function

    If VBA.Len(normalized) < 10 Then Exit Function
    If VBA.Left$(normalized, 1) <> "{" Then Exit Function
    If VBA.Right$(normalized, 1) <> "}" Then Exit Function
    If VBA.StrComp(VBA.Left$(normalized, 9), "{Binding ", VBA.vbTextCompare) <> 0 Then Exit Function

    private_IsBindingExpression = True
End Function
