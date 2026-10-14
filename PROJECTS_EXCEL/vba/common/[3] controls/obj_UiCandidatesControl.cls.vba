VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiCandidatesControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_tableControl As obj_UiTableControl
Private m_control As obj_IUiControl
Private m_definition As Object
Private m_context As obj_UiRenderContext
Private m_column As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiControl_Initialize() As Boolean
    obj_IUiControl_Initialize = Me.Initialize()
End Function

Private Sub obj_IUiControl_Dispose()
    Me.Dispose
End Sub

Private Function obj_IUiControl_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    Set m_definition = definition
    Set m_context = context
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    obj_IUiControl_Configure = m_control.Configure(definition, context, source, diagnostic)
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If Not private_PrepareSource(diagnostic) Then Exit Function
    obj_IUiControl_Measure = m_control.Measure(rows, columns, diagnostic)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If Not private_PrepareSource(diagnostic) Then Exit Function
    obj_IUiControl_Arrange = m_control.Arrange(row, column, diagnostic)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    obj_IUiControl_Render = m_control.Render(diagnostic)
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    obj_IUiControl_Validate = m_control.Validate(errors)
End Function

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_tableControl = New obj_UiTableControl
    Set m_control = m_tableControl
    If Not m_control.Initialize() Then
        Me.Dispose
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_control Is Nothing Then m_control.Dispose
    Set m_control = Nothing
    Set m_tableControl = Nothing
    Set m_definition = Nothing
    Set m_context = Nothing
End Sub

' //
' // Private
' //
Private Function private_PrepareSource(ByRef diagnostic As String) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim tableSource As obj_IUiTableSource
    Dim selectionSource As obj_IUiSelectionSource
    Dim table As obj_UiRawTable
    Dim candidatesSource As obj_UiCandidatesSource
    Dim columnMetadata As obj_IUiColumnMetadata
    Dim sourceAliases As Object
    Dim columnIndex As Long

    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
            private_ReadAttribute(m_definition, "itemsSource"), m_context.BindingContext, value, sourceObject, isObject) Then
        diagnostic = "Cannot resolve candidates source."
        Exit Function
    End If
    If Not isObject Then
        diagnostic = "Candidates source must be an object."
        Exit Function
    End If
    If Not TypeOf sourceObject Is obj_IUiTableSource Or Not TypeOf sourceObject Is obj_IUiSelectionSource Then
        diagnostic = "Candidates source requires table and selection interfaces."
        Exit Function
    End If
    Set tableSource = sourceObject
    Set selectionSource = sourceObject
    If tableSource.TableCount <> 1 Then
        diagnostic = "Candidates source must provide exactly one table."
        Exit Function
    End If
    Set table = tableSource.GetTable(1)
    If table Is Nothing Then
        diagnostic = "Candidates source returned no table."
        Exit Function
    End If
    Set sourceAliases = VBA.CreateObject("Scripting.Dictionary")
    If TypeOf sourceObject Is obj_IUiColumnMetadata Then
        Set columnMetadata = sourceObject
        For columnIndex = 1 To table.ColumnCount
            sourceAliases(VBA.CStr(table.HeaderAt(columnIndex))) = columnMetadata.SourceAliasAt(columnIndex)
        Next columnIndex
    End If
    If Not private_AlignToForm(table, diagnostic) Then Exit Function
    Set candidatesSource = New obj_UiCandidatesSource
    If Not candidatesSource.Initialize(table, selectionSource, sourceAliases) Then
        diagnostic = "Cannot initialize candidates source."
        Exit Function
    End If
    m_tableControl.SetSource candidatesSource
    private_PrepareSource = True
End Function

Private Function private_AlignToForm( _
    ByRef table As obj_UiRawTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim formName As String
    Dim formNode As Object
    Dim node As Object
    Dim fieldNode As Object
    Dim headers As Variant
    Dim values As Variant
    Dim caption As Variant
    Dim captionObject As Object
    Dim isObject As Boolean
    Dim projected As obj_UiRawTable
    Dim fieldCount As Long
    Dim position As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim matched As Object

    formName = private_ReadAttribute(m_definition, "alignToForm")
    On Error GoTo EH
    If VBA.Len(formName) = 0 Then Err.Raise 5, , "Candidates control requires alignToForm."
    If Not m_context.TryGetFormDefinition(formName, formNode) Then
        Err.Raise 5, , "Alignment form was not found: " & formName
    End If
    If private_ReadAttribute(formNode, "orientation") <> "horizontal" Or _
            private_ReadAttribute(formNode, "labelPosition") <> "top" Then
        Err.Raise 5, , "Table alignment requires a horizontal form with top labels: " & formName
    End If
    If private_ReadLong(formNode, "column", 1) <> m_column Then
        Err.Raise 5, , "Aligned table and form must start in the same column: " & formName
    End If
    For Each node In formNode.ChildNodes
        If node.NodeType = 1 Then
            If node.baseName <> "field" Then Err.Raise 5, , "Aligned form only supports fields: " & formName
            fieldCount = fieldCount + 1
        End If
    Next node
    If fieldCount = 0 Then Err.Raise 5, , "Alignment form has no fields: " & formName
    ReDim headers(0 To fieldCount - 1)
    If table.RowCount > 0 Then ReDim values(1 To table.RowCount, 1 To fieldCount)
    Set matched = VBA.CreateObject("Scripting.Dictionary")
    ' Match headers to field labels; leave extra form fields blank.
    For Each fieldNode In formNode.ChildNodes
        If fieldNode.NodeType = 1 Then
            If VBA.Len(private_ReadAttribute(fieldNode, "labelPosition")) > 0 And _
                    private_ReadAttribute(fieldNode, "labelPosition") <> "top" Then
                Err.Raise 5, , "Aligned field requires a top label: " & private_ReadAttribute(fieldNode, "name")
            End If
            position = position + 1
            headers(position - 1) = VBA.vbNullString
            If Not ex_UiBindingRuntime.fn_TryResolveValue( _
                    private_ReadAttribute(fieldNode, "label"), m_context.BindingContext, caption, captionObject, isObject) Then
                Err.Raise 5, , "Cannot resolve alignment field label: " & private_ReadAttribute(fieldNode, "name")
            End If
            If isObject Then Err.Raise 5, , "Alignment field label must be text."
            For columnIndex = 1 To table.ColumnCount
                If VBA.Len(VBA.CStr(caption)) > 0 And VBA.CStr(table.HeaderAt(columnIndex)) = VBA.CStr(caption) Then
                    If matched.Exists(columnIndex) Or VBA.Len(headers(position - 1)) > 0 Then
                        Err.Raise 5, , "Ambiguous alignment header: " & VBA.CStr(caption)
                    End If
                    matched.Add columnIndex, True
                    headers(position - 1) = table.HeaderAt(columnIndex)
                    For rowIndex = 1 To table.RowCount
                        values(rowIndex, position) = table.ValueAt(rowIndex, columnIndex)
                    Next rowIndex
                End If
            Next columnIndex
        End If
    Next fieldNode
    For columnIndex = 1 To table.ColumnCount
        If VBA.Len(VBA.CStr(table.HeaderAt(columnIndex))) > 0 Then
            If Not matched.Exists(columnIndex) Then Err.Raise 5, , "Candidate header has no matching form field: " & table.HeaderAt(columnIndex)
        End If
    Next columnIndex
    Set projected = New obj_UiRawTable
    If table.RowCount = 0 Then
        If Not projected.InitializeEmpty(headers, table.Title) Then Exit Function
    Else
        If Not projected.Initialize(values, headers, table.Title) Then Exit Function
    End If
    Set table = projected
    private_AlignToForm = True
    Exit Function
EH:
    diagnostic = "Candidates alignment (" & formName & "): " & VBA.Err.Description
End Function

Private Function private_ReadAttribute( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    private_ReadAttribute = ex_UiElementFactory.fn_Attribute(node, name)
End Function

Private Function private_ReadLong( _
    ByVal node As Object, _
    ByVal name As String, _
    ByVal defaultValue As Long _
) As Long
    private_ReadLong = ex_UiElementFactory.fn_Long(node, name, defaultValue)
End Function