Attribute VB_Name = "ex_UiPanelVisibilityTests"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Run(ByRef diagnostic As String) As Boolean
    Dim book As Workbook
    Dim previousBook As Workbook
    Dim document As Object
    Dim content As Object
    Dim definition As obj_UiPageDefinition
    Dim context As obj_UiRenderContext
    Dim bindings As obj_UiBindingContext
    Dim element As obj_IUiElement
    Dim fixtures As Variant
    Dim fixture As Variant

    On Error GoTo EH
    Set previousBook = Application.ActiveWorkbook
    Set book = Application.Workbooks.Add
    Set bindings = New obj_UiBindingContext
    If Not bindings.Initialize() Then
        Err.Raise 5, , "Cannot initialize fixture bindings."
    End If
    If Not bindings.SetValue("Form", "Visibility", "Visible") Then
        Err.Raise 5, , "Cannot initialize fixture visibility."
    End If
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    document.LoadXML "<page xmlns='urn:excelprototype:profiles'/>"
    Set definition = New obj_UiPageDefinition
    If Not definition.Initialize(document, ThisWorkbook.Path & "\visibility-test.xaml") Then
        Err.Raise 5, , "Cannot initialize fixture page."
    End If
    Set context = New obj_UiRenderContext
    If Not context.Initialize(book.Worksheets(1), definition, ThisWorkbook.Path, bindings) Then
        Err.Raise 5, , "Cannot initialize fixture context."
    End If
    fixtures = Array( _
        "<stackPanel xmlns='urn:excelprototype:profiles' dataContext='{Binding Path=Form}' " & _
        "orientation='vertical' row='4' column='5' visibility='{Binding Path=Visibility}'>" & _
        "<grid name='Nested'><stackPanel orientation='vertical' horizontalAlignment='stretch' " & _
        "columnSpan='2' rowSpan='3'/></grid></stackPanel>", _
        "<grid xmlns='urn:excelprototype:profiles' dataContext='{Binding Path=Form}' " & _
        "row='4' column='5' visibility='{Binding Path=Visibility}'>" & _
        "<grid.rowDefinitions><rowDefinition size='3'/></grid.rowDefinitions>" & _
        "<grid.columnDefinitions><columnDefinition size='2'/></grid.columnDefinitions>" & _
        "<stackPanel name='Nested' orientation='vertical'/></grid>")
    For Each fixture In fixtures
        document.LoadXML CStr(fixture)
        Set element = ex_UiElementFactory.fn_Create(document.documentElement, context, vbNullString, diagnostic)
        If element Is Nothing Then
            Err.Raise 5, , diagnostic
        End If
        private_Expect element, 6, 6
        If Not bindings.SetValue("Form", "Visibility", "Collapsed") Then
            Err.Raise 5, , "Cannot collapse fixture."
        End If
        private_Expect element, 0, 0
        private_Expect element, 0, 0
        If Not bindings.SetValue("Form", "Visibility", "Visible") Then
            Err.Raise 5, , "Cannot expand fixture."
        End If
        ' Recreating Nested also checks that collapsed descendants release their names.
        private_Expect element, 6, 6
        element.Dispose
        Set element = Nothing
    Next fixture
    Set content = CreateObject("MSXML2.DOMDocument.6.0")
    content.LoadXML "<fragment xmlns='urn:excelprototype:profiles'><grid name='ContentGrid'/></fragment>"
    If Not bindings.SetObject("Data", "Content", content.documentElement) Then
        Err.Raise 5, , "Cannot initialize fixture content."
    End If
    document.LoadXML "<stackPanel xmlns='urn:excelprototype:profiles' orientation='vertical' " & _
        "contentSource='Content' visibility='{Binding Path=Form.Visibility}'/>"
    Set element = ex_UiElementFactory.fn_Create(document.documentElement, context, vbNullString, diagnostic)
    If element Is Nothing Then
        Err.Raise 5, , diagnostic
    End If
    private_Expect element, 1, 1
    If Not bindings.SetValue("Form", "Visibility", "Collapsed") Then
        Err.Raise 5, , "Cannot collapse content fixture."
    End If
    private_Expect element, 0, 0
    content.LoadXML "<fragment xmlns='urn:excelprototype:profiles'><stackPanel name='ContentGrid' " & _
        "orientation='vertical' horizontalAlignment='stretch' columnSpan='4' rowSpan='2'/></fragment>"
    If Not bindings.SetObject("Data", "Content", content.documentElement) Then
        Err.Raise 5, , "Cannot replace hidden content."
    End If
    If Not bindings.SetValue("Form", "Visibility", "Visible") Then
        Err.Raise 5, , "Cannot expand content fixture."
    End If
    private_Expect element, 2, 4
    fn_Run = True
Cleanup:
    On Error Resume Next
    If Not element Is Nothing Then
        element.Dispose
    End If
    If Not context Is Nothing Then
        context.Dispose
    End If
    If Not bindings Is Nothing Then
        bindings.Dispose
    End If
    If Not definition Is Nothing Then
        definition.Dispose
    End If
    If Not book Is Nothing Then
        book.Close False
    End If
    If Not previousBook Is Nothing Then
        previousBook.Activate
    End If
    Exit Function
EH:
    diagnostic = Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_Expect( _
    ByVal element As obj_IUiElement, _
    ByVal expectedRows As Long, _
    ByVal expectedColumns As Long _
)
    Dim rows As Long
    Dim columns As Long
    Dim diagnostic As String

    If Not element.Measure(rows, columns, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If rows <> expectedRows Or columns <> expectedColumns Then
        Err.Raise 5, , "Unexpected container size: " & CStr(rows) & " x " & CStr(columns)
    End If
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------