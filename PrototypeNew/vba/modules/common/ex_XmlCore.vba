Attribute VB_Name = "ex_XmlCore"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const UI_NS As String = "urn:excelprototype:profiles"

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_XmlCore.fn_Module_Dispose"
#End If
End Sub
' //
' // API
' //
Public Function fn_CombineBasePath(ByVal wb As Workbook, ByVal relPath As String) As String
    Dim basePath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    basePath = wb.Path
    If VBA.Len(basePath) = 0 Then basePath = CurDir$

    fn_CombineBasePath = basePath & "\" & relPath
End Function


Public Function fn_CreateDom(Optional ByVal nsUri As String = VBA.vbNullString) As Object
    Dim doc As Object

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If VBA.Len(VBA.Trim$(nsUri)) > 0 Then
        doc.setProperty "SelectionNamespaces", "xmlns:p='" & nsUri & "'"
    End If

    Set fn_CreateDom = doc
End Function


Public Function fn_LoadDomByRelativePath( _
    ByVal wb As Workbook, _
    ByVal relPath As String, _
    ByVal missingPrefix As String, _
    ByVal parsePrefix As String, _
    Optional ByVal nsUri As String = UI_NS) As Object

    Dim filePath As String

    filePath = fn_CombineBasePath(wb, relPath)
    If VBA.Len(filePath) = 0 Then Exit Function

    Set fn_LoadDomByRelativePath = fn_LoadDomByFilePath(filePath, missingPrefix, parsePrefix, nsUri)
End Function


Public Function fn_LoadDomByFilePath( _
    ByVal filePath As String, _
    ByVal missingPrefix As String, _
    ByVal parsePrefix As String, _
    Optional ByVal nsUri As String = VBA.vbNullString) As Object

    Dim doc As Object

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function

    If Not private_FileExists(filePath) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError missingPrefix & filePath
#End If
        Exit Function
    End If

    Set doc = fn_CreateDom(nsUri)
    If Not doc.Load(filePath) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError parsePrefix & filePath
#End If
        Exit Function
    End If

    Set fn_LoadDomByFilePath = doc
End Function


Private Function private_FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object

    On Error GoTo EH
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Для путей вида `[4] controls` избегаем Dir(...), чтобы `[`/`]`
    ' трактовались как обычные символы имени папки.
    private_FileExists = fso.FileExists(filePath)
    Exit Function

EH:
    private_FileExists = False
End Function


Public Function fn_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    fn_NodeAttrText = VBA.CStr(node.Attributes.getNamedItem(attrName).Text)
    If Err.Number <> 0 Then
        Err.Clear
        fn_NodeAttrText = VBA.vbNullString
    End If
    On Error GoTo 0

    If VBA.Len(fn_NodeAttrText) > 0 Then Exit Function

    On Error Resume Next
    fn_NodeAttrText = VBA.CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        fn_NodeAttrText = VBA.vbNullString
    End If
    On Error GoTo 0
End Function


Public Function fn_XPathLiteral(ByVal value As String) As String
    Dim parts() As String
    Dim i As Long
    Dim dq As String

    dq = VBA.Chr$(34)

    If VBA.InStr(value, "'") = 0 Then
        fn_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If VBA.InStr(value, dq) = 0 Then
        fn_XPathLiteral = dq & value & dq
        Exit Function
    End If

    parts = VBA.Split(value, "'")
    fn_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            fn_XPathLiteral = fn_XPathLiteral & ", " & dq & "'" & dq & " , "
        End If
        fn_XPathLiteral = fn_XPathLiteral & "'" & parts(i) & "'"
    Next i
    fn_XPathLiteral = fn_XPathLiteral & ")"
End Function
