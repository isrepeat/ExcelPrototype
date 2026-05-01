Attribute VB_Name = "ex_XmlCore"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const UI_NS As String = "urn:excelprototype:profiles"

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_XmlCore.m_Module_Dispose"
#End If
End Sub
' //
' // API
' //
Public Function m_CombineBasePath(ByVal wb As Workbook, ByVal relPath As String) As String
    Dim basePath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    basePath = wb.Path
    If VBA.Len(basePath) = 0 Then basePath = CurDir$

    m_CombineBasePath = basePath & "\" & relPath
End Function


Public Function m_CreateDom(Optional ByVal nsUri As String = VBA.vbNullString) As Object
    Dim doc As Object

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If VBA.Len(VBA.Trim$(nsUri)) > 0 Then
        doc.setProperty "SelectionNamespaces", "xmlns:p='" & nsUri & "'"
    End If

    Set m_CreateDom = doc
End Function


Public Function m_LoadDomByRelativePath( _
    ByVal wb As Workbook, _
    ByVal relPath As String, _
    ByVal missingPrefix As String, _
    ByVal parsePrefix As String, _
    Optional ByVal nsUri As String = UI_NS) As Object

    Dim filePath As String

    filePath = m_CombineBasePath(wb, relPath)
    If VBA.Len(filePath) = 0 Then Exit Function

    Set m_LoadDomByRelativePath = m_LoadDomByFilePath(filePath, missingPrefix, parsePrefix, nsUri)
End Function


Public Function m_LoadDomByFilePath( _
    ByVal filePath As String, _
    ByVal missingPrefix As String, _
    ByVal parsePrefix As String, _
    Optional ByVal nsUri As String = VBA.vbNullString) As Object

    Dim doc As Object

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function

    If VBA.Len(Dir(filePath)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError missingPrefix & filePath
#End If
        Exit Function
    End If

    Set doc = m_CreateDom(nsUri)
    If Not doc.Load(filePath) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError parsePrefix & filePath
#End If
        Exit Function
    End If

    Set m_LoadDomByFilePath = doc
End Function


Public Function m_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    m_NodeAttrText = VBA.CStr(node.Attributes.getNamedItem(attrName).Text)
    If Err.Number <> 0 Then
        Err.Clear
        m_NodeAttrText = VBA.vbNullString
    End If
    On Error GoTo 0

    If VBA.Len(m_NodeAttrText) > 0 Then Exit Function

    On Error Resume Next
    m_NodeAttrText = VBA.CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        m_NodeAttrText = VBA.vbNullString
    End If
    On Error GoTo 0
End Function


Public Function m_XPathLiteral(ByVal value As String) As String
    Dim parts() As String
    Dim i As Long
    Dim dq As String

    dq = VBA.Chr$(34)

    If VBA.InStr(value, "'") = 0 Then
        m_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If VBA.InStr(value, dq) = 0 Then
        m_XPathLiteral = dq & value & dq
        Exit Function
    End If

    parts = VBA.Split(value, "'")
    m_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            m_XPathLiteral = m_XPathLiteral & ", " & dq & "'" & dq & " , "
        End If
        m_XPathLiteral = m_XPathLiteral & "'" & parts(i) & "'"
    Next i
    m_XPathLiteral = m_XPathLiteral & ")"
End Function
