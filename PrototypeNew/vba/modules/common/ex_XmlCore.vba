Attribute VB_Name = "ex_XmlCore"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"

Public Function m_CombineBasePath(ByVal wb As Workbook, ByVal relPath As String) As String
    Dim basePath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    basePath = wb.Path
    If Len(basePath) = 0 Then basePath = CurDir$

    m_CombineBasePath = basePath & "\" & relPath
End Function

Public Function m_CreateDom(Optional ByVal nsUri As String = vbNullString) As Object
    Dim doc As Object

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False

    If Len(Trim$(nsUri)) > 0 Then
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
    If Len(filePath) = 0 Then Exit Function

    Set m_LoadDomByRelativePath = m_LoadDomByFilePath(filePath, missingPrefix, parsePrefix, nsUri)
End Function

Public Function m_LoadDomByFilePath( _
    ByVal filePath As String, _
    ByVal missingPrefix As String, _
    ByVal parsePrefix As String, _
    Optional ByVal nsUri As String = vbNullString) As Object

    Dim doc As Object

    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then Exit Function

    If Len(Dir(filePath)) = 0 Then
        MsgBox missingPrefix & filePath, vbExclamation
        Exit Function
    End If

    Set doc = m_CreateDom(nsUri)
    If Not doc.Load(filePath) Then
        MsgBox parsePrefix & filePath, vbExclamation
        Exit Function
    End If

    Set m_LoadDomByFilePath = doc
End Function

Public Function m_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    m_NodeAttrText = CStr(node.Attributes.getNamedItem(attrName).Text)
    If Err.Number <> 0 Then
        Err.Clear
        m_NodeAttrText = vbNullString
    End If
    On Error GoTo 0

    If Len(m_NodeAttrText) > 0 Then Exit Function

    On Error Resume Next
    m_NodeAttrText = CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        m_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

Public Function m_XPathLiteral(ByVal value As String) As String
    Dim parts() As String
    Dim i As Long
    Dim dq As String

    dq = Chr$(34)

    If InStr(value, "'") = 0 Then
        m_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If InStr(value, dq) = 0 Then
        m_XPathLiteral = dq & value & dq
        Exit Function
    End If

    parts = Split(value, "'")
    m_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            m_XPathLiteral = m_XPathLiteral & ", " & dq & "'" & dq & " , "
        End If
        m_XPathLiteral = m_XPathLiteral & "'" & parts(i) & "'"
    Next i
    m_XPathLiteral = m_XPathLiteral & ")"
End Function
