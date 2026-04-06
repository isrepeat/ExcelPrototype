Attribute VB_Name = "ex_ProfilesStore"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const PROFILES_TEMPLATE As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><profiles xmlns=""" & PROFILES_NS & """ version=""1""/>"

Public Function m_GetProfilesFilePath(Optional ByVal modeKey As String = vbNullString, Optional ByVal wb As Workbook) As String
    Dim resolvedModeKey As String
    Dim defaultModeKey As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    resolvedModeKey = Trim$(modeKey)
    If Len(resolvedModeKey) = 0 Then
        defaultModeKey = Trim$(ex_UiXmlProvider.m_GetDefaultModeKey(wb))
        If Len(defaultModeKey) > 0 Then
            resolvedModeKey = defaultModeKey
        Else
            resolvedModeKey = Trim$(ex_UiXmlProvider.m_GetModeKeyByIndex(1, wb))
        End If
    End If

    m_GetProfilesFilePath = ex_UiXmlProvider.m_GetProfilesFilePathByMode(resolvedModeKey, wb, "profilesFileByMode")
End Function

Public Function m_LoadProfilesDom(ByVal filePath As String) As Object
    Dim parseErrorMessage As String
    Dim usedTemplateFallback As Boolean
    Dim isMissingFile As Boolean

    Set m_LoadProfilesDom = m_LoadProfilesDomWithStatus(filePath, parseErrorMessage, usedTemplateFallback, isMissingFile)
End Function

Public Function m_LoadProfilesDomWithStatus( _
    ByVal filePath As String, _
    ByRef parseErrorMessage As String, _
    ByRef usedTemplateFallback As Boolean, _
    ByRef isMissingFile As Boolean _
) As Object
    Dim doc As Object

    parseErrorMessage = vbNullString
    usedTemplateFallback = False
    isMissingFile = False

    If Len(Trim$(filePath)) = 0 Then Exit Function

    Set doc = ex_XmlCore.m_CreateDom(PROFILES_NS)

    If Len(Dir(filePath)) > 0 Then
        If Not doc.Load(filePath) Then
            parseErrorMessage = mp_BuildParseErrorMessage(doc, filePath)
            usedTemplateFallback = True
            doc.loadXML PROFILES_TEMPLATE
        End If
    Else
        isMissingFile = True
        usedTemplateFallback = True
        doc.loadXML PROFILES_TEMPLATE
    End If

    Set m_LoadProfilesDomWithStatus = doc
End Function

Private Function mp_BuildParseErrorMessage(ByVal doc As Object, ByVal filePath As String) As String
    Dim reasonText As String
    Dim sourceText As String
    Dim lineNumber As Long
    Dim linePosition As Long

    On Error Resume Next
    reasonText = Trim$(CStr(doc.parseError.reason))
    sourceText = Trim$(CStr(doc.parseError.srcText))
    lineNumber = CLng(doc.parseError.Line)
    linePosition = CLng(doc.parseError.linepos)
    On Error GoTo 0

    If Len(reasonText) = 0 Then reasonText = "Unknown XML parse error."

    mp_BuildParseErrorMessage = "Failed to parse profiles config file '" & filePath & "': " & reasonText
    If lineNumber > 0 Then
        mp_BuildParseErrorMessage = mp_BuildParseErrorMessage & " (line " & CStr(lineNumber)
        If linePosition > 0 Then
            mp_BuildParseErrorMessage = mp_BuildParseErrorMessage & ", pos " & CStr(linePosition)
        End If
        mp_BuildParseErrorMessage = mp_BuildParseErrorMessage & ")."
    End If
    If Len(sourceText) > 0 Then
        mp_BuildParseErrorMessage = mp_BuildParseErrorMessage & " Source: " & sourceText
    End If
End Function

Public Sub m_SaveProfilesDom(ByVal doc As Object, ByVal filePath As String)
    If doc Is Nothing Then Exit Sub
    If Len(Trim$(filePath)) = 0 Then Exit Sub

    If Len(Dir(filePath)) = 0 Then
        MsgBox "Profiles config file was not found: " & filePath, vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    mp_SaveXmlPretty doc, filePath
    Exit Sub
EH:
    MsgBox "Failed to save profiles config file '" & filePath & "': " & Err.Description, vbExclamation
End Sub

Public Function m_GetProfileNode(ByVal doc As Object, ByVal profileName As String, ByVal createIfMissing As Boolean) As Object
    Dim node As Object
    Dim root As Object

    If doc Is Nothing Then Exit Function
    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Function

    On Error Resume Next
    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PROFILES_NS & "'"
    On Error GoTo 0

    Set node = doc.selectSingleNode("/p:profiles/p:profile[@name=" & ex_XmlCore.m_XPathLiteral(profileName) & "]")
    If node Is Nothing Then
        Set node = doc.selectSingleNode("/*[local-name()='profiles']/*[local-name()='profile'][@name=" & ex_XmlCore.m_XPathLiteral(profileName) & "]")
    End If
    If node Is Nothing And createIfMissing Then
        Set root = doc.selectSingleNode("/p:profiles")
        If root Is Nothing Then
            Set root = doc.selectSingleNode("/*[local-name()='profiles']")
        End If
        If root Is Nothing Then Exit Function
        Set node = doc.createNode(1, "profile", PROFILES_NS)
        node.setAttribute "name", profileName
        root.appendChild node
    End If

    Set m_GetProfileNode = node
End Function

Private Sub mp_SaveXmlPretty(ByVal doc As Object, ByVal filePath As String)
    Dim reader As Object
    Dim writer As Object
    Dim stream As Object
    Dim xmlText As String

    Set writer = CreateObject("MSXML2.MXXMLWriter.6.0")
    writer.omitXMLDeclaration = False
    writer.indent = True
    writer.standalone = True
    writer.encoding = "UTF-8"

    Set reader = CreateObject("MSXML2.SAXXMLReader.6.0")
    Set reader.contentHandler = writer
    Set reader.dtdHandler = writer
    Set reader.errorHandler = writer
    On Error Resume Next
    reader.putProperty "http://xml.org/sax/properties/lexical-handler", writer
    On Error GoTo 0

    reader.parse doc.XML
    xmlText = CStr(writer.output)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText xmlText
    stream.Position = 0
    stream.SaveToFile filePath, 2
    stream.Close
End Sub
