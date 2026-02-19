Attribute VB_Name = "ex_ProfilesStore"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const PRESETS_TEMPLATE As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><presets xmlns=""" & PRESETS_NS & """ version=""1""/>"

Public Function m_GetProfilesFilePath(Optional ByVal modeName As String = vbNullString, Optional ByVal wb As Workbook) As String
    Dim resolvedMode As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    resolvedMode = Trim$(modeName)
    If Len(resolvedMode) = 0 Then
        resolvedMode = "Personal Card"
    End If

    m_GetProfilesFilePath = ex_UiXmlProvider.m_GetProfilesFilePathByMode(resolvedMode, wb, "profilesByMode")
End Function

Public Function m_LoadPresetsDom(ByVal filePath As String) As Object
    Dim doc As Object

    If Len(Trim$(filePath)) = 0 Then Exit Function

    Set doc = ex_XmlCore.m_CreateDom(PRESETS_NS)

    If Len(Dir(filePath)) > 0 Then
        If Not doc.Load(filePath) Then
            doc.loadXML PRESETS_TEMPLATE
        End If
    Else
        doc.loadXML PRESETS_TEMPLATE
    End If

    Set m_LoadPresetsDom = doc
End Function

Public Sub m_SavePresetsDom(ByVal doc As Object, ByVal filePath As String)
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

    Set node = doc.selectSingleNode("/p:presets/p:profile[@name=" & ex_XmlCore.m_XPathLiteral(profileName) & "]")
    If node Is Nothing And createIfMissing Then
        Set root = doc.selectSingleNode("/p:presets")
        If root Is Nothing Then Exit Function
        Set node = doc.createNode(1, "profile", PRESETS_NS)
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
