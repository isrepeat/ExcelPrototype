Attribute VB_Name = "ex_CustomXmlPartStore"
Option Explicit

Public Function m_TryFindPartByNamespace( _
    ByVal namespaceUri As String, _
    ByRef outPart As Object _
) As Boolean
    Dim parts As Object

    namespaceUri = Trim$(namespaceUri)
    If Len(namespaceUri) = 0 Then
        MsgBox "CustomXmlPartStore: namespace is empty.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_FIND
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(namespaceUri)
    On Error GoTo 0

    If Not parts Is Nothing Then
        If parts.Count > 0 Then
            Set outPart = parts(1)
        End If
    End If

    m_TryFindPartByNamespace = True
    Exit Function

EH_FIND:
    MsgBox "CustomXmlPartStore: failed to find XML part by namespace '" & namespaceUri & "': " & Err.Description, vbExclamation
End Function

Public Function m_TryLoadDomFromXml( _
    ByVal xmlText As String, _
    ByRef outDom As Object _
) As Boolean
    Dim dom As Object

    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.setProperty "SelectionLanguage", "XPath"

    If Not dom.LoadXML(CStr(xmlText)) Then
        MsgBox "CustomXmlPartStore: failed to parse XML.", vbExclamation
        Exit Function
    End If

    Set outDom = dom
    m_TryLoadDomFromXml = True
End Function

Public Function m_TryCreateEmptyDom( _
    ByVal rootNodeName As String, _
    ByVal namespaceUri As String, _
    ByRef outDom As Object _
) As Boolean
    Dim xmlText As String

    rootNodeName = Trim$(rootNodeName)
    namespaceUri = Trim$(namespaceUri)

    If Len(rootNodeName) = 0 Then
        MsgBox "CustomXmlPartStore: root node name is empty.", vbExclamation
        Exit Function
    End If
    If Len(namespaceUri) = 0 Then
        MsgBox "CustomXmlPartStore: namespace is empty.", vbExclamation
        Exit Function
    End If

    xmlText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
              "<" & rootNodeName & " xmlns=""" & namespaceUri & """></" & rootNodeName & ">"

    If Not m_TryLoadDomFromXml(xmlText, outDom) Then Exit Function
    m_TryCreateEmptyDom = True
End Function

Public Function m_TryLoadPartDom( _
    ByVal partObj As Object, _
    ByRef outDom As Object _
) As Boolean
    If partObj Is Nothing Then
        MsgBox "CustomXmlPartStore: part is not specified.", vbExclamation
        Exit Function
    End If

    If Not m_TryLoadDomFromXml(CStr(partObj.XML), outDom) Then Exit Function
    m_TryLoadPartDom = True
End Function

Public Function m_TrySaveDom( _
    ByVal dom As Object, _
    ByVal existingPart As Object _
) As Boolean
    Dim xmlText As String

    If dom Is Nothing Then
        MsgBox "CustomXmlPartStore: DOM is not specified.", vbExclamation
        Exit Function
    End If

    xmlText = CStr(dom.XML)
    If Len(Trim$(xmlText)) = 0 Then
        MsgBox "CustomXmlPartStore: state XML is empty.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH_SAVE
    If Not existingPart Is Nothing Then existingPart.Delete
    ThisWorkbook.CustomXMLParts.Add xmlText
    m_TrySaveDom = True
    Exit Function

EH_SAVE:
    MsgBox "CustomXmlPartStore: failed to persist state XML: " & Err.Description, vbExclamation
End Function

