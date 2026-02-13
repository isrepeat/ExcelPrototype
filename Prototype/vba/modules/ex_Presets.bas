Attribute VB_Name = "ex_Presets"
Option Explicit

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const PRESETS_TEMPLATE As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><presets xmlns=""" & PRESETS_NS & """ version=""1""/>"
Private Const PRESETS_REL_PATH As String = "profiles\config_profiles.xml"

' Apply selected profile values into column B (values only, no formulas).
Public Sub m_ApplyProfileFromDev(Optional ByVal profileName As String = vbNullString)

    Dim ws As Worksheet
    Dim startRow As Long
    Dim lastRow As Long
    Dim doc As Object
    Dim profileNode As Object
    Dim values() As Variant
    Dim r As Long
    Dim vNode As Object
    Dim prevEvents As Boolean

    On Error GoTo EH
    prevEvents = Application.EnableEvents

    Set ws = ws_Dev

    startRow = 3
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then Exit Sub

    If Len(profileName) = 0 Then
        profileName = CStr(ws.Range("D1").Value)
    End If
    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then Exit Sub

    ReDim values(1 To lastRow - startRow + 1, 1 To 1)

    For r = startRow To lastRow
        Set vNode = profileNode.selectSingleNode("p:v[@row='" & CStr(r) & "']")
        If Not vNode Is Nothing Then
            values(r - startRow + 1, 1) = vNode.Text
        Else
            values(r - startRow + 1, 1) = vbNullString
        End If
    Next r

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(startRow, 2), ws.Cells(lastRow, 2)).Value = values
EH:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEvents

End Sub

' Helper for assigning to a button on Dev sheet.
Public Sub m_ApplyProfile_Button()
    m_ApplyProfileFromDev
End Sub

' Save edited values in column B back to the active profile (stored in external XML file).
Public Sub m_SaveEditsToProfile(ByVal ws As Worksheet, ByVal targetRange As Range, Optional ByVal profileName As String = vbNullString)

    Dim startRow As Long
    Dim lastRow As Long
    Dim editRange As Range
    Dim doc As Object
    Dim profileNode As Object

    startRow = 3
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then Exit Sub

    If Len(profileName) = 0 Then
        profileName = CStr(ws.Range("D1").Value)
    End If
    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    Set editRange = Intersect(targetRange, ws.Range(ws.Cells(startRow, 2), ws.Cells(lastRow, 2)))
    If editRange Is Nothing Then Exit Sub

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, True)
    If profileNode Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    mp_WriteSheetValuesToProfile ws, doc, profileNode

    mp_SavePresetsDom doc
    m_RefreshProfileValidation ws

    Application.ScreenUpdating = True

End Sub

' Refresh D1 validation list from stored profiles.
Public Sub m_RefreshProfileValidation(Optional ByVal ws As Worksheet)

    Dim doc As Object
    Dim nodes As Object
    Dim i As Long
    Dim names() As String
    Dim listFormula As String
    Dim listSep As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    Set doc = mp_LoadPresetsDom(ws)
    Set nodes = doc.selectNodes("/p:presets/p:profile")

    On Error Resume Next
    ws.Range("D1").Validation.Delete
    On Error GoTo 0

    If nodes.Length = 0 Then
        ws.Range("D1").Value = vbNullString
        Exit Sub
    End If

    ReDim names(0 To nodes.Length - 1)
    For i = 0 To nodes.Length - 1
        names(i) = CStr(nodes.Item(i).getAttribute("name"))
    Next i

    listSep = Application.International(xlListSeparator)
    listFormula = Join(names, listSep)
    If Len(listFormula) = 0 Then Exit Sub

    With ws.Range("D1").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

End Sub

Private Sub mp_SeedProfileFromSheet(ByVal doc As Object, ByVal ws As Worksheet)

    Dim profileName As String
    Dim root As Object
    Dim profileNode As Object
    Dim startRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim vNode As Object

    profileName = Trim$(CStr(ws.Range("D1").Value))
    If Len(profileName) = 0 Then
        profileName = "Default"
        ws.Range("D1").Value = profileName
    End If

    Set root = doc.selectSingleNode("/p:presets")
    If root Is Nothing Then Exit Sub

    Set profileNode = doc.createNode(1, "profile", PRESETS_NS)
    profileNode.setAttribute "name", profileName
    root.appendChild profileNode

    startRow = 3
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then Exit Sub

    For r = startRow To lastRow
        Set vNode = doc.createNode(1, "v", PRESETS_NS)
        vNode.setAttribute "row", CStr(r)
        vNode.Text = CStr(ws.Cells(r, 2).Value)
        profileNode.appendChild vNode
    Next r

    mp_SavePresetsDom doc

End Sub

Public Sub m_EnsureProfileDropdown(Optional ByVal ws As Worksheet)
    Dim profiles As Variant
    Dim profileName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    profiles = mp_GetProfileNames(ws)
    m_RefreshProfileValidation ws

    If Not mp_ArrayHasItems(profiles) Then
        MsgBox "Profiles are missing or config file is empty: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, True)
    If Len(profileName) = 0 Then
        Exit Sub
    End If

    ws.Range("D1").Value = profileName
    m_ApplyProfileFromDev profileName
End Sub

Public Sub m_OnProfileChanged()
    Dim ws As Worksheet
    Dim profiles As Variant
    Dim profileName As String

    Set ws = ws_Dev
    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then
        MsgBox "Profiles are missing or config file is empty: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, True)
    If Len(profileName) = 0 Then Exit Sub

    ws.Range("D1").Value = profileName
    m_ApplyProfileFromDev profileName
End Sub

Public Sub m_EnsureProfileDropdown_UI()
    m_EnsureProfileDropdown ws_Dev
End Sub

Public Sub m_OpenProfilePicker_UI()
    m_EnsureProfileDropdown ws_Dev
End Sub

Public Sub m_OpenProfilePicker(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    m_EnsureProfileDropdown ws
End Sub

Public Sub m_SaveCurrentProfileToConfig(Optional ByVal ws As Worksheet)
    Dim profiles As Variant
    Dim profileName As String
    Dim doc As Object
    Dim profileNode As Object

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then
        MsgBox "Profiles are missing or config file is empty: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
    If Len(profileName) = 0 Then Exit Sub

    If Not mp_ArrayContains(profiles, profileName) Then
        MsgBox "Profile '" & profileName & "' was not found in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, True)
    If profileNode Is Nothing Then
        MsgBox "Failed to access profile '" & profileName & "' in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    mp_WriteSheetValuesToProfile ws, doc, profileNode
    mp_SavePresetsDom doc
    ws.Range("D1").Value = profileName
    Application.StatusBar = "Profiles config saved: " & profileName
End Sub

Private Function mp_GetProfileNames(Optional ByVal ws As Worksheet) As Variant
    Dim doc As Object
    Dim nodes As Object
    Dim names() As String
    Dim i As Long

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    Set doc = mp_LoadPresetsDom(ws)
    Set nodes = doc.selectNodes("/p:presets/p:profile")

    If nodes.Length = 0 Then
        mp_GetProfileNames = Array()
        Exit Function
    End If

    ReDim names(0 To nodes.Length - 1)
    For i = 0 To nodes.Length - 1
        names(i) = CStr(nodes.Item(i).getAttribute("name"))
    Next i

    mp_GetProfileNames = names
End Function

Private Function mp_ArrayHasItems(ByVal values As Variant) As Boolean
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayHasItems = (UBound(values) >= LBound(values))
    End If
    Exit Function
EH:
    mp_ArrayHasItems = False
End Function

Private Function mp_ArrayContains(ByVal values As Variant, ByVal needle As String) As Boolean
    Dim i As Long

    If Not mp_ArrayHasItems(values) Then Exit Function
    needle = Trim$(needle)
    If Len(needle) = 0 Then Exit Function

    For i = LBound(values) To UBound(values)
        If StrComp(CStr(values(i)), needle, vbTextCompare) = 0 Then
            mp_ArrayContains = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetSelectedProfileNameFromDropdown(ByVal ws As Worksheet, ByVal profiles As Variant, Optional ByVal syncItems As Boolean = False) As String
    Dim cf As Object
    Dim selectedIndex As Long
    Dim previousIndex As Long
    Dim previousName As String
    Dim matchedIndex As Long

    Set cf = mp_GetProfileDropdownControl(ws)
    If cf Is Nothing Then
        MsgBox "Profile control 'ddProfile' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    previousIndex = mp_GetControlIndex(cf)
    previousName = mp_GetControlItemByIndex(cf, previousIndex)

    If syncItems Then
        mp_SetDropdownItems cf, profiles
        matchedIndex = mp_FindProfileIndexByName(profiles, previousName)
        If matchedIndex > 0 Then
            On Error Resume Next
            cf.Value = matchedIndex
            On Error GoTo 0
        ElseIf previousIndex >= 1 And previousIndex <= mp_ArrayLength(profiles) Then
            On Error Resume Next
            cf.Value = previousIndex
            On Error GoTo 0
        End If
    End If

    selectedIndex = mp_GetControlIndex(cf)
    If selectedIndex < 1 Or selectedIndex > mp_ArrayLength(profiles) Then
        MsgBox "No active profile is selected in control 'ddProfile'.", vbExclamation
        Exit Function
    End If

    mp_GetSelectedProfileNameFromDropdown = CStr(profiles(selectedIndex - 1))
End Function

Private Function mp_GetControlItemByIndex(ByVal cf As Object, ByVal itemIndex As Long) As String
    On Error Resume Next
    If itemIndex >= 1 Then
        mp_GetControlItemByIndex = CStr(cf.List(itemIndex))
    End If
    On Error GoTo 0
End Function

Private Function mp_GetProfileDropdownControl(ByVal ws As Worksheet) As Object
    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes("ddProfile")
    On Error GoTo 0
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetProfileDropdownControl = shp.ControlFormat
    On Error GoTo 0
End Function

Private Sub mp_SetDropdownItems(ByVal cf As Object, ByVal profiles As Variant)
    Dim i As Long

    On Error Resume Next
    cf.RemoveAllItems
    If Err.Number <> 0 Then
        MsgBox "Failed to clear control items in 'ddProfile': " & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For i = LBound(profiles) To UBound(profiles)
        On Error Resume Next
        cf.AddItem CStr(profiles(i))
        If Err.Number <> 0 Then
            MsgBox "Failed to add profile '" & CStr(profiles(i)) & "' into 'ddProfile': " & Err.Description, vbExclamation
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    Next i
End Sub

Private Function mp_GetControlIndex(ByVal cf As Object) As Long
    On Error Resume Next
    mp_GetControlIndex = CLng(cf.Value)
    If Err.Number <> 0 Then
        Err.Clear
        mp_GetControlIndex = 0
    End If
    On Error GoTo 0
End Function

Private Function mp_ArrayLength(ByVal values As Variant) As Long
    If Not mp_ArrayHasItems(values) Then Exit Function
    mp_ArrayLength = UBound(values) - LBound(values) + 1
End Function

Private Function mp_FindProfileIndexByName(ByVal profiles As Variant, ByVal profileName As String) As Long
    Dim i As Long

    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Function
    If Not mp_ArrayHasItems(profiles) Then Exit Function

    For i = LBound(profiles) To UBound(profiles)
        If StrComp(CStr(profiles(i)), profileName, vbTextCompare) = 0 Then
            mp_FindProfileIndexByName = i - LBound(profiles) + 1
            Exit Function
        End If
    Next i
End Function

Private Function mp_LoadPresetsDom(Optional ByVal ws As Worksheet) As Object

    Dim filePath As String
    Dim doc As Object

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    filePath = mp_GetProfilesFilePath()
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")

    doc.async = False
    doc.validateOnParse = False

    If Len(Dir(filePath)) > 0 Then
        If Not doc.Load(filePath) Then
            doc.loadXML PRESETS_TEMPLATE
        End If
    Else
        doc.loadXML PRESETS_TEMPLATE
    End If

    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"

    Set mp_LoadPresetsDom = doc

End Function

Private Sub mp_SavePresetsDom(ByVal doc As Object)

    Dim filePath As String

    filePath = mp_GetProfilesFilePath()

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
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText xmlText
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

Private Sub mp_WriteSheetValuesToProfile(ByVal ws As Worksheet, ByVal doc As Object, ByVal profileNode As Object)

    Dim startRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim vNode As Object
    Dim child As Object

    startRow = 3
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then Exit Sub

    For Each child In profileNode.selectNodes("p:v")
        profileNode.removeChild child
    Next child

    For r = startRow To lastRow
        Set vNode = doc.createNode(1, "v", PRESETS_NS)
        vNode.setAttribute "row", CStr(r)
        vNode.Text = CStr(ws.Cells(r, 2).Value)
        profileNode.appendChild vNode
    Next r

End Sub

Private Function mp_GetProfilesFilePath() As String
    Dim basePath As String

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    mp_GetProfilesFilePath = basePath & "\" & PRESETS_REL_PATH
End Function

Private Function mp_GetProfileNode(ByVal doc As Object, ByVal profileName As String, ByVal createIfMissing As Boolean) As Object

    Dim node As Object
    Dim root As Object

    Set node = doc.selectSingleNode("/p:presets/p:profile[@name=" & mp_XPathLiteral(profileName) & "]")
    If node Is Nothing And createIfMissing Then
        Set root = doc.selectSingleNode("/p:presets")
        If root Is Nothing Then Exit Function
        Set node = doc.createNode(1, "profile", PRESETS_NS)
        node.setAttribute "name", profileName
        root.appendChild node
    End If

    Set mp_GetProfileNode = node

End Function

Private Function mp_XPathLiteral(ByVal value As String) As String

    If InStr(value, "'") = 0 Then
        mp_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If InStr(value, """") = 0 Then
        mp_XPathLiteral = """" & value & """"
        Exit Function
    End If

    Dim parts() As String
    Dim i As Long

    parts = Split(value, "'")
    mp_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            mp_XPathLiteral = mp_XPathLiteral & ", ""'"" , "
        End If
        mp_XPathLiteral = mp_XPathLiteral & "'" & parts(i) & "'"
    Next i
    mp_XPathLiteral = mp_XPathLiteral & ")"

End Function
