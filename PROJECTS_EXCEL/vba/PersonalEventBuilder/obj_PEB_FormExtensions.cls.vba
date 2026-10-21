VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_FormExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiBindingListener

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_page As obj_PageBase
Private m_rows As Collection
Private m_commands As Object
Private m_defaultStyles As Object
Private m_services As Object
Private m_profiles As Object
Private m_nextId As Long
Private m_form As String
Private m_active As Object
Private m_config As Object
Private m_headers As Object
Private m_defaultModes As Object
Private m_mode As String
Private m_viewport As String

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
Private Sub obj_IUiBindingListener_OnValueChanged( _
    ByVal sourceName As String, _
    ByVal bindingPath As String _
)
    Dim value As Variant
    Dim ignored As Object
    Dim isObject As Boolean

    If m_isDisposed Or Not m_isInitialized Then Exit Sub
    If sourceName <> "Form" Then Exit Sub
    If Not m_defaultStyles.Exists(bindingPath) Then Exit Sub
    If Not m_page.BindingContext.TryGetValue(sourceName, bindingPath, value, ignored, isObject) Then Exit Sub
    If isObject Or IsError(value) Or IsNull(value) Then Exit Sub
    If Len(Trim$(CStr(value))) > 0 Then Exit Sub
    If Not m_page.BindingContext.SetValue("Form", bindingPath & "Style", m_defaultStyles(bindingPath)) Then
        Err.Raise 5, , "Cannot reset extension field style."
    End If
End Sub

' //
' // API
' //
Public Function Initialize(ByVal page As obj_PageBase) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    If page Is Nothing Then Exit Function
    Set m_page = page
    Set m_rows = New Collection
    Set m_commands = CreateObject("Scripting.Dictionary")
    Set m_defaultStyles = CreateObject("Scripting.Dictionary")
    Set m_services = CreateObject("Scripting.Dictionary")
    Set m_profiles = CreateObject("Scripting.Dictionary")
    Set m_config = CreateObject("Scripting.Dictionary")
    Set m_headers = CreateObject("Scripting.Dictionary")
    Set m_defaultModes = CreateObject("Scripting.Dictionary")
    If Not m_page.BindingContext.SetValue("Form", "ExtCandidatesVisibility", "Collapsed") Then Exit Function
    m_isInitialized = True
    m_page.BindingContext.Subscribe Me
    Initialize = True
End Function

Public Sub Dispose()
    Dim command As obj_UiCommand
    Dim key As Variant

    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    ex_PEB_ExtRefresh.fn_Cancel Me
    If Not m_page Is Nothing Then m_page.BindingContext.Unsubscribe Me
    If Not m_commands Is Nothing Then
        For Each key In m_commands.Keys
            Set command = m_commands(key)
            command.Dispose
        Next key
    End If
    If Not m_services Is Nothing Then
        For Each key In m_services.Keys
            m_services(key).Dispose
            m_profiles(key).Dispose
        Next key
    End If
    Set m_commands = Nothing
    Set m_defaultStyles = Nothing
    Set m_services = Nothing
    Set m_profiles = Nothing
    Set m_rows = Nothing
    Set m_active = Nothing
    Set m_config = Nothing
    Set m_headers = Nothing
    Set m_defaultModes = Nothing
    Set m_page = Nothing
End Sub

Public Sub InvalidateCache()
    Dim key As Variant

    m_config.RemoveAll
    m_headers.RemoveAll
    For Each key In m_services.Keys
        m_services(key).InvalidateCache
    Next key
End Sub

Public Function HideCandidates() As Boolean
    Dim row As Object

    For Each row In m_rows
        If Not m_page.BindingContext.SetValue("Form", "ExVisible" & row("Id"), "Collapsed") Then Exit Function
    Next row
    HideCandidates = True
End Function

Public Function Reset(ByRef diagnostic As String) As Boolean
    Dim row As Object
    Dim field As Variant

    For Each row In m_rows
        For Each field In Split(private_Config(row("Type"), "fields"), "|")
            If Not m_page.BindingContext.SetValue("Form", private_Path(row, CStr(field)), vbNullString) Then Exit Function
        Next field
    Next row
    Set m_rows = New Collection
    m_defaultModes.RemoveAll
    Set m_active = Nothing
    If Not Me.HideCandidates() Then Exit Function
    Reset = True
End Function

Public Function Build( _
    ByVal formKind As String, _
    ByRef diagnostic As String, _
    Optional ByVal headerOnly As Boolean = False _
) As Boolean
    Dim document As Object
    Dim header As Object
    Dim body As Object
    Dim grid As Object
    Dim node As Object
    Dim row As Object
    Dim kind As Variant
    Dim options As Collection
    Dim payload As Object
    Dim key As String
    Dim column As Long
    Dim index As Long
    Dim groups As Object
    Dim anchor As Variant
    Dim kinds As Collection
    Dim cachedHeader As Boolean
    Dim headerKey As String
    Dim viewport As Variant
    Dim ignored As Object
    Dim isObject As Boolean
    Dim previousMode As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Extensions are not initialized."
    m_form = formKind
    previousMode = m_mode
    m_mode = private_ModeId()
    If previousMode <> m_mode Then
        If Not Me.HideCandidates() Then Exit Function
        Set m_active = Nothing
    End If
    If Not headerOnly Then private_ApplyDefaults
    ' Keep existing command objects alive when only extension rows are rebuilt.
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    If Not m_page.BindingContext.TryGetValue("Settings", "FormViewportStart", viewport, ignored, isObject) Then
        Err.Raise 5, , "Missing Settings.FormViewportStart."
    End If
    If m_viewport <> CStr(viewport) Then
        If Not Me.HideCandidates() Then Exit Function
        Set m_active = Nothing
    End If
    m_viewport = CStr(viewport)
    headerKey = m_mode & "|" & formKind & "|" & CStr(viewport)
    cachedHeader = m_headers.Exists(headerKey)
    If cachedHeader Then
        Set header = m_headers(headerKey).cloneNode(True)
    Else
        Set header = private_Node(document, "stackPanel")
        header.setAttribute "orientation", "vertical"
    End If
    Set body = private_Node(document, "stackPanel")
    body.setAttribute "orientation", "vertical"
    If Not cachedHeader Then
        Set grid = private_Node(document, "grid")
    End If
    Set groups = CreateObject("Scripting.Dictionary")
    For Each kind In Split(private_ReadConfig("extensions.types"), "|")
        If Not private_Visible(CStr(kind)) Then GoTo NextKind
        anchor = private_Config(CStr(kind), "anchor")
        If Not groups.Exists(anchor) Then
            Set kinds = New Collection
            groups.Add anchor, kinds
        End If
        groups(anchor).Add CStr(kind)
NextKind:
    Next kind
    index = 0
    For Each anchor In groups.Keys
        index = index + 1
        column = private_Column(CStr(anchor))
        Set kinds = groups(anchor)
        If column > 0 Then
            key = "ExtAdd" & CStr(index)
            Set options = New Collection
            Set kinds = groups(anchor)
            For Each kind In kinds
                options.Add private_Config(CStr(kind), "caption")
            Next kind
            If Not m_page.BindingContext.SetObject("Data", key, options) Then Exit Function
            If Not m_page.BindingContext.SetValue("Form", key, vbNullString) Then Exit Function
            Set payload = private_Payload("add", Nothing)
            Set payload("Types") = kinds
            payload("Path") = key
            private_MenuCommand key, payload
            If cachedHeader Then GoTo NextHeader
            Set node = private_Node(document, "controls:select")
            node.setAttribute "name", key
            node.setAttribute "itemsSource", "{Binding Path=Data." & key & "}"
            node.setAttribute "column", CStr(column)
            node.setAttribute "style", "secondaryButton"
            node.setAttribute "value", "{Binding Path=Form." & key & "}"
            node.setAttribute "placeholder", private_ReadConfig("extensions.addCaption")
            node.setAttribute "panelStyle", "formSelectPanel"
            node.setAttribute "itemStyle", "formSelectItem"
            node.setAttribute "onChange", "{Binding Path=Commands." & key & "}"
            grid.appendChild node
        End If
NextHeader:
    Next anchor
    If Not cachedHeader Then
        If grid.ChildNodes.Length > 0 Then header.appendChild grid
    End If
    If Not cachedHeader Then Set m_headers(headerKey) = header.cloneNode(True)
    If headerOnly Then GoTo PublishMarkup
    private_BuildRows document, body
    For Each row In m_rows
        If Not private_Active(row) Then GoTo NextCandidates
        If Not private_Visible(row("Type")) Then GoTo NextCandidates
        key = "ExSelect" & row("Id")
        private_Command key, private_Payload("select", row)
        Set node = private_Node(document, "controls:candidates")
        node.setAttribute "name", "ExCandidates" & row("Id")
        node.setAttribute "itemsSource", "{Binding Path=Data.ExCandidates" & row("Id") & "}"
        node.setAttribute "column", "1"
        node.setAttribute "style", "candidateTable"
        node.setAttribute "visibility", "{Binding Path=Form.ExVisible" & row("Id") & "}"
        private_Command "ExClose" & row("Id"), private_Payload("close", row)
        node.setAttribute "closeCommand", "{Binding Path=Commands.ExClose" & row("Id") & "}"
        node.setAttribute "closeCaption", "{Binding Path=Text.CloseTable}"
        node.setAttribute "selectionMode", "Confirm"
        node.setAttribute "selectedItem", "{Binding Path=Data.ExSelected" & row("Id") & "}"
        node.setAttribute "onSelect", "{Binding Path=Commands." & key & "}"
        node.setAttribute "alignToForm", "ExForm" & row("Id")
        node.setAttribute "workspaceFirstColumn", "1"
        node.setAttribute "workspaceLastColumn", private_ReadConfig("forms.columns.count")
        node.setAttribute "showHeaders", "true"
        body.appendChild node
NextCandidates:
    Next row
PublishMarkup:
    If Not m_page.BindingContext.SetObject("Data", "ExtensionHeader", header) Then Exit Function
    If Not headerOnly Then
        If Not m_page.BindingContext.SetObject("Data", "ExtensionBody", body) Then Exit Function
    End If
    Build = True
    Exit Function
EH:
    diagnostic = "Extensions: " & Err.Description
End Function

Public Function HandleAction( _
    ByVal payload As Object, _
    Optional ByVal selected As Object = Nothing _
) As Boolean
    Dim row As Object
    Dim field As Variant
    Dim mapping As Variant
    Dim updates As Object
    Dim value As Variant
    Dim record As obj_LookupRecord
    Dim diagnostic As String
    Dim index As Long
    Dim autoFill As Variant
    Dim ignored As Object
    Dim isObject As Boolean
    Dim kind As Variant

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ex_Core.fn_Diagnostic_WriteLog "EXTENSION_ACTION | Action=" & payload("Action")
    Select Case payload("Action")
        Case "close"
            Set row = payload("Row")
            HandleAction = m_page.BindingContext.SetValue("Form", "ExVisible" & row("Id"), "Collapsed")
            Exit Function
        Case "add"
            If m_rows.Count >= CLng(private_ReadConfig("extensions.limit")) Then Err.Raise 5, , private_ReadConfig("extensions.limitReached")
            If Not m_page.BindingContext.TryGetValue("Form", payload("Path"), value, ignored, isObject) Then Exit Function
            For Each kind In payload("Types")
                If private_Config(CStr(kind), "caption") = CStr(value) Then
                    If Not private_Visible(CStr(kind)) Then Err.Raise 5, , private_ReadConfig("extensions.unavailable")
                    Set row = private_Add(CStr(kind))
                End If
            Next kind
            If row Is Nothing Then Err.Raise 5, , "Unknown extension menu choice."
        Case "remove"
            Set row = payload("Row")
            If private_Required(row) Then Err.Raise 5, , private_ReadConfig("extensions.requiredCaption")
            If Len(row("RowId")) > 0 Then private_RemoveChildren row("RowId")
            For index = m_rows.Count To 1 Step -1
                If m_rows(index)("Id") = row("Id") Then m_rows.Remove index
            Next index
            Set m_active = Nothing
        Case "search"
            Set row = payload("Row")
            If Not private_Active(row) Or Not private_Visible(row("Type")) Then
                HandleAction = Me.HideCandidates()
                Exit Function
            End If
            If Not m_page.BindingContext.TryGetValue("Form", "AutoFillPersonnel", autoFill, ignored, isObject) Then Exit Function
            If Not CBool(autoFill) Then
                HandleAction = Me.HideCandidates()
                Exit Function
            End If
            If Not private_Search(row, diagnostic) Then GoTo Failed
            ' Candidate controls already exist; binding updates are flushed by the cell-change dispatcher.
            HandleAction = True
            Exit Function
        Case "select"
            If selected Is Nothing Then Exit Function
            Set record = selected
            Set row = payload("Row")
            If m_active Is Nothing Then Exit Function
            If m_active("Id") <> row("Id") Then
                HandleAction = True
                Exit Function
            End If
            If Not m_page.BindingContext.TryGetValue("Form", "AutoFillPersonnel", autoFill, ignored, isObject) Then Exit Function
            If Not CBool(autoFill) Then
                HandleAction = True
                Exit Function
            End If
            Set updates = CreateObject("Scripting.Dictionary")
            For Each field In Split(private_Config(row("Type"), "map"), "|")
                mapping = Split(field, "=")
                If Not record.TryGetValue(mapping(1), value) Then Err.Raise 5, , "Extension candidate field missing: " & mapping(1)
                updates("Form." & private_Path(row, mapping(0))) = value
                If Len(Trim$(CStr(value))) > 0 Then
                    updates("Form." & private_Path(row, mapping(0)) & "Style") = private_Style(row("Type"), CStr(mapping(0)), private_Config(row("Type"), "profile"))
                Else
                    updates("Form." & private_Path(row, mapping(0)) & "Style") = private_Style(row("Type"), CStr(mapping(0)), "default")
                End If
            Next field
            If Not m_page.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
            HandleAction = m_page.FlushUpdates()
            Exit Function
    End Select
    ' Never dispose the active dropdown or its binding while its callback is running.
    HandleAction = ex_PEB_ExtRefresh.fn_Request(Me)
    Exit Function
EH:
    diagnostic = Err.Description
Failed:
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, private_ReadConfig("extensions.errorTitle")
End Function

Public Function TryRefresh(ByRef diagnostic As String) As Boolean
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    startedAt = Timer
    If Not Me.Build(m_form, diagnostic) Then Exit Function
    ex_Core.fn_Diagnostic_WritePerf "Extensions.BuildMarkup", startedAt
    If Not Application.ActiveWorkbook Is ThisWorkbook Then
        TryRefresh = True
        Exit Function
    End If
    startedAt = Timer
    If Not m_page.RefreshContent("ExtensionMenus", diagnostic, True) Then Exit Function
    If Not m_page.RefreshContent("ExtensionRows", diagnostic, True) Then Exit Function
    TryRefresh = m_page.FlushUpdates()
    ex_Core.fn_Diagnostic_WritePerf "Extensions.RefreshContent", startedAt
End Function

Public Function Collect( _
    ByRef output As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim row As Object
    Dim item As Object
    Dim field As Variant
    Dim scalar As Variant
    Dim ignored As Object
    Dim isObject As Boolean
    Dim requiredFields As String

    Set output = New Collection
    For Each row In m_rows
        If Not private_Active(row) Then GoTo NextCollectedRow
        Set item = CreateObject("Scripting.Dictionary")
        item("RowType") = row("Type")
        item("RowId") = row("RowId")
        item("ParentRowId") = row("ParentRowId")
        item("ExtensionRole") = row("Role")
        For Each field In Split(private_Config(row("Type"), "fields"), "|")
            If Not m_page.BindingContext.TryGetValue("Form", private_Path(row, CStr(field)), scalar, ignored, isObject) Then Exit Function
            If isObject Or IsError(scalar) Or IsNull(scalar) Then Exit Function
            item(field) = CStr(scalar)
        Next field
        requiredFields = private_Config(row("Type"), "required")
        If private_Required(row) Then requiredFields = private_ReadConfig("extensions.role." & row("Role") & ".required")
        For Each field In Split(requiredFields, "|")
            If Len(Trim$(item(field))) = 0 Then
                diagnostic = Replace$(private_ReadConfig("extensions.invalid"), "{field}", CStr(field))
                Exit Function
            End If
        Next field
        output.Add item
NextCollectedRow:
    Next row
    Collect = True
End Function

Public Function TryLoad( _
    ByVal children As Collection, _
    ByRef diagnostic As String, _
    Optional ByVal preserveDrafts As Boolean = False _
) As Boolean
    Dim child As Object
    Dim row As Object
    Dim field As Variant
    Dim drafts As Collection

    Set drafts = New Collection
    If preserveDrafts Then
        If Not Me.HideCandidates() Then Exit Function
        For Each row In m_rows
            If Not private_Active(row) Then
                row("RowId") = vbNullString
                row("ParentRowId") = vbNullString
                drafts.Add row
            End If
        Next row
        Set m_rows = New Collection
        Set m_active = Nothing
        If m_defaultModes.Exists(m_mode) Then m_defaultModes.Remove m_mode
    Else
        If Not Me.Reset(diagnostic) Then Exit Function
    End If
    For Each child In children
        Set row = private_Add(child("RowType"))
        row("RowId") = child("RowId")
        row("ParentRowId") = child("ParentRowId")
        If child.Exists("ExtensionRole") Then row("Role") = child("ExtensionRole")
        For Each field In Split(private_Config(row("Type"), "fields"), "|")
            If Not m_page.BindingContext.SetValue("Form", private_Path(row, CStr(field)), child(field)) Then Exit Function
            If Len(Trim$(CStr(child(field)))) > 0 Then
                If InStr(1, private_Config(row("Type"), "map"), CStr(field) & "=", vbBinaryCompare) > 0 Then
                    If Not m_page.BindingContext.SetValue("Form", private_Path(row, CStr(field)) & "Style", private_Style(row("Type"), CStr(field), private_Config(row("Type"), "profile"))) Then Exit Function
                End If
            End If
        Next field
    Next child
    For Each row In drafts
        m_rows.Add row
    Next row
    TryLoad = True
End Function

Public Function TryAcceptSaved( _
    ByVal children As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim index As Long
    Dim row As Object
    Dim child As Object
    Dim activeRows As Collection

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    Set activeRows = New Collection
    For Each row In m_rows
        If private_Active(row) Then activeRows.Add row
    Next row
    If children.Count <> activeRows.Count Then
        diagnostic = "Saved extension count differs from the current form."
        Exit Function
    End If
    ' Validate the correspondence before updating identifiers used by existing commands.
    For index = 1 To activeRows.Count
        Set row = activeRows(index)
        Set child = children(index)
        If row("Type") <> child("RowType") Then
            diagnostic = "Saved extension type differs from the current form."
            Exit Function
        End If
        If Not child.Exists("RowId") Or Not child.Exists("ParentRowId") Then
            diagnostic = "Saved extension identifiers are missing."
            Exit Function
        End If
    Next index
    For index = 1 To activeRows.Count
        Set row = activeRows(index)
        Set child = children(index)
        row("RowId") = child("RowId")
        row("ParentRowId") = child("ParentRowId")
    Next index
    ' Inactive drafts were excluded from this saved event; detach their obsolete stored identifiers.
    For Each row In m_rows
        If Not private_Active(row) Then
            row("RowId") = vbNullString
            row("ParentRowId") = vbNullString
        End If
    Next row
    TryAcceptSaved = True
End Function

' //
' // Private
' //
Private Function private_ModeId() As String
    Dim value As Variant
    Dim ignored As Object
    Dim isObject As Boolean

    If Not m_page.BindingContext.TryGetValue("Form", "ModeId", value, ignored, isObject) Then Err.Raise 5, , "Missing Form.ModeId."
    If isObject Or IsNull(value) Or IsError(value) Then Err.Raise 5, , "Invalid Form.ModeId."
    private_ModeId = CStr(value)
End Function

Private Function private_Allowed(ByVal kind As String) As Boolean
    Dim fields As Object
    Dim field As Variant
    Dim index As Long
    Dim prefix As String

    Set fields = CreateObject("Scripting.Dictionary")
    prefix = "forms." & m_form & ".fields."
    For index = 1 To CLng(private_ReadConfig(prefix & "count"))
        fields(private_ReadConfig(prefix & index & ".name")) = True
    Next index
    For Each field In Split(private_Config(kind, "fields"), "|")
        If Not fields.Exists(CStr(field)) Then Exit Function
    Next field
    private_Allowed = True
End Function

Private Function private_Visible(ByVal kind As String) As Boolean
    Dim field As Variant

    If Not private_Allowed(kind) Then Exit Function
    For Each field In Split(private_Config(kind, "fields"), "|")
        If private_Column(CStr(field)) < 1 Then Exit Function
    Next field
    private_Visible = True
End Function

Private Function private_Active(ByVal row As Object) As Boolean
    If Len(row("Mode")) > 0 And row("Mode") <> m_mode Then Exit Function
    private_Active = private_Allowed(row("Type"))
End Function

Private Function private_Required(ByVal row As Object) As Boolean
    Dim roles As String

    If row("Mode") <> m_mode Or Len(row("Role")) = 0 Then Exit Function
    If InStr(1, "|" & private_ReadConfig("extensions.modes") & "|", "|" & m_mode & "|", vbBinaryCompare) = 0 Then Exit Function
    roles = private_ReadConfig("extensions.mode." & m_mode & ".required")
    private_Required = InStr(1, "|" & roles & "|", "|" & row("Role") & "|", vbBinaryCompare) > 0
End Function

Private Sub private_ApplyDefaults()
    Dim definition As Variant
    Dim pair As Variant
    Dim row As Object
    Dim selected As Object
    Dim modes As String

    If m_defaultModes.Exists(m_mode) Then Exit Sub
    modes = private_ReadConfig("extensions.modes")
    If InStr(1, "|" & modes & "|", "|" & m_mode & "|", vbBinaryCompare) = 0 Then Exit Sub
    For Each definition In Split(private_ReadConfig("extensions.mode." & m_mode & ".defaults"), "|")
        pair = Split(definition, "=")
        If UBound(pair) <> 1 Then Err.Raise 5, , "Invalid extension default: " & definition
        If Not private_Allowed(CStr(pair(0))) Then Err.Raise 5, , "Default extension has missing form fields: " & pair(0)
        Set selected = Nothing
        For Each row In m_rows
            If row("Type") = pair(0) Then
                If row("Role") = pair(1) And row("Mode") = m_mode Then
                    Set selected = row
                    Exit For
                End If
            End If
        Next row
        If selected Is Nothing Then
            For Each row In m_rows
                If row("Type") = pair(0) And Len(row("Mode")) = 0 Then
                    If row("Role") = pair(1) Or (Len(row("Role")) = 0 And Len(row("RowId")) > 0) Then
                        Set selected = row
                        Exit For
                    End If
                End If
            Next row
        End If
        If selected Is Nothing Then
            If m_rows.Count >= CLng(private_ReadConfig("extensions.limit")) Then Err.Raise 5, , private_ReadConfig("extensions.limitReached")
            Set selected = private_Add(CStr(pair(0)))
        End If
        selected("Role") = pair(1)
        selected("Mode") = m_mode
    Next definition
    m_defaultModes(m_mode) = True
End Sub

Private Sub private_MenuCommand( _
    ByVal name As String, _
    ByVal payload As Object _
)
    Dim command As obj_UiCommand

    ' Build runs before the deferred panel refresh, outside the dropdown callback.
    If m_commands.Exists(name) Then
        Set command = m_commands(name)
        command.Dispose
        m_commands.Remove name
    End If
    private_Command name, payload
End Sub

Private Sub private_BuildRows( _
    ByVal document As Object, _
    ByVal body As Object _
)
    Dim grid As Object
    Dim form As Object
    Dim node As Object
    Dim row As Object
    Dim occupied As Object
    Dim slots As Object
    Dim field As Variant
    Dim slot As Variant
    Dim lane As Long
    Dim column As Long
    Dim titleColumn As Long
    Dim removeColumn As Long
    Dim width As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    Dim conflict As Boolean
    Dim key As String
    Dim path As String
    Dim caption As String

    Set occupied = CreateObject("Scripting.Dictionary")
    Set grid = private_Node(document, "grid")
    width = CLng(private_ReadConfig("forms.columns.count"))
    For Each row In m_rows
        If Not private_Active(row) Then GoTo NextRow
        If Not private_Visible(row("Type")) Then GoTo NextRow
        Set slots = CreateObject("Scripting.Dictionary")
        firstColumn = width + 1
        lastColumn = 0
        For Each field In Split(private_Config(row("Type"), "fields"), "|")
            column = private_Column(CStr(field))
            slots(column) = True
            If column < firstColumn Then firstColumn = column
            If column > lastColumn Then lastColumn = column
        Next field
        ' Reserve two service cells without moving any mapped form field.
        titleColumn = firstColumn - 1
        If titleColumn < 1 Then titleColumn = 0
        removeColumn = lastColumn + 1
        If removeColumn > width Then removeColumn = 0
        For column = 1 To width
            If Not slots.Exists(column) Then
                If titleColumn = 0 And column <> removeColumn Then titleColumn = column
            End If
        Next column
        For column = width To 1 Step -1
            If Not slots.Exists(column) And column <> titleColumn Then
                If removeColumn = 0 Then removeColumn = column
            End If
        Next column
        If titleColumn < 1 Or removeColumn < 1 Then Err.Raise 5, , private_ReadConfig("extensions.noServiceSpace")
        slots(titleColumn) = True
        slots(removeColumn) = True
        lane = 1
        Do
            conflict = False
            For Each slot In slots.Keys
                If occupied.Exists(CStr(lane) & "|" & CStr(slot)) Then conflict = True
            Next slot
            If Not conflict Then Exit Do
            lane = lane + 1
        Loop
        For Each slot In slots.Keys
            occupied(CStr(lane) & "|" & CStr(slot)) = True
        Next slot
        caption = private_Config(row("Type"), "caption")
        If Len(row("Role")) > 0 Then caption = private_ReadConfig("extensions.role." & row("Role") & ".caption")
        Set node = private_Node(document, "controls:label")
        node.setAttribute "name", "ExTitle" & row("Id")
        node.setAttribute "row", CStr(lane * 2 - 1)
        node.setAttribute "column", CStr(titleColumn)
        node.setAttribute "rowSpan", "2"
        node.setAttribute "style", "statusHeading"
        node.setAttribute "text", caption
        grid.appendChild node
        Set node = private_Node(document, "controls:button")
        key = "ExRemove" & row("Id")
        If private_Required(row) Then
            Set node = private_Node(document, "controls:label")
            node.setAttribute "text", private_ReadConfig("extensions.requiredCaption")
        Else
            private_Command key, private_Payload("remove", row)
            node.setAttribute "caption", private_ReadConfig("extensions.removeCaption")
            node.setAttribute "command", "{Binding Path=Commands." & key & "}"
        End If
        node.setAttribute "name", key
        node.setAttribute "row", CStr(lane * 2 - 1)
        node.setAttribute "column", CStr(removeColumn)
        node.setAttribute "rowSpan", "2"
        node.setAttribute "style", "secondaryButton"
        grid.appendChild node
        Set form = private_Node(document, "controls:form")
        form.setAttribute "name", "ExForm" & row("Id")
        form.setAttribute "dataContext", "{Binding Path=Form}"
        form.setAttribute "row", CStr(lane * 2 - 1)
        form.setAttribute "column", "1"
        form.setAttribute "columnSpan", CStr(width)
        form.setAttribute "orientation", "horizontal"
        form.setAttribute "placement", "mapped"
        form.setAttribute "candidateLayout", "viewport"
        form.setAttribute "labelPosition", "top"
        form.setAttribute "labelStyle", "formLabel"
        form.setAttribute "fieldStyle", "formField"
        For Each field In Split(private_Config(row("Type"), "fields"), "|")
            path = private_Path(row, CStr(field))
            Set node = private_Node(document, "field")
            node.setAttribute "name", path
            node.setAttribute "type", "text"
            node.setAttribute "column", CStr(private_Column(CStr(field)))
            node.setAttribute "candidateColumn", CStr(private_Column(CStr(field)))
            node.setAttribute "value", "{Binding Path=" & path & "}"
            node.setAttribute "label", private_ReadConfig("text.Mode" & field & "Label")
            node.setAttribute "fieldStyle", "{Binding Path=Form." & path & "Style}"
            node.setAttribute "partStylesEnabled", "true"
            If InStr(1, "|" & private_Config(row("Type"), "readOnly") & "|", "|" & field & "|", vbBinaryCompare) > 0 Then
                node.setAttribute "readOnly", "{Binding Path=Form.AutoFillPersonnel}"
            End If
            If field = private_Config(row("Type"), "key") Then
                key = "ExSearch" & row("Id")
                private_Command key, private_Payload("search", row)
                node.setAttribute "onChange", "{Binding Path=Commands." & key & "}"
            End If
            form.appendChild node
        Next field
        grid.appendChild form
NextRow:
    Next row
    If grid.ChildNodes.Length > 0 Then body.appendChild grid
End Sub

Private Function private_Style( _
    ByVal kind As String, _
    ByVal field As String, _
    ByVal sourceAlias As String _
) As String
    Dim defaultStyle As String
    Dim style As String

    defaultStyle = private_ReadConfig("forms.sourceStyles.default")
    style = defaultStyle
    If private_ReadConfig("forms.dynamicStylesEnabled") = "true" Then
        style = private_ReadConfig("forms.sourceStyles." & sourceAlias)
    End If
    If field = private_Config(kind, "key") Then
        If style = defaultStyle Then
            style = private_ReadConfig("forms.lookupDefaultStyle")
        Else
            style = style & "Lookup"
        End If
    End If
    private_Style = style
End Function

Private Function private_Config( _
    ByVal kind As String, _
    ByVal key As String _
) As String
    private_Config = private_ReadConfig("extensions.type." & kind & "." & key)
End Function

Private Function private_ReadConfig(ByVal key As String) As String
    If Not m_config.Exists(key) Then m_config.Add key, ex_PEB_Export.fn_Config(key)
    private_ReadConfig = m_config(key)
End Function

Private Sub private_RemoveChildren(ByVal parentId As String)
    Dim index As Long
    Dim row As Object

    Do
        Set row = Nothing
        For index = 1 To m_rows.Count
            If m_rows(index)("ParentRowId") = parentId Then
                Set row = m_rows(index)
                Exit For
            End If
        Next index
        If row Is Nothing Then Exit Do
        private_RemoveChildren row("RowId")
        For index = m_rows.Count To 1 Step -1
            If m_rows(index) Is row Then m_rows.Remove index
        Next index
    Loop
End Sub

Private Function private_Path( _
    ByVal row As Object, _
    ByVal field As String _
) As String
    private_Path = "Ex" & row("Id") & "_" & field
End Function

Private Function private_Add(ByVal kind As String) As Object
    Dim row As Object
    Dim field As Variant
    Dim mapping As Variant
    Dim scalar As Variant
    Dim ignored As Object
    Dim isObject As Boolean
    Dim diagnostic As String

    m_nextId = m_nextId + 1
    Set row = CreateObject("Scripting.Dictionary")
    row("Id") = CStr(m_nextId)
    row("Type") = kind
    row("RowId") = vbNullString
    row("ParentRowId") = vbNullString
    row("Role") = vbNullString
    row("Mode") = vbNullString
    For Each field In Split(private_Config(kind, "fields"), "|")
        m_defaultStyles(private_Path(row, CStr(field))) = private_Style(kind, CStr(field), "default")
        If Not m_page.BindingContext.SetValue("Form", private_Path(row, CStr(field)) & "Style", m_defaultStyles(private_Path(row, CStr(field)))) Then Err.Raise 5, , "Cannot initialize extension style."
        If Not m_page.BindingContext.SetValue("Form", private_Path(row, CStr(field)), vbNullString) Then Err.Raise 5, , "Cannot initialize extension field."
    Next field
    For Each field In Split(private_Config(kind, "defaults"), "|")
        If field <> "-" Then
            mapping = Split(field, "=")
            If Not m_page.BindingContext.TryGetValue("Form", mapping(1), scalar, ignored, isObject) Then Err.Raise 5, , "Extension default not found."
            If IsDate(scalar) Then scalar = Format$(CDate(scalar), "dd.mm.yyyy")
            If Not m_page.BindingContext.SetValue("Form", private_Path(row, mapping(0)), scalar) Then Err.Raise 5, , "Cannot apply extension default."
        End If
    Next field
    If Not private_PublishCandidates(row, Nothing, diagnostic) Then Err.Raise 5, , diagnostic
    If Not m_page.BindingContext.SetValue("Form", "ExVisible" & row("Id"), "Collapsed") Then Err.Raise 5, , "Cannot initialize extension candidates visibility."
    m_rows.Add row
    Set m_active = Nothing
    Set private_Add = row
End Function

Private Function private_Column(ByVal field As String) As Long
    Dim value As Variant
    Dim source As Object
    Dim isObject As Boolean

    If Not m_page.BindingContext.TryGetValue("Data", "FormColumnMap", value, source, isObject) Then
        Err.Raise 5, , "Cannot resolve the active form column map: Data.FormColumnMap"
    End If
    If Not isObject Then Err.Raise 5, , "The active form column map must be a dictionary."
    If source.Exists(field) Then private_Column = CLng(source(field))
End Function

Private Function private_Node( _
    ByVal document As Object, _
    ByVal tag As String _
) As Object
    Dim uri As String

    uri = "urn:excelprototype:profiles"
    If Left$(tag, 9) = "controls:" Then uri = "urn:excelprototype:controls"
    Set private_Node = document.createNode(1, tag, uri)
End Function

Private Function private_Payload( _
    ByVal action As String, _
    ByVal row As Object _
) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload("Action") = action
    If Not row Is Nothing Then Set payload("Row") = row
    Set private_Payload = payload
End Function

Private Sub private_Command( _
    ByVal name As String, _
    ByVal payload As Object _
)
    Dim command As obj_UiCommand

    If m_commands.Exists(name) Then Exit Sub
    Set command = New obj_UiCommand
    If Not command.Initialize(Me, "HandleAction", payload) Then Err.Raise 5, , "Cannot initialize extension command."
    If Not m_page.BindingContext.SetObject("Commands", name, command) Then Err.Raise 5, , "Cannot publish extension command."
    m_commands.Add name, command
End Sub

Private Function private_Search( _
    ByVal row As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim profile As obj_LookupProfile
    Dim service As obj_LookupService
    Dim result As obj_LookupResult
    Dim scalar As Variant
    Dim ignored As Object
    Dim isObject As Boolean
    Dim key As String

    key = private_Config(row("Type"), "profile")
    If Not m_services.Exists(key) Then
        Set profile = New obj_LookupProfile
        If Not profile.Initialize("PersonalEventBuilder::lookup." & key, diagnostic) Then Exit Function
        Set service = New obj_LookupService
        If Not service.Initialize(profile) Then Exit Function
        m_services.Add key, service
        m_profiles.Add key, profile
    End If
    Set service = m_services(key)
    If Not m_page.BindingContext.TryGetValue("Form", private_Path(row, private_Config(row("Type"), "key")), scalar, ignored, isObject) Then Exit Function
    If Not service.TrySearch(CStr(scalar), result, diagnostic) Then Exit Function
    If row("Type") = "Report" Then
        If Not ex_PEB_Export.fn_TryApplyReporterPositions(result.Records, diagnostic) Then Exit Function
    End If
    If Not private_PublishCandidates(row, result, diagnostic) Then Exit Function
    If Not Me.HideCandidates() Then Exit Function
    If Not m_page.BindingContext.SetValue("Form", "CandidatesVisibility", "Collapsed") Then Exit Function
    If Not m_page.BindingContext.SetValue("Form", "ExVisible" & row("Id"), "Visible") Then Exit Function
    Set m_active = row
    private_Search = True
End Function

Private Function private_PublishCandidates( _
    ByVal row As Object, _
    ByVal result As obj_LookupResult, _
    ByRef diagnostic As String _
) As Boolean
    Dim projected As New obj_LookupResult
    Dim columns As New Collection
    Dim records As New Collection
    Dim column As Object
    Dim field As Variant
    Dim pair As Variant
    Dim mapping As Variant
    Dim hasMore As Boolean

    If Not result Is Nothing Then
        Set records = result.Records
        hasMore = result.HasMore
    End If
    For Each field In Split(private_Config(row("Type"), "fields"), "|")
        Set column = CreateObject("Scripting.Dictionary")
        column("name") = "Manual." & field
        For Each pair In Split(private_Config(row("Type"), "map"), "|")
            mapping = Split(pair, "=")
            If mapping(0) = field Then column("name") = mapping(1)
        Next pair
        column("formField") = private_Path(row, CStr(field))
        column("caption") = private_ReadConfig("text.Mode" & field & "Label")
        columns.Add column
    Next field
    If Not projected.Initialize(records, columns, hasMore, diagnostic) Then Exit Function
    ' Each form owns its projection so the previous table cannot consume another form's columns before the deferred rebuild.
    If Not m_page.BindingContext.SetObject("Data", "ExCandidates" & row("Id"), projected) Then Exit Function
    If Not m_page.BindingContext.SetValue("Data", "ExSelected" & row("Id"), vbNullString) Then Exit Function
    private_PublishCandidates = True
End Function