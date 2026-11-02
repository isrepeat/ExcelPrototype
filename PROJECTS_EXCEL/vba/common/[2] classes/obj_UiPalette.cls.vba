VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_palettes As Object
Private m_activeName As String
Private m_worksheet As Worksheet
Private m_paints As Collection
Private m_paintKeys As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get ActiveName() As String
    ActiveName = m_activeName
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal document As Object, _
    ByVal folder As String, _
    ByVal bindings As obj_UiBindingContext, _
    ByVal targetWorksheet As Worksheet _
) As Boolean
    Dim sources As String
    Dim source As Variant
    Dim firstPalette As Object
    Dim palette As Variant
    Dim key As Variant

    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_palettes = VBA.CreateObject("Scripting.Dictionary")
    m_palettes.CompareMode = VBA.vbTextCompare
    Set m_worksheet = targetWorksheet
    private_ResetPaints
    sources = VBA.Trim$(ex_UiElementFactory.fn_Attribute(document.documentElement, "paletteSources"))
    If VBA.Len(sources) > 0 Then
        For Each source In VBA.Split(sources, ";")
            private_LoadPalette folder, VBA.Trim$(VBA.CStr(source))
        Next source
        ' Every palette must implement the same resource contract.
        For Each palette In m_palettes.Items
            If firstPalette Is Nothing Then Set firstPalette = palette
            If palette.Count <> firstPalette.Count Then ex_UiStyleDiagnostics.fn_Raise "PaletteContract"
            For Each key In firstPalette.Keys
                If Not palette.Exists(key) Then ex_UiStyleDiagnostics.fn_Raise "PaletteColorMissing", key
            Next key
        Next palette
        If bindings Is Nothing Then ex_UiStyleDiagnostics.fn_Raise "PaletteBindings"
        If Not ex_UiBindingRuntime.fn_TryResolveText( _
                ex_UiElementFactory.fn_Attribute(document.documentElement, "palette"), bindings, m_activeName) Then
            ex_UiStyleDiagnostics.fn_Raise "PaletteResolve"
        End If
        If Not m_palettes.Exists(m_activeName) Then ex_UiStyleDiagnostics.fn_Raise "PaletteMissing", m_activeName
    ElseIf VBA.Len(ex_UiElementFactory.fn_Attribute(document.documentElement, "palette")) > 0 Then
        ex_UiStyleDiagnostics.fn_Raise "PaletteSources"
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Set m_palettes = Nothing
    Set m_worksheet = Nothing
    Set m_paints = Nothing
    Set m_paintKeys = Nothing
End Sub

Public Sub BeginRender()
    If Not m_isInitialized Or m_isDisposed Then ex_UiStyleDiagnostics.fn_Raise "PaletteInitialize"
    private_ResetPaints
End Sub

Public Function ResolveColor(ByVal value As String) As String
    Const PREFIX As String = "{ThemeResource "
    Dim key As String
    Dim palette As Object

    value = VBA.Trim$(value)
    ResolveColor = value
    If VBA.Left$(value, VBA.Len(PREFIX) - 1) <> "{ThemeResource" Then Exit Function
    If VBA.Left$(value, VBA.Len(PREFIX)) <> PREFIX Then ex_UiStyleDiagnostics.fn_Raise "PaletteResource", value
    If Not m_isInitialized Or m_isDisposed Then ex_UiStyleDiagnostics.fn_Raise "PaletteInitialize"
    If VBA.Right$(value, 1) <> "}" Then ex_UiStyleDiagnostics.fn_Raise "PaletteResource", value
    key = VBA.Trim$(VBA.Mid$(value, VBA.Len(PREFIX) + 1, VBA.Len(value) - VBA.Len(PREFIX) - 1))
    If Not m_palettes.Exists(m_activeName) Then ex_UiStyleDiagnostics.fn_Raise "PaletteMissing", m_activeName
    Set palette = m_palettes(m_activeName)
    If Not palette.Exists(key) Then ex_UiStyleDiagnostics.fn_Raise "PaletteRefMissing", m_activeName & "/" & key
    ResolveColor = palette(key)
End Function

Public Sub RememberColor( _
    ByVal area As Range, _
    ByVal shape As Object, _
    ByVal kind As String, _
    ByVal value As String, _
    Optional ByVal edge As Long = 0 _
)
    Dim entry As Object
    Dim key As String
    Dim shapeName As String

    If Not m_isInitialized Or m_isDisposed Then Exit Sub
    If m_palettes.Count = 0 Then Exit Sub
    If shape Is Nothing Then
        If area Is Nothing Then Exit Sub
        key = "cell|" & area.Address
    Else
        shapeName = shape.Name
        key = "shape|" & shapeName
    End If
    key = key & "|" & kind & "|" & VBA.CStr(edge)
    ' Keep the latest paint for an identical target at its actual position in the cascade.
    If m_paintKeys.Exists(key) Then
        m_paints.Remove key
    Else
        m_paintKeys.Add key, True
    End If
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    Set entry("area") = area
    entry.Add "shape", shapeName
    entry.Add "kind", kind
    entry.Add "value", value
    entry.Add "edge", edge
    entry.Add "key", key
    m_paints.Add entry, key
End Sub

Public Sub ForgetShape(ByVal name As String)
    Dim index As Long
    Dim entry As Object

    If m_paints Is Nothing Then Exit Sub
    For index = m_paints.Count To 1 Step -1
        Set entry = m_paints(index)
        If VBA.StrComp(entry("shape"), name, VBA.vbTextCompare) = 0 Then
            m_paintKeys.Remove entry("key")
            m_paints.Remove index
        End If
    Next index
End Sub

Public Sub ApplyTheme(ByVal name As String)
    Dim previousName As String
    Dim previousScreenUpdating As Boolean
    Dim previousEvents As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not m_isInitialized Or m_isDisposed Then ex_UiStyleDiagnostics.fn_Raise "PaletteInitialize"
    If Not m_palettes.Exists(name) Then ex_UiStyleDiagnostics.fn_Raise "PaletteMissing", name
    If VBA.StrComp(name, m_activeName, VBA.vbTextCompare) = 0 Then Exit Sub
    previousName = m_activeName
    previousScreenUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    m_activeName = name
    private_Repaint
CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If errorNumber <> 0 Then VBA.Err.Raise errorNumber, "UiPalette.ApplyTheme", errorDescription
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    m_activeName = previousName
    ' Best-effort rollback preserves the original failure for the caller.
    On Error Resume Next
    private_Repaint
    On Error GoTo 0
    GoTo CleanExit
End Sub

' //
' // Private
' //
Private Sub private_ResetPaints()
    Set m_paints = New Collection
    Set m_paintKeys = VBA.CreateObject("Scripting.Dictionary")
    m_paintKeys.CompareMode = VBA.vbTextCompare
End Sub

Private Sub private_LoadPalette( _
    ByVal folder As String, _
    ByVal source As String _
)
    Dim fileSystem As Object
    Dim document As Object
    Dim root As Object
    Dim node As Object
    Dim colors As Object
    Dim name As String
    Dim key As String
    Dim value As String
    Dim path As String

    If VBA.Len(source) = 0 Then ex_UiStyleDiagnostics.fn_Raise "PaletteFileEmpty"
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = VBA.Replace(source, "/", "\")
    If VBA.Len(fileSystem.GetDriveName(path)) = 0 And VBA.Left$(path, 1) <> "\" Then
        path = fileSystem.BuildPath(folder, path)
    End If
    If Not fileSystem.FileExists(path) Then ex_UiStyleDiagnostics.fn_Raise "PaletteFileMissing", path
    Set document = VBA.CreateObject("Msxml2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    document.validateOnParse = False
    document.setProperty "ProhibitDTD", True
    If Not document.Load(path) Then ex_UiStyleDiagnostics.fn_Raise "PaletteLoad", path & " | " & document.parseError.reason
    Set root = document.documentElement
    If root Is Nothing Then ex_UiStyleDiagnostics.fn_Raise "PaletteRoot", path
    If root.baseName <> "palette" Or root.namespaceURI <> "urn:excelprototype:profiles" Then
        ex_UiStyleDiagnostics.fn_Raise "PaletteRootInvalid", path
    End If
    If ex_UiElementFactory.fn_Attribute(root, "version") <> "1" Then ex_UiStyleDiagnostics.fn_Raise "PaletteVersion", path
    name = ex_UiElementFactory.fn_Attribute(root, "name")
    If VBA.Len(name) = 0 Then ex_UiStyleDiagnostics.fn_Raise "PaletteName", path
    If m_palettes.Exists(name) Then ex_UiStyleDiagnostics.fn_Raise "PaletteDuplicate", name
    Set colors = VBA.CreateObject("Scripting.Dictionary")
    colors.CompareMode = VBA.vbTextCompare
    For Each node In root.ChildNodes
        If node.NodeType = 1 Then
            If node.baseName <> "color" Or node.namespaceURI <> root.namespaceURI Then
                ex_UiStyleDiagnostics.fn_Raise "PaletteElements", path
            End If
            key = ex_UiElementFactory.fn_Attribute(node, "name")
            value = ex_UiElementFactory.fn_Attribute(node, "value")
            If VBA.Len(key) = 0 Or key Like "*[!A-Za-z0-9_.-]*" Then ex_UiStyleDiagnostics.fn_Raise "PaletteColorName", path
            If colors.Exists(key) Then ex_UiStyleDiagnostics.fn_Raise "PaletteColorDuplicate", key
            If (VBA.Len(value) <> 7 And VBA.Len(value) <> 9) Or VBA.Left$(value, 1) <> "#" Then
                ex_UiStyleDiagnostics.fn_Raise "PaletteColorValue", key & "=" & value
            End If
            If VBA.Mid$(value, 2) Like "*[!0-9A-Fa-f]*" Then ex_UiStyleDiagnostics.fn_Raise "PaletteColorValue", key & "=" & value
            colors.Add key, value
        End If
    Next node
    If colors.Count = 0 Then ex_UiStyleDiagnostics.fn_Raise "PaletteEmpty", path
    m_palettes.Add name, colors
End Sub

Private Sub private_Repaint()
    Dim shapes As Object
    Dim shape As Object
    Dim entry As Object
    Dim area As Range
    Dim value As String
    Dim color As Long
    Dim transparency As Double
    Dim alpha As Long
    Dim index As Long

    Set shapes = VBA.CreateObject("Scripting.Dictionary")
    shapes.CompareMode = VBA.vbTextCompare
    For Each shape In m_worksheet.Shapes
        Set shapes(shape.Name) = shape
    Next shape
    ' Popup controls can delete transient shapes without rebuilding the page.
    For index = m_paints.Count To 1 Step -1
        Set entry = m_paints(index)
        If VBA.Len(entry("shape")) > 0 Then
            If Not shapes.Exists(entry("shape")) Then
                m_paintKeys.Remove entry("key")
                m_paints.Remove index
            End If
        End If
    Next index
    ' Replay only colors in the original application order, including state overlays.
    For Each entry In m_paints
        Set shape = Nothing
        Set area = entry("area")
        If VBA.Len(entry("shape")) > 0 Then
            If Not shapes.Exists(entry("shape")) Then GoTo NextPaint
            Set shape = shapes(entry("shape"))
        End If
        value = Me.ResolveColor(entry("value"))
        If VBA.Left$(value, 1) = "#" Then value = VBA.Mid$(value, 2)
        transparency = 0
        If entry("kind") = "back" And VBA.Len(value) = 8 Then
            alpha = VBA.CLng("&H" & VBA.Left$(value, 2))
            transparency = 1# - alpha / 255#
            value = VBA.Mid$(value, 3)
        End If
        If VBA.Len(value) <> 6 Or value Like "*[!0-9A-Fa-f]*" Then ex_UiStyleDiagnostics.fn_Raise "PaletteResolvedColor", value
        color = VBA.RGB(VBA.CLng("&H" & VBA.Left$(value, 2)), _
            VBA.CLng("&H" & VBA.Mid$(value, 3, 2)), VBA.CLng("&H" & VBA.Right$(value, 2)))
        Select Case entry("kind")
            Case "back"
                If shape Is Nothing Then
                    If transparency = 1 Then
                        area.Interior.Pattern = xlNone
                    ElseIf transparency = 0 Then
                        area.Interior.Pattern = xlSolid
                        area.Interior.Color = color
                    Else
                        ex_UiStyleDiagnostics.fn_Raise "PaletteAlpha"
                    End If
                Else
                    shape.Fill.ForeColor.RGB = color
                    shape.Fill.Transparency = transparency
                End If
            Case "text"
                If shape Is Nothing Then
                    area.Font.Color = color
                Else
                    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color
                End If
            Case "border"
                If shape Is Nothing Then
                    If area.Borders(entry("edge")).LineStyle <> xlLineStyleNone Then
                        area.Borders(entry("edge")).Color = color
                    End If
                Else
                    shape.Line.ForeColor.RGB = color
                End If
            Case "shadow"
                shape.Shadow.ForeColor.RGB = color
        End Select
NextPaint:
    Next entry
End Sub