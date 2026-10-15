VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_DslRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_maxDepth As Long
Private m_sharedTemplates As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal sharedTemplates As Object) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If sharedTemplates Is Nothing Then
        Err.Raise 5, , "Shared DSL templates are required."
    End If
    Set m_sharedTemplates = sharedTemplates
    m_maxDepth = ex_PEB_Export.fn_Count("export.dsl.maxDepth")
    If m_maxDepth < 1 Or m_maxDepth > 64 Then
        Err.Raise 5, , "DSL maxDepth must be between 1 and 64."
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_sharedTemplates = Nothing
End Sub

Public Function TryRender( _
    ByVal template As String, _
    ByVal values As Object, _
    ByRef rendered As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim context As Object

    On Error GoTo EH
    rendered = VBA.vbNullString
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Template renderer is not initialized."
    End If
    If values Is Nothing Then
        Err.Raise 5, , "Template context is required."
    End If
    Set context = private_Clone(values)
    rendered = private_Normalize(private_Render(private_Layout(template), context, 0))
    TryRender = True
    Exit Function
EH:
    rendered = VBA.vbNullString
    diagnostic = "Template: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_Render( _
    ByVal template As String, _
    ByVal values As Object, _
    ByVal depth As Long _
) As String
    Dim cursor As Long
    Dim opening As Long
    Dim closing As Long
    Dim ending As Long
    Dim alternative As Long
    Dim afterBlock As Long
    Dim token As String
    Dim expression As String
    Dim body As String
    Dim output As String
    Dim name As String
    Dim collectionName As String
    Dim separator As Long
    Dim items As Collection
    Dim item As Variant
    Dim child As Object
    Dim key As Variant
    Dim index As Long
    Dim itemKey As String
    Dim letOpening As Long
    Dim letEnding As Long

    If depth > m_maxDepth Then
        Err.Raise 5, , "Template nesting/include limit exceeded."
    End If
    cursor = 1
    Do While cursor <= VBA.Len(template)
        opening = VBA.InStr(cursor, template, "{")
        letOpening = VBA.InStr(cursor, template, "#let ")
        If letOpening > 0 And (opening = 0 Or letOpening < opening) Then
            output = output & VBA.Mid$(template, cursor, letOpening - cursor)
            letEnding = private_FindUnquoted(template, ";", letOpening + 5)
            If letEnding = 0 Then
                Err.Raise 5, , "Missing semicolon after #let."
            End If
            private_Let VBA.Mid$(template, letOpening + 5, letEnding - letOpening - 5), values, depth
            cursor = letEnding + 1
            GoTo ContinueRender
        End If
        If opening = 0 Then
            output = output & VBA.Mid$(template, cursor)
            Exit Do
        End If
        output = output & VBA.Mid$(template, cursor, opening - cursor)
        closing = private_FindUnquoted(template, "}", opening + 1)
        If closing = 0 Then
            Err.Raise 5, , "Unclosed template token at " & VBA.CStr(opening)
        End If
        token = VBA.Trim$(VBA.Mid$(template, opening + 1, closing - opening - 1))
        cursor = closing + 1
        If VBA.Left$(token, 4) = "#if " Then
            private_FindBlock template, cursor, "#if ", "#endif", ending, alternative, afterBlock
            expression = VBA.Trim$(VBA.Mid$(token, 5))
            If private_Condition(expression, values) Then
                If alternative = 0 Then
                    alternative = ending
                End If
                body = VBA.Mid$(template, cursor, alternative - cursor)
            ElseIf alternative > 0 Then
                body = VBA.Mid$(template, alternative + VBA.Len("{#else}"), ending - alternative - VBA.Len("{#else}"))
            Else
                body = VBA.vbNullString
            End If
            output = output & private_Render(body, values, depth + 1)
            cursor = afterBlock
        ElseIf VBA.Left$(token, 5) = "#for " Then
            private_FindBlock template, cursor, "#for ", "#endfor", ending, alternative, afterBlock
            expression = VBA.Trim$(VBA.Mid$(token, 6))
            separator = VBA.InStr(1, expression, " in ", VBA.vbBinaryCompare)
            If separator < 2 Then
                Err.Raise 5, , "Invalid template loop: " & token
            End If
            name = VBA.Trim$(VBA.Left$(expression, separator - 1))
            collectionName = VBA.Trim$(VBA.Mid$(expression, separator + 4))
            If Not values.Exists(collectionName) Then
                Err.Raise 5, , "Unknown template collection: " & collectionName
            End If
            If Not VBA.IsObject(values(collectionName)) Then
                Err.Raise 5, , "Template collection required: " & collectionName
            End If
            Set items = values(collectionName)
            body = VBA.Mid$(template, cursor, ending - cursor)
            index = 0
            For Each item In items
                index = index + 1
                Set child = VBA.CreateObject("Scripting.Dictionary")
                child.CompareMode = VBA.vbTextCompare
                For Each key In values.Keys
                    If VBA.Left$(VBA.CStr(key), VBA.Len(name) + 1) <> name & "." Then
                        If VBA.IsObject(values(key)) Then
                            Set child(key) = values(key)
                        Else
                            child(key) = values(key)
                        End If
                    End If
                Next key
                If Not VBA.IsObject(item) Then
                    Err.Raise 5, , "Template loop items must be dictionaries."
                End If
                For Each key In item.Keys
                    itemKey = name & "." & VBA.CStr(key)
                    If VBA.IsObject(item(key)) Then
                        Set child(itemKey) = item(key)
                    Else
                        child(itemKey) = item(key)
                    End If
                Next key
                child(name & ".__first") = (index = 1)
                child(name & ".__last") = (index = items.Count)
                child(name & ".__index") = index
                output = output & private_Render(body, child, depth + 1)
            Next item
            cursor = afterBlock
        ElseIf VBA.Left$(token, 9) = "#include " Then
            name = VBA.Trim$(VBA.Mid$(token, 10))
            If Not m_sharedTemplates.Exists(name) Then
                Err.Raise 5, , "Shared XML template not found: " & name
            End If
            body = private_Layout(VBA.CStr(m_sharedTemplates(name)))
            output = output & private_Render(body, values, depth + 1)
        ElseIf token = "#newline" Then
            output = output & VBA.vbCrLf
        ElseIf token = "#space" Then
            output = output & " "
        ElseIf VBA.Left$(token, 1) = "#" Then
            Err.Raise 5, , "Unsupported or misplaced DSL directive: " & token
        Else
            output = output & private_Value(token, values)
        End If
ContinueRender:
    Loop
    private_Render = output
End Function

Private Sub private_FindBlock( _
    ByVal template As String, _
    ByVal cursor As Long, _
    ByVal startToken As String, _
    ByVal endToken As String, _
    ByRef ending As Long, _
    ByRef alternative As Long, _
    ByRef afterBlock As Long _
)
    Dim opening As Long
    Dim closing As Long
    Dim token As String
    Dim nesting As Long

    ending = 0
    alternative = 0
    nesting = 1
    Do
        opening = VBA.InStr(cursor, template, "{")
        If opening = 0 Then
            Err.Raise 5, , "Missing DSL terminator: " & endToken
        End If
        closing = private_FindUnquoted(template, "}", opening + 1)
        If closing = 0 Then
            Err.Raise 5, , "Unclosed DSL token."
        End If
        token = VBA.Trim$(VBA.Mid$(template, opening + 1, closing - opening - 1))
        If VBA.Left$(token, VBA.Len(startToken)) = startToken Then
            nesting = nesting + 1
        End If
        If token = endToken Then
            nesting = nesting - 1
            If nesting = 0 Then
                ending = opening
                afterBlock = closing + 1
                Exit Sub
            End If
        End If
        If token = "#else" And startToken = "#if " And nesting = 1 Then
            If alternative > 0 Then
                Err.Raise 5, , "Duplicate #else in template."
            End If
            If VBA.Mid$(template, opening, closing - opening + 1) <> "{#else}" Then
                Err.Raise 5, , "Use {#else} without extra spaces."
            End If
            alternative = opening
        End If
        cursor = closing + 1
    Loop
End Sub

Private Function private_Condition( _
    ByVal expression As String, _
    ByVal values As Object _
) As Boolean
    Dim separator As Long
    Dim value As Variant
    Dim result As Boolean

    separator = VBA.InStr(1, expression, " #or ", VBA.vbBinaryCompare)
    If separator > 0 Then
        private_Condition = private_Condition(VBA.Left$(expression, separator - 1), values) Or _
            private_Condition(VBA.Mid$(expression, separator + 5), values)
        Exit Function
    End If
    separator = VBA.InStr(1, expression, " #and ", VBA.vbBinaryCompare)
    If separator > 0 Then
        private_Condition = private_Condition(VBA.Left$(expression, separator - 1), values) And _
            private_Condition(VBA.Mid$(expression, separator + 6), values)
        Exit Function
    End If
    If VBA.Left$(expression, 5) = "#not " Then
        private_Condition = Not private_Condition(VBA.Trim$(VBA.Mid$(expression, 6)), values)
        Exit Function
    End If
    value = ex_PEB_Export.fn_Required(values, private_Key(expression))
    If VBA.VarType(value) = VBA.vbBoolean Then
        result = VBA.CBool(value)
    Else
        result = (VBA.Len(ex_PEB_Export.fn_Text(value)) > 0)
    End If
    private_Condition = result
End Function

Private Function private_Key(ByVal expression As String) As String
    expression = VBA.Trim$(expression)
    expression = VBA.Replace$(expression, "[", VBA.vbNullString)
    expression = VBA.Replace$(expression, "]", VBA.vbNullString)
    private_Key = expression
End Function

Private Function private_Value( _
    ByVal expression As String, _
    ByVal values As Object _
) As String
    Dim parts As Variant
    Dim value As Variant
    Dim formatter As String
    Dim argument As String
    Dim colon As Long
    Dim index As Long
    Dim offset As Double
    Dim arguments As Collection
    Dim rx As Object
    Dim parsedDate As Date

    Set arguments = private_Split(expression, "|")
    ReDim parts(0 To arguments.Count - 1)
    For index = 1 To arguments.Count
        parts(index - 1) = arguments(index)
    Next index
    value = ex_PEB_Export.fn_Required(values, private_Key(VBA.CStr(parts(0))))
    For index = 1 To UBound(parts)
        formatter = VBA.Trim$(VBA.CStr(parts(index)))
        argument = VBA.vbNullString
        colon = VBA.InStr(formatter, ":")
        If colon > 0 Then
            argument = VBA.Trim$(VBA.Mid$(formatter, colon + 1))
            formatter = VBA.Left$(formatter, colon - 1)
        End If
        Select Case formatter
            Case "upper"
                value = VBA.UCase$(ex_PEB_Export.fn_Text(value))
            Case "lower"
                value = VBA.LCase$(ex_PEB_Export.fn_Text(value))
            Case "upperFirstLetter"
                value = ex_PEB_Export.fn_Text(value)
                value = VBA.UCase$(VBA.Left$(value, 1)) & VBA.Mid$(value, 2)
            Case "lowerFirstLetter"
                value = ex_PEB_Export.fn_Text(value)
                value = VBA.LCase$(VBA.Left$(value, 1)) & VBA.Mid$(value, 2)
            Case "dateformat"
                argument = private_Unquote(argument)
                If VBA.Len(argument) = 0 Then
                    Err.Raise 5, , "dateformat requires a format."
                End If
                parsedDate = private_Date(value)
                If VBA.InStr(argument, "\") > 0 Then
                    value = private_DateFormat(parsedDate, argument)
                Else
                    value = VBA.Format$(parsedDate, argument)
                End If
            Case "dateoffset"
                argument = private_Unquote(argument)
                If Not VBA.IsNumeric(argument) Then
                    Err.Raise 5, , "Invalid dateoffset."
                End If
                offset = VBA.CDbl(argument)
                If offset <> VBA.Fix(offset) Then
                    Err.Raise 5, , "dateoffset requires whole days."
                End If
                value = VBA.DateAdd("d", offset, private_Date(value))
            Case "regexreplace"
                Set arguments = private_Split(argument, ",")
                If arguments.Count <> 2 Then
                    Err.Raise 5, , "regexreplace requires two quoted arguments."
                End If
                Set rx = VBA.CreateObject("VBScript.RegExp")
                rx.Global = True
                rx.IgnoreCase = False
                rx.MultiLine = True
                rx.Pattern = private_Unquote(VBA.CStr(arguments(1)))
                value = rx.Replace(ex_PEB_Export.fn_Text(value), private_Unquote(VBA.CStr(arguments(2))))
            Case Else
                Err.Raise 5, , "Unsupported DSL formatter: " & formatter
        End Select
    Next index
    private_Value = ex_PEB_Export.fn_Text(value)
End Function

Private Function private_Clone(ByVal values As Object) As Object
    Dim result As Object
    Dim key As Variant

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    For Each key In values.Keys
        If VBA.IsObject(values(key)) Then
            Set result(key) = values(key)
        Else
            result(key) = values(key)
        End If
    Next key
    Set private_Clone = result
End Function

Private Function private_Layout(ByVal text As String) As String
    Dim rx As Object

    ' Обрабатываем разметку до подстановок, чтобы не менять пользовательские данные.
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "(\{#\^\}|#\^)[ \t]*(\r\n|\r|\n)[ \t]*"
    text = rx.Replace(text, VBA.vbNullString)
    text = VBA.Replace$(text, "{#^}", VBA.vbNullString)
    text = VBA.Replace$(text, "#^", VBA.vbNullString)
    rx.Pattern = "(\{#_\}|#_)[ \t]*"
    private_Layout = VBA.Trim$(rx.Replace(text, VBA.vbNullString))
End Function

Private Function private_Normalize(ByVal text As String) As String
    Dim lines As Variant
    Dim index As Long
    Dim rx As Object

    text = VBA.Replace$(text, VBA.vbCrLf, VBA.vbLf)
    text = VBA.Replace$(text, VBA.vbCr, VBA.vbLf)
    lines = VBA.Split(text, VBA.vbLf)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "[ \t]+"
    For index = LBound(lines) To UBound(lines)
        lines(index) = VBA.Trim$(rx.Replace(VBA.CStr(lines(index)), " "))
    Next index
    text = VBA.Join(lines, VBA.vbCrLf)
    Do While VBA.Left$(text, 2) = VBA.vbCrLf
        text = VBA.Mid$(text, 3)
    Loop
    Do While VBA.Right$(text, 2) = VBA.vbCrLf
        text = VBA.Left$(text, VBA.Len(text) - 2)
    Loop
    private_Normalize = text
End Function

Private Function private_FindUnquoted( _
    ByVal text As String, _
    ByVal delimiter As String, _
    ByVal first As Long _
) As Long
    Dim index As Long
    Dim quoted As Boolean
    Dim character As String

    index = first
    Do While index <= VBA.Len(text)
        character = VBA.Mid$(text, index, 1)
        If character = """" Then
            If quoted And VBA.Mid$(text, index + 1, 1) = """" Then
                index = index + 1
            Else
                quoted = Not quoted
            End If
        ElseIf Not quoted And character = delimiter Then
            private_FindUnquoted = index
            Exit Function
        End If
        index = index + 1
    Loop
    If quoted Then
        Err.Raise 5, , "Unclosed quoted DSL argument."
    End If
End Function

Private Function private_Split( _
    ByVal text As String, _
    ByVal delimiter As String _
) As Collection
    Dim result As New Collection
    Dim first As Long
    Dim separator As Long

    first = 1
    Do
        separator = private_FindUnquoted(text, delimiter, first)
        If separator = 0 Then
            result.Add VBA.Trim$(VBA.Mid$(text, first))
            Exit Do
        End If
        result.Add VBA.Trim$(VBA.Mid$(text, first, separator - first))
        first = separator + 1
    Loop
    Set private_Split = result
End Function

Private Function private_Unquote(ByVal text As String) As String
    text = VBA.Trim$(text)
    If VBA.Len(text) < 2 Then
        Err.Raise 5, , "Quoted DSL argument required."
    End If
    If VBA.Left$(text, 1) <> """" Or VBA.Right$(text, 1) <> """" Then
        Err.Raise 5, , "Quoted DSL argument required."
    End If
    text = VBA.Mid$(text, 2, VBA.Len(text) - 2)
    private_Unquote = VBA.Replace$(text, """""", """")
End Function

Private Sub private_Let( _
    ByVal statement As String, _
    ByVal values As Object, _
    ByVal depth As Long _
)
    Dim separator As Long
    Dim opening As Long
    Dim name As String
    Dim expression As String
    Dim functionName As String
    Dim arguments As Collection
    Dim leftValue As String
    Dim rightValue As String
    Dim rx As Object

    separator = VBA.InStr(statement, "=")
    If separator = 0 Then
        Err.Raise 5, , "Invalid #let assignment."
    End If
    name = VBA.Trim$(VBA.Left$(statement, separator - 1))
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^__[A-Za-z][A-Za-z0-9_]*$"
    If Not rx.Test(name) Then
        Err.Raise 5, , "#let requires an internal variable starting with __."
    End If
    expression = VBA.Trim$(VBA.Mid$(statement, separator + 1))
    opening = VBA.InStr(expression, "(")
    If opening = 0 Or VBA.Right$(expression, 1) <> ")" Then
        Err.Raise 5, , "Invalid #let function call."
    End If
    functionName = VBA.Left$(expression, opening - 1)
    Set arguments = private_Split(VBA.Mid$(expression, opening + 1, VBA.Len(expression) - opening - 1), ",")
    If arguments.Count <> 2 Then
        Err.Raise 5, , "#let function requires two arguments."
    End If
    leftValue = private_Render(private_Unquote(VBA.CStr(arguments(1))), values, depth + 1)
    Select Case functionName
        Case "$ex_Helpers.m_RegexIsMatch"
            ' Разрешённая DSL-операция, без вызова произвольного VBA из XML.
            rx.Pattern = private_Unquote(VBA.CStr(arguments(2)))
            rx.IgnoreCase = True
            values(name) = rx.Test(leftValue)
        Case "$ex_Helpers.fn_IsDateOnOrAfter"
            rightValue = private_Render(private_Unquote(VBA.CStr(arguments(2))), values, depth + 1)
            If VBA.Len(leftValue) = 0 Or VBA.Len(rightValue) = 0 Then
                values(name) = False
            Else
                values(name) = (private_Date(leftValue) >= private_Date(rightValue))
            End If
        Case Else
            Err.Raise 5, , "Unsupported #let function: " & functionName
    End Select
End Sub

Private Function private_Date(ByVal value As Variant) As Date
    Dim parts As Variant
    Dim text As String
    Dim parsed As Date
    Dim rx As Object

    If VBA.VarType(value) = VBA.vbDate Then
        private_Date = VBA.CDate(value)
        Exit Function
    End If
    text = ex_PEB_Export.fn_Text(value)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^\d{2}\.\d{2}\.\d{4}$"
    If Not rx.Test(text) Then
        Err.Raise 5, , "DSL date requires DD.MM.YYYY: " & text
    End If
    parts = VBA.Split(text, ".")
    parsed = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(parsed, "dd.mm.yyyy") <> text Then
        Err.Raise 5, , "Invalid DSL calendar date: " & text
    End If
    private_Date = parsed
End Function

Private Function private_DateFormat( _
    ByVal value As Date, _
    ByVal pattern As String _
) As String
    pattern = VBA.Replace$(pattern, "\month", ex_PEB_Export.fn_Config("export.dsl.months." & VBA.CStr(VBA.Month(value))))
    pattern = VBA.Replace$(pattern, "\yyyy", VBA.Format$(value, "yyyy"))
    pattern = VBA.Replace$(pattern, "\yy", VBA.Format$(value, "yy"))
    pattern = VBA.Replace$(pattern, "\dd", VBA.Format$(value, "dd"))
    pattern = VBA.Replace$(pattern, "\d", VBA.CStr(VBA.Day(value)))
    pattern = VBA.Replace$(pattern, "\mm", VBA.Format$(value, "mm"))
    pattern = VBA.Replace$(pattern, "\m", VBA.CStr(VBA.Month(value)))
    If VBA.InStr(pattern, "\") > 0 Then
        Err.Raise 5, , "Unsupported DSL date-format token."
    End If
    private_DateFormat = pattern
End Function