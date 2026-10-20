Attribute VB_Name = "ex_PEB_ReturnTvo"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Read(ByVal values As Object) As Collection
    Dim result As New Collection
    Dim period As Object
    Dim plan As Object
    Dim statuses As Object
    Dim names As Collection
    Dim people As Collection
    Dim positions As Collection
    Dim seen As Object
    Dim row As Object
    Dim index As Long
    Dim arrivalDate As Double
    Dim eventDate As Double

    Set fn_Read = result
    If ex_PEB_EventRoutes.fn_Value(values("EventId"), "direction") <> "Arrival" Then Exit Function
    Set plan = ex_PEB_MovementRegistry.fn_Event(values("EventId"))
    If plan("operation") <> "Arrival" Then Exit Function
    Set period = ex_PEB_MovementSource.fn_ReadLatest(values("PersonId"), True)
    If period Is Nothing Then Exit Function
    Set statuses = ex_PEB_MovementRegistry.fn_Statuses()
    If VBA.StrComp(ex_PEB_Export.fn_Text(period("EventName")), statuses(plan("sourceStatus")), VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoPeriodMismatch")
    End If
    eventDate = ex_PEB_MovementSource.fn_OrderDate(values("EventDate"))
    arrivalDate = ex_PEB_MovementSource.fn_OrderDate(period("ArrivalDate"))
    If ex_PEB_MovementSource.fn_OrderDate(period("DepartureDate")) > eventDate Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoPeriodMismatch")
    End If
    ' A previously exported arrival may already have closed this exact period.
    If arrivalDate <> 0 Then
        If arrivalDate <> eventDate Or _
                ex_PEB_Export.fn_Text(period("ArrivalOrder")) <> ex_PEB_Export.fn_Text(values("AcceptedOrderNumber")) Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoPeriodMismatch")
        End If
    End If
    Set names = private_Lines(period("TvoName"))
    Set people = private_Lines(period("TvoPersonId"))
    Set positions = private_Lines(period("TvoPositionCode"))
    If names.Count <> people.Count Or names.Count <> positions.Count Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoInvalidChain")
    End If
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    seen.CompareMode = VBA.vbTextCompare
    seen.Add VBA.Trim$(VBA.CStr(values("PersonId"))), True
    For index = 1 To people.Count
        If seen.Exists(people(index)) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoInvalidChain")
        seen.Add people(index), True
        Set row = VBA.CreateObject("Scripting.Dictionary")
        row("PersonId") = people(index)
        row("PersonName") = names(index)
        row("PositionCode") = positions(index)
        result.Add row
    Next index
End Function

Public Function fn_Encode( _
    ByVal values As Object, _
    ByVal rows As Collection _
) As String
    Dim document As Object
    Dim root As Object
    Dim node As Object
    Dim row As Object
    Dim field As Variant

    Set document = private_Document()
    Set root = document.createElement("returnTvo")
    document.appendChild root
    root.setAttribute "version", "1"
    root.setAttribute "person", values("PersonId")
    root.setAttribute "event", values("EventId")
    root.setAttribute "date", VBA.Format$(ex_PEB_History.fn_Date(values("EventDate")), "dd.mm.yyyy")
    For Each row In rows
        Set node = document.createElement("person")
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("returnTvo.snapshot.fields"), "|")
            node.setAttribute VBA.CStr(field), ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(row, VBA.CStr(field)))
        Next field
        root.appendChild node
    Next row
    fn_Encode = document.XML
    If VBA.Len(fn_Encode) > 32767 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoInvalidSnapshot")
End Function

Public Function fn_Decode(ByVal values As Object) As Collection
    Dim result As New Collection
    Dim document As Object
    Dim root As Object
    Dim node As Object
    Dim row As Object
    Dim field As Variant
    Dim snapshot As String
    Dim seen As Object

    Set fn_Decode = result
    If ex_PEB_EventRoutes.fn_Value(values("EventId"), "direction") <> "Arrival" Then Exit Function
    If Not values.Exists("TvoReturnSnapshot") Then Exit Function
    snapshot = ex_PEB_Export.fn_Text(values("TvoReturnSnapshot"))
    If VBA.Len(snapshot) = 0 Then Exit Function
    Set document = private_Document()
    If Not document.LoadXML(snapshot) Then GoTo InvalidSnapshot
    Set root = document.documentElement
    If root.nodeName <> "returnTvo" Then GoTo InvalidSnapshot
    For Each field In VBA.Array("version", "person", "event", "date")
        If Not root.hasAttribute(VBA.CStr(field)) Then GoTo InvalidSnapshot
    Next field
    If root.getAttribute("version") <> "1" Or root.getAttribute("person") <> values("PersonId") Or _
            root.getAttribute("event") <> values("EventId") Or _
            root.getAttribute("date") <> VBA.Format$(ex_PEB_History.fn_Date(values("EventDate")), "dd.mm.yyyy") Then GoTo InvalidSnapshot
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    seen.CompareMode = VBA.vbTextCompare
    seen.Add VBA.Trim$(VBA.CStr(values("PersonId"))), True
    For Each node In root.childNodes
        If node.nodeType <> 1 Or node.nodeName <> "person" Then GoTo InvalidSnapshot
        Set row = VBA.CreateObject("Scripting.Dictionary")
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("returnTvo.snapshot.fields"), "|")
            If Not node.hasAttribute(VBA.CStr(field)) Then GoTo InvalidSnapshot
            row(field) = node.getAttribute(VBA.CStr(field))
            If VBA.Len(ex_PEB_Export.fn_Text(row(field))) = 0 Then GoTo InvalidSnapshot
        Next field
        If seen.Exists(row("PersonId")) Then GoTo InvalidSnapshot
        seen.Add row("PersonId"), True
        result.Add row
    Next node
    Exit Function
InvalidSnapshot:
    Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoInvalidSnapshot")
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Document() As Object
    Dim document As Object

    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    document.setProperty "ProhibitDTD", True
    Set private_Document = document
End Function

Private Function private_Lines(ByVal value As Variant) As Collection
    Dim result As New Collection
    Dim text As String
    Dim line As Variant

    text = ex_PEB_Export.fn_Text(value)
    text = VBA.Replace$(VBA.Replace$(text, VBA.vbCrLf, VBA.vbLf), VBA.vbCr, VBA.vbLf)
    Do While VBA.Left$(text, 1) = VBA.vbLf
        text = VBA.Mid$(text, 2)
    Loop
    Do While VBA.Right$(text, 1) = VBA.vbLf
        text = VBA.Left$(text, VBA.Len(text) - 1)
    Loop
    Set private_Lines = result
    If VBA.Len(text) = 0 Then Exit Function
    For Each line In VBA.Split(text, VBA.vbLf)
        If VBA.Len(VBA.Trim$(VBA.CStr(line))) = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTvoInvalidChain")
        result.Add VBA.Trim$(VBA.CStr(line))
    Next line
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------