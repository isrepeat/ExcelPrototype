Attribute VB_Name = "ex_PEB_WordStoreTests"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Run() As Boolean
    Dim store As obj_PEB_WordEventStore
    Dim reopened As obj_PEB_WordEventStore
    Dim section As Object
    Dim values As Object
    Dim keys As Collection
    Dim events As Collection
    Dim item As Object
    Dim uid As String
    Dim firstUid As String
    Dim secondUid As String
    Dim previous As String
    Dim diagnostic As String
    Dim rejected As Boolean

    On Error GoTo EH
    Set store = New obj_PEB_WordEventStore
    private_Assert store.Initialize(VBA.vbNullString, "order-fixture", diagnostic), diagnostic
    Set section = store.BindSection("Return", "FromAnnualMainVacation", "FromAnnualMainVacation", "fixture" & VBA.vbCrLf & "plan", "fixture" & VBA.vbCrLf & "shared")
    private_Assert store.Text(section, "plan") = "fixture" & VBA.vbLf & "plan", "Plan line endings were not normalized."
    Set section = store.BindSection("Return", "FromAnnualMainVacation", "FromAnnualMainVacation", "fixture" & VBA.vbCrLf & "plan", "fixture" & VBA.vbCrLf & "shared")
    Set values = private_Values("1000000001", "PERSON A", "17.10.2026")
    Set keys = private_Keys(values)
    store.Change "INSERT", section, values, keys, uid
    firstUid = uid
    private_Assert VBA.Len(uid) > 0, "INSERT did not allocate a UID."
    Set events = store.Events(section)
    private_Assert events.Count = 1, "INSERT did not add one event."

    ' A repeat, including a changed report number, must leave the store untouched.
    previous = store.Xml
    values("IncomingNo") = "changed-report"
    uid = VBA.vbNullString
    On Error Resume Next
    store.Change "INSERT", section, values, keys, uid
    rejected = (VBA.Err.Number <> 0)
    VBA.Err.Clear
    On Error GoTo EH
    private_Assert rejected And store.Xml = previous And uid = VBA.vbNullString, "Duplicate INSERT changed the store."
    Set section = store.Sections.Item(0)

    ' UPDATE addresses the saved UID instead of guessing from the new fields.
    uid = firstUid
    store.Change "UPDATE", section, values, keys, uid
    Set events = store.Events(section)
    Set item = events(1)
    private_Assert uid = firstUid And item("IncomingNo") = "changed-report", "UPDATE did not preserve the UID."

    values("DateFromShort") = "18.10.2026"
    values("IncomingNo") = "second-report"
    uid = VBA.vbNullString
    store.Change "INSERT", section, values, keys, uid
    secondUid = uid
    private_Assert firstUid <> secondUid, "UID was reused."
    Set events = store.Events(section)
    private_Assert events.Count = 2, "Two distinct events did not accumulate."
    private_Assert events(1)("PersonKey") = events(2)("PersonKey"), "IPN did not reuse the person map."

    values("RankGenitive") = "new shared rank"
    uid = secondUid
    store.Change "UPDATE", section, values, keys, uid
    Set events = store.Events(section)
    private_Assert events(1)("RankGenitive") = "new shared rank", "Rank is not shared within the document."
    private_Assert events(1)("IncomingNo") = "changed-report", "UPDATE changed another event's report."

    Set reopened = New obj_PEB_WordEventStore
    private_Assert reopened.Initialize(store.Xml, "order-fixture", diagnostic), diagnostic
    private_Assert reopened.Xml = store.Xml, "Save/reopen changed the store."
    reopened.Dispose
    Set reopened = New obj_PEB_WordEventStore
    private_Assert Not reopened.Initialize(store.Xml, "other-order", diagnostic), "Different order was accepted."

    uid = firstUid
    store.Change "DELETE", Nothing, Nothing, Nothing, uid
    Set events = store.Events(section)
    private_Assert events.Count = 1 And events(1)("WordEventUid") = secondUid, "DELETE removed the wrong event."
    uid = secondUid
    store.Change "DELETE", Nothing, Nothing, Nothing, uid
    private_Assert store.Events(section).Count = 0, "Deleting the last event did not empty its section."

    Set values = private_Values(VBA.vbNullString, "PERSON WITHOUT IPN", "19.10.2026")
    Set keys = private_Keys(values)
    uid = VBA.vbNullString
    store.Change "INSERT", section, values, keys, uid
    values("DateFromShort") = "20.10.2026"
    uid = VBA.vbNullString
    store.Change "INSERT", section, values, keys, uid
    Set events = store.Events(section)
    private_Assert events(1)("PersonKey") = events(2)("PersonKey"), "FIO/position fallback did not reuse the person."

    previous = store.Xml
    uid = "missing-uid"
    rejected = False
    On Error Resume Next
    store.Change "DELETE", Nothing, Nothing, Nothing, uid
    rejected = (VBA.Err.Number <> 0)
    VBA.Err.Clear
    On Error GoTo EH
    private_Assert rejected And store.Xml = previous, "Missing UID changed the store."
    fn_Run = True
Cleanup:
    If Not store Is Nothing Then
        store.Dispose
    End If
    If Not reopened Is Nothing Then
        reopened.Dispose
    End If
    Exit Function
EH:
    ex_Core.fn_Diagnostic_WriteError "WordStore.Tests", VBA.Err.Number, VBA.Err.Source, VBA.Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Values( _
    ByVal ipn As String, _
    ByVal fio As String, _
    ByVal eventDate As String _
) As Object
    Dim values As Object

    Set values = VBA.CreateObject("Scripting.Dictionary")
    values.Add "EventId", "FromAnnualMainVacation"
    values.Add "TemplateName", "FromAnnualMainVacation"
    values.Add "PersonId", ipn
    values.Add "PersonName", fio
    values.Add "PositionCode", "fixture-position"
    values.Add "AcceptedOrderNumber", "300"
    values.Add "OrderDateShort", "17.10.2026"
    values.Add "DateFromShort", eventDate
    values.Add "RankGenitive", "fixture-rank"
    values.Add "PositionGenitive", "fixture-position-name"
    values.Add "IncomingNo", "first-report"
    values.Add "IncomingDateShort", "17.10.2026"
    Set private_Values = values
End Function

Private Function private_Keys(ByVal values As Object) As Collection
    Dim result As New Collection
    Dim key As Variant

    For Each key In values.Keys
        result.Add VBA.CStr(key)
    Next key
    Set private_Keys = result
End Function

Private Sub private_Assert( _
    ByVal success As Boolean, _
    ByVal diagnostic As String _
)
    If Not success Then
        Err.Raise 5, , diagnostic
    End If
End Sub