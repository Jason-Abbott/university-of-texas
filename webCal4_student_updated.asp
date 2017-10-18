<!--#include file="data/webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' updated 09/16/1999

' if cancel is hit then send student back to calendar

if Request.Form("cancel") = "Cancel" then
	response.redirect "webCal4_student_week.asp?type=" & Request.Form("grade")
end if

dim intStaff, instStudent, strTitle, strStart, strDescription, strEnd
dim strDate, db, strQuery, rsEvent, intID

intStaff = Request.Form("staff_id")
intStudent = Request.Form("student_id")
strTitle = Request.Form("title")
strDate = DateValue(Request.Form("date"))
strDescription = "phone: " & Request.Form("phone") & VbCrLf _
	& "e-mail: " & Request.Form("email") & VbCrLf _
	& "reason: " & replace(Request.Form("reason"), "'", "''")

strStart = TimeValue(Request.Form("date"))
	
Select Case CStr(Request.Form("type"))
	case "15"
		strEnd = DateAdd("n", 15, strStart)
	case "30"
		strEnd = DateAdd("n", 30, strStart)
	case "mock"
		strEnd = DateAdd("n", 60, strStart)
End Select

'------------------------------------
' check db for three types of conflicting events
'------------------------------------

' find events on this day for this staff person

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D" _
	& " ON E.event_id = D.event_id) WHERE" _
	& " (E.staff_id=" & intStaff & ")" _
	& " AND (D.event_date=" & strDelim & strDate & strDelim & ")"
	
' match existing events that begin during the new event
	
strQuery = strQuery _
	& " AND ((E.time_start BETWEEN " & strDelim & strStart & strDelim _
	& " AND " & strDelim & strEnd & strDelim & ")"
	
' or those that end during the new event

strQuery = strQuery _
	& " OR (E.time_end BETWEEN " & strDelim & strStart & strDelim _
	& " AND " & strDelim & strEnd & strDelim & ")"

' or begin before and end after the new event
	
strQuery = strQuery _
	& " OR (E.time_start < " & strDelim & strStart & strDelim _
	& " AND E.time_end > " & strDelim & strEnd & strDelim & ")"
	
strQuery = strQuery & ")"

Set rsEvents = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001

if rsEvents.RecordCount > 0 then
	rsEvents.Close
	Set rsEvents = nothing
	response.redirect "webCal4_student_week.asp?error=conflict&staff_id=" _
		& intStaff & "&type=" & Request.Form("grade")
else

' otherwise update db with the new event

	Set db = Server.CreateObject("ADODB.Connection")
	db.Open strDSN

'------------------------------------
' update cal_events
'------------------------------------

	strQuery = "INSERT INTO cal_events (" _
		& "event_title, event_description, student_id, " _
		& "staff_id, event_type, event_recur, time_start, " _
		& "time_end, show_student, show_staff" _
		& ") VALUES ('" _
		& strTitle & "', '" _
		& strDescription & "', " _
		& intStudent & ", " _
		& intStaff & ", '" _
		& Request.Form("type") & "', '" _
		& "none', '" _
		& strStart & "', '" _
		& strEnd & "', " _
		& "0, 1)"
	
	db.Execute strQuery,,&H0001
	
' event dates and views are keyed to event info by the event ID,
' so find out what auto-id was assigned
	
	strQuery = "SELECT event_id, event_title FROM cal_events " _
		& "WHERE event_title='" & strTitle _
		& "' ORDER BY event_id DESC"
	Set rsEvent = db.Execute(strQuery,,&H0001)
	intID = rsEvent("event_id")
	rsEvent.Close
	Set rsEvent = nothing

'------------------------------------
' update cal_dates
'------------------------------------

	strQuery = "INSERT INTO cal_dates (" _
		& "event_id, event_date) VALUES ('" _
		& intID & "', '" & strDate & "')"
	db.Execute strQuery,,&H0001

	db.Close
	Set db = nothing

' with the data updated send user back to calendar
' or to the edit page again, if requested

	response.redirect "webCal4_student_week.asp?type=" & Request.Form("grade")
end if
%>