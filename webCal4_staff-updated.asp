<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' updated 10/11/1999

' if cancel is hit then send back to calendar

if Request.Form("cancel") = "Cancel" then
	response.redirect "webCal4_staff-" & Request.QueryString("view") _
		& ".asp?date=" & Request.Form("start_date")
end if

' otherwise begin populating variables

dim strStart, strEnd, strDate, arDates(), intCount, intHide
dim strTitle, strDescription, intID, strQuery, queryEvent
dim strViews, arTemp, arViews, db

'------------------------------------
' normalize some values
'------------------------------------

strDate = DateValue(Request.Form("start_date"))

intID = Request.Form("event_id")

strStart = TimeValue(Request.Form("start_hour") & ":" _
	& Request.Form("start_min"))
strEnd = TimeValue(Request.Form("end_hour") & ":" _
	& Request.Form("end_min"))

if Request.Form("staff") = "on" then
	intStaff = 1
else
	intStaff = 0
end if

if Request.Form("bba") = "on" then
	intBBA = 1
else
	intBBA = 0
end if

if Request.Form("mba") = "on" then
	intMBA = 1
else
	intMBA = 0
end if

strTitle = replace(Request.Form("title"), "'", "''")
strDescription = replace(Request.Form("description"), "'", "''")

' this generates an array of dates on which the event is supposed
' to occur, based on the recurrence type
' we need this list before we can check for conflicting events

intCount = 0
if	Request.Form("event_recur") <> "none" then
	Select Case Request.Form("event_recur")
		Case "daily"
			addType = "d"
			addNum = 1
		Case "weekly"
			addType = "d"
			addNum = 7
		Case "2weeks"
			addType = "d"
			addNum = 14
		Case "monthly"
			addType = "m"
			addNum = 1
		Case "yearly"
			addType = "yyyy"
			addNum = 1
	end Select		

' populate the array with dates, according to the above
' addition, until the end date for the event

	While DateDiff("d", strDate, Request.Form("end_date")) >= 0
		if Request.Form("skip") <> "on" _
			OR (WeekDay(strDate) > 1 AND WeekDay(strDate) < 7) then

			ReDim Preserve arDates(intCount)
			arDates(intCount) = strDate
			intCount = intCount + 1
		end if
		strDate = DateAdd(addType, addNum, strDate)
	Wend

' if there was no recurrence selected then put the single
' date into the array

else
	ReDim Preserve arDates(intCount)
	arDates(intCount) = strDate
end if

' also generate a list of the dates for the subsequent
' query for conflicts

for each x in arDates
	strDates = strDates & strDelim & x & strDelim & ", "
next

strDates = Left(strDates, Len(strDates) - 2)

' when checking for conflicts we have to skip the
' present event

if intID <> "" then
	strQuery = " AND (D.event_id<>" & intID & ")"
else
	strQuery = ""
end if

'------------------------------------
' clear old values out of cal_dates in preparation for new ones
'------------------------------------

Set db = Server.CreateObject("ADODB.Connection")
db.Open strDSN

if Request.Form("edit_type") <> "new" then
	Select Case Request.Form("edit_type")
		Case "one"
		
' erase single date

			strQuery = " AND event_date BETWEEN " & strDelim _
				& strDate & strDelim & " AND " & strDelim _
				& DateAdd("d", 1, strDate) & strDelim
		Case "future"

' erase current and all future dates

			strQuery = " AND event_date >= " & strDelim _
				& strDate & strDelim
		Case "all"

' erase all event dates without limitation

			strQuery = ""
	end Select
	
	strQuery = "DELETE FROM cal_dates" _
		& " WHERE event_id=" & intID _
		& strQuery	

' 0001 is the hex value for adCmdText which tells the connection
' object that we're sending a text command, which is speedier

	db.Execute strQuery,,&H0001
end if

'------------------------------------
' update cal_events and cal_views as needed
'------------------------------------
' only update values if all occurrences of that event
' were selected for modification, otherwise create new
' entries

if Request.Form("edit_type") = "all" then

' update existing event

	strQuery = "UPDATE cal_events SET " _
		& "event_title = '" & strTitle & "', " _
		& "event_description = '" & strDescription & "', " _
		& "event_recur = '" & Request.Form("event_recur") & "', " _
		& "time_start = '" & strStart & "', " _
		& "time_end = '" & strEnd & "', " _
		& "in_bba = " & intBBA & ", " _
		& "in_mba = " & intMBA & ", " _
		& "in_staff = " & intStaff & " " _
		& "WHERE (event_id)=" & intID

	db.Execute strQuery,,&H0001
else

' add new event

	strQuery = "INSERT INTO cal_events (" _
		& "event_title, event_description, event_type, " _
		& "staff_id, event_recur, time_start, time_end, " _
		& "in_staff, in_bba, in_mba" _
		& ") VALUES ('" _
		& strTitle & "', '" _
		& strDescription & "', '" _
		& "other', " _
		& Session("StudentID") & ", '" _
		& Request.Form("event_recur") & "', '" _
		& strStart & "', '" _
		& strEnd & "', " _
		& intStaff & ", " _
		& intBBA & ", " _
		& intMBA & ")"

	db.Execute strQuery,,&H0001
	
' event dates and views are keyed to event info by the event ID,
' so find out what auto-id was assigned
	
	strQuery = "SELECT * FROM cal_events " _
		& "WHERE event_title='" & strTitle & "' AND " _
		& "staff_id=" & Session("StudentID") & " AND " _
		& "time_start='" & strStart & "'"
	Set rsEvent = db.Execute(strQuery,,&H0001)
	intID = rsEvent("event_id")
	rsEvent.Close
	Set rsEvent = nothing
end if

'------------------------------------
' update cal_dates as needed
'------------------------------------
' now go through everything inserted into the dates array
' and insert it into the event dates table

for each d in arDates
	strQuery = "INSERT INTO cal_dates (" _
		& "event_id, event_date) VALUES ('" _
		& intID & "', '" & d & "')"
	db.Execute strQuery,,&H0001
next

'	db.CommitTrans
db.Close
Set db = nothing

' with the data updated send user back to calendar
' or to the edit page again, if requested

if Request.Form("save") = "Save" then
	response.redirect "webCal4_staff-" & Request.QueryString("view") _
		& ".asp?date=" & Request.Form("start_date")
elseif Request.Form("saveadd") = "Save & Add Another" then
	response.redirect "webCal4_staff-edit.asp?date=" _
		& Request.Form("start_date") _
		& "&view=" & Request.QueryString("view")
end if
%>