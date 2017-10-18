<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 09/16/1999

' if the cancel button was hit then return the user
' to the event detail page

if Request.Form("cancel") = "No" then
	response.redirect "webCal4_detail.asp?date=" & Request.Form("date") _
		& "&event_id=" & Request.Form("event_id") _
		& "&view=" & Request.Form("view") _
		& "&type=" & Request.Form("type")
end if
%>
<!--#include file="webCal4_data.inc"-->
<%
' otherwise prepare to delete some event dates
' figure out which event dates need to be purged

dim strQuery, db

Set db = Server.CreateObject("ADODB.Connection")
db.Open strDSN

Select Case Request.Form("scope")
	Case "one"

' if deleting only one occurrence then erase only
' a single day, leaving event info intact

		strQuery = " AND event_date BETWEEN " & strDelim _
			& Request.Form("date") & " 12:00:00 AM" & strDelim & " AND " & strDelim _
			& Request.Form("date") & " 11:59:59 PM" & strDelim
	Case "future"

' if deleting all future events then erase today
' and all after today, leaving event info intact

		strQuery = " AND event_date >= " & strDelim _
			& Request.Form("date") & strDelim
	Case Else

' if erasing all occurrences then delete not only the dates
' but the event information itself
	
			db.Execute "DELETE FROM cal_events WHERE (event_id)=" _
				& Request.QueryString("event_id"),,&H0001
end Select

' put the query together

strQuery = "DELETE FROM cal_dates" _
	& " WHERE event_id=" & Request.Form("event_id") _
	& strQuery

' and run it

db.Execute strQuery,,&H0001
db.Close
Set db = nothing

' send the user back to the calendar

response.redirect "webCal4_" & Request.Form("type") & "-" _
	& Request.Form("view") & ".asp?date=" & Request.Form("date")
%>