<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/06/1999

' if the cancel button was hit then return
' to the user admin page

if Request.Form("cancel") = "Cancel" then
	response.redirect "webCal4_admin.asp?view=" & Request.Form("view")
end if
%>
<!--#include file="data/webCal4_data.inc"-->
<%

dim query

' first erase the user

query = "DELETE FROM cal_users WHERE user_id=" & Request.Form("user_id")

' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

db.Execute query,,&H0001

' then deal with that user's events, if any

if Request.Form("event_count") > 0 then
	select case Request.Form("do")
		case "delete"
' delete all events created by this user
			query = "DELETE FROM cal_events " _
				& "WHERE user_id=" & Request.Form("user_id")
		case "some"
' first delete all private events
			query = "DELETE FROM cal_events " _
				& "WHERE user_id=" & Request.Form("user_id") _
				& " AND private=" & True
			db.Execute query,,&H0001
' then move the remaining public events to new user
			query = "UPDATE cal_events SET " _
				& "user_id = " & Request.Form("recipient") _
				& " WHERE user_id=" & Request.Form("user_id")
		case "move"
' move all events to new user
			query = "UPDATE cal_events SET " _
				& "user_id = " & Request.Form("recipient") _
				& " WHERE user_id=" & Request.Form("user_id")
	end select
	db.Execute query,,&H0001
end if

db.Close
Set db = nothing

' send the user back to user management

response.redirect "webCal4_user-admin.asp?view=" _
	& Request.Form("view")
%>