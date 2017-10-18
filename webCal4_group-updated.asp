<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' updated 08/06/1999

' if cancel is hit then send back to user admin page

if Request.Form("cancel") = "Cancel" then
	response.redirect "webCal4_admin.asp?view=" & Request.Form("view")
end if
%>
<!--#include file="data/webCal4_data.inc"-->
<%
dim query

if Request.Form("edit_type") = "update" then

' update existing user

	query = "UPDATE cal_groups SET " _
		& "name_first = '" & Request.Form("name_first") & "', " _
		& "name_last = '" & Request.Form("name_last") & "', " _
		& "email_name = '" & Request.Form("email_name") & "', " _
		& "email_site = '" & Request.Form("email_site") & "', " _
		& "login = '" & Request.Form("login") & "', " _
		& "password = '" & Request.Form("password") & "', " _
		& "user_groups = '" & Request.Form("dbgroups") & "' " _
		& "WHERE (user_id)=" & Request.Form("user_id")
else

' add new user

	query = "INSERT INTO cal_users (" _
		& "name_first, name_last, email_name, " _
		& "email_site, login, password, user_groups" _
		& ") VALUES ('" _
		& Request.Form("name_first") & "', '" _
		& Request.Form("name_last") & "', '" _
		& Request.Form("email_name") & "', '" _
		& Request.Form("email_site") & "', '" _
		& Request.Form("login") & "', '" _
		& Request.Form("password") & "', '" _
		& Request.Form("dbgroups") & "')"
end if

' 0001 is the hex value for adCmdText which tells the connection
' object that we're sending a text command, which is speedier

db.Execute query,,&H0001
db.Close
Set db = nothing

' with the data updated send user back to user admin
' or to the edit page again, if requested

if Request.Form("save") = "Save" then
	response.redirect "webCal4_user-admin.asp?view=" _
		& Request.Form("view")
elseif Request.Form("saveadd") = "Save & Add Another" then
	response.redirect "webCal4_user-edit.asp?view=" _
		& Request.Form("view")
end if
%>