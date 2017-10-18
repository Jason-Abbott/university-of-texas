<!--#include file="data/webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/18/1999

' if the user is logging out then erase relevant
' session values and return to calendar

if Request.Form("logout") = "Logout" then
	Session(dataName & "User") = ""
	Session(dataName & "Groups") = ""
	Session(dataName & "Query") = ""
	Session(dataName & "Scopes") = ""
	response.redirect "webCal4_month.asp"
end if

' arScope will contains list of visibility levels

dim arScopes, arGroups, intChecked, strChanged, strScopes

arGroups = Session(dataName & "Groups")
arScopes = Session(dataName & "Scopes")

' 0 = group id
' 1 = group name
' 2 = group permissions
' 3 = group visibility
' 4 = group e-mail setting

' to avoid changing the setting back and forth
' keep track of which groups have already been
' set to visible (1)

strChanged = ""

' go through each of the user's groups and update
' the visibility setting based on form selection

for each x in Request.Form("groups")
	for y = 0 to UBound(arGroups)
		if arGroups(y,0) = CInt(x) then
			arGroups(y,3) = 1
			strChanged = strChanged & "," & y
		elseif InStr(strChanged,y) = 0 then
			arGroups(y,3) = 0
		end if
	next
next
Session(dataName & "Groups") = arGroups

' cycle through each checkbox visibility option
' and update the arScopes values

for x = 0 to 2
	if Request.Form(CStr(x)) = "on" then
		arScopes(x) = 1
		intChecked = intChecked + 1
	else
		arScopes(x) = 0
	end if
next
Session(dataName & "Scopes") = arScopes
%>
<!--#include file="webCal4_makesql.inc"-->
<%

' update database if "make default" was checked
' first update group visibility settings

if Request.Form("remember") = "on" then
	for x = 0 to UBound(arGroups)
		strQuery = "UPDATE cal_permissions SET " _
			& "visible = " & arGroups(x,3) _
			& " WHERE user_id = " & Session(dataName & "User") _
			& " AND group_id = " & arGroups(x,0)

		db.Execute strQuery,,&H0001
	next

' now update scope settings
' generate new scopes string for database

	strScopes = arScopes(0) & "," & arScopes(1) & "," & arScopes(2)

	strQuery = "UPDATE cal_users SET " _
		& "user_scopes = '" & strScopes & "' " _
		& "WHERE (user_id)=" & Session(dataName & "User")
	db.Execute strQuery,,&H0001
end if

db.Close
Set db = nothing

response.redirect "webCal4_month.asp"
%>