<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/10/1999

' This include assumes that arScopes and arGroups have already
' been defined and populated, and we're being sent the total
' number of groups in the "total" form element

' JavaScript form validation should ensure that something was selected

dim strQuery, intCount

strQuery = ""
strGroups = ""
strScopes = ""

' 0 = group id
' 1 = group name
' 2 = group permissions
' 3 = group visibility
' 4 = group e-mail setting

' list the groups we want to view

intCount = 0
for x = 0 to UBound(arGroups)
	if arGroups(x,3) then
		strGroups = strGroups & arGroups(x,0) & ","
		intCount = intCount + 1
	end if
next
strGroups = Left(strGroups, Len(strGroups)-1)

if intCount = 1 then
	strGroups = "V.group_id=" & strGroups
elseif intCount <> UBound(arGroups)+1 AND intCount > 0 then
	strGroups = "V.group_id IN (" & strGroups & ")"
else
	strGroups = ""
end if

' list the scopes we want to view

intCount = 0
for x = 0 to 2
	if arScopes(x) then
		strScopes = strScopes & x & ","
		intCount = intCount + 1
	end if
next
strScopes = Left(strScopes, Len(strScopes)-1)

if intCount = 1 then
	strScopes = "V.event_scope=" & strScopes
elseif intCount <> 3 AND intCount > 0 then
	strScopes = "V.event_scope IN (" & strScopes & ")"
else
	strScopes = ""
end if

' now put the two together

if strGroups <> "" AND strScopes <> "" then
	strQuery = " AND "
end if

strQuery = strGroups & strQuery & strScopes

if strQuery <> "" then
	strQuery = "AND E.event_id IN " _
		& "(SELECT V.event_id FROM cal_views V " _
		& "WHERE " & strQuery & ") "
end if

strQuery = strQuery & "ORDER BY D.event_date, E.time_start"

'response.write strQuery

Session(dataName & "Query") = strQuery
%>