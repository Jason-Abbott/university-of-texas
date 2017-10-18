<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 05/27/1999

' check to see if the user has already logged in
' if not, send them to the login page

if Session(dataName & "User") = "" OR Session(dataName & "User") = 0 then
	response.redirect "webCal4_login.asp?url=" _
		& Request.ServerVariables("URL") & "?" _
		& Server.URLEncode(Request.ServerVariables("QUERY_STRING"))
end if
%>