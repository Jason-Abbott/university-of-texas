<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/18/1999

dim strError, strStatus, rsUsers, arTemp, arGroups, strQueryTemp

strStatus = "This action is available only to registered users"
strError = "The information you entered could not be validated. " _
	& "Please try again."
%>
<!--#include file="data/webCal4_data.inc"-->
<%
if Request.Form("login") <> "" then
	strQuery = "SELECT * FROM cal_users WHERE " _
		& "login = '" & Request.Form("login") & "'"
	Set rsUsers = db.Execute(strQuery,,&H0001)
	if rsUsers.EOF = -1 then

' if no records match the login information then show the
' login form again with an error
	
		strStatus = strError
	else
		if rsUsers("password") = Request.Form("password") then
			Session(dataName & "User") = rsUsers("user_id")

' now make a 2D Session array that lists groups and properties
' relevant to the logged in user
' 0 = group id
' 1 = group name
' 2 = group permissions
' 3 = group visibility
' 4 = group e-mail setting (unused feature)

			strQuery = "SELECT * FROM cal_permissions P INNER JOIN" _
				& " cal_groups G ON (P.group_id = G.group_id)" _
				& " WHERE user_id =" & rsUsers("user_id")
				
			Set rsAccess = Server.CreateObject("ADODB.RecordSet")

' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

			rsAccess.Open strQuery, DSN, 3, 1, &H0001

			intCount = CInt(rsAccess.Recordcount - 1)
			ReDim arGroups(intCount,4)

			for x = 0 to intCount
				arGroups(x,0) = rsAccess("group_id")
				arGroups(x,1) = rsAccess("group_name")
				arGroups(x,2) = rsAccess("access_level")
				arGroups(x,3) = rsAccess("visible")
				arGroups(x,4) = rsAccess("send_email")

				strQuery = strQuery & " OR event_views LIKE '%," _
					& arGroups(x,0) & "%'"

				rsAccess.MoveNext
			next

' set scope options
			
			dim arScopes
			arScopes = Split(rsUsers("user_scopes"), ",")
			Session(dataName & "Scopes") = arScopes
			Session(dataName & "Groups") = arGroups
%>
<!--#include file="webCal4_makesql.inc"-->
<%
			rsUsers.Close
			rsAccess.Close
			db.Close
			Set rsUsers = nothing
			Set rsAccess = nothing
			Set db = nothing
			response.redirect Request.Form("url")
		else
			strStatus = strError
		end if
	end if
end if

' obtain the administrator's e-mail address
query = "SELECT email_name, email_site, user_id " _
	& "FROM cal_users WHERE (user_id) = 1"
Set rs = db.Execute(query,,&H0001)
%>

<html>
<!--#include file="webCal4_themes.inc"-->
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" width="60%" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=3 cellspacing=0 width="100%">
<form action="webCal4_login.asp" method="post">
<tr bgcolor="#<%=arColor(4)%>" valign="bottom">
	<td colspan=4><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Login</b></font></td>
<tr>
	<td colspan=4 align="center"><font face="Arial, Helvetica" size=2>
	<%=strStatus%><br></font></td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>" align="right"><font face="Arial, Helvetica">Username:&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>"><input type="text" name="login" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>" align="right"><font face="Arial, Helvetica">Password:&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>"><input type="password" name="password" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td colspan=4 align="center"><font face="Verdana, Arial, Helvetica" size=2>
	<a href="mailto:<%=rs("email_name")%>@<%=rs("email_site")%>">Request an account</a></font>
	<br>
	<input type="submit" value="Continue"></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<%
rs.Close
db.Close
Set rs = nothing
Set db = nothing

response.write "<input type=""hidden"" name=""url"" value="""
if Request.Form("url") <> "" then
	response.write Request.Form("url")
else
	response.write Request.QueryString("url")
end if
response.write """>"
%>
</form>

</center>
</body>
</html>