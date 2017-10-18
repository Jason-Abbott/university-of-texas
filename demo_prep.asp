<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/06/1999

' send the user to the appropriate week view

if Request.Form("signup") = "on" then
	strGoto = "signup"
else
	strgoto = "student"
end if

if Request.Form("bba_id") <> "" then
	Session("StudentID") = Request.Form("bba_id")
	response.redirect "webCal4_" & strGoto & "-week.asp?type=BBA"
elseif Request.Form("mba_id") <> "" then
	Session("StudentID") = Request.Form("mba_id")
	response.redirect "webCal4_" & strGoto & "-week.asp?type=MBA"
elseif Request.Form("staff_id") <> "" then
	Session("StudentID") = Request.Form("staff_id")
	response.redirect "webCal4_counsel-week.asp"
elseif Request.Form("other_id") <> "" then
	Session("StudentID") = Request.Form("other_id")
	response.redirect "webCal4_staff-week.asp"
end if
%>

<html>
<body bgcolor="#c0c0c0">
<font face="Tahoma, Arial, Helvetica" size=5>
Test the calendar using:<br>
</font>

<table>

<!-- undergraduate -->

<tr>
	<form action="demo_prep.asp" method="post">
	<td align="right">
	<input type="submit" value="BBA student">
	</td>
	
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<select name="bba_id">
<%
dim strQuery, rsBBA, rsMBA

strQuery = "SELECT ID_NUMBER, NAME_FIRST, NAME_LAST FROM " _
	& "tblStudents WHERE C_DEGREE_ < 3 AND NAME_LAST IS NOT NULL " _
	& "ORDER BY NAME_LAST, NAME_FIRST"

Set rsUsers = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsUsers.Open strQuery, strDSN, 0, 1, &H0001
do while not rsUsers.EOF
	response.write "<option value=""" & rsUsers("ID_NUMBER") _
		& """>" & rsUsers("NAME_LAST") & ", " & rsUsers("NAME_FIRST") _
		& VBCrLf
	rsUsers.MoveNext
loop
rsUsers.Close
%>
	</select>
	<input type="checkbox" name="signup">set counseling appointment
	</font>
	</td>
	</form>

<!-- graduate -->
	
<tr>
	<form action="demo_prep.asp" method="post">
	<td align="right">
	<input type="submit" value="MBA student">
	</td>
	
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<select name="mba_id">
<%
strQuery = "SELECT ID_NUMBER, NAME_FIRST, NAME_LAST FROM " _
	& "tblStudents WHERE C_DEGREE_ > 2 AND NAME_LAST IS NOT NULL " _
	& "ORDER BY NAME_LAST, NAME_FIRST"
rsUsers.Open strQuery, strDSN, 0, 1, &H0001
do while not rsUsers.EOF
	response.write "<option value=""" & rsUsers("ID_NUMBER") _
		& """>" & rsUsers("NAME_LAST") & ", " & rsUsers("NAME_FIRST") _
		& VBCrLf
	rsUsers.MoveNext
loop
rsUsers.Close
%>
	</select>
	<input type="checkbox" name="signup">set counseling appointment
	</font>
	</td>
	</form>
	
<!-- counselor -->
	
<tr>
	<form action="demo_prep.asp" method="post">
	<td align="right">
	<input type="submit" value="Counselor">
	</td>
	
	<td>
	<select name="staff_id">
<%
strQuery = "SELECT pwid, First_Name, Last_Name FROM " _
	& "tblSTAFF WHERE CSOgroup LIKE '%Counselors' " _
	& "ORDER BY Last_Name, First_Name"
rsUsers.Open strQuery, strDSN, 0, 1, &H0001
do while not rsUsers.EOF
	response.write "<option value=""" & rsUsers("pwid") _
		& """>" & rsUsers("Last_Name") & ", " & rsUsers("First_Name") _
		& VBCrLf
	rsUsers.MoveNext
loop
rsUsers.Close
%>
	</select>
	</td>
	</form>

<!-- other staff -->
	
<tr>
	<form action="demo_prep.asp" method="post">
	<td align="right">
	<input type="submit" value="Other Staff">
	</td>
	
	<td>
	<select name="other_id">
<%
strQuery = "SELECT pwid, First_Name, Last_Name FROM " _
	& "tblSTAFF WHERE CSOgroup NOT LIKE '%Counselors' " _
	& "ORDER BY Last_Name, First_Name"
rsUsers.Open strQuery, strDSN, 0, 1, &H0001
do while not rsUsers.EOF
	response.write "<option value=""" & rsUsers("pwid") _
		& """>" & rsUsers("Last_Name") & ", " & rsUsers("First_Name") _
		& VBCrLf
	rsUsers.MoveNext
loop
rsUsers.Close
Set rsUsers = nothing
%>
	</select>
	</td>
	</form>
	
</table>

</body>
</html>
