<html>
<head>
<!--#include file="webCal4_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form method="post" action="webCal4_rules-edit.asp?view=<%=Request.QueryString("view")%>&action=form">
<tr>
	<td bgcolor="#<%=arColor(3)%>" colspan=2>
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>Appointment Rules</b></font>
	</td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<input type="submit" name="add" value="Add">
	</td>
	
	<td><font face="Tahoma, Arial, Helvetica" size=2>a new rule</font></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<input type="submit" name="edit" value="Edit">
	</td>
	
	<td><font face="Tahoma, Arial, Helvetica" size=2>the selected rule</font></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<input type="submit" name="delete" value="Delete">
	</td>
	
	<td><font face="Tahoma, Arial, Helvetica" size=2>the selected rule</font></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>select:</font></td>

	<td>
	<select name="rule_id">
	
<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/25/1999

dim strQuery, rsRules, rsLimit, strLimit

strQuery = "SELECT * FROM cal_rules" _
	& " WHERE staff_id=" & Session("StudentID") _
	& " ORDER BY time_start"

Set rsRules = Server.CreateObject("ADODB.RecordSet")
rsRules.Open strQuery, strDSN, 3, 1, &H0001
do while not rsRules.EOF
	response.write "<option value=" & rsRules("rule_id") _
		& ">" & rsRules("rule_name") & VbCrLF
	rsRules.MoveNext
loop

rsRules.Close
Set rsRules = nothing
%>
	</select>
	</td>
	
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<input type="submit" name="limit" value="Limit">
	</td>

<%
strQuery = "SELECT * FROM cal_staff" _
	& " WHERE staff_id=" & Session("StudentID")

Set rsLimit = Server.CreateObject("ADODB.RecordSet")
rsLimit.Open strQuery, strDSN, 3, 1, &H0001
intCount = CInt(rsLimit.RecordCount)
if intCount > 0 then
	strLimit = " value=""" & rsLimit("mock_limit") & """"
	intID = rsLimit("cal_staffID")
else
	strLimit = ""
end if
rsLimit.Close
Set rsLimit = nothing

%>
	
	<td><input name="mock_limit" type="text" size=2<%=strLimit%>>
	<font face="Tahoma, Arial, Helvetica" size=2>mock(s) per week</font></td>
	
<% if intCount > 0 then %>
	<input type="hidden" name="type" value="edit">
	<input type="hidden" name="limit_id" value="<%=intID%>">
<% end if %>

</form>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

</body>
</html>