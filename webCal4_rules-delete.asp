<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

dim rsRules, strQuery, strView

strQuery = "SELECT * FROM cal_rules WHERE " _
	& "(rule_id)=" & Request.QueryString("rule_id")
	
Set rsRules = Server.CreateObject("ADODB.RecordSet")
		
'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsRules.Open strQuery, strDSN, 3, 1, &H0001
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form action="webCal4_rules-deleted.asp" method="post">
<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=3><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Rule Deletion</b></font></td>
<tr>
	<td align="center" colspan=3><font face="Arial, Helvetica">
	Are you sure you want to erase <i><%=rsRules("rule_name")%></i>?</font></td>
<tr>
	<td colspan=3 align="center" bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="delete" value="Yes">
		<input type="submit" name="cancel" value="No">
	</td>
<tr>
	<td align="center" colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased rules cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="rule_id" value="<%=Request.QueryString("rule_id")%>">
</form>

<%
rsRules.Close
Set rsRules = nothing
%>

</center>
</body>
</html>