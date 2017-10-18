<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

dim rsEvents, strQuery, strView

dim strSay
strQuery = "SELECT * FROM cal_events WHERE " _
	& "(event_id)=" & Request.QueryString("event_id")
	
Set rsEvents = Server.CreateObject("ADODB.RecordSet")
		
'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001
	
' how many did they want to delete?

Select Case Request.QueryString("scope")
	Case "one"
		strSay = "this (" & Request.QueryString("date") & ") instance of"
	Case "future"
		strSay = "this (" & Request.QueryString("date") & ") and <b>all future</b> instances of"
	Case "all"
		strSay = "<b>all " & Request.QueryString("count") & "</b> instances of"
	Case else
		strSay = ""
End Select
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form action="webCal4_deleted.asp" method="post">
<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=3><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Event Deletion</b></font></td>
<tr>
	<td align="center" colspan=3><font face="Arial, Helvetica">
	Are you sure you want to erase <%=strSay%> <i><%=rsEvents("event_title")%></i>?</font></td>
<tr>
	<td colspan=3 align="center" bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="delete" value="Yes">
		<input type="submit" name="cancel" value="No">
	</td>
<tr>
	<td align="center" colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased events cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="event_id" value="<%=Request.QueryString("event_id")%>">
<input type="hidden" name="date" value="<%=Request.QueryString("date")%>">
<input type="hidden" name="scope" value="<%=Request.QueryString("scope")%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
<input type="hidden" name="type" value="<%=Request.QueryString("type")%>">
</form>

<%
rsEvents.Close
Set rsEvents = nothing
%>

</center>
</body>
</html>