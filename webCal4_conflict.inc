<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' updated 09/13/1999

dim strSay, strFont, strStaff, strStudent

if Request.Form("staff") = "on" then
	strStaff = " checked"
else
	strStaff = ""
end if

if Request.Form("student") = "on" then
	strStudent = " checked"
else
	strStudent = ""
end if

if Request.Form("skip") = "on" then
	strSkip = " checked"
else
	strSkip = ""
end if

strFont = "<font face=""Tahoma, Arial, Helvetica"" size=2>"

if rsEvents.RecordCount > 1 then
	strSay = "events"
else
	strSay = "event"
end if

if Request.Form("event_recur") <> "none" then
	strSay = ", recurring between " & Request.Form("start_date") _
		& " and " & Request.Form("end_date") & " from " _
		& strStart & " to " & strEnd & ", is in conflict with the following " _
		& strSay & ":"
else
	strSay = " on " & Request.Form("start_date") & " from " _
		& strStart & " to " & strEnd & " is in conflict with the following " _
		& strSay & ":"
end if
%>

<html>
<head>
<!--#include file="webCal4_rollovers.inc"-->
<!--#include file="webCal4_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" width="60%" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form action="webCal4_edit.asp?action=conflict&view=<%=Request.QueryString("view")%>" method="post">
<tr>
	<td bgcolor="#<%=arColor(3)%>">
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>Conflict detected</b></font></td>
<tr>
	<td>

<%=strFont%>
The event you attempted to schedule<%=strSay%>
<p>
<table width="100%" cellspacing=0 cellpadding=2 border=0>
<tr>
	<td align="center" bgcolor="#<%=arColor(12)%>"><%=strFont%><b>Title</b></font></td>
	<td align="center" bgcolor="#<%=arColor(12)%>"><%=strFont%><b>Date</b></font></td>
	<td align="center" bgcolor="#<%=arColor(12)%>"><%=strFont%><b>Start</b></font></td>
	<td align="center" bgcolor="#<%=arColor(12)%>"><%=strFont%><b>End</b></font></td>
<tr>

<%
' this assumes that rsEvents has been populated with
' the conflicting events

do while not rsEvents.EOF
	response.write "<td align=""center"">" & strFont _
		& "<a href=""webCal4_detail.asp?event_id=" & rsEvents("event_id") _
		& "&date=" & rsEvents("event_date") & "&view=" & Request.QueryString("view") _
		& """ " & showStatus("View '" & rsEvents("event_title") & "' details") & ">" _
		& rsEvents("event_title") & "</a></font></td>" _
		& "<td align=""center"">" & strFont & DateValue(rsEvents("event_date")) _
		& "</font></td>" _
		& "<td align=""center"">" & strFont & TimeValue(rsEvents("time_start")) _
		& "</font></td>" _
		& "<td align=""center"">" & strFont & TimeValue(rsEvents("time_end")) _
		& "</font></td>"
	rsEvents.MoveNext
loop
%>
</table>
<p>
If this error is unexpected it may indicate that the conflicting event was added to your calendar since you last viewed it.
	</td>
<tr>
	<td align="center" bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="resched" value="Reschedule">
		<input type="submit" name="cancel" value="Cancel">
	</td>
</table>
</center>

<input type="hidden" name="title" value="<%=Request.Form("title")%>">
<input type="hidden" name="description" value="<%=Request.Form("description")%>">
<input type="hidden" name="start_date" value="<%=Request.Form("start_date")%>">
<input type="hidden" name="end_date" value="<%=Request.Form("end_date")%>">
<input type="hidden" name="event_recur" value="<%=Request.Form("event_recur")%>">
<input type="hidden" name="start_hour" value="<%=Request.Form("start_hour")%>">
<input type="hidden" name="start_min" value="<%=Request.Form("start_min")%>">
<input type="hidden" name="end_hour" value="<%=Request.Form("end_hour")%>">
<input type="hidden" name="end_min" value="<%=Request.Form("end_min")%>">
<input type="hidden" name="staff" value="<%=strStaff%>">
<input type="hidden" name="student" value="<%=strStudent%>">
<input type="hidden" name="skip" value="<%=strSkip%>">
</form>

</html>
</body>