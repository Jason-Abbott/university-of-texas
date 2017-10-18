<!--#include file="data/webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 06/04/1999

dim dayFirst, dayLast, dayNames, d, col
dim events(7), dayShow, prevFirst, nextFirst

daySelect = Request.QueryString("date")
dayFirst = DateAdd("d", 1 - WeekDay(daySelect), daySelect)
dayLast = DateAdd("d", 6, dayFirst)

' ---------------------------------------------------------
' build an array of event data for selected week
' ---------------------------------------------------------

query = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
	& " ON (E.event_id = D.event_id) " _
	& " WHERE event_date BETWEEN #" _
	& dayFirst & " 12:00:00 AM# AND #" & dayLast & " 11:59:59 PM#" _
	& " AND (user_id = " & Session(dataName & "User") & " OR " _
	& "private = " & False & ") ORDER BY event_date"
	
' put all matching events in an array indexed by day number

%>
<!--#include file="show_status.inc"-->
<%
' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

Set rs = db.Execute(query,,&H0001)
do while not rs.EOF
	index = WeekDay(rs("event_date"))
	events(index) = events(index) _
		& "<img src=""./images/arrow_right_black.gif"" width=4 height=7>" _
		& rs("event_title") & "<br>" & VbCrLf
	rs.movenext
loop

rs.Close
db.Close
set rs = nothing
set db = nothing
%>

<html>
<head>
</head>
<body bgcolor="#ffffff">

<font face="Verdana, Arial, Helvatica" size=5>
<b><nobr><%=MonthName(Month(dayFirst)) & " " & Year(dayFirst)%></nobr></b>
</font>

<table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">
<%
' generate heading

dayShow = dayFirst
dim spanArray(1)
m = 0
for col = 1 to 7
	dayNames = dayNames _
		& "<td width=""14.3%"" align=""center"">" _
 		& "<font face=""Verdana, Arial, Helvetica"" size=1>" _
		& WeekDayName(col,0) & "</font></td>"

	spanArray(m) = spanArray(m) + 1
	dayNext = DateAdd("d", 1, dayShow)
	if Month(dayNext) <> Month(dayShow) then
		m = 1
	end if
	dayShow = dayNext
next

if spanArray(0) < 7 then
%>
<tr>
	<td bgcolor="#cccccc" colspan=<%=spanArray(0)%> align="center">
	<font face="Verdana, Arial, Helvetica" size=2>
	<b><%=MonthName(Month(dayFirst)) & " " & Year(dayLast)%></b></font></td>

	<td bgcolor="#cccccc" colspan=<%=spanArray(1)%> align="center">
	<font face="Verdana, Arial, Helvetica" size=2>
	<b><%=MonthName(Month(dayLast)) & " " & Year(dayLast)%></b></font></td>
<% end if %>	
	
<tr><%=dayNames%>
<tr>

<%
dayShow = dayFirst
for col = 1 to 7
%>
	<td height=200 valign="top"
<%	if col = 1 or col = 7 then %>
	bgcolor="#dddddd"
<%	end if %>
	><font face="Tahoma, Arial, Helvetica" size=2><b>
	<%=Day(dayShow)%></b></font>
	<br>
	<font face="Arial, Helvetica" size=1>
	<%=events(col)%></font></td>
<%
	dayShow = DateAdd("d", 1, dayShow)
next
%>

	</table>
	</td>
</table>

<br>
<font face="Verdana, Arial, Helvetica" size=2>
Generated <%=Now%></font>
</font>

</body>
</html>