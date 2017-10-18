<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 06/04/1999

dim dayFirst, dayLast, d, cal, col
dim events(31), rowCurrent, rowTotal, m, y, mLoop, yLoop
dim mNext, mPrev, yNext, yPrev

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

' determine how this page was called and assign values
' for month and year accordingly

m = Month(Request.QueryString("date"))
y = Year(Request.QueryString("date"))

if m < 12 then
	mNext = m + 1
else
	mNext = 1
end if

' find the numeric value of the first day of the month
' ie Sunday = 1, Wednesday = 4

dayFirst = WeekDay(Dateserial(y, m, 1))

' find the last day by subtracting 1 day from the first day
' of the next month (no need for yNext here)

dayLast = Day(Dateserial(y, mNext, 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

dayLastMonth = Day(Dateserial(y, m, 1) - 1)

' ---------------------------------------------------------
' build an array of event data for selected month
' ---------------------------------------------------------
' find all events occuring between the first and
' last second of the selected month

if Session(dataName & "User") = "" then
 Session(dataName & "User") = 0
end if
%>
<!--#include file="data/webCal4_data.inc"-->
<%
query = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
	& " ON (E.event_id = D.event_id) " _
	& " WHERE event_date BETWEEN #" _
	& m & "/1/" & y & " 12:00:00 AM# " _
	& "AND #" & m & "/" & dayLast & "/" & y _
	& " 11:59:59 PM# AND (user_id = " _
	& Session(dataName & "User") & " OR " _
	& "private = " & False & ") ORDER BY event_date"

' put all matching events in an array indexed by day number

' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

Set rs = db.Execute(query,,&H0001)
do while not rs.EOF
	events(Day(rs("event_date"))) = events(Day(rs("event_date"))) _
		& "<img src=""./images/arrow_right.gif"" width=4 height=7>" & VbCrLf _
		& rs("event_title") & "<br>" & VbCrLf
	rs.movenext
loop
rs.Close
db.Close
Set rs = nothing
Set db = nothing
%>

<html>
<head>
<title>Print <%=MonthName(m) & " " & y%></title></head>
<body bgcolor="#ffffff" link="#330033" vlink="#330033" alink="#330033">

<!-- calendar table -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<tr>
	<td><font face="Verdana, Arial, Helvatica" size=5>
		<b><%=MonthName(m) & " " & y%></b></font>
	</td>
<tr>
	<td align="center">
	<table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">
	<tr>

<%
' print all the day names as headings

for column = 1 to 7
	response.write "   <td width=""14.3%"" align=""center"">" _
		& VbCrLf & "<font face=""Verdana, Arial, Helvetica"" size=1>" _
		& WeekDayName(column,0) & "</font></td>" & VbCrLf
next

response.write "<tr>"

' ---------------------------------------------------------
' now generate calendar body
' ---------------------------------------------------------

' the column variable keeps constant track of the
' current calendar column

column = 0

font = "<font face=""Tahoma, Arial, Helvetica"" size=2><b>"
nondayFormat = "<td valign=""top"" bgcolor=""#d0d0d0"">" & font _
	& "</b><font size=1 color=""#909090"">" & VbCrLf

' cycle through all the days previous to the first
' day of the active month

for d = 1 to dayFirst - 1
	response.write nondayFormat & dayLastMonth - dayFirst + d + 1 _
		& "</font></font></td>" & VbCrLf
	column = column + 1
next

' now cycle through all the days of the current month

for d = 1 to dayLast
	column = column + 1
	response.write "<td height=45 valign=""top"""
	if column = 1 or column = 7 then
		response.write " bgcolor=""#e0e0e0"""
	end if
	response.write ">" & font & VbCrLf _
		& VbCrLf & d & "</b></font><br>" _
		& "<font face=""Arial, Helvetica"" size=1>" _
		& events(d) & "</font></td>" & VbCrLf
	if column = 7 AND d < dayLast then
		response.write "<tr>" & VbCrLf
		column = 0
	end if
next

' finally, cycle through as many days of the next month as
' necessary to fill the calendar grid through column 7

if column > 0 then
	d = 1
	do while column < 7
		response.write nondayFormat & d & "</font></font></td>" & VbCrLf
		d = d + 1
		column = column + 1
	loop
end if
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