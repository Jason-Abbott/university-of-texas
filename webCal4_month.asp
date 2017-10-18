<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_rollovers.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_buttons.js"></SCRIPT>
<!--#include file="data/webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 09/17/1999

dim strFirst, strLast, d, intCol, strDescription, intRow
dim arEvents(31), m, y
dim intIndex, strQuery, strView

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

' determine how this page was called and assign values
' for month and year accordingly

if Request.Form("month") <> "" then
	m = CDbl(Request.Form("month"))
	y = CDbl(Request.Form("year"))
elseif Request.QueryString("date") <> "" then
	m = Month(Request.QueryString("date"))
	y = Year(Request.QueryString("date"))
else
	m = Month(Date)
	y = Year(Date)
end if

strNext = DateAdd("m", 1, DateSerial(y, m, 1))
strPrev = DateAdd("m", -1, DateSerial(y, m, 1))
strView = "month"

' ---------------------------------------------------------
' build an array of event data for selected month
' ---------------------------------------------------------
' find the numeric value of the first day of the month
' ie Sunday = 1, Wednesday = 4

strFirst = WeekDay(Dateserial(y, m, 1))

' find the last day by subtracting 1 day from the first day
' of the next month

strLast = Day(Dateserial(y, Month(strNext), 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

strLastMonth = Day(Dateserial(y, m, 1) - 1)

' are we viewing someone else's calendar?

if Request("staff_id") <> "" then
	intID = Request("staff_id")
else
	intID = Session("StudentID")
end if

' find all events occuring between the first and
' last second of the selected month

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) WHERE (D.event_date " _
	& "BETWEEN " & strDelim _
	& m & "/1/" & y & " 12:00:00 AM" & strDelim & " AND " _
	& strDelim & m & "/" & strLast & "/" & y _
	& " 11:59:59 PM" & strDelim & ") " _
	& "AND E.staff_id=" & intID _
	& " ORDER BY E.time_start"
	
' put all matching events in an array indexed by day number

Set rsEvents = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001

do while not rsEvents.EOF
	intIndex = Day(rsEvents("event_date"))

	Select Case rsEvents("event_type")
		Case "30"
			strColor = "green"
		Case "15"
			strColor = "blue"
		Case "mock"
			strColor = "red"
		Case else
			strColor = "black"
	End Select

	strDescription = Replace(TimeValue(rsEvents("time_start")), ":00 ", " ") _
		& " to " & Replace(TimeValue(rsEvents("time_end")), ":00 ", " ")

' determine the level of event detail to reveal

	if intID = Session("StudentID") then
		arEvents(intIndex) = arEvents(intIndex) _
			& "<img src=""./images/arrow_right_" & strColor _
			& ".gif"" width=4 height=7> " & VbCrLf _
			& "<a href=""webCal4_detail.asp?event_id=" & rsEvents("event_id") _
			& "&date=" & rsEvents("event_date") & "&view=month"" " _
			& showStatus(strDescription) & ">" _
			& rsEvents("event_title") & "</a><br>" & VbCrLf
	elseif rsEvents("show_staff") = 1 then
		arEvents(intIndex) = arEvents(intIndex) _
			& "<img src=""./images/arrow_right_" & strColor _
			& ".gif"" width=4 height=7> " & rsEvents("event_title") _
			& "<br>" & VbCrLf
	else	
		arEvents(intIndex) = arEvents(intIndex) _
			& "<img src=""./images/arrow_right_black.gif"" " _
			& "width=4 height=7> <i>" & strDescription & "</i><br>" & VbCrLf
	end if

	rsEvents.MoveNext
loop

rsEvents.Close
set rsEvents = nothing

' create the text to display at the top of the calendar

strQuery = "SELECT Last_Name, First_Name FROM tblSTAFF WHERE " _
	& "pwid=" & intID

Set rsUser = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsUser.Open strQuery, strDSN, 0, 1, &H0001
strTitle = "<font face=""Verdana, Arial, Helvetica"" color=""#" _
	& arColor(6) & """ size=4><b>" _
	& rsUser("First_Name") & " " & rsUser("Last_Name") & "</b></font>"
rsUser.Close
Set rsUser = nothing
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<!-- calendar table -->

<!-- display query while debugging -->
<!-- <font face="Tahoma" size=1><%=strQuery%></font><p> -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<!--#include file="webCal4_buttons.inc"-->
<tr>
	<td bgcolor="#<%=arColor(6)%>" align="center" colspan=2>
	<table width="100%" border=0 cellspacing=1 cellpadding=1>
	<tr>
		<td bgcolor="#<%=arColor(2)%>" colspan=8 align="center">
		<font face="Verdana, Arial, Helvetica" size=2 color="#<%=arColor(6)%>">
		<b><%=MonthName(m) & " " & y%></b></font></td>
	<tr>

<%
' print all the day names as headings

for intCol = 1 to 7
	response.write "   <td width=""14%"" align=""center"" bgcolor=""#" & arColor(2) & """>" _
		& VbCrLf & "<font face=""Verdana, Arial, Helvetica"" size=1>" _
		& WeekDayName(intCol,0) & "</font></td>" & VbCrLf
next

response.write "<td></td><tr>"

' ---------------------------------------------------------
' now generate calendar body
' ---------------------------------------------------------

' the column variable keeps constant track of the
' current calendar column

intCol = 0

strFont = "<font face=""Tahoma, Arial, Helvetica"" size=2><b>"
strNonDay = "<td valign=""top"" bgcolor=""#" & arColor(11) _
	& """>" & strFont & "<font color=""#" & arColor(13) & """>" & VbCrLf

' cycle through all the days previous to the first
' day of the active month

for d = 1 to strFirst - 1
	response.write strNonDay & strLastMonth - strFirst + d + 1 _
		& "</b></font></font></td>" & VbCrLf
	intCol = intCol + 1
next

' now cycle through all the days of the current month

intRow = 1
for d = 1 to strLast
	intCol = intCol + 1
	response.write "<td height=45 valign=""top"""

	if y & m & d = Year(now) & Month(now) & Day(now) then
		response.write " bgcolor=""#" & arColor(8) & """"
	elseif intColn = 1 or intCol = 7 then
		response.write " bgcolor=""" & arColor(10) & """"
	else
		response.write " bgcolor=""" & arColor(9) & """"
	end if

' only allow calendar owners to see link for adding events
	
	if intID = Session("StudentID") then
		response.write ">" & strFont & VbCrLf _
			& VbCrLf & "<a href=""webCal4_edit.asp?action=new&" _
			& "date=" & Dateserial(y, m, d) & "&view=month"" " _
			& VbCrLf _
			& showStatus("Add a new event to " & DateSerial(y, m, d)) _
			& ">" & d & "</a></b></font>"
	else
		response.write ">" & strFont & d & "</b></font>" & VbCrLf
	end if
		
' if the day contains events then generate link to day detail
		
	if arEvents(d) <> "" then
		response.write " <a href=""webCal4_day.asp?date=" _
			& Dateserial(y, m, d) & "&staff_id=" & intID & """ " _
			& switchIcon("Day" & d, "Day", "View " & Dateserial(y, m, d) & " detail") _
			& "><img name=""Day" & d & """ src=""./images/day_grey.gif"" border=0></a>"
	end if

	response.write "<br>" & VbCrLf & "<font face=""Arial, Helvetica"" size=1>" _
		& arEvents(d) & "</font></td>" & VbCrLf
	if intCol = 7 AND d <= strLast then
	
' if we're at the last column then generate link to week view
	
		response.write "<td valign=""center""><a href=""webCal4_week.asp?date=" _
			& DateSerial(y, m, d) & "&staff_id=" & intID & """ " _
			& switchIcon("Week" & intRow, "Week", "View week " & intRow) _
			& "><img name=""Week" & intRow _
			& """ src=""./images/week_grey.gif"" border=0></a></td>"
		
' only start a new row if days of the month remain

		if d < strLast then
			response.write "<tr>"
		end if
		response.write VbCrLf
		
		intCol = 0
		intRow = intRow + 1
	end if
next

' finally, cycle through as many days of the next month as
' necessary to fill the calendar grid through column 7

if intCol > 0 then
	d = 1
	do while intCol < 7
		response.write strNonDay & d & "</font></b></font></td>" & VbCrLf
		d = d + 1
		intCol = intCol + 1
	loop
	response.write "<td valign=""center""><a href=""webCal4_week.asp?date=" _
		& strNext & """ " & switchIcon("Week", "Week", "View week " & intRow) _
		& "><img name=""Week"" src=""./images/week_grey.gif"" border=0></a></td>"
end if
%>
	</table>
	</td>
</table>

<!--#include file="webCal4_switch.inc"-->

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://boise.uidaho.edu/jason/webCal.html" target="_top">
webCal 4.0</a>
</font>

</body>
</html>