<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_rollovers.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_buttons.js"></SCRIPT>
<!--#include file="data/webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 09/17/1999

dim strFirst, strLast, d, intCol, strColor, tfWeekends
dim dayShow, strQuery, intTime, strView, intID

' ---------------------------------------------------------
' set parameters
' ---------------------------------------------------------
' view type

strView = "week"

' display weekends?

tfWeekends = True

' this defines the intervals in minutes

intErval = 15
intFactor = 60/intErval

' these define the range of time to display
' the first number is the 24-hour time of day

intRange1 = 6 * intFactor
intRange2 = 22 * intFactor - 1

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

' calculate the number of days to display

if tfWeekends then
	intFirst = 1
	intLast = 7
else
	intFirst = 2
	intLast = 6
end if

if Request.QueryString("date") <> "" then
	strSelect = Request.QueryString("date")
	strFirst = DateAdd("d", intFirst - WeekDay(strSelect), strSelect)
	strLast = DateAdd("d", intLast, strFirst)
else
	strFirst = DateAdd("d", intFirst - WeekDay(Date), Date)
	strLast = DateAdd("d", intLast - WeekDay(Date), Date)
end if

strPrev = DateAdd("d", -intLast, strFirst)
strNext = DateAdd("d", intFirst, strLast)

intDays = intLast - intFirst
intRatio = Round(90/(intDays + 1), 2)

' this function takes a time and converts it to the
' proper number of table segments based on the
' specified interval

intSegments = intFactor

function segments(strTime)
	intMin = Minute(strTime)
	intAdd = intErval/2
	for z = 0 to 60/intErval - 1
		if intMin < intAdd then
			intSegments = z
			exit for
		end if
		intAdd = intAdd + intErval
	next
	segments = intSegments + (Hour(strTime) * intFactor)
end function

' ZERO-BASED count of time segements/day

intTotal = (1440/intErval) - 1
intHeight = 24/intFactor - 1

' are we viewing someone else's calendar?

if Request("staff_id") <> "" then
	intID = Request("staff_id")
else
	intID = Session("StudentID")
end if

' ---------------------------------------------------------
' build an array of event data for selected week
' ---------------------------------------------------------

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) WHERE (D.event_date " _
	& "BETWEEN " & strDelim _
	& strFirst & " 12:00:00 AM" & strDelim & " AND " _
	& strDelim & strLast & " 11:59:59 PM" & strDelim & ") " _
	& "AND E.staff_id=" & intID _
	& " ORDER BY E.time_start"

' put all matching events in an array indexed by day number
' and time segement

dim arEvents()
ReDim arEvents(intDays,intTotal,2)

' this keeps track of which days have events to generate
' the link to day detail

dim arDays()
ReDim arDays(intDays)

Set rsEvents = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001
do while not rsEvents.EOF

' Weekday is 0 based so subtract start day

	d = WeekDay(rsEvents("event_date")) - intFirst
	intTime = segments(rsEvents("time_start"))

' assign appropriate color
	
	Select Case rsEvents("event_type")
		Case "30"
			arEvents(d,intTime,2) = "99ff99"
		Case "15"
			arEvents(d,intTime,2) = "9999ff"
		Case "mock"
			arEvents(d,intTime,2) = "ff9999"
		Case else
			arEvents(d,intTime,2) = arColor(0)
	End Select

	strDescription =  Replace(TimeValue(rsEvents("time_start")), ":00 ", " ") _
			& " to " & Replace(TimeValue(rsEvents("time_end")), ":00 ", " ")
	
' determine the level of event detail to reveal
		
	if intID = Session("StudentID") then
		arEvents(d,intTime,0) = "<a href=""webCal4_detail.asp?event_id=" _
			& rsEvents("event_id") & "&date=" & rsEvents("event_date") _
			& "&view=week"" " & showStatus(strDescription) & ">" _
			& rsEvents("event_title") & "</a>" & VbCrLf
	elseif rsEvents("show_staff") = 1 then
		arEvents(d,intTime,0) = rsEvents("event_title")
	else	
		arEvents(d,intTime,0) = Replace(strDescription, "to ", "to<br>")
		arEvents(d,intTime,2) = arColor(6)
	end if
		
	arEvents(d,intTime,1) = segments(rsEvents("time_end")) - intTime

' signify that events occur on this day	

	arDays(d) = True
	
	rsEvents.MoveNext
loop

rsEvents.Close
set rsEvents = nothing

' now generate the title to display at the top of the calendar

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

<!-- display query while debugging -->
<!-- <font face="Tahoma" size=1><%=strQuery%><p></font> -->

<!-- heading table -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<!--#include file="webCal4_buttons.inc"-->
<tr>
	<td bgcolor="#<%=arColor(6)%>" align="center" colspan=2>
	<table width="100%" border=0 cellspacing=1 cellpadding=0>
	<tr>
		<td rowspan=3 width="10%">&nbsp;</td>
<%
' generate the heading

strShow = strFirst
dim arSpan()
ReDim arSpan(1)
m = 0
for intCol = intFirst to intLast
	if strShow = Date then
		strColor = arColor(8)
	elseif intCol = 1 or intCol = 7 then
		strColor = arColor(10)
	else
		strColor = arColor(9)
	end if
	
	strDays = strDays & "<td width=""" & intRatio & "%"" bgcolor=""#" _
		& strColor & """>" & VbCrLf _
		& "<table cellspacing=0 cellpadding=1 width=""100%""><tr>" & VbCrLf _
		& "<td align=""center"" bgcolor=""#000000"" width=""20%"" valign=""top"">" & VbCrLf _
		& "<font face=""Verdana, Arial, Helvetica"" color=""#ffffff"" size=2><b>" _
		& Day(strShow) & "</b></font></td>" _
		& VbCrLf _
		& "<td align=""center"" rowspan=2>" _
		& "<font face=""Verdana, Arial, Helvetica"" size=1>"
		
' only allow calendar owners to see link for adding events
		
	if intID = Session("StudentID") then
		strDays = strDays _
			& "<a href=""webCal4_edit.asp?action=new&view=week&date=" & strShow & """ " _
			& showStatus("Add a new event to " & strShow) & ">" _
			& WeekDayName(intCol,0) & "</a></td>" _
			& VbCrLf _
			& "<tr>"
	else
		strDays = strDays _
			& WeekDayName(intCol,0) & "</td>" & VbCrLf & "<tr>"
	end if

' if this day has events, display the link to day detail
		
	if arDays(intCol-intFirst) then
		strDays = strDays & "<td>" _
		& "<a href=""webCal4_day.asp?date=" & strShow _
		& "&staff_id=" & intID & """ " _
		& switchIcon("Day" & intCol, "Day", "View " & strShow & " detail") _
		& "><img name=""Day" & intCol & """ src=""./images/day_grey.gif"" width=18 height=7 border=0>"
	else
		strDays = strDays & "<td bgcolor=""#000000"">" _
			& "<img src=""./images/tiny_blank.gif"" height=7>"
	end if
	
	strDays = strDays & "</td>" & VbCrLf & "</table></td>" & VbCrLf
		
	arSpan(m) = arSpan(m) + 1
	strNext = DateAdd("d", 1, strShow)
	if Month(strNext) <> Month(strShow) then
		m = 1
	end if
	strShow = strNext
next
%>
	<tr>
		<td bgcolor="#<%=arColor(2)%>" colspan=<%=arSpan(0)%> align="center">
		<font face="Tahoma, Arial, Helvetica" size=2 color="#<%=arColor(6)%>">
		<a href="webCal4_month.asp?date=<%=strFirst%>&staff_id=<%=intID%>"
		<%=showStatus("View all of " & MonthName(Month(strFirst)))%>>
		<b><%=MonthName(Month(strFirst)) & " " & Year(strFirst)%></b></a></font></td>
<% if arSpan(0) < intDays + 1 then %>
		<td bgcolor="#<%=arColor(2)%>" colspan=<%=arSpan(1)%> align="center">
		<font face="Tahoma, Arial, Helvetica" size=2 color="#<%=arColor(6)%>">
		<a href="webCal4_month.asp?date=<%=strLast%>&staff_id=<%=intID%>"
		<%=showStatus("View all of " & MonthName(Month(strLast)))%>>
		<b><%=MonthName(Month(strLast)) & " " & Year(strLast)%></b></a></font></td>
<% end if %>	
	
	<tr><%=strDays%>

<%

ReDim arSpan(intDays)
for x = 0 to intDays
	arSpan(x) = 1
next

' go through each time segment of the day

for intTime = intRange1 to intRange2

	response.write "<tr>" & VbCrLf
	
	if (intTime Mod intFactor = 0 OR intTime = 0) AND intTime <> intRange2 then

' insert an extra row to mark the hour break
	
'	response.write "<td colspan=8 height=1><img src=""./images/tiny_blank.gif"" height=1>" _
'		& "</td><tr>" & VbCrLf

		response.write "<td rowspan=" & intFactor & " align=""right"" bgcolor=""#" _
			& arColor(1) & """ height=" & intHeight * intFactor & ">" _
			& "<font face=""Tahoma, Arial, Helvetica"" size=2><nobr>"
			
		if intTime > 0 then
			intHour = intTime/intFactor
		else
			intHour = 0
		end if

		if intHour = 0 then
			response.write "<b>midnight</b>"
		elseif intHour < 12 then
			response.write intHour & ":00 AM"
		elseif intHour = 12 then
			response.write "<b>noon</b>"
		else
			response.write intHour - 12 & ":00 PM"
		end if
	
' alternate hour colors
	
		if intHour Mod 2 = 0 then
			strColor = "ffffff"
		else
			strColor = "dfdfdf"
		end if

		response.write "</nobr></font></td>" & VbCrLf
		
'	elseif intTime Mod intFactor = intFactor/2 then

' insert an extra row to mark the half-hour break

'		response.write "<td colspan=8 height=1><img src=""./images/tiny_blank.gif"" height=1>" _
'			& "</td><tr>" & VbCrLf
	end if
	
	for d = 0 to intDays
		if arSpan(d) = 1 then
			response.write "<td bgcolor=""#"
			
			if arEvents(d,intTime,0) <> "" then
				arSpan(d) = arEvents(d,intTime,1)
				
' this spans the additional rows inserted to mark hours and
' half hours
				
'				intExtra = round(arSpan(d) / 2 - 1)

' this calculates the number of time segments into the hour
' at which this event begins
				
'				intFraction = intTime - intHour * intFactor

' if the event begins right before an hour or half hour break
' then we need to add one to the extra spanning				
				
'				if intFraction < intFactor / 2 AND intFraction > 0 then
'					intExtra = intExtra + 1
'				end if
								
				response.write arEvents(d,intTime,2) _
					& """ rowspan=" & arSpan(d) & " align=""center"">" _
					& "<font face=""Tahoma, Arial, Helvetica"" size=1>" _
					& arEvents(d,intTime,0) & "</font>"
				
				intExtra = 0
			else

				response.write strColor & """ height=" & intHeight & ">" _
					& "<img src=""./images/tiny_blank.gif"">"
			end if
			response.write "</td>" & VbCrLf
		else
			arSpan(d) = arSpan(d) - 1
		end if
	next
next
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