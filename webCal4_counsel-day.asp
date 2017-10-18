<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_rollovers.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_buttons.js"></SCRIPT>
<!--#include file="webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

dim dayShow, strQuery, intTime, strView
dim strThisDay, strCal

' ---------------------------------------------------------
' set parameters
' ---------------------------------------------------------

strCal = "counsel"
strView = "day"

' this defines the intervals in minutes
' must go evenly into 60

intErval = 5
intHourStart = 6
intHourEnd = 23
%>
<!--#include file="webCal4_define-segments.inc"-->
<%
' are we viewing someone else's calendar?

if Request("staff_id") <> "" then
	intID = Request("staff_id")
else
	intID = Session("StudentID")
end if

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

if Request.QueryString("date") <> "" then
	strThisDay = Request.QueryString("date")
else
	strThisDay = Date
end if

strPrev = DateAdd("d", -1, strThisDay)
strNext = DateAdd("d", 1, strThisDay)

' ---------------------------------------------------------
' build an array of event data for the selected day
' ---------------------------------------------------------

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) WHERE (D.event_date " _
	& "BETWEEN " & strDelim _
	& strThisDay & " 12:00:00 AM" & strDelim _
	& " AND " & strDelim & strThisDay _
	& " 11:59:59 PM" & strDelim & ") " _
	& "AND E.staff_id=" & intID _
	& " ORDER BY E.time_start"

dim arEvents()
ReDim arEvents(intTotal,2)
	
Set rsEvents = Server.CreateObject("ADODB.RecordSet")

' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001

do while not rsEvents.EOF

	intTime = segments(rsEvents("time_start"))

' assign appropriate color
	
	Select Case rsEvents("event_type")
		Case "30"
			arEvents(intTime,2) = "99ff99"
		Case "15"
			arEvents(intTime,2) = "9999ff"
		Case "mock"
			arEvents(intTime,2) = "ff9999"
		Case else
			arEvents(intTime,2) = arColor(0)
	End Select
	
	strDescription =  Replace(TimeValue(rsEvents("time_start")), ":00 ", " ") _
		& " to " & Replace(TimeValue(rsEvents("time_end")), ":00 ", " ")	
		
' determine the level of event detail to reveal
		
	if intID = Session("StudentID") then
		arEvents(intTime,0) = "<a href=""webCal4_" & strCal & "-detail.asp?event_id=" _
			& rsEvents("event_id") & "&date=" & rsEvents("event_date") _
			& "&view=day"" " & showStatus(strDescription) & ">" _
			& rsEvents("event_title") & "</a>" & VbCrLf
	elseif rsEvents("show_staff") = 1 then
		arEvents(intTime,0) = rsEvents("event_title")
	else	
		arEvents(intTime,0) = strDescription
		arEvents(intTime,2) = arColor(6)
	end if

	arEvents(intTime,1) = segments(rsEvents("time_end")) - intTime

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

<!-- heading table -->
<%=strTitle%><br>
<table width="100%" border=0 cellspacing=0 cellpadding=1>
<tr>
	<td width="90%" bgcolor="#<%=arColor(6)%>">
	
<!-- calendar body -->
	
	<table width="100%" border=0 cellspacing=1 cellpadding=0>
<%
' track the row spanning

intSpan = 1

' go through each time segment of the day

for intTime = intRange1 to intRange2

	response.write "<tr>" & VbCrLf

	if (intTime Mod intFactor = 0 OR intTime = 0) AND intTime <> intRange2 then
		response.write "<td rowspan=" & intFactor & " align=""right"" bgcolor=""#" _
			& arColor(1) & """ height=" & intHeight * intFactor & " width=""10%"">" _
			& "<font face=""Tahoma, Arial, Helvetica"" size=2><nobr>"
			
		if intTime > 0 then
			intHour = intTime/intFactor
		else
			intHour = 0
		end if

		if intHour = 0 then
			strHour = "<b>midnight</b>"
		elseif intHour < 12 then
			strHour = intHour & ":00 AM"
		elseif intHour = 12 then
			strHour = "<b>noon</b>"
		else
			strHour = intHour - 12 & ":00 PM"
		end if

		
' only allow calendar owners to see link for adding events
		
		if intID = Session("StudentID") then
			response.write "<a href=""webCal4_" & strCal & "-edit.asp?action=new&view=day&date=" _
				& strThisDay & """ " _
				& showStatus("Add a new event to " & strThisDay) & ">" _
				& strHour & "</a>"
		else
			response.write strHour
		end if
	
		response.write "</nobr></font></td>" & VbCrLf
	
' alternate hour colors
	
		if intHour Mod 2 = 0 then
			strColor = "ffffff"
		else
			strColor = "dfdfdf"
		end if
	end if
	
' 0 = event id
' 1 = event title
' 2 = start time
' 3 = end time
' 4 = column
' 5 = remaining time segments

	if intSpan = 1 then
		response.write "<td bgcolor=""#"
			
		if arEvents(intTime,0) <> "" then
			intSpan = arEvents(intTime,1)
				
			response.write arEvents(intTime,2) _
				& """ rowspan=" & intSpan & " align=""center"">" _
				& "<font face=""Tahoma, Arial, Helvetica"" size=1>" _
				& arEvents(intTime,0) & "</font>"
		else
			response.write strColor & """ height=" & intHeight & " width=""89%"">" _
				& "<img src=""./images/tiny_blank.gif"" height=1>"
		end if
		response.write "</td>" & VbCrLf
	else
		intSpan = intSpan - 1
	end if

	if intHour < 8 OR intHour > 16 OR (intHour < 13 AND intHour > 11) then
		strColor2 = arColor(1)
	else
		strColor2 = arColor(6)
	end if
	
	response.write "<td bgcolor=""#" & strColor2 & """ height=" _
		& intHeight & " width=""1%"">" _
		& "<img src=""./images/tiny_blank.gif"" height=1></td>" & VbCrLf
	
next
%>

	</table>
	</td>

<!-- end calendar body -->

	<td valign="top" align="center">
<!--#include file="webCal4_day-nav.inc"-->
	</td>
</table>

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://boise.uidaho.edu/jason/webCal.html" target="_top">
webCal 4.0</a>
</font>

</body>
</html>