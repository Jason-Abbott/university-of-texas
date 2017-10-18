<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<script language="javascript"><!--
if (document.images) {
// add icon
	var iconFifteen = new Image();
	iconFifteen.src = "images/tiny_add_blue_light.gif";
	var iconFifteenOn = new Image();
	iconFifteenOn.src = "images/tiny_add_blue.gif";

	var iconThirty = new Image();
	iconThirty.src = "images/tiny_add_green_light.gif";
	var iconThirtyOn = new Image();
	iconThirtyOn.src = "images/tiny_add_green.gif";
	
	var iconMock = new Image();
	iconMock.src = "images/tiny_add_red_light.gif";
	var iconMockOn = new Image();
	iconMockOn.src = "images/tiny_add_red.gif";
}

//-->
</script>
<!--#include file="webCal4_rollovers.inc"-->
<!--#include file="webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/26/1999

dim strFirst, strLast, dayNames, d, intCol, strColor, intStaff
dim strShow, strQuery, intTime, strView, tfAdd, tfWeekends, strType
dim tfProspect, tfAlumni, strCal

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

' these are dummy values that will need to indicate whether
' the user is alumni or student

tfProspect = False
tfAlumni = False

strCal = "signup"
strType = Request.QueryString("type")
tfWeekends = False
intErval = 15
intHourStart = 8
intHourEnd = 17

if Request.Form("week") = "next" then
	strSelect = DateAdd("d", 7, Date)
	strWeek1 = ""
	strWeek2 = " checked"
else
	strSelect = Date
	strWeek1 = " checked"
	strWeek2 = ""
end if

%>
<!--#include file="webCal4_define-week.inc"-->
<!--#include file="webCal4_define-segments.inc"-->
<%

dim arEvents()
ReDim arEvents(intDays,intTotal,2)

' three values of the third dimension:
' 0 = description
' 1 = duration
' 2 = color

' the filter array defines the types of appointments
' allowed for each time segment

dim arRules()
ReDim arRules(intDays,intTotal, 2)

' three values of the third dimension:
' 0 = 15 min session t/f
' 1 = 30 min session t/f
' 2 = mock interview t/f

' if a staff member was selected then parse their rules

if Request("staff_id") <> "" then

	intStaff = Request.Form("staff_id")

' ---------------------------------------------------------
' if the student already has an appointment then allow no more
' ---------------------------------------------------------

' define the common parts of the queries we'll be using

	strCommon = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
		& "ON E.event_id = D.event_id) WHERE (D.event_date " _
		& "BETWEEN " & strDelim _
		& strFirst & " 12:00:00 AM" & strDelim & " AND " _
		& strDelim & strLast & " 11:59:59 PM" & strDelim & ") "
		
	strQuery = strCommon &  "AND E.student_id=" & Session("StudentID")
		
	Set rsEvents = Server.CreateObject("ADODB.RecordSet")
		
'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

	rsEvents.Open strQuery, strDSN, 3, 1, &H0001
	intCount = CInt(rsEvents.RecordCount)
	rsEvents.Close
	Set rsEvents = nothing

	if intCount > 0 then

' set every time segment event option to false

		for x = 0 to UBound(arRules,1)
			for y = 0 to UBound(arRules,2)
				for z = 0 to UBound(arRules,3)
					arRules(x,y,z) = False
				next
			next
		next
	else

' initialize array by setting every time segment to true

		for x = 0 to UBound(arRules,1)
			for y = 0 to UBound(arRules,2)
				for z = 0 to UBound(arRules,3)
					arRules(x,y,z) = True
				next
			next
		next
	
' ---------------------------------------------------------
' build an array of rules data for the selected week
' ---------------------------------------------------------

' find out if the counselor has set limits on mock interviews
	
		strQuery = "SELECT * FROM cal_staff WHERE " _
			& "staff_ID=" & intStaff
		Set rsLimit = Server.CreateObject("ADODB.RecordSet")
		rsLimit.Open strQuery, strDSN, 3, 1, &H0001
		if CInt(rsLimit.RecordCount) > 0 then
			intLimit = rsLimit("mock_limit")
		else
			intLimit = 0
		end if
		rsLimit.Close
		Set rsLimit = nothing
	
' if a mock interview limit has been set then we need to
' count the total mock interviews already scheduled
	
		tfMocks = True
		if intLimit > 0 then
			strQuery = strCommon _
				& "AND E.staff_id=" & intStaff & " AND E.event_type='mock'"
			Set rsMocks = Server.CreateObject("ADODB.RecordSet")
			rsMocks.Open strQuery, strDSN, 3, 1, &H0001
			intMocks = CInt(rsMocks.RecordCount)
			rsMocks.Close
			Set rsMocks = nothing
			if intMocks >= intLimit then

' set mocks to false for whole week

				for x = 0 to UBound(arRules,1)
					for y = 0 to UBound(arRules,2)
						arRules(x,y,2) = False
					next
				next
				tfMocks = False
			end if
		end if
			
' if mocks aren't limited then continue with parsing
' other rules
' make sure longer appointment types begin in time to finish
' before the end of the day
' (this presently assumes 15' time segments)
		
		for x = 0 to UBound(arRules,1)
			for y = intRange2 to intRange2 - 60/intErval + 2 step - 1
				arRules(x,y,2) = False
			next
	
			for y = intRange2 to intRange2 - 30/intErval + 2 step - 1
				arRules(x,y,1) = False
			next	
		next

' order the query by time so that more recent rules override
' earlier ones in the case of conflict

		strQuery = "SELECT * FROM (cal_rules R INNER JOIN cal_rule_dates D " _
			& "ON R.rule_id = D.rule_id) WHERE (D.rule_date " _
			& "BETWEEN " & strDelim _
			& strFirst & " 12:00:00 AM" & strDelim & " AND " _
			& strDelim & strLast & " 11:59:59 PM" & strDelim & ") " _
			& "AND R.staff_id=" & intStaff _
			& " ORDER BY R.time_start"

' now go through rules defined for today
	
		Set rsRules = Server.CreateObject("ADODB.RecordSet")
		rsRules.Open strQuery, strDSN, 3, 1, &H0001
		do while not rsRules.EOF
		
' define indexes into array
' weekday is 0 based so subtract start day

			d = WeekDay(rsRules("rule_date")) - intFirst
			intTime = segments(rsRules("time_start"))
		
' this counts the total number of segments spanned by the rule
		
			intCount = segments(rsRules("time_end")) - intTime

' don't bother with excluding further mocks if they're
' already precluded by the mock limitation
			
			if rsRules("no_mock") = 1 AND tfMocks then
				for x = 0 to intCount - 1
					arRules(d,intTime + x,2) = False
				next
			end if

' if the user is a prospect or an alumnus then
' all types of appointments may be restricted
		
			if rsRules("no_prospect") = 1 AND tfProspect then
				for x = 0 to intCount - 1
					arRules(d,intTime + x,0) = False
					arRules(d,intTime + x,1) = False
					arRules(d,intTime + x,2) = False
				next
			end if
		
			if rsRules("no_alumni") = 1 AND tfAlumni then
				for x = 0 to intCount - 1
					arRules(d,intTime + x,0) = False
					arRules(d,intTime + x,1) = False
					arRules(d,intTime + x,2) = False
				next
			end if
	
			rsRules.MoveNext
		loop
	
		rsRules.Close
		Set rsRules = nothing
	end if
	
' ---------------------------------------------------------
' build an array of event data for selected week
' ---------------------------------------------------------

	strQuery = strCommon & "AND E.staff_id=" & intStaff _
		& " ORDER BY E.time_start"

' put all matching events in an array indexed by day number

	Set rsEvents = Server.CreateObject("ADODB.RecordSet")
	rsEvents.Open strQuery, strDSN, 3, 1, &H0001
	do while not rsEvents.EOF

' define indexes into array
' Weekday is 0 based so subtract start day

		d = WeekDay(rsEvents("event_date")) - intFirst
		intTime = segments(rsEvents("time_start"))

' determine the level of detail to reveal
		
		if rsEvents("show_student") = 1 then
			arEvents(d,intTime,0) = "<font face=""Tahoma, Arial, Helvetica"" " _
				& "size=1>" & rsEvents("event_title") & "</font>"
			arEvents(d,intTime,2) = arColor(0)
		elseif rsEvents("student_id") = Session("StudentID") then
			arEvents(d,intTime,0) = "<font face=""Tahoma, Arial, Helvetica"" " _
				& "size=1>your appointment</font>"
			arEvents(d,intTime,2) = "ff9999"
		else	
			arEvents(d,intTime,0) = "<img src=""./images/tiny_blank.gif"">"
			arEvents(d,intTime,2) = arColor(6)
		end if
		
		arEvents(d,intTime,1) = segments(rsEvents("time_end")) - intTime
		
' now adjust the appointment filter
' first for the 30 minute appointments

		for x = 1 to 30/intErval - 1
			if intTime - x > 0 then
				arRules(d,intTime - x,1) = False
			end if
		next

' now for the 60 minute mock interviews
		
		for x = 1 to 60/intErval - 1
			if intTime - x > 0 then
				arRules(d,intTime - x,2) = False
			end if
		next
		

		rsEvents.movenext
	loop

	rsEvents.Close
	set rsEvents = nothing
	
else

' write a notice to the middle of the week view

	intTime = segments("10:00")
	d = 4 - intFirst
	arEvents(d,intTime,0) = "<center><font face=""Tahoma, Arial, Helvetica"" " _
		& "size=2><b>Select a<br>counselor</b><font size=2>" _
		& "<br>to check their schedule for free time</font></font></center>"
	arEvents(d,intTime,1) = segments("15:00") - intTime
	arEvents(d,intTime,2) = "ff9999"
	tfAdd = False
end if

'response.write arEvents(4,intTime,1)

' now generate the title to display at the top of the calendar

strQuery = "SELECT NAME_LAST, NAME_FIRST FROM tblStudents WHERE " _
	& "ID_NUMBER=" & Session("StudentID")

Set rsUser = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsUser.Open strQuery, strDSN, 0, 1, &H0001
	strTitle = "<font face=""Verdana, Arial, Helvetica"" color=""#" _
		& arColor(6) & """ size=4><b>" _
		& rsUser("NAME_FIRST") & " " & rsUser("NAME_LAST") & "</b></font>"
rsUser.Close
Set rsUser = nothing
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<!-- heading table -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<tr>
	<td><font face="Verdana, Arial, Helvetica" size=4><%=strTitle%></font></td>
	<form action="webCal4_<%=strCal%>-week.asp?type=<%=Request.QueryString("type")%>" method="post">
	<td bgcolor="#<%=arColor(6)%>">
	<!--#include file="webCal4_signup-options.inc"-->
	</td>
	</form>
<tr>
	<td bgcolor="#<%=arColor(6)%>" align="center" colspan=2>
	
<!-- body of calendar -->
	
	<table width="100%" border=0 cellspacing=1 cellpadding=0>
	<tr>
		<td rowspan=3 width="10%"></td>
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
		& strColor & """><table cellspacing=0 cellpadding=1 width=""100%""><tr>" _
		& "<td align=""center"" bgcolor=""#000000"" width=""20%"">" _
		& "<font face=""Verdana, Arial, Helvetica"" color=""#ffffff"" size=2><b>" _
		& Day(strShow) & "</b></font></td><td align=""center"">" _
		& "<font face=""Verdana, Arial, Helvetica"" size=1>" _
		& WeekDayName(intCol,0) & "</td></table></td>" & VbCrLf
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
		<b><%=MonthName(Month(strFirst)) & " " & Year(strFirst)%></b></font></td>
<% if arSpan(0) < intDays + 1 then %>
		<td bgcolor="#<%=arColor(2)%>" colspan=<%=arSpan(1)%> align="center">
		<font face="Tahoma, Arial, Helvetica" size=2 color="#<%=arColor(6)%>">
		<b><%=MonthName(Month(strLast)) & " " & Year(strLast)%></b></font></td>
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
	end if
	
' go through each day of the week
' 0 = description
' 1 = duration
' 2 = color
	
	for d = 0 to intDays
		if arSpan(d) = 1 then
		
' no event is currently being displayed on this day
		
			response.write "<td bgcolor=""#"
			
			if arEvents(d,intTime,0) <> "" then
			
' an event starts at this time segment

				arSpan(d) = arEvents(d,intTime,1)
		
				response.write arEvents(d,intTime,2) _
					& """ rowspan=" & arSpan(d) _
					& " align=""center"">" _
					& arEvents(d,intTime,0)
			else
			
' there are no events here so display correct color
			
				if intFirst <> 2 then
					if d > 0 AND d < 6 then
						response.write strColor
					else
						response.write arColor(12)
					end if
				else
					response.write strColor
				end if

				response.write """ height=" & intHeight & ">"

' display icon to add appointment if a counselor was selected
				
				strTime = TimeValue(intHour & ":" & (intTime Mod intFactor) * intErval)
				strDate = DateAdd("d", d, strFirst)

' 15 minute appointment icon
					
				if arRules(d,intTime,0) then
				
				response.write "<a href=""webCal4_" & strCal & "-edit.asp?" _
					& "time=" & strTime & "&date=" & strDate _
					& "&staff_id=" & intStaff _
					& "&type=" & strType & "&appt=15"" " _
					& switchIcon("Fifteen" & d & intTime, "Fifteen", "Add a 15 minute appointment on " & strDate & " at " & strTime) _
					& "><img name=""Fifteen" & d & intTime _
					& """ src=""./images/tiny_add_blue_light.gif"" width=10 height=5 border=0></a>"
				end if

' 30 minute appointment icon
					
				if arRules(d,intTime,1) then
					
				response.write "<a href=""webCal4_" & strCal & "-edit.asp?" _
					& "time=" & strTime & "&date=" & strDate _
					& "&staff_id=" & intStaff _
					& "&type=" & strType & "&appt=30"" " _
					& switchIcon("Thirty" & d & intTime, "Thirty", "Add a 30 minute appointment on " & strDate & " at " & strTime) _
					& "><img name=""Thirty" & d & intTime _
					& """ src=""./images/tiny_add_green_light.gif"" width=10 height=5 border=0></a>"
				end if
						
' mock interview icon

				if arRules(d,intTime,2) then

				response.write "<a href=""webCal4_" & strCal & "-edit.asp?" _
					& "time=" & strTime & "&date=" & strDate _
					& "&staff_id=" & intStaff _
					& "&type=" & strType & "&appt=mock"" " _
					& switchIcon("Mock" & d & intTime, "Mock", "Add a mock interview on " & strDate & " at " & strTime) _
					& "><img name=""Mock" & d & intTime _
					& """ src=""./images/tiny_add_red_light.gif"" width=10 height=5 border=0></a>"
				end if
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

<% if Request.QueryString("error") = "conflict" then %>
<center>
<font face="Verdana, Arial, Helvetica" size=2 color="#aa0000"><b>
Your appointment conflicts with an existing event<br></b>
Please select another time
</font>
</center>
<% end if %>

<!--#include file="webCal4_legend.inc"-->

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://boise.uidaho.edu/jason/webCal.html" target="_top">
webCal 4.0</a>
</font>

</body>
</html>