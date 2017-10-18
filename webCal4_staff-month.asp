<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_rollovers.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_buttons.js"></SCRIPT>
<!--#include file="webCal4_data.inc"-->
<!--#include file="webCal4_define-month.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/05/1999

strCal = "staff"
strQueryString = "staff_id=" & intID

' find all events occuring between the first and
' last second of the selected month

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) WHERE (D.event_date " _
	& "BETWEEN " & strDelim _
	& m & "/1/" & y & " 12:00:00 AM" & strDelim & " AND " _
	& strDelim & m & "/" & strLast & "/" & y _
	& " 11:59:59 PM" & strDelim & ") " _
	& "AND E.in_staff=1 " _
	& "ORDER BY E.time_start"
	
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
			& "<a href=""webCal4_" & strCal & "-detail.asp?event_id=" & rsEvents("event_id") _
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

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<!--#include file="webCal4_buttons.inc"-->
<tr>
	<td bgcolor="#<%=arColor(6)%>" align="center" colspan=2>
<!--#include file="webCal4_layout-month.inc"-->
	</td>
</table>

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://boise.uidaho.edu/jason/webCal.html" target="_top">
webCal 4.0</a>
</font>

</body>
</html>