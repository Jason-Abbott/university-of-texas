<html>
<head>
<script language="javascript"><!--
//preload images and text for faster operation

if (document.images) {
// back to calendar icon
	var iconMonth = new Image();
	iconMonth.src = "images/icon_calprev_grey.gif";
	var iconMonthOn = new Image();
	iconMonthOn.src = "images/icon_calprev.gif"
}

//-->
</script>
<!--#include file="webCal4_rollovers.inc"-->
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/26/1999

dim rs, url, strDescription, strRecur, strTitle, intID
dim strDate, strStart, strEnd

' these are the variables used by the included webCal4_showrecur:

dim arMonths(12), strYear, x, objYears

' pull the event information and dates from db

strQuery = "SELECT * FROM cal_events E INNER JOIN tblStudents S " _
	& "ON E.student_id = S.ID_NUMBER " _
	& "WHERE (event_id)=" & Request.QueryString("event_id")

Set rsEvent = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsEvent.Open strQuery, strDSN, 3, 1, &H0001
	strDescription = rsEvent("event_description")
	strRecur = rsEvent("event_recur")
	strTitle = rsEvent("event_title")
	intID = rsEvent("event_id")
	strDate = Request.QueryString("date")
	strStart = TimeValue(rsEvent("time_start"))
	strEnd = TimeValue(rsEvent("time_end"))
	intUser = rsEvent("staff_id")
	strType = rsEvent("event_type")
	intStudent = rsEvent("student_id")
	
	if intStudent <> 0 then
		strPhone = rsEvent("CADDRPHONE")
		strEmail = rsEvent("EMAIL_")
	end if
rsEvent.Close
Set rsEvent = nothing
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<br>
<center>
<table border=0 cellspacing=0 cellpadding=3>
<tr>
	<td rowspan=2 align="center" valign="top">
		<font face="Tahoma, Arial, Helvatica" color="#<%=arColor(4)%>">
		<b><font size=2><%=WeekdayName(WeekDay(strDate))%></font><br>
		<font size=7><%=Day(strDate)%></font><br>
		<font size=5><%=MonthName(Month(strDate),1)%></font></b><br>
		<font size=4><%=Year(strDate)%></font>
		</font>
	</td>

	<td valign="top">
	<table cellspacing=0 cellpadding=3 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=arColor(4)%>">
			<font face="Verdana, Arial, Helvetica" size=4><b>
			<a href="webCal4_<%=Request.QueryString("view")%>.asp?date=<%=eventDate%>"
			<%=switchIcon("Month","","Return to Calendar")%>><img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>

			<%=strTitle%></b></font>&nbsp;
			
		</td>
	<tr>
		<td><font face="Tahoma, Arial, Helvetica" size=2>
		
<%
if strDescription <> "" then
	response.write Replace(strDescription, VbCrLf, "<br>") & "<p>"
end if

if intStudent <> 0 then
	response.write "<b>Contact</b><br>" & strPhone & "<br>" _
		& "<a href=""mailto:" & strEmail & """>" & strEmail & "</a>"
end if
%>
		</font>
		</td>
	</table>

<%
' if this is a recurring event then get list of all
' dates on which it occurs

if strRecur <> "none" then
	dim intCount, arDates()

	strQuery = "SELECT * FROM cal_dates" _
		& " WHERE event_id=" & intID _
		& " ORDER BY event_date"

	Set rsDates = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

	rsDates.Open strQuery, strDSN, 0, 1, &H0001

' generate array of event dates

	intCount = 0
	do while not rsDates.EOF
		ReDim preserve arDates(intCount)
		arDates(intCount) = rsDates("event_date")
		intCount = intCount + 1
		rsDates.MoveNext
	loop
	rsDates.Close
	Set rsDates = nothing
	
' if the event recurs then display the dates on
' which it occurs, invoking the include file
' that does the special formatting
%>
	<table bgcolor="#<%=arColor(2)%>" width="100%"><form><tr><td>
	<!--#include file="webCal4_showrecur.inc"-->
	</td></form></table>
<%
end if
response.write "</td>" & VbCrLf

'-----------------------------------
' display time range if one was entered for this event
'-----------------------------------

if strStart <> "" then
	dim intHrStart, intHrEnd, intSpan, strHrCurrent, strHrColor, strTxtColor

' the Hour function formats to military time
	
	intHrStart = Hour(strStart)
	intHrEnd = Hour(strEnd)

	response.write "<td rowspan=2 valign=""top"">" _
		& "<table cellspacing=1 cellpadding=0 border=0>"

' calculate the hours spanned by the event

	intSpan = (strHrEnd - strHrStart) + 1

	for h = 0 to 23
		if h = intHrStart then
			strHrCurrent = "<b>" & Replace(strStart, ":00 ", " ") & "</b>"
		elseif h = intHrEnd then
			strHrCurrent = "<b>" & Replace(strEnd, ":00 ", " ") & "</b>"
		else

' otherwise insert the array value with regular clock notation
' appended, changing 12PM to noon for the temporally challenged
			
			if h = 0 then
				strHrCurrent = "<b>midnight</b>"
			elseif h < 12 then
				strHrCurrent = h & ":00 AM"
			elseif h = 12 then
				strHrCurrent = "<b>noon</b>"
			else
				strHrCurrent = h - 12 & ":00 PM"
			end if
		end if

' make the hours covered by the event a different color
		
		if h >= intHrStart AND h <= intHrEnd then
			strHrColor = "ffffff"
			strTxtColor = "000000"
		else
			strHrColor = arColor(2)
			strTxtColor = arColor(5)
		end if

		response.write "<tr><td bgcolor=""#" & strHrColor _
			& """ align=""right"" nowrap><font face=""Tahoma, Arial, Helvetica""" _
			& "size=1 color=""#" & strTxtColor & """>" _
			& strHrCurrent & "</font></td>"
	next
	response.write "</td></table>"
end if

' from here display the management buttons
%>

<!-- buttons -->

<% if strType = "counselor" then %>

<tr>
	<td valign="bottom">
	
	<!-- framing table -->
	<table bgcolor="#<%=arColor(5)%>" width="100%" cellspacing=0 cellpadding=2 border=0><tr><td>
	<!-- end framing table -->
	
	<table cellspacing=0 cellpadding=2 border=0 width="100%">
	<tr>
		<td align="right" bgcolor="#<%=arColor(12)%>">
		<form action="webCal4_edit.asp?action=form" method="post">
		<input type="submit" name="edit" value="Edit"><br>
		<input type="submit" name="delete" value="Delete"></td>

<% 
' NOTE that this excludes events which might be listed
' as recurring but now occur on just one date (count=1)

	if strRecur <> "none" AND intCount > 1 then
%>

		<td bgcolor="#<%=arColor(12)%>"><font face="Tahoma, Arial, Helvetica" size=2>
		<input type="radio" name="scope" value="one">only this occurence<br>
		<input type="radio" name="scope" value="future">this and all future occurences<br>
		<input type="radio" name="scope" value="all" checked>all <%=count%> occurences
		</font>
		
<%		end if %>

		</td>

	</table>

	<!-- framing table -->
	</td></table>
	<!-- end framing table -->
	
	<input type="hidden" name="event_id" value="<%=intID%>">
	<input type="hidden" name="date" value="<%=strDate%>">
	<input type="hidden" name="count" value="<%=intCount%>">
	<input type="hidden" name="url" value="<%=url%>">
	<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
	</form>

<% end if %>

	</td>
</table>

</center>
</body>
</html>