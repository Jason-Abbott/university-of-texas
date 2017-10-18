<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 04/23/1999

dim dayFirst, dayLast, dayLastMonth, m, y, mLoop, yLoop, e, d, column

' this is the form element index that should be
' dynamically updated with this mini calendar

 e = Request.QueryString("element")

' assign month and year

if Request.QueryString("date") <> "" then
	m = Month(Request.QueryString("date"))
	y = Year(Request.QueryString("date"))
else
	m = Month(now)
	y = Year(now)
end if
%>

<html>
<head>
<script lang="javascript">
<!--
// this is the script that actually updates the
// correct form element

function day_onclick(m,d,y) {
	window.opener.document.forms[0].elements[<%=e%>].value = m + "/" + d + "/" + y;
	self.close();
}

// this refreshes the display with the selected month or year
// keeping track of the element to target on the original page

function date_change() {
	var selMonth = document.calMini.Month.options[document.calMini.Month.selectedIndex].value + '';
	var selYear = document.calMini.Year.options[document.calMini.Year.selectedIndex].value + '';
	var date = selMonth + "/1/" + selYear;
	location.href = "webCal4_mini.asp?date=" + date + "&element=<%=e%>";
}

// -->
</script>

<style type="text/css">
	A:hover { color:#000066; }
</style>

<title><%=MonthName(m)%> &nbsp; &nbsp; &nbsp;</title>
</head>
<body bgcolor="#c0c0c0" link="#000044" vlink="#000044" alink="#000066" leftmargin=2 topmargin=2>

<center>
<table border=0 cellspacing=1 cellpadding=0>
<tr>
	<td colspan=12 align="center"><nobr>
	<form name="calMini">
	<select name="Month" onChange="date_change();">
<%
' this creates the form list of month names

for mLoop = 1 to 12
	response.write "<option value=" & mLoop
	if mLoop = m then response.write " selected"
	response.write ">" & MonthName(mLoop,1) & VbCrLf
next
%>
	</select><select name="Year" onChange="date_change();">
<%
' this creates the form list of 20 years

for yLoop = year(Now) - 10 to year(Now) + 10
	response.write "<option value=" & yLoop
	if yLoop = y then response.write " selected"
	response.write ">" & yLoop & VbCrLf
next
%>
	</select></nobr>
	</td>
</form>
<tr height=14>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>S</b></font></td>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>M</b></font></td>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>T</b></font></td>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>W</b></font></td>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>T</b></font></td>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>F</b></font></td>
	<td width="14.3%" align="center" bgcolor="#404040">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#ffffff"><b>S</b></font></td>
<tr>

<%

' calculate the numeric value of the next month

if m < 12 then
	mNext = m + 1
else
	mNext = 1
end if

' the column variable keeps constant track of the
' current calendar column

column = 0

' get the first column of the first day

dayFirst = WeekDay(Dateserial(y, m, 1))

' get the total days of the month by subtracting one
' day from the last day of next month

dayLast = Day(Dateserial(y, mNext, 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

dayLastMonth = Day(Dateserial(y, m, 1) - 1)

' cycle through all the days previous to the first
' day of the active month

for d = 1 to dayFirst - 1
	response.write "<td align=""right""><font face=""Tahoma, Arial, Helvetica"" size=1 color=""#777777"">" _
		& dayLastMonth - dayFirst + d + 1 & "</td>"
	column = column + 1
next

' now cycle through all the days of the current month

for d = 1 to dayLast
	column = column + 1
	response.write "<td align=""right"""
	if y & m & d = Year(now) & Month(now) & Day(now) then
		response.write " bgcolor=""#e0e0e0"""
	elseif column = 1 or column = 7 then
		response.write " bgcolor=""#999999"""
	end if
	response.write "><font face=""Tahoma, Arial, Helvetica"" size=1>" _
		& "<a href='#' name='day' onclick=""day_onclick(" _
		& m & "," & d & "," & y & ")"" " _
		& "onMouseOver=""window.opener.status='Set to " & m & "/" & d & "/" & y _
		& "'; return true;""	onMouseOut=""window.opener.status='';"">" _
		& d & "</a></font></td>"
	if column = 7 AND d < dayLast then
		response.write "<tr>"
		column = 0
	end if
next

' finally, cycle through as many days of the next month as
' necessary to fill the calendar grid through column 7

if column > 0 then
	d = 1
	do while column < 7
		response.write "<td align=""right""><font face=""Tahoma, Arial, Helvetica"" size=1 color=""#777777"">" _
			& d & "</font></td>"
		d = d + 1
		column = column + 1
	loop
end if


%>

</table>
</center>
</body>
</html>