<table width="100%" border=0 cellspacing=1 cellpadding=1>
<tr>
	<td bgcolor="#<%=arColor(2)%>" colspan=8 align="center">
	<font face="Verdana, Arial, Helvetica" size=2 color="#<%=arColor(6)%>">
	<b><%=MonthName(m) & " " & y%></b></font></td>
<tr>

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/06/1999

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

' only allow staff to see link for adding events
	
	if strCal = "staff" then
		response.write ">" & strFont & VbCrLf _
			& VbCrLf & "<a href=""webCal4_" & strCal & "-edit.asp?action=new&" _
			& "date=" & Dateserial(y, m, d) & "&view=month"" " _
			& VbCrLf _
			& showStatus("Add a new event to " & DateSerial(y, m, d)) _
			& ">" & d & "</a></b></font>"
	else
		response.write ">" & strFont & d & "</b></font>" & VbCrLf
	end if
		
' if the day contains events then generate link to day detail
		
	if arEvents(d) <> "" then
		response.write " <a href=""webCal4_" & strCal & "-day.asp?date=" _
			& Dateserial(y, m, d) & "&" & strQueryString & """ " _
			& switchIcon("Day" & d, "Day", "View " & Dateserial(y, m, d) & " detail") _
			& "><img name=""Day" & d & """ src=""./images/day_grey.gif"" border=0></a>"
	end if

	response.write "<br>" & VbCrLf & "<font face=""Arial, Helvetica"" size=1>" _
		& arEvents(d) & "</font></td>" & VbCrLf
	if intCol = 7 AND d <= strLast then
	
' if we're at the last column then generate link to week view
	
		response.write "<td valign=""center""><a href=""webCal4_" & strCal _
			& "-week.asp?date=" _
			& DateSerial(y, m, d) & "&" & strQueryString & """ " _
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
	response.write "<td valign=""center""><a href=""webCal4_" & strCal _
		& "-week.asp?date=" & strNext & "&" & strQueryString & """ " _
		& switchIcon("Week", "Week", "View week " & intRow) _
		& "><img name=""Week"" src=""./images/week_grey.gif"" border=0></a></td>"
end if
%>

</table>
