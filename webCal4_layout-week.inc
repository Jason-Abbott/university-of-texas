<table width="100%" border=0 cellspacing=1 cellpadding=0>

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/06/1999

' generate the heading

strShow = strFirst
dim arSpan()
ReDim arSpan(1)
m = 0
for intCol = 0 to 6
	if strShow = Date then
		strColor = arColor(8)
	elseif intCol = 0 or intCol = 6 then
		strColor = arColor(10)
	else
		strColor = arColor(9)
	end if
	
	strDays = strDays & "<td width=""14.28%"" bgcolor=""#" _
		& strColor & """>" & VbCrLf _
		& "<table cellspacing=0 cellpadding=1 border=0 width=""100%""><tr>" & VbCrLf _
		& "<td align=""center"" bgcolor=""#000000"" width=""20%"" valign=""top"">" & VbCrLf _
		& "<font face=""Verdana, Arial, Helvetica"" color=""#ffffff"" size=2><b>" _
		& Day(strShow) & "</b></font></td>" _
		& VbCrLf _
		& "<td align=""center"" rowspan=2>" _
		& "<font face=""Verdana, Arial, Helvetica"" size=1>"

' include option to add events if viewed by staff

	if strCal = "staff" then
		strDays = strDays _
			& "<a href=""webCal4_" & strCal & "-edit.asp?action=new&view=week&date=" & strShow & """ " _
			& showStatus("Add a new event to " & strShow) & ">" _
			& WeekDayName(intCol + 1,0) & "</a></td>" & VbCrLf & "<tr>"
	else
		strDays = strDays _
			& WeekDayName(intCol + 1,0) & "</td>" & VbCrLf & "<tr>"
	end if

' show link to day view if there are events on this day
		
	if arEvents(intCol) <> "" then
		strDays = strDays & "<td>" _
			& "<a href=""webCal4_" & strCal & "-day.asp?date=" & strShow _
			& "&" & strQueryString & """ " _
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
	<a href="webCal4_<%=strCal%>-month.asp?date=<%=strFirst%>&<%=strQueryString%>"
	<%=showStatus("View all of " & MonthName(Month(strFirst)))%>>
	<b><%=MonthName(Month(strFirst)) & " " & Year(strFirst)%></b></a></font></td>
<% if arSpan(0) < 7 then %>
	<td bgcolor="#<%=arColor(2)%>" colspan=<%=arSpan(1)%> align="center">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#<%=arColor(6)%>">
	<a href="webCal4_<%=strCal%>-month.asp?date=<%=strLast%>&<%=strQueryString%>"
	<%=showStatus("View all of " & MonthName(Month(strLast)))%>>
	<b><%=MonthName(Month(strLast)) & " " & Year(strLast)%></b></a></font></td>
<% end if %>	
	
<tr><%=strDays%>
<tr>

<% for intCol = 0 to 6 %>
<td height="200" bgcolor="#<%=arColor(9)%>" valign="top">
<font face="Verdana, Arial, Helvetica" size=1>
<%		if arEvents(intCol) <> "" then %>
<%=arEvents(intCol)%>
<%		else %>
&nbsp;
<%		end if %>
</font></td>
<% next %>

</table>