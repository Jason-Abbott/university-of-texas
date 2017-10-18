<font face="Tahoma, Arial, Helvatica" color="#<%=arColor(4)%>">
<b><font size=2><%=WeekdayName(WeekDay(strThisDay))%></font><br>
<font size=7><%=Day(strThisDay)%></font><br>
<font size=5><%=MonthName(Month(strThisDay),1)%></font></b><br>
<font size=4><%=Year(strThisDay)%></font>
</font>
<p>

<!-- mini calendar -->

<font face="Tahoma, Arial, Helvetica" size=1>Go to Day:

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" cellspacing=0 cellpadding=2 border=0><tr><td>
<!-- end framing table -->

<table border=0 cellspacing=1 cellpadding=0 bgcolor="#<%=arColor(4)%>">
<tr height=14>

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/06/1999

dim arDays(31), m, y, mNext, strFirst, strLast, strLastMonth, x

for x = 1 to 7
	response.write "<td width=""14.3%"" align=""center"">" _
		& "<font face=""Tahoma, Arial, Helvetica"" size=1 color=""#ffffff"">" _
		& "<b>" & Left(WeekDayName(x),1) & "</b></font></td>" & VbCrLf
next

response.write "<tr>" & VbCrLf

m = Month(strThisDay)
y = Year(strThisDay)

' calculate the numeric value of the next month

if m < 12 then
	mNext = m + 1
else
	mNext = 1
end if

' get the first column of the first day

strFirst = WeekDay(Dateserial(y, m, 1))

' get the total days of the month by subtracting one
' day from the last day of next month

strLast = Day(Dateserial(y, mNext, 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

strLastMonth = Day(Dateserial(y, m, 1) - 1)

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) WHERE (D.event_date " _
	& " BETWEEN " & strDelim & m & "/1/" & y & " 12:00:00 AM" & strDelim _
	& " AND " & strDelim & m & "/" & strLast & "/" & y _
	& " 11:59:59 PM" & strDelim _
	& ") AND E.in_" & strType & "=1"

Set db = Server.CreateObject("ADODB.Connection")
db.Open strDSN

Set rsEvents = db.Execute(strQuery,,&H0001)
do while not rsEvents.EOF
	arDays(Day(rsEvents("event_date"))) = True
	rsEvents.MoveNext
loop
rsEvents.Close
Set rsEvents = nothing

' the column variable keeps constant track of the
' current calendar column

intCol = 0

strNonDay = "<td align=""right""><font face=""Tahoma, Arial, Helvetica"" size=1 color=""#" _
		& arColor(11) & """>"

' cycle through all the days previous to the first
' day of the active month
		
for d = 1 to strFirst - 1
	response.write strNonDay & strLastMonth - strFirst + d + 1 & "</font></td>"
	intCol = intCol + 1
next

' now cycle through all the days of the current month

for d = 1 to strLast
	intCol = intCol + 1
	response.write "<td align=""right"" bgcolor=""#"
	if y & m & d = Year(now) & Month(now) & Day(now) then
		response.write arColor(8)
	elseif intCol = 1 OR intCol = 7 then
		response.write arColor(3)
	else
		response.write arColor(2)
	end if
	response.write """ ><font face=""Tahoma, Arial, Helvetica"" size=1>"
	
	if arDays(d) then
		response.write "<a href=""webCal4_" & strCal & "-day.asp?date=" & DateSerial(y,m,d) _
			& "&" & strQueryString & """ " _
			& showStatus("View " & DateSerial(y,m,d) & " details") & "><b>" _
			& d & "</b></a>"
	else
		response.write d
	end if
	
	response.write "</font></td>"
	
	if intCol = 7 AND d < strLast then
		response.write "<tr>"
		intCol = 0
	end if
next

' finally, cycle through as many days of the next month as
' necessary to fill the calendar grid through column 7

if intCol > 0 then
	d = 1
	do while intCol < 7
		response.write strNonDay & d & "</font></td>"
		d = d + 1
		intCol = intCol + 1
	loop
end if
%>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<a href="webCal4_<%=strCal%>-month.asp?date=<%=strThisDay%>&<%=strQueryString%>">Go to Month</a><br>
<a href="webCal4_<%=strCal%>-week.asp?date=<%=strThisDay%>&<%=strQueryString%>">Go to Week</a><br>