<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/05/1999

' setup common values for the month view

dim strFirst, strLast, d, intCol, strDescription, intRow
dim arEvents(31), m, y
dim intIndex, strQuery, strView, strCal, strQueryString

intID = Session("StudentID")

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

' find the numeric value of the first day of the month
' ie Sunday = 1, Wednesday = 4

strFirst = WeekDay(Dateserial(y, m, 1))

' find the last day by subtracting 1 day from the first day
' of the next month

strLast = Day(Dateserial(y, Month(strNext), 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

strLastMonth = Day(Dateserial(y, m, 1) - 1)

%>