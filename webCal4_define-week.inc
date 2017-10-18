<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/05/1999

' this include file calculates the variables needed
' to generate an HTML table for the selected week

strView = "week"

if tfWeekends then
	intFirst = 1
	intLast = 7
else
	intFirst = 2
	intLast = 6
end if

strFirst = DateAdd("d", intFirst - WeekDay(strSelect), strSelect)
strLast = DateAdd("d", intLast, strFirst)
strPrev = DateAdd("d", -intLast, strFirst)
strNext = DateAdd("d", intFirst, strLast)

intDays = intLast - intFirst
intRatio = Round(90/(intDays + 1), 2)

%>