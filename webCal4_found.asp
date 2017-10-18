<!--#include file="show_status.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 06/04/1999

dim rs, query, events, dateList, k, combine
dim s, field, wordList, word, eventID, eventYear

' these are the variables used by the included webCal4_showrecur
' they can't be declared within the include because the include
' is called repeatedly by this script

dim monthList(12), yKey, x, years

combine = ""
query = ""

' this subroutine converts the field information into
' proper SQL

sub parse(field, s)
	query = query & combine & "("
	
' if it begins and ends with a quote
' use Mid to cut out what is between the quotes

	if Left(s,1) = """" and Right(s,1) = """" then
		query = query & field & " LIKE '%" _
			& Mid(s, 2, Len(s)-2) & "%'"

' or if it has a plus (+) in it

	elseif InStr(s,"+") > 0 then
		wordList = Split(s, "+")
		for word = 0 to UBound(wordList)
			query = query & field & " LIKE '%" & wordList(word) & "%'"
			if word < UBound(wordList) then
				query = query & " AND "
			end if
		next
		
' or if it starts with a minus (-)

	elseif Left(s,1) = "-" then
		query = query & field & " NOT LIKE '%" _
			& Right(s, Len(s)-1) & "%'"
	
' otherwise split on spaces
	
	else
		wordList = Split(s)
		for word = 0 to UBound(wordList)
			query = query & field & " LIKE '%" & wordList(word) & "%'"
			if word < UBound(wordList) then
				query = query & " OR "
			end if
		next
	end if
	
	query = query & ")"
	
end sub	

' now go through all the form elements to figure out which 
' ones have values to generate concise SQL

if Request.Form("title") <> "" then
	query = query & combine
	Call parse("event_title", Request.Form("title"))
	combine = " AND "
end if

if Request.Form("description") <> "" then
	query = query & combine
	Call parse("event_description", Request.Form("description"))
	combine = " AND "
end if

' if a start or end date was entered, but not both, then
' use the current date as the missing date

if Request.Form("date_start") <> "" OR Request.Form("date_end") <> "" then
	query = query & combine & "(event_date BETWEEN #"
	if Request.Form("date_start") <> "" then
		query = query & Request.Form("date_start")
	else
		query = query & Date
	end if
	query = query & "# AND #"
	if Request.Form("date_end") <> "" then
		query = query & Request.Form("date_end")
	else
		query = query & Date
	end if
	query = query & "#)"
	combine = " AND "
end if
%>
<!--#include file="data/webCal4_data.inc"-->
<%
' now build the full query and run it

if Session(dataName & "User") = "" then
 Session(dataName & "User") = 0
end if

query = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
	& " ON (E.event_id = D.event_id) " _
	& " WHERE " & query & " AND (user_id = " _
	& Session(dataName & "User") & " OR " _
	& "private = " & False & ") ORDER BY event_date"

' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

Set rs = db.Execute(query,,&H0001)
Set dates = CreateObject("Scripting.Dictionary")

' break out event dates by event id
' (how I long for Perl's arrays at this point)

' use associative array (dictionary object) to generate
' a unique list of events

dateCount = 0
do while not rs.EOF
	dateCount = dateCount + 1
	eventID = rs("event_id")
	if dates.Exists(eventID) then

' if the event is already in the array then we need to
' temporarily save the already entered event dates,
' erase the key and then add in the old dates plus the
' new one (save me Perl!)

		dateList = dates.Item(eventID)
		dates.Remove(eventID)
		dates.Add eventID, dateList & " " & rs("event_date")
		
' event info and dates are delimited so that we can later
' split them out into other arrays
		
	else
	
' otherwise we can just add a new key with the single date
	
 		dates.Add eventID, rs("event_title") & "|" _
			& rs("event_recur") & "|" & rs("event_date")
'		response.write dates.Item(eventID) & "<br>"
	end if
	rs.movenext
loop
rs.Close
db.Close
Set rs = nothing
Set db = nothing

' now, after all that trouble, let's see if we ended up with
' any matches; if no matches then send back to search form
' with notice

if dateCount = 0 then
	response.redirect "webCal4_find.asp?retry=1" _
		& "&view=" & Request.Form("view")
end if

' if we're still here then there must have been matches
' let's display them:
%>

<html>
<!--#include file="webCal4_themes.inc"-->
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<table border=0 cellpadding=1 cellspacing=0 width="60%">
<tr bgcolor="#<%=arColor(4)%>">
	<td><nobr><font face="Tahoma, Arial, Helvetica" size=4><b>
<%
response.write "Found " & dates.Count & " match"
if dates.Count > 1 or dates.Count < 1 then
	response.write "es"
end if
response.write "</b> covering " & dateCount & " date"
if dateCount > 1 then
	response.write "s"
end if
response.write "</font></nobr></td>"

' now go through every matched event and display
' its information

for each eventID in dates.Keys
	tempList = split(dates.Item(eventID), "|")
	eventTitle = tempList(0)
	eventRecur = tempList(1)
	dateList = split(tempList(2))
	Erase tempList

' this creates a zero based array, the first two items being
' the title and recur type, the rest being the event dates
	
	response.write "<tr><td bgcolor=""#" & arColor(2) & """>" _
	& "<font face=""Verdana, Arial, Helvetica"" size=2><b>" _
	& eventTitle & "</b></font></td>" & VbCrLf _
	& "<tr><form><td>"
	
' now that we have an associative array based on the
' event id we need to break out the years and then
' months to nicely display the event information,
' which is handled by the following include:
%>
<!--#include file="webCal4_showrecur.inc"-->
<%
	response.write "</td></form>" & VbCrLf & "<tr><td>&nbsp;</td>" & VbCrLf

' next event key (eKey) and set of years
next
dates.RemoveAll
%>

	</td>
</table>
</center>
</body>
</html>