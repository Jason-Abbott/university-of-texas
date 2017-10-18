<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

' ---------------------------------------------------------
' build an array of event data for the selected day
' ---------------------------------------------------------

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) WHERE (D.event_date " _
	& "BETWEEN " & strDelim _
	& strThisDay & " 12:00:00 AM" & strDelim & " AND " _
	& strDelim & strThisDay & " 11:59:59 PM" & strDelim & ") " _
	& "AND E.in_" & strType & "=1 " _
	& "ORDER BY E.time_start"

Set rsEvents = Server.CreateObject("ADODB.RecordSet")

' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001
intCount = CInt(rsEvents.RecordCount) - 1

ReDim arTemp(intCount,4)

' 0 = description
' 1 = start time
' 2 = column
' 3 = time segments spanned
' 4 = color

dim arLast()
ReDim arLast(0)
arLast(0) = 0

' this keeps track of the last used time in each column

x = 0
do while not rsEvents.EOF

	strDescription = "From " _
		& FormatDateTime(rsEvents("time_start"),3) & " to " _
		& FormatDateTime(rsEvents("time_end"),3)
		
	arTemp(x,0) = "<a href=""webCal4_detail.asp?event_id=" _
		& rsEvents("event_id") & "&date=" & rsEvents("event_date") _
		& "&view=day"" " & showStatus(strDescription) & ">" _
		& rsEvents("event_title") & "</a>" & VbCrLf

	arTemp(x,1) = segments(rsEvents("time_start"))

	arTemp(x,2) = ""

	intEnd = segments(rsEvents("time_end"))
	
	arTemp(x,3) = intEnd - arTemp(x,1)

	Select Case rsEvents("event_type")
		Case "30"
			arTemp(x,4) = "99ff99"
		Case "15"
			arTemp(x,4) = "9999ff"
		Case "mock"
			arTemp(x,4) = "ff9999"
		Case else
			arTemp(x,4) = arColor(0)
	End Select
	
' go through each column to find the first fit
' update arLast to new end segment for that column
	
	for y = 0 to UBound(arLast)
		if arTemp(x,1) > arLast(y) then
			arTemp(x,2) = y
			arLast(y) = intEnd
			exit for
		end if
	next
	
	if arTemp(x,2) = "" then
	
' there was no room in any existing columns
' so add a new column and set end segment
	
		y = UBound(arLast) + 1
		arTemp(x,2) = y
		ReDim Preserve arLast(y)
		arLast(y) = intEnd
	end if

	x = x + 1
	rsEvents.MoveNext
loop
rsEvents.Close
set rsEvents = nothing

' get the total number (ZERO-BASED) of columns

intCols = UBound(arLast)

dim arEvents()
ReDim arEvents(intCols,intTotal,2)

' DIMENSIONS
' 0 = column
' 1 = time segment of day
' 2 = event information

for y = 0 to intCols

' go through each column

	for x = 0 to intCount
	
' go through each event
	
		if arTemp(x,2) = y then
		
' the event occurs in this column

			arEvents(y,arTemp(x,1),0) = arTemp(x,0)
			arEvents(y,arTemp(x,1),1) = arTemp(x,3)
			arEvents(y,arTemp(x,1),2) = arTemp(x,4)
		end  if
	next
next
%>