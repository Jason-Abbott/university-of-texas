<table width="100%" border=0 cellspacing=1 cellpadding=1>

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

' this is the width of each event column

intRatio = 90/(intCols + 1)

' track row spanning for each column

dim arSpan()
ReDim arSpan(intCols)
for x = 0 to intCols
	arSpan(x) = 1
next

' go through each time segment of the day

for intTime = intRange1 to intRange2
	response.write "<tr>" & VbCrLf

	if (intTime Mod intFactor = 0 OR intTime = 0) AND intTime <= intRange2 then
	
' we need to display the next hour label
	
		response.write "<td rowspan=" & intFactor & " align=""right"" bgcolor=""#" _
			& arColor(1) & """ height=" & intHeight * intFactor & " width=""15%"">" _
			& VbCrLf & "<font face=""Verdana, Arial, Helvetica"" size=2><nobr>"
			
		if intTime > 0 then
			intHour = intTime/intFactor
		else
			intHour = 0
		end if

		if intHour = 0 then
			response.write "<b>midnight</b>"
		elseif intHour < 12 then
			response.write intHour & ":00 AM"
		elseif intHour = 12 then
			response.write "<b>noon</b>"
		else
			response.write intHour - 12 & ":00 PM"
		end if

	' alternate hour colors
	
		if intHour Mod 2 = 0 then
			strColor = "ffffff"
		else
			strColor = "dfdfdf"
		end if
		
		response.write "</nobr></font></td>" & VbCrLf
	end if

' 0 = event id
' 1 = event title
' 2 = start time
' 3 = end time
' 4 = column
' 5 = remaining time segments

' now go through each column of events

	for x = 0 to intCols
		if arSpan(x) = 1 then
		
' no event is already being displayed in the current column
		
			intSpan = 1
			response.write "<td bgcolor=""#"
			
			if arEvents(x,intTime,0) <> "" then
			
' an event starts at this time; record the number of ROWS it spans
			
				arSpan(x) = arEvents(x,intTime,1)
				intTall = intHeight * arSpan(x)
				
				if x < intCols then
					for y = 0 to arSpan(x) - 1
					
' go through each time segment (row) spanned by this event
					
						for z = x + 1 to intCols
						
' go through each remaining event column

							if arSpan(z) = 1 AND arEvents(z,intTime + y,0) = "" then
								if intSpan > 1 AND intSpan > z - x then
									intSpan = z - x
								end if
							else
								exit for
							end if
						next
					next
				end if
				intRatio2 = intRatio * intSpan
			
				response.write arEvents(x,intTime,2) & """ rowspan=" & arSpan(x) _
					& " align=""center"" width=""" & intRatio2 & "%"" " _
					& "height=" & intTall & "colspan=" & intSpan & ">" _
					& "<font face=""Tahoma, Arial, Helvetica"" size=1>" _
					& arEvents(x,intTime,0) & "</font>"

				x = x + intSpan - 1
			else
			
' if there's no event in this time slot then display a blank
' figure out how many COLUMNS to span, if any
		
				intSpan = 1
				if x < intCols then	
					for y = x + 1 to intCols

' cycle through the remaining columns to find the blank ones

						if arSpan(y) = 1 AND arEvents(y,intTime,0) = "" then
							intSpan = intSpan + 1
						else
							exit for
						end if
					next
				end if
				
' jump ahead as many columns as we're spanning, and adjust the
' width% to include multiple columns
				
				x = x + intSpan - 1
				intRatio2 = intRatio * intSpan
			
				response.write strColor & """ width=""" & intRatio2 _
					& "%"" colspan=" & intSpan & " height=" & intHeight & ">" _
					& "<img src=""./images/tiny_blank.gif"">"
			end if
'			response.write intSpan
			response.write "</td>" & VbCrLf
		else
		
' as we move down through the time segements, decrement row span count
		
			arSpan(x) = arSpan(x) - 1
		end if
	next

' display one last, tiny column for clarity
	
	if intHour < 8 OR intHour > 17 OR (intHour < 13 AND intHour > 11) then
		strColor2 = arColor(11)
	else
		strColor2 = arColor(6)
	end if
	
	response.write "<td bgcolor=""#" & strColor2 _
		& """ width=""1%"" height=" & intHeight & "><img src=""./images/tiny_blank.gif""></td>"
next
%>

</table>