<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 06/04/1999

' this script creates specialized views of the different
' recurrence types

' it expects the following things:
' 1) the pre-existence of arDates
' 2) the pre-existence of intID, strRecur
' 3) the showStatus function already defined by container
' 4) these declarations should be in the container script:
'    dim arMonths(12), strYear, x, objYears, arDates


Set objYears = CreateObject("Scripting.Dictionary")

for x = 0 to UBound(arDates)
	strYear = Year(arDates(x))
	if objYears.Exists(strYear) then
		arYears = objYears.Item(strYear)
		objYears.Remove(strYear)
		objYears.Add strYear, arYears & " " & arDates(x)
			
' insert delimiter for subsequent splitting by month
			
	else
		objYears.Add strYear, arDates(x)
	end if
next

' now generate the text that describes the type of recurrence
' for each event and a view recurrence dates

select case strRecur
'------------------------------------------------	
	case "yearly"
'------------------------------------------------
' if there's yearly recurrence then create a list
' of those years

		response.write "<font face=""Verdana, Arial, Helvetica"" size=2>" _
			& "scheduled yearly on " _
			& MonthName(Month(arDates(0))) _
			& " " & Day(arDates(0)) _
			& " during these years:" _
			& "<select onChange=""location.href='" _
			& "webCal4_detail.asp?event_id=" _
			& intID & "&view=" & Request.Form("view") _
			& "&date=" & Month(arDates(0)) & "/" _
			& Day(arDates(0)) & "/' + " _
			& "options[selectedIndex].value;"">" & VbCrLf
		for each strYear in objYears.Keys
			response.write "<option value=" & strYear
			if strYear = Year(strDate) then
				response.write " selected"
			end if
			response.write ">" & strYear & VbCrLf
		next
		response.write "</select></font>" & VbCrLf
			
'------------------------------------------------
	case "monthly"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled for the " & Day(arDates(0))

' generate the correct wording for the day number
' (no detail overlooked)
				
		select case Right(Day(arDates(0)), 1)
			case 1
				response.write "st"
			case 2
				response.write "nd"
			case 3
				response.write "rd"
			case else
				response.write "th"
		end select
		response.write " of these months:<br>"

' go through each year
			
		for each strYear in objYears.Keys
			
' for each year create an array of scheduled months
				
			for each x in split(objYears.Item(strYear))
				arMonths(Month(x)) = MonthName(Month(x))
			next
				
' then display each item in the month array

			response.write "<table><tr><td>" & VbCrLf _
				& "<select onChange=""location.href=" _
				& "'webCal4_detail.asp?event_id=" _
				& intID & "&view=" & Request.Form("view") _
				& "&date=' + options[selectedIndex].value + " _
				& "'/" & Day(arDates(0)) & "/" & strYear & "';"">" & VbCrLf
			for each x in arMonths
				if x <> "" then
					response.write "<option value=" & x
					if x = MonthName(Month(strDate)) then
						response.write " selected"
					end if
					response.write ">" & x & VbCrLf
				end if
			next
			Erase arMonths
			response.write "</select></td>" & VbCrLf _
				& "<td><font face=""Verdana, Arial, Helvetica"" size=2>" _
				& strYear & "</font></td>" & VbCrLf & "</table>"
		next
		
'------------------------------------------------
	case "2weeks"
'------------------------------------------------
' the formatting for the three day based recurrence types
' (2weeks, weekly and daily) are handled together, below,
' so there is only basic formatting in this case statement

		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled for every other " _
			& WeekDayName(WeekDay(arDates(0))) _
			& " on these dates:<br>"
			
'------------------------------------------------
	case "weekly"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled for every " _
			& WeekDayName(WeekDay(arDates(0))) _
			& " on these dates:<br>"
			
'------------------------------------------------
	case "daily"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled daily for these dates:<br>"
			
'------------------------------------------------
	case "none"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=2>" _
			& "<a href=""webCal4_detail.asp?event_id=" _
			& intID & "&view=" & Request.Form("view") _
			& "&date=" & arDates(0) & """ " _
			& showStatus("View details") & ">" _
			& MonthName(Month(arDates(0))) & " " _
			& Day(arDates(0)) & ", " _
			& Year(arDates(0)) & "</a></font>"
end select

'------------------------------------------------
' now finish the formatting for the day based events
'------------------------------------------------
if strRecur = "daily" OR _
	strRecur = "weekly" OR _
	strRecur = "2weeks" then
' this is the biggie
' for every event, go through each year of occurence and break
' out dates by month

	for each strYear in objYears.Keys
		response.write "<table width=""100%""><tr>" & VbCrLf _
			& "<td valign=""top"" align=""center"">" _
			& "<font face=""Verdana, Arial, Helvetica"" size=2>" _
			& strYear & "</font></td>" & VbCrLf
		
' go through each event date in that year and create an array
' of days for each month

		for each x in split(objYears.Item(strYear))
			if arMonths(Month(x)) <> "" then
				strCombine = "<br>"
			else
				strCombine = ""
			end if
				
' now use that array to make a simple list of the days on
' which the event occurs for every month
				
			arMonths(Month(x)) = arMonths(Month(x)) & strCombine _
				& "<a href=""webCal4_detail.asp?event_id=" _
				& intID & "&view=" & Request.Form("view") _
				& "&date=" & Month(x) & "/" & Day(x) _
				& "/" & strYear & """ " & showStatus("View " _
					& MonthName(Month(x)) & " " & Day(x) & ", " _
					& strYear & " detail") _
				& ">" & Day(x) & "</a>"
		next

' create month headings and make them link to the calendar
' view for that month and year
		
		response.write "<td><table width=""100%"" " _
			& "cellpadding=2 cellspacing=0 border=0>" _
			& "<tr>" & VbCrLf
				
		for x = 1 to 12
			response.write "   <td align=""center"" " _
				& "width=""8.3%"" bgcolor=""#" & arColor(0) & """>" _
				& "<font face=""Tahoma, Arial, Helvetica"" size=1"
			if arMonths(x) <> "" then
				response.write "><a href=""webCal4_month.asp?date=" _
					& x & "/1/" & strYear & """ " _
					& showStatus("View " _
						& MonthName(x) & " " & strYear _
						& " calendar") _
					& ">" & MonthName(x,1) & "</a>"
			else
				response.write " color=""#" & arColor(5) & """>" & MonthName(x,1)
			end if
			response.write "</font></td>" & VbCrLf
		next
			
		response.write "<tr>" & VbCrLf
		for x = 1 to 12
			response.write "   <td valign=""top"" align=""right"">" _
				& "<font face=""Verdana, Arial, Helvetica"" size=1>" _
				& arMonths(x) & "</font></td>" & VbCrLf
		next
		response.write "</table></td>" & VbCrLf & "</table>"
		Erase arMonths
		Erase arDates
	next
end if
objYears.RemoveAll	
%>