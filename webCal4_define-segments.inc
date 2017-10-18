<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/05/1999

' this include file calculates the variables needed
' to create an HTML table displaying the proper
' number of time segments each day

intFactor = 60/intErval

' these define the range of time to display
' the first number is the 24-hour time of day

intRange1 = intHourStart * intFactor
intRange2 = intHourEnd * intFactor - 1

' this function takes a time and converts it to the
' proper number of table segments based on the
' specified interval

intSegments = intFactor

function segments(strTime)
	intMin = Minute(strTime)
	intAdd = intErval/2
	for z = 0 to 60/intErval - 1
		if intMin < intAdd then
			intSegments = z
			exit for
		end if
		intAdd = intAdd + intErval
	next
	segments = intSegments + (Hour(strTime) * intFactor)
end function

' ZERO-BASED count of time segements/day

intTotal = (1440/intErval) - 1
intHeight = 24/intFactor - 1

%>