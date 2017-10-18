<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' updated 10/25/1999

' if cancel is hit then send back to rule admin page

if Request.Form("cancel") = "Cancel" then
	response.redirect "webCal4_rules-admin.asp"
end if

' otherwise begin populating variables

dim strStart, strEnd, strDate, arDates(), intCount, intHide
dim strTitle, strDescription, intID, strQuery, queryEvent
dim strViews, arTemp, arViews, db

'------------------------------------
' normalize some values
'------------------------------------

strDate = DateValue(Request.Form("start_date"))

intID = Request.Form("rule_id")

strStart = TimeValue(Request.Form("start_hour") & ":" _
	& Request.Form("start_min"))
strEnd = TimeValue(Request.Form("end_hour") & ":" _
	& Request.Form("end_min"))

strName = replace(Request.Form("name"), "'", "''")

if Request.Form("mock") = "on" then
	intMock = 1
else
	intMock = 0
end if

if Request.Form("alumni") = "on" then
	intAlumni = 1
else
	intAlumni = 0
end if

if Request.Form("prospects") = "on" then
	intProspect = 1
else
	intProspect = 0
end if

'------------------------------------
' update tables
'------------------------------------

Set db = Server.CreateObject("ADODB.Connection")
db.Open strDSN

if Request.Form("edit_type") <> "new" then

' erase old dates from cal_rule_dates

	strQuery = "DELETE FROM cal_rule_dates" _
		& " WHERE rule_id=" & intID _
		& strQuery	
	db.Execute strQuery,,&H0001
	
' upate rule information in cal_rules

	strQuery = "UPDATE cal_rules SET " _
		& "rule_name = '" & strName & "', " _
		& "rule_recur = '" & Request.Form("rule_recur") & "', " _
		& "time_start = '" & strStart & "', " _
		& "time_end = '" & strEnd & "', " _
		& "no_mock = " & intMock & ", " _
		& "no_alumni = " & intAlumni & ", " _
		& "no_prospect = " & intProspect & " " _
		& "WHERE (rule_id)=" & intID

	db.Execute strQuery,,&H0001

else 

' add a new rule

	strQuery = "INSERT INTO cal_rules (" _
		& "rule_name, staff_id, rule_recur, time_start, time_end, " _
		& "no_mock, no_alumni, no_prospect" _
		& ") VALUES ('" _
		& strName & "', " _
		& Session("StudentID") & ", '" _
		& Request.Form("rule_recur") & "', '" _
		& strStart & "', '" _
		& strEnd & "', " _
		& intMock & ", " _
		& intAlumni & ", " _
		& intProspect & ")"

	db.Execute strQuery,,&H0001
	
' rule dates are keyed to rule info by the rule ID,
' so find out what auto-id was assigned
	
	strQuery = "SELECT * FROM cal_rules " _
		& "WHERE rule_name='" & strName & "' AND " _
		& "staff_id=" & Session("StudentID") & " AND " _
		& "time_start='" & strStart & "'"
	Set rsRule = db.Execute(strQuery,,&H0001)
	intID = rsRule("rule_id")
	rsRule.Close
	Set rsRule = nothing
end if

'------------------------------------
' update cal_rule_dates as needed
'------------------------------------
' this generates an array of dates for which the rule should apply

intCount = 0
if	Request.Form("rule_recur") <> "none" then
	Select Case Request.Form("rule_recur")
		Case "daily"
			addType = "d"
			addNum = 1
		Case "weekly"
			addType = "d"
			addNum = 7
		Case "2weeks"
			addType = "d"
			addNum = 14
		Case "monthly"
			addType = "m"
			addNum = 1
		Case "yearly"
			addType = "yyyy"
			addNum = 1
	end Select		

' populate the array with dates, according to the above
' addition, until the end date for the event

	While DateDiff("d", strDate, Request.Form("end_date")) >= 0
		if Request.Form("skip") <> "on" _
			OR (WeekDay(strDate) > 1 AND WeekDay(strDate) < 7) then

			ReDim Preserve arDates(intCount)
			arDates(intCount) = strDate
			intCount = intCount + 1
		end if
		strDate = DateAdd(addType, addNum, strDate)
	Wend

' if there was no recurrence selected then put the single
' date into the array

else
	ReDim Preserve arDates(intCount)
	arDates(intCount) = strDate
end if

' now go through everything inserted into the dates array
' and insert it into the event dates table

for each d in arDates
	strQuery = "INSERT INTO cal_rule_dates (" _
		& "rule_id, rule_date) VALUES ('" _
		& intID & "', '" & d & "')"
	db.Execute strQuery,,&H0001
next

db.Close
Set db = nothing

' with the data updated send user back to calendar
' or to the edit page again, if requested

if Request.Form("save") = "Save" then
	response.redirect "webCal4_rules-admin.asp"
elseif Request.Form("saveadd") = "Save & Add Another" then
	response.redirect "webCal4_rules-edit.asp?view=" _
		& Request.QueryString("view")
end if
%>