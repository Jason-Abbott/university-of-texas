<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

' if the cancel button was hit then return the user
' to the rule admin page

if Request.Form("cancel") = "No" then
	response.redirect "webCal4_rules-admin.asp"
end if

' otherwise delete rule

dim strQuery, db

Set db = Server.CreateObject("ADODB.Connection")
db.Open strDSN

' erase from cal_rules

strQuery = "DELETE FROM cal_rules WHERE (rule_id)=" _
	& Request.Form("rule_id")

db.Execute strQuery,,&H0001

' erase from cal_rule_dates

strQuery = "DELETE FROM cal_rule_dates" _
	& " WHERE rule_id=" & Request.Form("rule_id")
	
db.Execute strQuery,,&H0001
db.Close
Set db = nothing

' send the user back to rules admin

response.redirect "webCal4_rules-admin.asp"
%>