<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/19/1999

dim dayShow, strQuery, intTime, strView
dim strThisDay

if Request.QueryString("date") <> "" then
	strThisDay = Request.QueryString("date")
else
	strThisDay = Date
end if

intID = Session("StudentID")
%>

<html>
<head>
<title>Daily Detail</title>
</head>
<body bgcolor="#FFFFFF">

<table border="1" cellspacing="1" cellpadding="2" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#000000">
<tr>
	<td colspan=5 align="center" bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica"><b>
	<%=FormatDateTime(strThisDay,1)%></b></font></td>
<tr>
	<td align="center" bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica" size=2><b>
	Time</b></font></td>
	<td align="center" bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica" size=2><b>
	Student</b></font></td>
	<td align="center" bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica" size=2><b>
	Contact</b></font></td>
	<td align="center" bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica" size=2><b>
	Reason</b></font></td>
	<td align="center" bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica" size=2><b>
	Resume</b></font></td>

<!--#include file="webCal4_data.inc"-->
<%
' ---------------------------------------------------------
' build an array of event data for the selected day
' ---------------------------------------------------------

strQuery = "SELECT * FROM (cal_events E INNER JOIN cal_dates D " _
	& "ON E.event_id = D.event_id) INNER JOIN tblStudents S " _
	& "ON E.student_id = S.ID_NUMBER WHERE (D.event_date " _
	& "BETWEEN " & strDelim _
	& strThisDay & " 12:00:00 AM" & strDelim _
	& " AND " & strDelim & strThisDay _
	& " 11:59:59 PM" & strDelim & ") " _
	& "AND (E.staff_id=" & intID _
	& " AND E.student_id<>0)" _
	& " ORDER BY E.time_start"

'response.write strQuery
	
Set rsEvents = Server.CreateObject("ADODB.RecordSet")

' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

rsEvents.Open strQuery, strDSN, 3, 1, &H0001

do while not rsEvents.EOF
	Select Case rsEvents("event_type")
		Case "15"
			strType = "15 minute session"
		Case "30"
			strType = "30 minute session"
		Case "mock"
			strType = "Mock interview"
	End Select

%>

<tr>
	<td bgcolor="#c0c0c0">
	<font face="Tahoma, Arial, Helvetica" size=2><b>
	<%=TimeValue(rsEvents("time_start"))%></b></font></td>
	<td><%=rsEvents("NAME_FIRST")%> <%=rsEvents("NAME_LAST")%><br>
	<%=rsEvents("student_id")%></td>
	
	<td><%=rsEvents("CADDRPHONE")%><br>
	<%=rsEvents("EMAIL_")%></td>
	
	<td><b><%=strType%></b>:<br><%=Replace(rsEvents("event_description"), VbCrLf, "<br>")%></td>
	
	<td><%=rsEvents("student_id")%></td>

<%
	rsEvents.MoveNext
loop
rsEvents.Close
Set rsEvents = nothing
%>
</table>
<font face="Verdana, Arial, Helvetica" size=1>webCal 4.0</font>

</body>
</html>