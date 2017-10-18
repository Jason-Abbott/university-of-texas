<html>
<head>
<!--#include file="webCal4_themes.inc"-->
<!--#include file="webCal4_rollovers.inc"-->
<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/11/1999

dim strThisDay, strQuery, strView, intErval
dim intTime, intMin, strType, strCal

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

strView = "day"
strCal = "staff"
strType = "staff"
intID = Session("StudentID")

intErval = 15
intHourStart = 5
intHourEnd = 23

if Request.QueryString("date") <> "" then
	strThisDay = Request.QueryString("date")
else
	strThisDay = Date
end if

%>
<!--#include file="webCal4_define-segments.inc"-->
<!--#include file="webCal4_define-day.inc"-->
	
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_buttons.js"></SCRIPT>
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<table width="100%" border=0 cellspacing=5 cellpadding=1>
<tr>
	<td width="90%" bgcolor="#<%=arColor(6)%>">
<!--#include file="webCal4_layout-day.inc"-->
	</td>

	<td valign="top" align="center">
<!--#include file="webCal4_day-nav.inc"-->
	</td>
</table>

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://boise.uidaho.edu/jason/webCal.html" target="_top">
webCal 4.0</a>
</font>

</body>
</html>