<!--#include file="data/webCal4_data.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 09/08/1999

dim strQuery, rsStaff, strStaff, rsUser, strDate

' retrieve the counselor's name

strQuery = "SELECT First_Name, Last_Name FROM " _
	& "tblSTAFF WHERE pwid=" & Request.QueryString("staff_id")

Set rsStaff = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsStaff.Open strQuery, strDSN, 0, 1, &H0001
	strStaff = rsStaff("First_Name") & " " & rsStaff("Last_Name")
rsStaff.Close
Set rsStaff = nothing

' retrieve the user's name

strQuery = "SELECT NAME_LAST, NAME_FIRST FROM tblStudents WHERE " _
	& "ID_NUMBER=" & Session("StudentID")

Set rsUser = Server.CreateObject("ADODB.RecordSet")
rsUser.Open strQuery, strDSN, 0, 1, &H0001
	strTitle = rsUser("NAME_FIRST") & " " & rsUser("NAME_LAST")
rsUser.Close
Set rsUser = nothing

' normalize other values

strDate = Request.QuerySTring("time") & " " & Request.QueryString("date")
%>

<html>
<head>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_popup.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.apptForm.phone.value.length <= 0) {
		alert("You must enter your phone number");
		document.apptForm.phone.select();
		document.apptForm.phone.focus();
		return false;
	}
	if (document.apptForm.email.value.length <= 0) {
		alert("You must enter your e-mail address");
		document.apptForm.email.select();
		document.apptForm.email.focus();
		return false;
	}
	if (document.apptForm.email.value.length > 0) {
		var email = document.apptForm.email.value;
		if (email.indexOf("@") < 3) {
		   alert("The e-mail address you entered seems incorrect.  Please check it")
			document.apptForm.email.select();
			document.apptForm.email.focus();
			return false;
		}
	}
}

//-->
</SCRIPT>
<!--#include file="webCal4_themes.inc"-->
</head>
<body onload="init();" bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="apptForm" method="post" action="webCal4_student_updated.asp">

<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=2 align="center"><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Schedule appointment</b></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Name:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<%=strTitle%></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Phone:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<input type="text" name="phone"></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>e-mail:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<input type="text" name="email"></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Counselor:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<%=strStaff%></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Date:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<%=FormatDateTime(DateValue(strDate),1)%></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Time:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<%=TimeValue(strDate)%></font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Type:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>

<%
Select Case Request.QueryString("appt")
	Case "15"
		response.write "15 minute session"
	Case "30"
		response.write "30 minute session"
	Case "mock"
		response.write "mock interview"
End Select
%>
	<input type="hidden" name="type" value="<%=Request.QueryString("appt")%>">
	</font></td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="right" valign="top">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<b>Reason:</b></font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
		<textarea cols="25" name="reason" type="text" rows="5" wrap="virtual"></textarea>
	</td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" align="center" colspan=2>
		<input type="submit" name="save" value="Save" onClick="return Validate();">
      <input type="submit" name="cancel" value="Cancel">
	</td>
	
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="staff_id" value="<%=Request.QueryString("staff_id")%>">
<input type="hidden" name="grade" value="<%=Request.QueryString("type")%>">
<input type="hidden" name="student_id" value="<%=Session("StudentID")%>">
<input type="hidden" name="title" value="<%=strTitle%>">
<input type="hidden" name="date" value="<%=strDate%>">



</form>

</center>
</body>
</html>