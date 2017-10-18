<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/25/1999

dim strQuery

' if adding a weekly mock interview limit then

if Request.Form("limit") = "Limit" then
	if Request.Form("type") = "edit" then
		strQuery = "UPDATE cal_staff SET " _
			& "mock_limit = " & Request.Form("mock_limit") _
			& " WHERE (cal_staffID)=" & Request.Form("limit_id")
	else
		strQuery = "INSERT INTO cal_staff (" _
			& "staff_id, mock_limit" _
			& ") VALUES (" _
			& Session("StudentID") & "," _
			& Request.Form("mock_limit") & ")"
	end if

	Set db = Server.CreateObject("ADODB.Connection")
	db.Open strDSN
	db.Execute strQuery,,&H0001
	db.Close
	Set db = nothing
	
	response.redirect "webCal4_counsel-" & Request.QueryString("view") & ".asp"
end if

dim rsEvents, strView
dim strTitle, strDescription, strRecur, startDate, endDate
dim startHour, startMin, endHour, endMin, recurType, recurName
dim x, hourName, editType, noTime, showTime
dim arGroups, arTemp, arViews, strColor, strSkip

' define default values

recurType = Array("none","daily","weekly","2weeks","monthly","yearly")
recurName = Array("None","Daily","Weekly","Every other wk","Monthly","Yearly")

intStartHour = 8
intEndHour = 16

' use military time

startHour = 8
startMin = "00"
endHour = 16
endMin = "45"
showTime = ""
strSkip = ""

' determine the type of edit

if Request.Form("delete") = "Delete" then
		
' delete the event
		
	response.redirect "webCal4_rules-delete.asp?rule_id=" & Request.Form("rule_id")
elseif Request.Form("edit") = "Edit" then

' edit the event
		
	strQuery = "SELECT * FROM cal_rules R INNER JOIN cal_rule_dates D" _
		& " ON (R.rule_id = D.rule_id)" _
		& " WHERE (R.rule_id)=" & Request.Form("rule_id") _
		& " ORDER BY D.rule_date"

	Set rsRules = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

	rsRules.Open strQuery, strDSN, 3, 1, &H0001

	strName = rsRules("rule_name")
		
' these need to be broken out for separate form fields
		
	startHour = Hour(rsRules("time_start"))
	startMin = Minute(rsRules("time_start"))
	endHour = Hour(rsRules("time_end"))
	endMin = Minute(rsRules("time_end"))
	
	if rsRules("skip_weekends") = 1 then
		strSkip = " checked"
	else
		strSkip = ""
	end if
		
	if rsRules("no_mock") = 1 then
		strMock = " checked"
	else
		strMock = ""
	end if
		
	if rsRules("no_prospect") = 1 then
		strProspect = " checked"
	else
		strProspect = ""
	end if

	if rsRules("no_alumni") = 1 then
		strAlumni = " checked"
	else
		strAlumni = ""
	end if
			
	strStartDate = rsRules("rule_date")
	strRecur = rsRules("rule_recur")
			
	if strRecur <> "none" then
		rsRules.MoveLast
		strEndDate = DateValue(rsRules("rule_date"))
	else
		strEndDate = ""
	end if
	
	intID = rsRules("rule_id")

	rsRules.Close
	Set rsRules = nothing

else

' ----------------------------------
' adding a new event
' ----------------------------------

	strName = ""
	strRecur = "none"
	strStartDate = Date
	strEndDate = ""
	strType = "new"
		
end if

' now include the JavaScript for the popup calendar
' and populate the edit form with values
%>

<html>
<head>

<!--#include file="webCal4_themes.inc"-->

<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal4_popup.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.ruleForm.name.value.length <= 0) {
		alert("You must give your rule a name");
		document.ruleForm.name.select();
		document.ruleForm.name.focus();
		return false;
	}

// we need to make sure that something was checked here
//	alert(document.ruleForm.mock.value);
	
}

function updateEnd() {
//	if (document.ruleForm.end_date.value == "") {
		var r = document.ruleForm.rule_recur.options[document.ruleForm.rule_recur.selectedIndex].value;
		var d = document.ruleForm.start_date.value;
		var day = d.split("/")[1];
		var month = d.split("/")[0];
		var year = d.split("/")[2];

		if (r == "none") {
			d = "";
			document.ruleForm.skip.disabled=1;
			document.ruleForm.end_date.disabled=1;
		} else {
			document.ruleForm.skip.disabled=0;
			document.ruleForm.end_date.disabled=0;
		}
		
		if (r == "monthly") {
			if (month != 12) {
				month = month - 1 + 2;
			} else {
				month = 1;
				year = year - 1 + 2;
			}
			d = month + "/" + day + "/" + year;
		}
		if (r == "yearly") {
			year = year - 1 + 2;
			d = month + "/" + day + "/" + year;
		}
		document.ruleForm.end_date.value = d;
//	}
}

//-->
</SCRIPT>
</head>
<body onload="init();" bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="ruleForm" method="post" id="event" action="webCal4_rules-updated.asp?view=<%=Request.QueryString("view")%>">

<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Rule Parameters</b></font></td>
<tr>
	<td valign="top"><b><font face="Tahoma, Arial, Helvetica" color="#<%=arColor(14)%>" size=3>Name</font></b><br>
		<input name="name" type="text" size="35" max="50" value="<%=strName%>">
	</td>
<tr>
	<td valign="top">

<!-- timing table -->

	<table cellpadding=2 cellspacing=2 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=arColor(12)%>"><font face="Tahoma, Arial, Helvetica">
			<font color="#<%=arColor(14)%>" size=3><b>Date</b></font>
			<br>
			<input name="start_date" id="date" type="text" size="10" value="<%=strStartDate%>"><font size=2><input type="button" value="&gt;" onClick="calpopup(1);">
			<br>
			Recurrence<br>
<%
' generate the recurrence options
' select the option that matches the current event

	response.write "<select name=""rule_recur"" " _
		& "onChange=""updateEnd();"">"
	for x = 0 to UBound(recurType)
		response.write("<option value=""" & recurType(x) & """")
		if recurType(x) = strRecur then
			response.write(" selected")
		end if
		response.write(">" & recurName(x) & VbCrLf)
	next
%>
			</select><br>
			until</font><br>
			<input name="end_date" id="recurend" type="text" size="10" value="<%=strEndDate%>"><font size=2><input type="button" value="&gt;" onClick="calpopup(4);"></font>
		</font>
		</td>
		
		<td valign="top" bgcolor="#<%=arColor(12)%>"><font face="Tahoma, Arial, Helvetica">
			<font color="#<%=arColor(14)%>" size=3><b>Time</b></font>
			<font size=2>
			<br>
			<nobr>
			<select name="start_hour"<%=showTime%>>
<%
' generate the hours form list and select the
' one assigned above

	for x = intStartHour to intEndHour
		response.write "<option value=" & x
		if x = startHour then
			response.write(" selected")
		end if
		response.write ">"
		
		if x < 12 then
			response.write x & " AM"
		else
			response.write x - 12 & " PM"
		end if
		
		response.write VbCrLf
	next
%>	
			</select>
			<select name="start_min"<%=showTime%>>
<%
' generate the minutes form list and select the
' one assigned above

	for x = 0 to 55 step 5
		if x < 10 then
			x = "0" & x
		end if
		response.write("<option value=""" & x & """")
		if x = startMin then
			response.write(" selected")
		end if
		response.write(">:" & x & VbCrLf)
	next
%>
			</select>
			</nobr>
			<br>until<br>
		
			<nobr>
			<select name="end_hour"<%=showTime%>>
<%
' hours list

	for x = intStartHour to intEndHour
		response.write "<option value=" & x
		if x = endHour then
			response.write(" selected")
		end if
		response.write ">"
		
		if x < 12 then
			response.write x & " AM"
		else
			response.write x - 12 & " PM"
		end if
		
		response.write VbCrLf
	next
%>
			</select>
			<select name="end_min"<%=showTime%>>
<%
' minutes list

	for x = 0 to 55 step 5
		if x < 10 then
			x = "0" & x
		end if
		response.write("<option value=""" & x & """")
		if x = endMin then
			response.write(" selected")
		end if
		response.write(">:" & x & VbCrLf)
	next
%>
			</select>
			</nobr>
			<p>
			<input type="checkbox" name="skip"<%=strSkip%>><font size=2>Skip weekends</font></font>
		</td>

<!-- end timing table -->

	<tr>
		<td bgcolor="#<%=arColor(12)%>" colspan=2>
		<font face="Tahoma, Arial, Helvetica" color="#<%=arColor(14)%>">
		<b>Exclusions</b></font>

		<table width="100%">
		<tr>
			<td valign="top"><font face="Tahoma, Arial, Helvetica" size=2>
			<input type="checkbox" name="mock"<%=strMock%>>mock interviews<br>
			<!-- <input type="checkbox" name="thirty"<%=strStudent%>>30 minute sessions -->
			</font></td>

			<td valign="top"><font face="Tahoma, Arial, Helvetica" size=2>
			<input type="checkbox" name="prospects"<%=strProspects%>>Prospects<br>
			<input type="checkbox" name="alumni"<%=strAlumni%>>Alumni
			</font></td>
		</table>
		</td>

		
	</table>
	</td>
<tr>
	<td colspan=2 align="center">
		<input type="submit" name="save" value="Save" onClick="return Validate();">
		<input type="submit" name="saveadd" value="Save & Add Another" onClick="return Validate();">
      <input type="submit" name="cancel" value="Cancel">
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="rule_id" value="<%=intID%>">
<input type="hidden" name="edit_type" value="<%=strType%>">
</form>

<script lang="javascript"><!--
// focus on the title form element and initialize weekend checkbox

var r = document.ruleForm.rule_recur.options[document.ruleForm.rule_recur.selectedIndex].value;
if (r == "none") {
	document.ruleForm.skip.disabled=1;
}	
// document.ruleForm.name.focus();

// -->
</script>

<font face="Tahoma, Arial, Helvetica" size=2><b>Note:</b> the system does not prevent you from creating conflicting sets of rules.  It will process your rules chronologically.  New rules will override old rules if there is a conflict.
</font>

</center>
</body>
</html>