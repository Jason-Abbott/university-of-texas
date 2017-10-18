<!--#include file="webCal4_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/25/1999

dim rsEvents, strQuery, strView
dim strTitle, strDescription, strRecur, startDate, endDate
dim startHour, startMin, endHour, endMin, recurType, recurName
dim x, hourName, editType, noTime, showTime
dim arGroups, arTemp, arViews, strColor, strSkip, strCal

strCal = "counsel"

' define default values

recurType = Array("none","daily","weekly","2weeks","monthly","yearly")
recurName = Array("None","Daily","Weekly","Every other wk","Monthly","Yearly")
hourName = Array("12 AM","1 AM","2 AM","3 AM","4 AM","5 AM","6 AM","7 AM","8 AM","9 AM","10 AM","11 AM","12 PM","1 PM","2 PM","3 PM","4 PM","5 PM","6 PM","7 PM","8 PM","9 PM","10 PM","11 PM")

' use military time

startHour = 8
startMin = "00"
endHour = 17
endMin = "00"
showTime = ""
strSkip = ""
strColor = "black"

' determine the type of edit

Select Case Request.QueryString("action")
	Case "form"

' ----------------------------------
' dealing with an event from the detail form
' ----------------------------------

		if Request.Form("delete") = "Delete" then
		
' delete the event
		
			response.redirect "webCal4_delete.asp?event_id=" & Request.Form("event_id") _
				& "&type=" & Request.Form("type") & "&view=" & Request.Form("view") _
				& "&scope=" & Request.Form("scope") & "&date=" & Request.Form("date")
		elseif Request.Form("edit") = "Edit" then

' edit the event
		
			strQuery = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
				& " ON (E.event_id = D.event_id)" _
				& " WHERE (E.event_id)=" & Request.Form("event_id") _
				& " ORDER BY D.event_date"

			Set rsEvents = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

			rsEvents.Open strQuery, strDSN, 3, 1, &H0001

			strTitle = rsEvents("event_title")
			strDescription = rsEvents("event_description")
		
' these need to be broken out for separate form fields
		
			startHour = Hour(rsEvents("time_start"))
			startMin = Minute(rsEvents("time_start"))
			endHour = Hour(rsEvents("time_end"))
			endMin = Minute(rsEvents("time_end"))
	
			if rsEvents("skip_weekends") = 1 then
				strSkip = " checked"
			else
				strSkip = ""
			end if
		
			if rsEvents("show_student") = 1 then
				strStudent = " checked"
			else
				strStudent = ""
			end if
		
			if rsEvents("show_staff") = 1 then
				strStaff = " checked"
			else
				strStaff = ""
			end if

			Select Case Request.Form("scope")
				Case "future"
					strRecur = rsEvents("event_recur")
					startDate = Request.Form("date")
					rsEvents.MoveLast
					endDate = DateValue(rsEvents("event_date"))
				Case "all"
					strRecur = rsEvents("event_recur")
					startDate = DateValue(rsEvents("event_date"))
					rsEvents.MoveLast
					endDate = DateValue(rsEvents("event_date"))
				Case else
					strRecur = "none"
					startDate = Request.Form("date")
					endDate = ""
					skipWE = ""
			End Select
		
' if no scope was sent then we're editing an event that
' doesn't recur, in which case we want to edit "all"
' instances
		
			if Request.Form("scope") <> "" then
				strType = Request.Form("scope")
			else
				strType = "all"
			end if
		
			strView = Request.Form("view")
			rsEvents.Close
			Set rsEvents = nothing
		end if

	Case	"new"

' ----------------------------------
' adding a new event
' ----------------------------------

		strTitle = ""
		strDescription = ""
		strRecur = "none"
		startDate = Request.QueryString("date")
		endDate = ""
		strType = "new"
		
	Case "conflict"
		
' ----------------------------------
' rescheduling an event that conflicted
' ----------------------------------

		strTitle = Request.Form("title")
		strDescription = Request.Form("description")
		strRecur = "none"
		startDate = Request.Form("start_date")
		endDate = Request.Form("end_date")
		startHour = CInt(Request.Form("start_hour"))
		startMin = Request.Form("start_min")
		endHour = CInt(Request.Form("end_hour"))
		endMin = Request.Form("end_min")
		strRecur = Request.Form("event_recur")
		strStaff = Request.Form("staff")
		strStudent = Request.Form("student")
		strSkip = Request.Form("skip")
		
		strType = "new"

End Select

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
	if (document.editForm.title.value.length <= 0) {
		alert("You must enter a title for the event");
		document.editForm.title.select();
		document.editForm.title.focus();
		return false;
	}
}

function updateEnd() {
//	if (document.editForm.end_date.value == "") {
		var r = document.editForm.event_recur.options[document.editForm.event_recur.selectedIndex].value;
		var d = document.editForm.start_date.value;
		var day = d.split("/")[1];
		var month = d.split("/")[0];
		var year = d.split("/")[2];

		if (r == "none") {
			d = "";
			document.editForm.skip.disabled=1;
		} else {
			document.editForm.skip.disabled=0;
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
		document.editForm.end_date.value = d;
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
<form name="editForm" method="post" id="event" action="webCal4_<%=strCal%>-updated.asp?view=<%=Request.QueryString("view")%>">

<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Event Details</b></font></td>
<tr>
	<td valign="top"><b><font face="Tahoma, Arial, Helvetica" color="#<%=arColor(14)%>" size=3>Title</font></b><br>
		<input name="title" id="title" type="text" size="35" max="50" value="<%=strTitle%>">
	</td>
	<td rowspan=2 width=256 valign="top"><font face="Tahoma, Arial, Helvetica" color="#<%=arColor(14)%>"><b>Description</b></font><br>
		<textarea cols="24" name="description" type="text" rows="12" wrap="virtual"><%=strDescription%></textarea>
	</td>
<tr>
	<td valign="top">

<!-- timing table -->

	<table cellpadding=2 cellspacing=2 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=arColor(12)%>"><font face="Tahoma, Arial, Helvetica">
			<font color="#<%=arColor(14)%>" size=3><b>Date</b></font>
			<br>
			<input name="start_date" id="date" type="text" size="10" value="<%=startDate%>"><font size=2><input type="button" value="&gt;" onClick="calpopup(2);">
			<br>
			Recurrence<br>
<%
' generate the recurrence options
' select the option that matches the current event

	response.write "<select name=""event_recur"" " _
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
			<input name="end_date" id="recurend" type="text" size="10" value="<%=endDate%>"><font size=2><input type="button" value="&gt;" onClick="calpopup(5);"></font>
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

	for x = 0 to 23
		response.write("<option value=" & x)
		if x = startHour then
			response.write(" selected")
		end if
		response.write(">" & hourName(x) & VbCrLf)
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

	for x = 0 to 23
		response.write("<option value=" & x)
		if x = endHour then
			response.write(" selected")
		end if
		response.write(">" & hourName(x) & VbCrLf)
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
		<td bgcolor="#<%=arColor(12)%>" colspan=2 align="center">
		<font face="Tahoma, Arial, Helvetica" size=2>
		Reveal details to
		<input type="checkbox" name="staff"<%=strStaff%>>staff
		<input type="checkbox" name="student"<%=strStudent%>>students
		</font></td>
		
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

<input type="hidden" name="edit_type" value="<%=strType%>">
<input type="hidden" name="event_id" value="<%=Request.Form("event_id")%>">
<input type="hidden" name="url" value="<%=Request.Form("url")%>">
</form>

<script lang="javascript"><!--
// focus on the title form element and initialize weekend checkbox

var r = document.editForm.event_recur.options[document.editForm.event_recur.selectedIndex].value;
if (r == "none") {
	document.editForm.skip.disabled=1;
}	
document.editForm.title.focus();

// -->
</script>
</center>
</body>
</html>