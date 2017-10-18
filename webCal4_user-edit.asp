<!--#include file="data/webCal4_data.inc"-->
<!--#include file="webCal4_verify.inc"-->

<html>
<head>
<!--#include file="webCal4_themes.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/23/1999

dim rsUsers, rsEvents, strQuery, arTemp, strOptions, jsGroups
dim eventCount, showName, strChanged
dim nameFirst, nameLast, accessLevel, emailName, emailSite
dim login, password, editType

if Request.Form("delete") = "Delete" then
' ----------------------------------
' DELETION FORM
' ----------------------------------
' if the delete button was hit then display the deletion form
' get the info on the user to be deleted

	strQuery = "SELECT * FROM cal_users WHERE " _
		& "(user_id)=" & Request.Form("user_id")
	Set rsUsers = db.Execute(strQuery,,&H0001)
		nameFirst = rsUsers("name_first")
	rsUsers.Close
	Set rsUsers = nothing
	
	strQuery = "SELECT user_id FROM cal_events WHERE " _
		& "user_id=" & Request.Form("user_id")
	Set rsEvents = Server.CreateObject("ADODB.RecordSet")

' DSN was defined by data include
' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

	rsEvents.Open strQuery, DSN, 3, 1, &H0001
	eventCount = rsEvents.Recordcount
	rsEvents.Close
	Set rsEvents = nothing
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form action="webCal4_user-deleted.asp" method="post">
<tr>
	<td bgcolor="#<%=arColor(4)%>" colspan=3>
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>User Deletion</b></font></td>
<tr>
	<td colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	
<% if eventCount > 0 then %>
	
	<b>What should happen to the <%=eventCount%> events scheduled by <%=nameFirst%>?</b><br>
	<input type="radio" name="do" value="delete">erase them all<br>
	<input type="radio" name="do" value="some" checked>erase the private but transfer the public events<br>
	<input type="radio" name="do" value="move">transfer them all<br>
	<center>transfer to <select name="recipient">
<%
		strQuery = "SELECT user_id, name_first, name_last FROM cal_users" _
			& " WHERE (user_id)<>" & Request.Form("user_id") _
			& " ORDER BY name_last, name_first"
		Set rsUsers = db.Execute(strQuery,,&H0001)
		do while not rsUsers.EOF
			if rsUsers("name_last") <> "" then
				showName = rsUsers("name_last") & ", " & rsUsers("name_first")
			else
				showName = rsUsers("name_first")
			end if
			response.write "<option value=" & rs("user_id") _
				& ">" & showName & VbCrLf
			rsUsers.MoveNext
		loop
		rsUsers.Close
		Set rsUsers = nothing
%>
	</select></center>
	</font></td>
	
<% else %>

Are you sure you want to erase the user <%=nameFirst%>?

<% end if %>
	
<tr>
	<td colspan=3 align="center" bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="delete" value="Continue">
		<input type="submit" name="cancel" value="Cancel">
	</td>
<tr>
	<td align="center" colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased users and events cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="user_id" value="<%=Request.Form("user_id")%>">
<input type="hidden" name="event_count" value="<%=eventCount%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<%
else
' ----------------------------------
' EDIT FORM
' ----------------------------------
' if any button other than delete was hit, display the edit form

	if Request.Form("edit") = "Edit" then

' we're editing an existing event

		editType = "update"
		strQuery = "SELECT * FROM cal_users WHERE " _
			& "(user_id)=" & Request.Form("user_id")
		Set rsUsers = db.Execute(strQuery,,&H0001)
			nameFirst = rsUsers("name_first")
			nameLast = rsUsers("name_last")
			emailName = rsUsers("email_name")
			emailSite = rsUsers("email_site")
			login = rsUsers("login")
			password = rsUsers("password")
		rsUsers.Close
		Set rsUsers = nothing



'		arTemp = Split(Right(rsUsers("user_groups"), Len(rsUsers("user_groups"))-1), ",")
	else

' otherwise create a new event

		editType = "new"
		nameFirst = ""
		nameLast = ""
		accessLevel = "user"
		emailName = ""
		emailSite = ""
		login = ""
		password = ""
'		arTemp = Array(0000)
	end if

' get a list of all groups to generate JavaScript hash
' and form option list
	
	strQuery = "SELECT * FROM cal_groups G INNER JOIN " _
		& "cal_permissions P ON G.group_id = P.group_id " _
		& "WHERE P.user_id=" & Request.Form("user_id") _
		& " ORDER BY G.group_name"
	
	Set rsGroups = db.Execute(strQuery,,&H0001)
	strOptions = ""
	jsGroups = ""
	strChanged = ""

	do while not rsGroups.EOF
		strOptions = strOptions & "<option value=""" & rsGroups("group_id") _
			& """>" & rsGroups("group_name")

' the key is the group id and the value is the option
' box index, also the access level number
	
		jsGroups = jsGroups & "group[""" & rsGroups("group_id") & """] = "
	
		for x = 0 to UBound(arTemp)
			if CInt(Left(arTemp(x),1)) = rsGroups("group_id") then
				jsGroups = jsGroups & Mid(arTemp(x),2,1) & ";" & VbCrLf
				strChanged = strChanged & "," & x
				exit for
			elseif InStr(strChanged, "," & x) = 0 then
				jsGroups = jsGroups & "0;" & VbCrLf
			else
				jsGroups = jsGroups & "fail" & VbCrLf
			end if
		next
	
		rsGroups.MoveNext
	loop
	rsGroups.Close
	Set rsGroups = nothing
%>

<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.editform.name_first.value.length <= 0) {
		alert("You must enter a first name");
		document.editform.name_first.select();
		document.editform.name_first.focus();
		return false;
	}
	if (document.editform.login.value.length <= 0) {
		alert("You must enter a login name");
		document.editform.login.select();
		document.editform.login.focus();
		return false;
	}
	if (document.editform.password.value.length <= 0) {
		alert("You must enter a password");
		document.editform.password.select();
		document.editform.password.focus();
		return false;
	}
	if (document.editform.password.value != document.editform.confirm.value) {
		alert("The password values do not match");
		document.editform.confirm.select();
		document.editform.confirm.focus();
		return false;
	}
}

// generate hash of groups and access levels

group = new Object();
<%=jsGroups%>

// update the scope option list when the group selection is changed

function updateLevel() {
	var selGroup = document.editForm.group.options[document.editForm.group.selectedIndex].value;
	document.editForm.level.options[group[selGroup]].selected=true;
}

// save scope settings whenever it is changed

function saveLevel() {
	var selGroup = document.editForm.group.options[document.editForm.group.selectedIndex].value;
	var selShow = document.editForm.level.options[document.editForm.level.selectedIndex].value;
	group[selGroup] = selShow;
	var levels = ""
	for (var i in group) {
		if (group[i] != 0) {
			levels += "," + i + group[i] + "01";
		}
	}
	alert(levels);
	document.editForm.dbgroups.value = levels;
}	
	
//-->
</SCRIPT>
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="editForm" method="post" id="event" action="webCal4_user-updated.asp">
<tr bgcolor="#<%=arColor(4)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>User Details</b></font></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">First Name</font></td>
	<td><input type="text" name="name_first" value="<%=nameFirst%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Last Name</font></td>
	<td><input type="text" name="name_last" value="<%=nameLast%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>e-mail</td>
	<td><input type="text" name="email_name" value="<%=emailName%>" size=10>@<input type="text" name="email_site" value="<%=emailSite%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">Login Name</font></td>
	<td><input type="text" name="login" value="<%=login%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">Password<br>
	confirm</font></td>
	<td><input type="password" name="password" value="<%=password%>" size=15><br>
	<input type="password" name="confirm" value="<%=password%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Permissions</font></td>
	<td><select name="level" onChange="saveLevel();">
		<option value="0">no access to
		<option value="1">read
		<option value="2">add to
		<option value="3">edit
		<option value="4">manage
		<option value="5">admin
		</select>
	
		<select name="group" onChange="updateLevel();">
		<%=strOptions%>
		</select>
	</td>
<%
db.Close
Set db = nothing
%>

<tr>
	<td colspan=2 align="center" bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="save" value="Save" onClick="return Validate();">
		<input type="submit" name="saveadd" value="Save & Add Another" onClick="return Validate();">
      <input type="submit" name="cancel" value="Cancel">
	</td>

</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="dbgroups" value="0">
<input type="hidden" name="edit_type" value="<%=editType%>">
<input type="hidden" name="user_id" value="<%=Request.Form("user_id")%>">
<input type="hidden" name="url" value="<%=Request.Form("url")%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000"><b>Red fields are required</b></font>

<%	if jsGroups <> "" then %>

<script lang="javascript"><!--
// initialize group access levels
	
	var selGroup = document.editForm.group.options[document.editForm.group.selectedIndex].value;
	document.editForm.level.options[group[selGroup]].selected=true;

// -->
</script>

<%
	end if
' ----------------------------------
' END OF SEPERATE FORMS
' ----------------------------------
end if
%>

</center>
</body>
</html>