<!--#include file="data/webCal4_data.inc"-->
<!--#include file="webCal4_verify.inc"-->

<html>
<head>
<!--#include file="webCal4_themes.inc"-->

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/23/1999

dim rsGroups, rsEvents, strQuery, arTemp, strOptions, jsGroups
dim strCount, showName, strChanged
dim strName, strType

if Request.Form("delete") = "Delete" then
' ----------------------------------
' DELETION FORM
' ----------------------------------
' if the delete button was hit then display the deletion form
' get the info on the group to be deleted

	strQuery = "SELECT * FROM cal_groups WHERE " _
		& "(group_id)=" & Request.Form("group_id")
	Set rsGroups = db.Execute(strQuery,,&H0001)
		strName = rsGroups("group_name")
	rsGroups.Close
	Set rsGroups = nothing

' count the users who are members of this group
	
	strQuery = "SELECT user_id FROM cal_users WHERE " _
		& "user_groups LIKE '%," & Request.Form("group_id") & "%'"
	Set rsUsers = Server.CreateObject("ADODB.RecordSet")

' DSN was defined by data include
' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

	rsUsers.Open strQuery, DSN, 3, 1, &H0001
	strCount = rsUsers.Recordcount
	rsUsers.Close
	Set rsUsers = nothing
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
	<b>Group Deletion</b></font></td>
<tr>
	<td colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	
<% if strCount > 0 then %>
	
	<b>What should happen to the <%=strCount%> user(s) belonging to <%=strName%>?</b><br>
	<input type="radio" name="do" value="delete">erase them all<br>
	<input type="radio" name="do" value="some" checked>erase the private but transfer the public events<br>
	<input type="radio" name="do" value="move">transfer them all<br>
	<center>transfer to <select name="recipient">
<%
		strQuery = "SELECT group_id, group_name FROM cal_groups" _
			& " WHERE (group_id)<>" & Request.Form("group_id") _
			& " ORDER BY group_name"
		Set rsGroups = db.Execute(strQuery,,&H0001)
		do while not rsGroups.EOF
			response.write "<option value=" & rsGroups("group_id") _
				& ">" & rsGroups("group_name") & VbCrLf
			rsGroups.MoveNext
		loop
		rsGroups.Close
		Set rsGroups = nothing
%>
	</select></center>
	</font></td>
	
<% else %>

Are you sure you want to erase the group <%=strName%>?

<% end if %>
	
<tr>
	<td colspan=3 align="center" bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="delete" value="Continue">
		<input type="submit" name="cancel" value="Cancel">
	</td>
<tr>
	<td align="center" colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased groups cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="user_id" value="<%=Request.Form("group_id")%>">
<input type="hidden" name="event_count" value="<%=strCount%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<%
else
' ----------------------------------
' EDIT FORM
' ----------------------------------
' if any button other than delete was hit, display the edit form

	strQuery = "SELECT * FROM cal_users WHERE user_id<>1"
	Set rsUsers = db.Execute(strQuery,,&H0001)

	if Request.Form("edit") = "Edit" then

' get the information on the selected group

		strType = "update"
		strQuery = "SELECT * FROM cal_groups WHERE " _
			& "(group_id)=" & Request.Form("group_id")
		Set rsGroups = db.Execute(strQuery,,&H0001)
			strName = rsGroups("group_name")
			strDesc = rsGroups("group_description")
		rsGroups.Close
		Set rsGroups = nothing

' create a list of members for the current group

		strUsers = ","
		
		strQuery = "SELECT user_id, access_level FROM cal_permissions " _
			& "WHERE group_id=" & Request.Form("group_id")
		Set rsAccess = db.Execute(strQuery,,&H0001)
		do while not rsAccess.EOF
			strUsers = strUsers & rsAccess("user_id") & "-" _
				& rsAccess("access_level") & ","
			rsAccess.MoveNext
		loop
		rsAccess.Close
		Set rsAccess = nothing
		
' get a list of all users

		do while not rsUsers.EOF
			intPos = InStr(strUsers, "," & rsUsers("user_id") & "-")
			if intPos <> 0 then
				strMembers = strMembers & "<option value='" & rsUsers("user_id") _
					& "'>" & rsUsers("name_first") & VbCrLf
				jsUsers = jsUsers & "user[""" & rsUsers("user_id") & """] = " _
					& Mid(strUsers, intPos + Len(rsUsers("user_id")) + 2, 1) & ";" & VbCrLf
			else
				strNons = strNons & "<option value='" & rsUsers("user_id") _
					& "'>" & rsUsers("name_first") & VbCrLf
				jsUsers = jsUsers & "user[""" & rsUsers("user_id") & """] = 0;" & VbCrLf
			end if
			rsUsers.MoveNext
		loop
	else

' otherwise create a new event

		do while not rsUsers.EOF
			strNons = strNons & "<option value='" & rsUsers("user_id") _
				& "'>" & rsUsers("name_first") & VbCrLf
			jsUsers = jsUsers & "user[""" & rsUsers("user_id") & """] = 0;" & VbCrLf
'			jsOld = jsUsers & "old[""" & rsUsers("user_id") & """] = 0;" & VbCrLf
			rsUsers.MoveNext
		loop

		strMembers = ""
		strType = "new"
		strName = ""
		strDesc = ""
	end if
	rsUsers.Close
	Set rsUsers = nothing
%>

<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.editform.group_name.value.length <= 0) {
		alert("You must enter a group name");
		document.editform.group_name.select();
		document.editform.group_name.focus();
		return false;
	}
}

user = new Object();
old = new Object(user);
<%=jsUsers%>
//old = user;

function Move(dir) {
	if (dir=="remove") {
		var from = "members";
		var to = "nons";
	} else {
		var from = "nons";
		var to = "members";
	}

// retrieve current values
	
	var menuFrom = eval("document.editForm." + from);
	var menuTo = eval("document.editForm." + to);
	var sel = menuFrom.selectedIndex;
	var length = menuTo.length;
	
	if (sel == -1) {
		alert("You must select a user to " + dir);
		return false;
	}

	var userid = menuFrom.options[sel].value;
	var moved = new Option(menuFrom.options[sel].text, userid);
	
	menuTo.options[length] = moved;
	menuTo.options[length].selected = true;
	
	var x = 0;
	option = new Array(menuFrom.length - 1);

// create an array copy of the previous list minus the selection
	
	for (var i=0; i < menuFrom.length; i++) {
		if (menuFrom.options[i].value != menuFrom.options[sel].value) {
			option[x] = new Option(menuFrom.options[i].text, menuFrom.options[i].value);
			x =+ 1;
		}
	}
	
// erase the previous option list and generate a new one with our array

	menuFrom.length = 0;
	for (var i=0; i < option.length; i++) {
		menuFrom.options[i] = option[i];
	}

// set default permissions if adding new user, and update display
	
	if (dir=="add" && user[userid]==0) {
		user[userid] = 1;
		document.editForm.level.options[0].selected=true;
	}	
	return true;
}

// save the user's access level to this group

function saveLevel() {
	var selUser = document.editForm.members.options[document.editForm.members.selectedIndex].value;
	var selLevel = document.editForm.level.options[document.editForm.level.selectedIndex].value;
	if (selUser == -1 || selLevel == -1) {
		return false;
	}
	user[selUser] = selLevel;
	var levels = "";
	for (var i in user) {
		alert(i + " " + user[i] + " " + old[i]);
		if (user[i] != old[i]) {
			levels += "," + user[i] + "|" + i;
		}
	}
//	alert(levels);
//	document.editForm.foo.value = levels;
}	

// update the access level to match the selected member

function updateLevel() {
	var selUser = document.editForm.members.options[document.editForm.members.selectedIndex].value;
	document.editForm.level.options[user[selUser]-1].selected=true;
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
<tr>
	<td colspan=4 bgcolor="#<%=arColor(4)%>"><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Group Details</b></font></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">Name</font></td>
	<td colspan=3><input type="text" name="group_name" value="<%=strName%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Description</font></td>
	<td colspan=3><input type="text" name="group_description" value="<%=strDesc%>" size=25></td>
<tr>
	<td valign="top" align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Membership</font></td>
	
	<td valign="top" align="center" bgcolor="#<%=arColor(13)%>">
	<font face="Tahoma, Arial, Helvetica" size=1>Members</font><br>
	<select name="members" size="5" onChange="updateLevel();"><%=strMembers%></select></td>
	
	<td valign="center" align="center">
	<input type="button" name="add" value="&lt;-" onClick="Move('add');">
	<p>
	<input type="button" name="remove" value="-&gt;" onClick="Move('remove');">
	</td>
	
	<td width="30%" valign="top" align="center">
	<font face="Tahoma, Arial, Helvetica" size=1>Non-members</font><br>
	<select name="nons" size="5"><%=strNons%></select></td>
<tr>
	<td align="right" bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Access</font></td>
	<td align="center" bgcolor="#<%=arColor(13)%>">
		<select name="level" onChange="saveLevel();">
		<option value="1">read
		<option value="2">add to
		<option value="3">edit
		<option value="4">manage
		</select></td>
	<td></td><td></td>
	
<%
db.Close
Set db = nothing
%>

<tr>
	<td colspan=4 align="center" bgcolor="#<%=arColor(12)%>">
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
<input type="hidden" name="group_id" value="<%=Request.Form("group_id")%>">
<input type="hidden" name="url" value="<%=Request.Form("url")%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
<input type="hidden" name="foo" value="">
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