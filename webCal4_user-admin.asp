<!--
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 07/06/1999
 -->
 
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="webCal4_verify.inc"-->

<html>
<head>
<script language="javascript"><!--
//preload images and text for faster operation

if (document.images) {
// back to calendar icon
	var iconMonth = new Image();
	iconMonth.src = "images/icon_calprev_grey.gif";
	var iconMonthOn = new Image();
	iconMonthOn.src = "images/icon_calprev.gif"
	statusMonth = "Return to calendar";
}

function iconOver(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+"On.src");
		status=eval("status"+name);
	}
}

function iconOut(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+".src");
		status="";
	}
}
//-->
</script>

<%
dim strQuery, rsUsers, strUserList, strGroupList

strQuery = "SELECT user_id, name_first, name_last FROM cal_users " _
	& "WHERE user_id <> 1 ORDER BY name_last, name_first"
Set rsUsers = db.Execute(strQuery,,&H0001)

if rsUsers.EOF = -1 then
	strUserList = ""
else
	do while not rsUsers.EOF
		if rsUsers("name_last") <> "" then
			strUserList = strUserList & rsUsers("name_last") & ", " & rsUsers("name_first") & VbCrLf
		else
			strUserList = strUserList & rsUsers("name_first") & VbCrLf
		end if
		strUserList = "<option value=" & rsUsers("user_id") _
			& ">" & strUserList & VbCrLf
		rsUsers.MoveNext
	loop
end if
rsUsers.Close
Set rsUsers = nothing

strQuery = "SELECT group_id, group_name FROM cal_groups " _
	& "ORDER BY group_name"
Set rsGroups = db.Execute(strQuery,,&H0001)

if rsGroups.EOF = -1 then
	strGroupList = ""
else
	do while not rsGroups.EOF
		strGroupList = strGroupList & "<option value=" & rsGroups("group_id") _
			& ">" & VbCrLf & rsGroups("group_name") & VbCrLf
		rsGroups.MoveNext
	loop
end if
rsGroups.Close
Set rsGroups = nothing
db.Close
Set db = nothing
%>

<!--#include file="webCal4_themes.inc"-->
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="adminform" action="webCal4_user-edit.asp?view=<%=Request.QueryString("view")%>" method="post">
<tr>
	<td bgcolor="#<%=color(4)%>" colspan=2>
	<a href="webCal4_<%=Request.QueryString("view")%>.asp"
	onMouseOver="iconOver('Month'); return true;" 
   onMouseOut="iconOut('Month'); return true;">
	<img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>User Management</b></font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="add" value="Add">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	a new user</font></td>

<% if strUserList <> "" then %>
	
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="edit" value="Edit">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	the selected user</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="delete" value="Delete">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	the selected user</font></td>
<tr>
	<td bgcolor="#<%=color(12)%>" align="right">
	<font face="Verdana, Arial, Helvetica" size=2>user:</font></td>
	<td bgcolor="#<%=color(12)%>">
	<select name="user_id">
	<%=strUserList%>
	</select>
	</td>

<% else %>

<tr>
	<td bgcolor="#<%=color(5)%>"><font size=1>&nbsp;</td>
	<td bgcolor="#<%=color(5)%>"><font size=1>&nbsp;</td>
	
<% end if %>

<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
	<input type="submit" name="add" value="Add">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	a new group</font></td>

<% if strGroupList <> "" then %>
	
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="edit" value="Edit">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	the selected group</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="delete" value="Delete">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	the selected group</font></td>
<tr>
	<td bgcolor="#<%=color(12)%>" align="right">
	<font face="Verdana, Arial, Helvetica" size=2>group:</font></td>
	<td bgcolor="#<%=color(12)%>">
	<select name="group_id">
	<%=strGroupList%>
	</select>
	</td>

<% end if %>

	
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

</center>
</body>
</html>