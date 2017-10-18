<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 07/15/1999
 -->
 
<script language="javascript">
<!--
// go through each division and hide children

/*
function init() {
	var allDIVs = document.all.tags("DIV");
	for (i=0; i<allDIVs.length; i++) {
		if (allDIVs(i).className == "child") {
			allDIVs(i).style.display = "none";
		}
	}
}
*/

min = new Image(); min.src = "./images/icon_min.gif";
max = new Image(); max.src = "./images/icon_max.gif";
minClick = new Image(); minClick.src = "./images/icon_min_click.gif";
maxClick = new Image(); maxClick.src = "./images/icon_max_click.gif";

// change style visibility and image

function minIt() {
	options.style.display = "none";
	document.min_img.src=minClick.src;
	document.max_img.src=max.src;
   window.event.cancelBubble = true;
}

function maxIt() {
	options.style.display = "block";
	document.max_img.src=maxClick.src;
	document.min_img.src=min.src;
   window.event.cancelBubble = true;
}

//onload = init;

//-->
</script>

<form action="webCal4_prep.asp?date=<%=DateSerial(y,m,1)%>" method="post">

<!-- framing table -->
<table bgcolor="#<%=color(5)%>" border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><font face="Verdana, Arial, Helvetica" size=2 color="#ffffff"><b>Options</b></font></td>
	<td align="right"><a href="#" onClick="minIt(); return true;" <%=ShowStatus("hide options")%>><img name="min_img" src="images/icon_min.gif" width=16 height=14 alt="" border="0"></a><a href="#" onClick="maxIt(); return true;" <%=ShowStatus("display options")%>><img name="max_img" src="images/icon_max_click.gif" width=16 height=14 alt="" border="0"></a></td>
<tr>
	<td colspan=2>
<!-- end framing table -->

<div id="options" class="child">
	
<table bgcolor="#<%=color(12)%>">
<tr>
	<td>
	<select name="groups" multiple size=3>
<%
' display the list of groups available to the current user

	if Session(dataName & "Groups") <> "" then
		arGroups = Session(dataName & "Groups")
		response.write "bad"
	else
	
' if no user has logged in then, for now, just display the two
' available groups at maximum permissions (5)

' the student session variable will later be used to determine
' which groups should show up in this list

		dim arGroups(2,1)
	
		arGroups(1,0) = 1
		arGroups(1,1) = 5
		arGroups(2,0) = 2
		arGroups(2,1) = 5
		response.write "good"
	end if

	for x = 1 to UBound(arGroups)
		if arGroups(x,1) >= 2 then
			strGroups = strGroups & arGroups(x,0) & ", "
		end if
	next

	strQuery = "SELECT group_name, group_id FROM cal_groups" _
		& " WHERE group_id IN (" & strGroups & ")"
	Set rsGroups = db.Execute(strQuery,,&H0001)
	do while not rsGroups.EOF
		response.write "<option value=""" & rsGroups("group_id") _
			& """>" & rsGroups("group_name")
		rsGroups.MoveNext
	loop
	rsGroups.Close
	Set rsGroups = nothing
%>
	</select>
	</td>
	<td>
	<font face="Tahoma, Arial, Helvetica" size=2>
	<input type="checkbox" name="public">Public<br>
	<input type="checkbox" name="group">Group<br>
	<input type="checkbox" name="personal" checked>Personal
	</font>
	</td>
	<td align="right">
	<input type="submit" name="change" value="Update"><p>
	<input type="submit" name="logout" value="Logout">
	</td>
</table>
	
<!-- framing table -->
<font face="Verdana, Arial, Helvetica" size=2 color="#ffffff">
<input type="checkbox" name="remember">make default</font>
</td></table>
<!-- end framing table -->

</form>
</div>