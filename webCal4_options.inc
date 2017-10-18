<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 08/10/1999
 -->
 
<script language="javascript">
<!--

if (document.all) {
// this browser can do what we need

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
}

//-->
</script>

<form action="webCal4_prep.asp?date=<%=DateSerial(y,m,1)%>" method="post">

<!-- framing table -->
<table bgcolor="#<%=arColor(5)%>" border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><font face="Verdana, Arial, Helvetica" size=2 color="#ffffff"><b>Options</b></font></td>
	<td align="right"><a href="#" onClick="minIt(); return true;" <%=ShowStatus("hide options")%>><img name="min_img" src="images/icon_min.gif" width=16 height=14 alt="" border="0"></a><a href="#" onClick="maxIt(); return true;" <%=ShowStatus("display options")%>><img name="max_img" src="images/icon_max_click.gif" width=16 height=14 alt="" border="0"></a></td>
<tr>
	<td colspan=2>
<!-- end framing table -->

<div id="options" class="child">
	
<table bgcolor="#<%=arColor(12)%>" border=0 cellspacing=0 cellpadding=3>
<tr>
	<td>
	<select name="groups" multiple size=3>
<%
' get the array of groups available to current user

if Session(dataName & "User") <> "" then
	arGroups = Session(dataName & "Groups")
	arScopes = Session(dataName & "Scopes")
'else
	
' if no user has logged in then, for now, just display the
' first two available groups at maximum permissions (5)

'	dim arGroups(2,1)
	
'	arGroups(1,0) = 1
'	arGroups(1,1) = 5
'	arGroups(2,0) = 2
'	arGroups(2,1) = 5
'end if

	for x = 0 to UBound(arGroups)
	
' go through each group visible to current user
' if the user has permission >= 1 then display group
' form value is the group id
	
		if arGroups(x,2) >= 1 then
			response.write "<option value=""" & arGroups(x,0) & """"
			if arGroups(x,3) then
				response.write " selected"
			end if
			response.write ">" & arGroups(x,1)
		end if
	next
		
%>
	</select>
	</td>
	<td>
	<font face="Tahoma, Arial, Helvetica" size=2>
	<input type="checkbox" name="2"<%if arScopes(2) then%> checked<%end if%>>Public<br>
	<input type="checkbox" name="1"<%if arScopes(1) then%> checked<%end if%>>Group<br>
	<input type="checkbox" name="0"<%if arScopes(0) then%> checked<%end if%>>Personal
	</font>
	</td>
<input type="hidden" name="total" value="<%=UBound(arGroups)+1%>">
<% end if %>
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