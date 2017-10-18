<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 09/30/1999

dim mLoop, yLoop
%>

<tr>
	<td><%=strTitle%></td>
	<form method="post" action="webCal4_staff-month.asp?other_id=<%=intID%>">
	<td align="right" valign="bottom"><nobr>

		<a href="webCal4_<%=strType%>-<%=strView%>.asp?date=<%=strPrev%>"
		<%=switchIcon("Prev","", "Previous " & strView)%>><img name="Prev" src="./images/icon_calprev_grey.gif"
		 width=15 height=16 alt="Previous <%=strView%>" border=0></a>
		&nbsp;
		<a href="webCal4_<%=strType%>-<%=strView%>.asp?date=<%=strNext%>"
		<%=switchIcon("Next","", "Next " & strView)%>><img name="Next" src="./images/icon_calnext_grey.gif"
		 width=15 height=16 alt="Next <%=strView%>" border=0></a>
		&nbsp;
<!-- 		<a href="webCal4_rules-admin.asp?action=new&view=<%=strView%>"
		<%=switchIcon("Rules","", "Edit Rules")%>><img name="Rules" src="./images/icon_users_grey.gif"
		 width=15 height=16 alt="Edit Rules" border=0></a>
		&nbsp; -->
		<a href="javascript:document.forms[0].submit();" 
		<%=switchIcon("Goto","", "Goto the selected date")%>><img name="Goto" src="./images/icon_goto_grey.gif"
		 width=18 height=15 alt="Goto the selected date" border=0></a>
		
		<select name="month">
<%
' this creates the form list of month names

for mLoop = 1 to 12
	response.write "<option value='" & mLoop & "'"
	if mLoop = Month(Date) then response.write " selected"
	response.write ">" & MonthName(mLoop,1) & VbCrLf
next
%>
		</select>
		<select name="year">
<%
' this creates the form list of 20 years

for yLoop = Year(Date) - 10 to Year(Date) + 10
	response.write "<option"
	if yLoop = Year(Now) then response.write " selected"
	response.write ">" & yLoop & VbCrLf
next
%>
		</select>
	</nobr>
	</td>
	</form>	