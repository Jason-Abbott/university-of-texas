<table align="right">
<tr>
	<td align="right">
	<!--#include file="webCal4_legend.inc"-->
	</td>
<tr>
<form name="user" action="webCal4_<%=strView%>.asp" method="post">
	<td align="right">
<font face="Tahoma, Arial, Helvetica" size=1>view:</font>
<select name="staff_id" onChange="document.user.submit();" >
<option value="<%=Session("StudentID")%>">my calendar
<%
strQuery = "SELECT pwid, First_Name, Last_Name FROM " _
	& "tblSTAFF WHERE CSOgroup LIKE '%Counselors' " _
	& "AND pwid<>" & Session("StudentID") _
	& " ORDER BY Last_Name, First_Name"

Set rsUsers = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenStatic = 3
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsUsers.Open strQuery, strDSN, 3, 1, &H0001
do while not rsUsers.EOF
	response.write "<option value=""" & rsUsers("pwid") & """"
	if CInt(rsUsers("pwid")) = CInt(intID) then
		response.write " selected"
	end if
	response.write ">" & rsUsers("Last_Name") & ", " & rsUsers("First_Name") _
		& VBCrLf
	rsUsers.MoveNext
loop
rsUsers.Close
Set rsUsers = nothing
%>
</select>
	</td>
</form>
</table>
