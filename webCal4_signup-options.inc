<font face="Tahoma, Arial, Helvetica" size=2 color="#ffffff">
<input type="submit" value="Check">
<select name="staff_id">
<%
strQuery = "SELECT pwid, First_Name, Last_Name FROM " _
	& "tblSTAFF WHERE CSOgroup='" & Request.QueryString("type") _
	& " Counselors' ORDER BY Last_Name, First_Name"

Set rsStaff = Server.CreateObject("ADODB.RecordSet")

'  cursor: adOpenForwardOnly = 0
' locking: adLockReadOnly = 1
' command: adCmdText = &H0001

rsStaff.Open strQuery, strDSN, 0, 1, &H0001
do while not rsStaff.EOF
	response.write "<option value=""" & rsStaff("pwid") & """"
	
	if CInt(rsStaff("pwid")) = CInt(Request.Form("staff_id")) then
		response.write " selected"
	end if
	
	response.write ">" & rsStaff("Last_Name") & ", " _
		& rsStaff("First_Name") & VBCrLf
	rsStaff.MoveNext
loop
rsStaff.Close
Set rsStaff = nothing
%>
</select>'s schedule

<input type="radio" name="week" value="this"<%=strWeek1%>>this week or
<input type="radio" name="week" value="next"<%=strWeek2%>>next week
