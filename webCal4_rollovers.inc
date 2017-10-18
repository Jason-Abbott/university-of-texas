<%
' this updates the status bar with a message

function showStatus(strMsg)
	strMsg = Replace(strMsg, "'", "\'")
	showStatus = "onMouseOver=""status='" & strMsg & "'; " _
		& "return true;"" " & VbCrLf _
		& "onMouseOut=""status=''; return true;"""
end function

' this switches an icon

function switchIcon(strName, strSource, strMsg)
	if strSource = "" then
		strSource = strName
	end if

	switchIcon = "onMouseOver=""" _
		& "iconOver('" & strName & "','" & strSource & "','" & strMsg & "'); return true;"" " _
		& "onMouseOut=""" _
		& "iconOut('" & strName & "','" & strSource & "'); return true;"""
end function
%>

<script language="javascript"><!--

function iconOver(name, source, message){
	if (document.images) {
  		document.images[name].src=eval("icon"+source+"On.src");
		status=message;
	}
}

function iconOut(name, source){
	if (document.images) {
  		document.images[name].src=eval("icon"+source+".src");
		status="";
	}
}
//-->
</script>