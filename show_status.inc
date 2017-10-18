<%
' this updates the status bar with a message

function showStatus(msg)
	msg = Replace(msg, "'", "\'")
	showStatus = "onMouseOver=""status='" & msg & "'; " _
		& "return true;"" " & VbCrLf _
		& "onMouseOut=""status=''; return true;"""
end function

' this switches an icon

function switchIcon(name)
	switchIcon = "onMouseOver=""iconOver('" & name _
		& "'); return true;"" onMouseOut=""iconOut('" _
		& name & "'); return true;"""
end function
%>

