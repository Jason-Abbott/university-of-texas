<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/17/1999

dim strDSN, strDelim

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

'dataName = "newCal"

' Access uses # to delimit dates while SQL uses '
' so switch as necessary

strDelim = "'"

' This opens the database stored in the
' the directory /data under the main webCal files

'dim db
'Set db = Server.CreateObject("ADODB.Connection")
'DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
'	& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")
strDSN="DSN=Struct;UID=csosql;PWD=bus1ness"
'db.Open DSN

' make sure the containing file does
' db.Close
' Set db = nothing
%>