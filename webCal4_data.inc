<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/17/1999

dim strDSN, strDelim

' Access uses # to delimit dates while SQL uses '
' so switch as necessary

strDelim = "'"
strDSN="DSN=Struct;UID=csosql;PWD=bus1ness"
%>