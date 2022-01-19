<!--#include file="..\..\netcom/connect.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
</head>

<%
ORGANIZATION_ID = Request.QueryString("ORGANIZATION_ID")
if Request.QueryString("idSite")<>nil then
		sqlStr="update USERS set active=1-active where USER_ID=" & Request.QueryString("idsite")
		con.GetRecordSet(sqlStr)
end if
%>

<meta http-equiv="refresh" content="0;url=workers.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>">  
<%
set con = nothing
%>
</body>
</html>
