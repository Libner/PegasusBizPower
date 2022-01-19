<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->

<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
</head>

<%
if Request.QueryString("idSite")<>nil then
		sqlStr="update ORGANIZATIONS set active=1-active where ORGANIZATION_ID=" & Request.QueryString("idsite")
		con.GetRecordSet(sqlStr)
end if
%>

<meta http-equiv="refresh" content="0;url=default.asp">  
<%
set con = nothing
%>
</body>
</html>
