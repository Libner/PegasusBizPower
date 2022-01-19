<!--#include file="../../connect.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title></title>
<meta charset="windows-1255">
</head>
<%
if Request.QueryString("idSite")<>nil then
		sqlStr="update USERS set active=1-active where USER_ID=" & Request.QueryString("idsite") &";update USERS set email=null where USER_ID=" & Request.QueryString("idsite")&" and active=0;"
		con.GetRecordSet(sqlStr)
		
end if


Response.Redirect "default.asp"
set con = nothing
%>
</body>
</html>
