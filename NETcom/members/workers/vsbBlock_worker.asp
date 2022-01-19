<!--#include file="../../connect.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title></title>
<meta charset="windows-1255">
</head>
<%
if Request.QueryString("idSite")<>nil then
		sqlStr="update USERS set User_Bloked=1-User_Bloked where USER_ID=" & Request.QueryString("idsite") 
		con.GetRecordSet(sqlStr)
		sqlStr1="delete from FailedLogins where  USER_ID=" & Request.QueryString("idsite") 
		con.GetRecordSet(sqlStr1)
end if


Response.Redirect "default.asp"
set con = nothing
%>
</body>
</html>
