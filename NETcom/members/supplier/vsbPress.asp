<!--#include file="../../connect.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title></title>
<meta charset="windows-1255">
</head>
<%
if Request.QueryString("idSite")<>nil then
		sqlStr="update Suppliers set isVisible=1-isVisible where supplier_Id=" & Request.QueryString("idsite")
		con.GetRecordSet(sqlStr)
set con = nothing

end if
if Request.QueryString("isAllowed")<>nil then
		sqlStr="update Suppliers set isAllowedSalesReport=1-isAllowedSalesReport where supplier_Id=" & Request.QueryString("isAllowed")
		con.GetRecordSet(sqlStr)
set con = nothing

end if
'response.Write sqlStr
Response.Redirect "default.asp"
%>
</body>
</html>
