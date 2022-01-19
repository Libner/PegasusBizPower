<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title></title>
<meta charset="windows-1255">
</head>
<%
if Request.QueryString("idSite")<>nil then
    sql="Select company_id from PROJECTS where PROJECT_ID=" & Request.QueryString("idsite")
    set pr=con.GetRecordSet(sql)
    
 
    
	sqlStr="Update PROJECTS set active=1-active where PROJECT_ID=" & Request.QueryString("idsite")
	con.GetRecordSet(sqlStr)
end if
If isNumeric(wizard_id) Then
	Response.Redirect "../wizard/wizard_" & wizard_id & "_3.asp"
Else
if trim(pr("company_id"))="" or pr("company_id")=0 then
   Response.Redirect "default_action.asp"
else
	Response.Redirect "default.asp"
end if	
End If	
set con = nothing
%>
</body>
</html>
