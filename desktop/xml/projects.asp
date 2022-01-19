<%@ Language=VBScript %>
<%
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,OrgID,strLocal,sqlstr,companyID
dim rs_tmp, arr_proj, pp

function validname(txt)
	dim outtxt
	outtxt=txt
	outtxt=replace(outtxt,"&"," ")
	outtxt=replace(outtxt,"<"," ")
	outtxt=replace(outtxt,">"," ")
	outtxt=replace(outtxt,chr(9)," ")
	outtxt=replace(outtxt,chr(10)," ")
	outtxt=replace(outtxt,chr(13)," ")
	outtxt=replace(outtxt,chr(34),"''")
	'outtxt=replace(outtxt,cstr(cint(x9)),"_")
	'outtxt=replace(outtxt,cstr(cint(xA)),"_")
	'outtxt=replace(outtxt,cstr(cint(xD)),"_")
	validname=trim(outtxt)
end function
%>
<!--#include file="../../include/connect.asp"-->
<%
	Server.ScriptTimeOut =120
	
	'Declare variables		
	XMLFileText = ""
	OrgID = ""
	companyID = ""	
	OrgID = Request.QueryString("OrgID")
	companyID = Request.QueryString("companyID")
	
	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<projects>"
	if OrgID <> "" and companyID <> "" then
		
		sqlstr = "Select PROJECT_ID, PROJECT_NAME FROM PROJECTS WHERE company_id IN (0," & companyID &_
		") AND ORGANIZATION_ID = " & OrgID & " AND active = 1 AND status = '2' Order BY PROJECT_NAME"
		set rs_tmp = con.Execute(sqlstr)
		If not rs_tmp.eof Then
			arr_proj = rs_tmp.getRows()
		End If
		Set rs_tmp = Nothing
		
		If isArray(arr_proj) Then
		For pp=0 To Ubound(arr_proj,2)
			
			XMLFileText = XMLFileText & "<project id=""" & arr_proj(0,pp) & """ name=""" & validname(arr_proj(1,pp)) & """></project>"
			
		Next
		End If
	
	end if 'OrgID <> ""
	
	XMLFileText = XMLFileText & "</projects>"
	Response.Write XMLFileText
	
	set con = nothing%>
