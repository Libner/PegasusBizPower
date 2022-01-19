<%@ Language=VBScript %>
<%
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,projectId,strLocal,sqlstr,UserID,date_,companyId,mechanismId,minutes,OrgId
dim rs_comp,hour_id
%>
<!--#include file="../include/connect.asp"-->
<%
	Server.ScriptTimeOut =120
	
	UserID = Request.QueryString("UserID")
	OrgId = Request.QueryString("OrgId")
	date_ = Request.QueryString("date")
	companyId = Request.QueryString("companyId")
	projectId = Request.QueryString("projectId")
	mechanismId = Request.QueryString("mechanismId")
		sqlstr = "SELECT minuts FROM hours WHERE user_id=" & UserID & " and organization_id=" & OrgId & " and date='" & date_ & "' and project_id=" & projectId & " and company_id=" & companyId &" and mechanism_Id = " &mechanismId
'		Response.Write sqlstr
'	Response.end
		    
	
			
	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<checkhour>"
	if date_ <> "" and companyId <> "" and UserID <> "" and projectId <> "" and mechanismId <> "" then
		
		con.Execute("SET DATEFORMAT dmy")
		
		sqlstr = "SELECT minuts FROM hours WHERE user_id=" & UserID & " and organization_id=" & OrgId & " and date='" & date_ & "' and project_id=" & projectId & " and company_id=" & companyId &" and mechanism_Id = " &mechanismId
		   
		set rs_comp = con.Execute(sqlstr)
		if not rs_comp.eof then
			XMLFileText = XMLFileText & "<minutes>" & rs_comp("minuts") & "</minutes>"
		else	
			XMLFileText = XMLFileText & "<minutes>0</minutes>"
		end if
		set rs_comp = Nothing
	
	end if 'username <> ""
	
	XMLFileText = XMLFileText & "</checkhour>"
	Response.Write XMLFileText
	
	set con = nothing%>
