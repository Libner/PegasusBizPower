<%@ Language=VBScript %>
<%
Session.LCID = 2057
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,projectId,strLocal,sqlstr,UserID,date_,companyId,minutes,OrgId,mechanismId
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
	minutes = Request.QueryString("min")	
'		sqlstr = "SELECT hour_id FROM hours WHERE user_id=" & UserID & " and organization_id=" & OrgId & " and date='" & date_ & "' and project_id=" & projectId & " and mechanism_Id= "& mechanismId &" and company_id=" & companyId
	
'Response.Write("yes="& sqlstr	)
'Response.end

	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<addhour>"
	if date_ <> "" and companyId <> "" and UserID <> "" and projectId <> "" and mechanismId<> "" and minutes <> "" then
		
		con.Execute("SET DATEFORMAT dmy")
		
		sqlstr = "SELECT hour_id FROM hours WHERE user_id=" & UserID & " and organization_id=" & OrgId & " and date='" & date_ & "' and project_id=" & projectId & " and mechanism_Id= "& mechanismId &" and company_id=" & companyId
		   
		set rs_comp = con.Execute(sqlstr)
		if not rs_comp.eof then
			' update minutes
			hour_id = rs_comp("hour_id")
			con.Execute("Update hours Set minuts = " & minutes & " WHERE hour_id = " & hour_id)
			XMLFileText = XMLFileText & "<ok>update</ok>"			
		else	
			'insert new record
		'	Response.Write("Insert")
			
			sqlstr = "SET DATEFORMAT DMY; Insert Into hours (user_id,organization_id,date,project_id,mechanism_Id,company_id,minuts) VALUES (" &_
					UserID & "," & OrgId & ",'" & date_ & "'," & projectId &"," &mechanismId & "," & companyId & "," & minutes & ")"
			con.Execute(sqlstr)
			
			XMLFileText = XMLFileText & "<ok>insert</ok>"			
		end if
		set rs_comp = Nothing
	
	end if 'username <> ""
	
	XMLFileText = XMLFileText & "</addhour>"
	Response.Write XMLFileText
	
	set con = nothing%>
