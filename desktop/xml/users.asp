<%@ Language=VBScript %>
<%
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,username,strLocal,sqlstr,password,UserId
dim rs_comp
%>
<!--#include file="../../include/connect.asp"-->
<!--#include file="../../include/reverse.asp"-->
<%
	Server.ScriptTimeOut =120

function sFix(myString)
	dim quote, newQuote, temp
	temp = myString
	if (isnull(temp) = False) Or (Len(trim(temp))>0) then
		quote = "'"
		newQuote = "''"
		if instr(myString,"'") then
			temp = Replace(temp, quote, newQuote)
		end if	
	sFix = temp
	else 	
		sFix = ""
	end if
end function
	
	'Declare variables		
	XMLFileText = ""
	username = cStr(Request.QueryString("u"))
	password = cStr(Request.QueryString("p"))
		
	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<list>"
	if CStr(username) <> "" and cStr(password) <> "" then
		
		sqlstr = "SELECT USER_ID, ORGANIZATION_ID, CHIEF ,FIRSTNAME,LASTNAME FROM USERS WHERE ACTIVE = 1 " &_
		" and loginName='"& sFix(username) &"' AND Password='"& sFix(password) &"'"
		   
		set rs_comp = con.Execute(sqlstr)
		if not rs_comp.eof then
		
			UserId = cStr(rs_comp("USER_ID"))
			
			XMLFileText = XMLFileText & "<employee><id>" & UserId & "</id><logon>" & username & "</logon><organizationId>" & rs_comp("ORGANIZATION_ID") & "</organizationId><status>" & rs_comp("CHIEF") & "</status></employee>"
			
		end if
		set rs_comp = Nothing
	
	end if 'username <> ""
	
	XMLFileText = XMLFileText & "</list>"
	Response.Write XMLFileText
	
	set con = nothing	%>
