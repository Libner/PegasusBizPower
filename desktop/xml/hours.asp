<%@ Language=VBScript %>
<%
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,OrgID,strLocal,sqlstr,mechanismID,companyID,projectID,UserId,rs_tmp,sum_minuts,sum_hours,sum_my_hours


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
	mechanismID = ""	
	UserId = ""
	OrgID = Request.QueryString("OrgID")
	mechanismID = Request.QueryString("mechanismID")
	companyID = Request.QueryString("companyID")
	projectID = Request.QueryString("projectID")
	UserId = Request.QueryString("UserId")		

	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<hours>"
	'if OrgID <> "" and mechanismID <> "" then
	if OrgID <> ""  then	
	    sum_hours = 0 : sum_my_hours = 0
		sqlstr = "Select Sum(minuts) FROM hours WHERE company_ID = " & companyID & " And project_ID = " & projectID &_
		" And ORGANIZATION_ID = " & OrgID & " AND mechanism_id = " & mechanismID
		set rs_tmp = con.Execute(sqlstr)
		If not rs_tmp.eof Then
			sum_minuts = trim( rs_tmp(0))
			If isNUll(sum_minuts) = false And len(sum_minuts) > 0 Then
				sum_hours = Round(cDbl(sum_minuts) / 60, 1)
			End If			
		End If
		Set rs_tmp = Nothing
		
		sqlstr = "Select Sum(minuts) FROM hours WHERE company_ID = " & companyID & " And project_ID = " & projectID &_
		" And ORGANIZATION_ID = " & OrgID & " AND mechanism_id = " & mechanismID & " And user_id = " & UserId	
		set rs_tmp = con.Execute(sqlstr)
		If not rs_tmp.eof Then
			sum_minuts = trim(rs_tmp(0))
			If isNUll(sum_minuts) = false And len(sum_minuts) > 0 Then
				sum_my_hours = Round(cDbl(sum_minuts) / 60, 1)
			End If			
		End If
		Set rs_tmp = Nothing				
		
		XMLFileText = XMLFileText & "<hour mechanism_id=""" & mechanismID &	""" user_id=""" & UserId & _
		""" sum_hours=""" & sum_hours & """ sum_my_hours=""" & sum_my_hours & """></hour>"
	
	end if 'OrgID <> ""
	
	XMLFileText = XMLFileText & "</hours>"
	
	Response.Write XMLFileText
	
	set con = nothing%>
