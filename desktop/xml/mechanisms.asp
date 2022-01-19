<%@ Language=VBScript %>
<%
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,OrgID,strLocal,sqlstr,companyID,projectID,rs_tmp,arr_mech,mm,budget_hours


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
	projectID = ""
	OrgID = Request.QueryString("OrgID")
	companyID = Request.QueryString("companyID")
	projectID = Request.QueryString("projectID")		

	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<mechanisms>"
	'if OrgID <> "" and companyID <> "" then
	if OrgID <> ""  then	
		sqlstr = "Select mechanism_ID, mechanism_NAME, budget_hours FROM mechanism WHERE " &_
		" ORGANIZATION_ID = " & OrgID & " AND project_Id = " &projectID & " Order BY mechanism_NAME"	
		set rs_tmp = con.Execute(sqlstr)
		If not rs_tmp.eof Then
			arr_mech = rs_tmp.getRows()
		End If
		Set rs_tmp = Nothing
		
		If isArray(arr_mech) Then
		For mm=0 To Ubound(arr_mech,2)
		
			budget_hours = trim(arr_mech(2,mm))
			If isNull(budget_hours) Or Len(budget_hours) = 0 Then
				budget_hours = 0
			End If
			
			XMLFileText = XMLFileText & "<mechanism id=""" & arr_mech(0,mm) &_
			""" name=""" & validname(arr_mech(1,mm)) & """ budget_hours=""" & budget_hours & """></mechanism>"
			
		Next
		End If
	
	end if 'OrgID <> ""
	
	XMLFileText = XMLFileText & "</mechanisms>"
	
	Response.Write XMLFileText
	
	set con = nothing%>
