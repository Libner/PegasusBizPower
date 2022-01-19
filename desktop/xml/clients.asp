<%@ Language=VBScript %>
<%  
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.ContentType = "text/xml"
response.buffer=true
Response.Expires = 0 
dim con
dim XMLFileText,OrgID,strLocal,sqlstr
dim rs_tmp, arr_comp, cc
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
	validname=trim(outtxt)
end function
%>
<!--#include file="../../include/connect.asp"-->
<%
	Server.ScriptTimeOut =120
	
	'Declare variables		
	XMLFileText = ""
	OrgID = ""
		
	OrgID = Request.QueryString("OrgID")
	
	XMLFileText = "<?xml version=""1.0"" encoding=""iso-8859-8"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<clients>"
	if OrgID <> "" then	
	
	sqlstr = "Select COMPANY_ID, COMPANY_NAME FROM COMPANIES Where ORGANIZATION_ID = "& OrgID &_
	" And status = 2 Order BY COMPANY_NAME"
	set rs_tmp = con.Execute(sqlstr)
	If not rs_tmp.eof Then
		arr_comp = rs_tmp.getRows()
	End If
	Set rs_tmp = Nothing
	
	If isArray(arr_comp) Then
	For cc=0 To Ubound(arr_comp,2)
		
		XMLFileText = XMLFileText & "<client id=""" & arr_comp(0,cc) & """ name=""" & validname(arr_comp(1,cc)) & """></client>"
		
	Next
	End If
	
	end if 'OrgID <> ""
	
	XMLFileText = XMLFileText & "</clients>"
	Response.Write XMLFileText
	
	set con = nothing%>
