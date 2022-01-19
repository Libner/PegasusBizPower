<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%
	sqlstr = "Select ORGANIZATION_ID,LANG_ID FROM ORGANIZATIONS Order BY ORGANIZATION_ID"
	set rs_org = con.getRecordSet(sqlstr)
	while not rs_org.eof
		orgID = trim(rs_org(0))
		langID = trim(rs_org(1))
		If trim(langID) =  "2" Then
			company_name = "No company"
		Else
			company_name = "ללא חברה"
		End If	
		sqlstr = "Insert into companies (Organization_ID,company_name,private_flag) values (" &_
		trim(OrgID) & ",'" & sFix(company_name) & "',1)"
		con.executeQuery(sqlstr)
	rs_org.moveNext
	Wend
	set rs_org = Nothing
%>