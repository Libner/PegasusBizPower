<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%
	sqlstr = "Select ORGANIZATION_ID,LANG_ID FROM ORGANIZATIONS Order BY ORGANIZATION_ID"
	set rs_org = con.getRecordSet(sqlstr)
	while not rs_org.eof
		orgID = trim(rs_org(0))
		langID = trim(rs_org(1))
		If trim(langID) =  "2" Then
			mechanism_name = "General"
		Else
			mechanism_name = "כללי"
		End If
		
		sqlstr = "Select project_id, company_id From projects Where ORGANIZATION_ID = " & orgID
		set rs_pr = con.getRecordSet(sqlstr)
		While not rs_pr.eof
			prID =  trim(rs_pr(0))
			compID =  trim(rs_pr(1))	    
		        	
			sqlstr = "SET NOCOUNT ON; Insert Into mechanism (project_id,company_id,mechanism_name,Organization_ID) values (" &_
			prID & "," & compID & ",'" & sFix(mechanism_name) & "'," & orgID & "); SELECT @@IDENTITY AS NewID"
			set rs_tmp = con.getRecordSet(sqlstr)
				mechID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing
			
			sqlstr = "Update hours Set mechanism_id = " & mechID & " Where company_id = " & compID &_
			" And project_id = " & prID & " And Organization_ID = " & orgID
			con.executeQuery(sqlstr)
			
			sqlstr = "Update pricing_to_projects Set mechanism_id = " & mechID & " Where project_id = " & prID
			con.executeQuery(sqlstr)			
				
			rs_pr.moveNext
		Wend
		set rs_pr = Nothing
			
	rs_org.moveNext
	Wend
	set rs_org = Nothing
%>