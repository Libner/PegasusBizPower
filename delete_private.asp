<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%
    sqlstr = "Select company_id, organization_id FROM companies WHERE private = '1'"
    set rs_comp = con.getRecordSet(sqlstr)
    while not rs_comp.eof		
		company_id = rs_comp(0)
		orgID = rs_comp(1)
		
		sqlstr = "Select company_id From companies WHERE Organization_ID = " & orgID & " And private_flag = 1"
		set rs_private = con.getRecordSet(sqlstr)
		if not rs_private.eof then
		
			compPrivID = rs_private(0)
					
			If trim(company_id) <> "" Then
			
			sqlstr = "Select contact_id From contacts WHERE company_id = " & company_id
			set rs_contacts = con.getRecordSet(sqlstr)
			while not rs_contacts.eof
				contact_id = rs_contacts(0)
				sqlstr = "Update Appeals set company_id = "&compPrivID&", contact_id = " & contact_id & " WHERE company_id = " & company_id
				con.executeQuery sqlstr
				
				sqlstr = "Update Tasks set company_id ="&compPrivID&", contact_id = " & contact_id & " WHERE company_id = " & company_id
				con.executeQuery sqlstr
				
				sqlstr = "Update Projects set company_id ="&compPrivID&" WHERE company_id = " & company_id
				con.executeQuery sqlstr		
				
				sqlstr = "Update company_documents set company_id ="&compPrivID&" WHERE company_id = " & company_id
				con.executeQuery sqlstr									
			
				sqlstr = "Update contacts set company_id = "&compPrivID&" WHERE company_id = " & company_id
				con.executeQuery sqlstr			
				
			
			rs_contacts.moveNext
			wend
			set rs_contacts = Nothing
			
			sqlstr = "Delete From Companies WHERE company_id = " & company_id
			con.executeQuery sqlstr
			End If
		End If
		set rs_private = Nothing	
		
    rs_comp.moveNext
    Wend
    set rs_comp = Nothing
%>