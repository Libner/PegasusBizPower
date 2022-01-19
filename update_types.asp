<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%
    sqlstr = "Select contact_id,company_id FROM contacts Order BY contact_id"
    set rs_contacts = con.getRecordSet(sqlstr)
    while not rs_contacts.eof
		contact_id = rs_contacts(0)
		company_id = rs_contacts(1)
		
		If trim(company_id) <> "" Then
		sqlstr = "Select type_id From company_to_types WHERE company_id = " & company_id
		set rs_types = con.getRecordSet(sqlstr)
		while not rs_types.eof
			con.executeQuery "Insert Into contact_to_types (contact_id, type_id) Values (" &_
			contact_id & "," & rs_types(0) & ")"
			rs_types.moveNext
		wend
		set rs_types = Nothing	
        End If
    rs_contacts.moveNext
    Wend
    set rs_contacts = Nothing
%>