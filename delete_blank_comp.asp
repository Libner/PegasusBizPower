<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%
    sqlstr = "Select company_id FROM companies WHERE Len(company_name) = 0 Order BY company_id"
    set rs_companies = con.getRecordSet(sqlstr)
    while not rs_companies.eof	
		company_id = rs_companies(0)
		
		If trim(company_id) <> "" Then
		sqlstr = "Update Appeals Set company_id = NULL WHERE company_id = " & company_id
		con.executeQuery sqlstr
		
		sqlstr = "Update Tasks Set company_id = NULL WHERE company_id = " & company_id
		con.executeQuery sqlstr
		
		sqlstr = "Delete From companies WHERE company_id = " & company_id
		con.executeQuery sqlstr
		
        End If
    rs_companies.moveNext
    Wend
    set rs_companies = Nothing
%>
