<!--#include file="connect.asp"-->
<%
    sqlstr = "Select ORGANIZATION_ID FROM ORGANIZATIONS ORDER BY ORGANIZATION_ID"
    set rs_orgs = con.getRecordSet(sqlstr)
    while not rs_orgs.eof
        orgID = trim(rs_orgs(0))
        sqlstr = "Select product_id from products WHERE ORGANIZATION_ID = " & orgID & " AND PRODUCT_TYPE = '1'"
        set rs_check = con.getRecordSet(sqlstr)
        If rs_check.eof Then
			sqlstr = "Insert Into Products (ORGANIZATION_ID,PRODUCT_TYPE) values ("&trim(ORGANIZATION_ID) & ",1)"
			con.executeQuery(sqlstr)
		End If
		set rs_check = nothing
		rs_orgs.moveNext
    wend
    set rs_orgs = nothing

%>