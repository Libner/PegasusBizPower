<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%
	sqlstr = "Select ORGANIZATION_ID,LANG_ID FROM ORGANIZATIONS Order BY ORGANIZATION_ID"
	set rs_org = con.getRecordSet(sqlstr)
	while not rs_org.eof
		orgID = trim(rs_org(0))
		langID = trim(rs_org(1))
		If trim(langID) =  "2" Then
			lang_ = "_Eng"
		Else
			lang_ = ""
		End If	
		sqlUpdate="Update COMPANIES SET COMPANIES.city_Name = cities.city_Name" & lang_ &_
        " FROM cities WHERE COMPANIES.city_ID = cities.city_ID AND "&_
        " COMPANIES.ORGANIZATION_ID = " & trim(OrgID)    
		con.executeQuery(sqlUpdate)   

	rs_org.moveNext
	Wend
	set rs_org = Nothing		

%>