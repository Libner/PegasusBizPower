<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%

    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

	for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "compId" then compId=x.text  
	next
	
	strTemp = ""
	
	If trim(compId) <> "" Then    
		sqlstr = "Select PROJECT_ID, PROJECT_NAME FROM PROJECTS WHERE ORGANIZATION_ID = " & OrgID & " AND company_Id IN (" & compId & ", 0) Order BY PROJECT_NAME"
		set rs_projects = con.getRecordSet(sqlstr)	
		if rs_projects.eof Then
			strTemp = ""
		else
			'separate rows using ";;" and columns using  ";"                         
			strTemp = rs_projects.Getstring(,,",",";")
		end if
    Else
		strTemp = ""
    End If
    
    set rs_projects = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write strTemp
    
%>