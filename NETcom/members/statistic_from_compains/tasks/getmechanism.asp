<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%

    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

	for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "projectID" then projectID=x.text  
	next
	
	strTemp = ""
	
	If trim(projectID) <> "" Then    
		sqlstr = "Select mechanism_ID, mechanism_NAME FROM mechanism WHERE ORGANIZATION_ID = " & OrgID & " AND Project_ID = " & projectID & " Order BY mechanism_NAME"
		set rs_mechanisms = con.getRecordSet(sqlstr)	
		if rs_mechanisms.eof Then
			strTemp = ""
		else
			'separate rows using ";;" and columns using  ";"                         
			strTemp = rs_mechanisms.Getstring(,,",",";")
		end if
    Else
		strTemp = ""
    End If
    
    set rs_mechanisms = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write strTemp
    
%>