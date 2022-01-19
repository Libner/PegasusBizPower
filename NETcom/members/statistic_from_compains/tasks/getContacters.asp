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
		sqlstr = "SELECT contact_ID,contact_Name FROM contacts WHERE company_id = " & compId & " ORDER BY contact_Name"       
		set rs_contacters = con.getRecordSet(sqlstr)   
		if rs_contacters.eof Then
			strTemp = ""
		else
			'separate rows using ";;" and columns using  ";"                         
			strTemp = rs_contacters.Getstring(,,",",";") 			
		end if
    Else
		strTemp = ""
    End If
    
    set rs_contacters = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write strTemp
    
%>