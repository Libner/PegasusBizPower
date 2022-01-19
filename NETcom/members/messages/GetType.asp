<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%

    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

	for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "typeId" then typeId=x.text  
	next
	
	strTemp = ""
	
	If trim(typeId) <> "" Then    
		sqlstr = "SELECT messages_type_Content from messages_types where messages_type_id = " & typeId
		set rs_type = con.getRecordSet(sqlstr)	
		if rs_type.eof Then
			strTemp = ""
		else
			'separate rows using ";;" and columns using  ";"                         
			strTemp = rs_type.Getstring(,,",")
		end if
    Else
		strTemp = ""
    End If
    
    set rs_type = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write strTemp
    
%>