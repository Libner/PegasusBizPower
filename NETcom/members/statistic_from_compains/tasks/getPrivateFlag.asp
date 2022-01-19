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
	
	private_flag = 0
	
	If trim(compId) <> "" Then    
		sqlstr = "Select private from companies WHERE company_Id = " & compId
		set rs_flag = con.getRecordSet(sqlstr)   
		if rs_flag.eof Then
			private_flag = 0
		else        
			private_flag = rs_flag(0)        
		end if    
    Else
		private_flag = 1
    End If
    
    set rs_flag = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write private_flag
    
%>