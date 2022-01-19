<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))

	found_flag = -1   : phone = ""  :  cellular  = ""  : ContactId = ""
	set rs_tmp = Server.CreateObject("ADODB.recordset")       

	set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

	For each x In xmldoc.documentElement.childNodes
		if x.NodeName = "cellular" Then cellular=x.text  
		if x.NodeName = "phone" Then phone=x.text     
		if x.NodeName = "ContactId" Then ContactId=x.text     
	Next
 
	If trim(ContactId) = "" Or IsNumeric(ContactId) = false Or trim(ContactId) = "0" Then
		ContactId = 0
	End If	

	found_flag = 1		: found_contact_id = 0

	If OrgID <> "" Then
		sqlstr = "EXEC dbo.contacts_chk_phone	@OrgId='" & OrgID & "', @EditContactId='" & ContactId & "', " & _
		" @cp='" & cellular & "',	@pn='" & phone & "'"	
		set rs_tmp = con.getRecordSet(sqlstr)	
		If rs_tmp.eof = False Then
			found_contact_id = trim(rs_tmp(0))
		End If					
		set rs_tmp = Nothing
		set con = Nothing    
		
		If found_contact_id > 0 Then
			found_flag = 1
		Else
			found_flag = 0	
		End If
	Else 
		found_flag = -1
	End If	 		      
	    
	Response.Write found_flag
 
	set xmldoc = Nothing %>    