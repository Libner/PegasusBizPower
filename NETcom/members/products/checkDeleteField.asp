<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%    
    Response.CharSet = "windows-1255"
    Session.LCID = 1037
    set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

	for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "FieldID" then FieldID=x.text  
	next
	
	is_answered = ""
	sqlstr = "Select Top 1 Field_Id FROM FORM_VALUE WHERE FIELD_ID="&FieldID&" And ltrim(rtrim(FIELD_VALUE)) <> ''"
	set rs_check = con.getRecordSet(sqlstr)
	If not rs_check.eof Then
		is_answered = 1
	Else
		is_answered = 0	
	End If
	set rs_check = Nothing
    
    Response.Write is_answered
    
%>