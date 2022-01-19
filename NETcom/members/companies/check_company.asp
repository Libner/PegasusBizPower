<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%

 OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))

 found_flag = -1 
 set rs_name = Server.CreateObject("ADODB.recordset")       

 set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
 xmldoc.async=false
 xmldoc.load(request)

 for each x in xmldoc.documentElement.childNodes
   if x.NodeName = "cName" then cName=x.text  
 next

 If cName <> "" And OrgID <> "" Then
	sqlstr = "Select * FROM companies WHERE company_name Like '"& sFix(cName) &"' And ORGANIZATION_ID = " & OrgID
	'Response.Write sqlStr
	set rs_name = con.getRecordSet(sqlstr)	
	If rs_name.eof = true Then
		found_flag = 0
	Else
		found_flag = 1		
	End If					
 Else 
	  found_flag = -1
 End If	 		      
    
 Response.Write found_flag
 set rs_name = Nothing
 set con = Nothing     
 set xmldoc = Nothing
%>    