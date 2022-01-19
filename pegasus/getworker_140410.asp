<!--#INCLUDE FILE="../Netcom/reverse.asp"-->
<!--#include file="../Netcom/connect.asp"-->
<%Response.AddHeader "cache-control", "private"
	 Response.AddHeader "pragma", "no-cache"
	 Response.ContentType = "text/xml"
	 response.buffer=true
	 Response.Expires = 0 
	 
	function validname(txt)
		dim outtxt
		outtxt=txt
		outtxt=replace(outtxt,"&"," ")
		outtxt=replace(outtxt,"<"," ")
		outtxt=replace(outtxt,">"," ")
		outtxt=replace(outtxt,chr(9)," ")
		outtxt=replace(outtxt,chr(10)," ")
		outtxt=replace(outtxt,chr(13)," ")
		outtxt=replace(outtxt,chr(34),"''")
		validname=trim(outtxt)
	end function	 

	phone = trim(Request.QueryString("phone"))	
    phone = Replace(phone, " ", "")
    phone = Replace(phone, "_", "")
    phone = Replace(phone, "-", "")
    phone = Replace(phone, ".", "")
    phone = Replace(phone, Chr(39), "")
    phone = Replace(phone, Chr(34), "")	
	
	ContactId = 0 : responsible_id = 0 : responsible_name = "" : is_exists = 0
	
	XMLFileText = "<?xml version=""1.0"" encoding=""windows-1255"" ?>" & vbCrLf 
	XMLFileText = XMLFileText & "<Result>"
	
	If Len(phone)>0 Then'
		sqlstr = "SELECT TOP 1 C.CONTACT_ID, responsible_id,  U1.LOGINNAME as responsible_name " & _
		" FROM dbo.CONTACTS C  LEFT OUTER JOIN Users U1 ON C.Responsible_ID = U1.User_ID " & _
		" WHERE (C.ORGANIZATION_ID  = 264) " & _ 
		" AND ((dbo.fnc_remove_chr(C.phone) = '" & sFix(phone) & "') " & _
		" OR (dbo.fnc_remove_chr(C.cellular) = '" & sFix(phone) & "')) ORDER BY date_update DESC"
	    set rs = con.GetRecordSet(sqlstr) 
		If Not rs.EOF then		
			ContactId = cLng(rs("CONTACT_ID"))
			responsible_id = trim(rs("responsible_id"))
			If Not isNULL(rs("responsible_name")) Then
				responsible_name = trim(cStr(rs("responsible_name")))
			Else
				responsible_name = ""
			End If
			is_exists = 1
		End If
		Set rs = Nothing
		
		XMLFileText = XMLFileText & "<Worker WorkerId=""" & responsible_id & _
		""" WorkerName=""" & validname(responsible_name) & """ ClientExists='" & is_exists & "'></Worker>"
     
	End If
	
	XMLFileText = XMLFileText & "</Result>"
	Response.Write XMLFileText
	
	Set con = Nothing 	%>