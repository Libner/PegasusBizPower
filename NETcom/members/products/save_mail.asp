<%@ Language=VBScript%>
<%
	Response.CharSet = "windows-1255"
	Response.Buffer = false
	Server.ScriptTimeout = 600
%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--	
	window.opener.document.location.href="products.asp";		
//-->
</SCRIPT>
</head>
<BODY onload="window.focus();" style="margin:0px">
<%
   UserID=trim(Request.Cookies("bizpegasus")("UserID"))
   OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
   set upl=Server.CreateObject("SoftArtisans.FileUp")
   prodId = upl.Form("prodId")   
   groups_id = trim(upl.Form("groups_id"))	
   
   If prodID = "" Or IsNull(prodID) Then
		If groups_id <> nil Then
			sqlStr = "Select Count(DISTINCT PEOPLE_ID) "&_ 
			" FROM PEOPLES where ORGANIZATION_ID="& OrgID &_
			" AND GROUPID IN ("& groups_id &")"				
			set rsc = con.getRecordSet(sqlStr)
			If not rsc.eof Then
				totalCount = trim(rsc(0))
			Else
				totalCount = 0	
			End If
			set rsc = Nothing	
		Else
			totalCount = 0
		End If
	 Else
		sqlStr = "Select Count(DISTINCT PEOPLE_ID) "&_ 
		" FROM PEOPLES where ORGANIZATION_ID="& OrgID &_
		" AND GROUPID IN ("& groups_id &") "&_
		" And People_ID Not IN (Select People_ID From PRODUCT_CLIENT WHERE PRODUCT_ID = " & prodID & ")"			
		set rsc = con.getRecordSet(sqlStr)
		If not rsc.eof Then
			totalCount = trim(rsc(0))
		Else
			totalCount = 0	
		End If
		set rsc = Nothing		
   End If
	
   If prodId = "" or IsNull(prodId) then 	
	If upl.Form("editFlag") <> nil Then    
	         
			If (upl.TotalBytes > 204800) Then
			%>
			<div style="background-color:#D4D0C8;width:100%;height:100%;color:red">
			<br><br>
			<p align=center>You can not attach a file larger then 200 KB</p> 
			<br>
			<p align=center><input type=button value="Ok" onclick="window.close();" style="width:60"></p> 
			</div>			
			<%
			Set upl = Nothing
			Response.End
			End If 
	  
   			If upl.Form("attachment") <> nil Then   		
   				attachment=trim(upl.Form("attachment"))    		
   			Else
   				attachment=""	
   			End if   	   	
	 
			dateend = "NULL"		
			datestart = " getDate()"			
			
			emailSubject = sFix(trim(upl.Form("EMAIL_SUBJECT")))
			
			If upl.Form("page_id") <> nil And IsNumeric(upl.Form("page_id")) Then
				page_id = trim(upl.Form("page_id"))
			Else
				page_id = "NULL"
			End If			
			
			If IsNull(upl.Form("quest_id")) =  false And IsNumeric(trim(upl.Form("quest_id"))) Then
				quest_id = trim(upl.Form("quest_id"))
			ElseIf IsNull(page_id) = false And IsNumeric(page_id) Then
				sqlstr = "Select Product_Id FROM Pages WHERE Page_Id = " & page_id
				set rs_qst = con.getRecordSet(sqlstr)
				If not rs_qst.eof Then
					quest_id = trim(rs_qst(0))
					If trim(quest_id) = "" Or IsNull(quest_id) Then
						quest_id = "NULL"
					End If
				Else
					quest_id = "NULL"
				End If	
				set rs_qst = Nothing
			Else
				quest_id = "NULL"
			End If	
				
		If  trim(upl.UserFilename) <> ""  Then
				uploadsDirVar = Server.MapPath("../../../download/products/")					
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   				attachment=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
		   			
   				file_path=uploadsDirVar & "/" & attachment
				if fs.FileExists(file_path) then
					set a = fs.GetFile(file_path)
					a.delete			
				end if			
									
				upl.Form("attachment").SaveAs uploadsDirVar & "/" & attachment
				if Err <> 0 Then
					Response.Write("An error occurred when saving the file on the server.")			 
					set upl = Nothing
					Response.End
				end if	
			Else
				attachment = ""
			End If			
	
			sql="SET NOCOUNT ON; INSERT INTO PRODUCTS (ORGANIZATION_ID,USER_ID,PRODUCT_NUMBER,DATE_START,DATE_END,"&_
			"PAGE_ID,QUESTIONS_ID,FROM_MAIL,EMAIL_SUBJECT,FILE_ATTACHMENT,PRODUCT_TYPE) Values ("&_
			OrgID & "," & UserId & ",'111'," & datestart & "," & dateend & "," & page_id &_
			"," & quest_id & ",'','" & emailSubject & "','" & sFix(attachment) & "',3);" &_
			"SELECT @@IDENTITY AS NewID"
			set rs_tmp = con.getRecordSet(sql)
				prodId = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing
			
			Set upl = Nothing			
			
			sqlstr = "Select GroupID From Groups Where GroupID IN (" & groups_id & ")"
			Set rs_groups = con.getRecordSet(sqlstr)
			If not rs_groups.eof Then
				arr_groups = rs_groups.getRows()
			End If
			Set rs_groups = Nothing
			
			If isArray(arr_groups) Then
			For gg=0 To Ubound(arr_groups,2)
			
				send_groupId = trim(arr_groups(0,gg))
			
				sqlStr = "INSERT INTO PRODUCT_GROUPS (PRODUCT_ID,groupID,USER_ID,Organization_ID,"&_
				"DATE_SEND) VALUES ("& prodId &","& send_groupId &","& UserID & "," & OrgID &", getDate())"
				con.ExecuteQuery(sqlStr)
				
				sqlStr = "INSERT INTO PRODUCT_CLIENT (PRODUCT_ID, PEOPLE_ID, PEOPLE_NAME, " &_
				" PEOPLE_PHONE, PEOPLE_FAX, PEOPLE_CELL, PEOPLE_EMAIL, PEOPLE_OFFICE, " &_
				" PEOPLE_COMPANY, GROUPID, User_ID, ORGANIZATION_ID, CONTACT_ID, COMPANY_ID, DATE_SEND)" &_
				" Select " & prodId & ", PEOPLE_ID, PEOPLE_NAME, PEOPLE_PHONE, PEOPLE_FAX, PEOPLE_CELL, " &_
				" PEOPLE_EMAIL, PEOPLE_OFFICE, PEOPLE_COMPANY, GROUPID, " & UserID & "," & OrgID & "," &_
				" CONTACT_ID, COMPANY_ID, getDate()"&_
				" FROM PEOPLES WHERE ORGANIZATION_ID=" & OrgID & " AND GROUPID = " & send_groupId &_
				" AND PEOPLE_ID NOT IN (SELECT PEOPLE_ID FROM PRODUCT_CLIENT Where ORGANIZATION_ID=" & OrgID &_
				" AND PRODUCT_ID = "& prodId &")"
				
				'Response.Write sqlstr
				'Response.End 
				
				con.ExecuteQuery(sqlStr)
			
            Next
            End If
		End If					
	End If
%>
<table cellpadding=2 cellspacing=2 width="100%" border=0>
<tr><td height="40" nowrap></td></tr>
<tr><td align="center" class="td_subj" dir=ltr><font color="#736BA6" dir=rtl>תיעוד שליחת דואר הסתיים</td></tr>
<tr><td align="center" class="td_subj" dir=ltr><font color="#736BA6"><b><%=totalCount%></b></font> מספר נמענים </td></tr>
<tr><td height="15" nowrap></td></tr>
<tr><td nowrap align="center">
<INPUT class="but_menu" style="width:90"  type="button" onclick="javascript:window.close()" value="סגור חלון">
</td></tr>
<tr><td height="15" nowrap></td></tr>
</table>
</body>
</html>
<%Set con = nothing%>