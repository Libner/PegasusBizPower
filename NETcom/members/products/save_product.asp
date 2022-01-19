<%@ Language=VBScript%>
<%Response.Buffer = False%>
<%       Function RandomPW(ByVal myLength)
            'These constant are the minimum and maximum length for random
            'length passwords.  Adjust these values to your needs.
            Const minLength = 6
            Const maxLength = 20

            Dim X, Y, strPW

            If myLength = 0 Then
                Randomize()
                myLength = Int((maxLength * Rnd()) + minLength)
            End If


            For X = 1 To myLength
                'Randomize the type of this character
                Y = Int((3 * Rnd()) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase

                Select Case Y
                    Case 1
                        'Numeric character
                        Randomize()
                        strPW = strPW & Chr(Int((9 * Rnd()) + 48))
                    Case 2
                        'Uppercase character
                        Randomize()
                        strPW = strPW & Chr(Int((25 * Rnd()) + 65))
                    Case 3
                        'Lowercase character
                        Randomize()
                        strPW = strPW & Chr(Int((25 * Rnd()) + 97))

                End Select
            Next

            RandomPW = strPW

        End Function

%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<BODY onload="window.focus();" style="margin:0px">
<% UserID = trim(Request.Cookies("bizpegasus")("UserID"))
   OrgID = trim(trim(Request.Cookies("bizpegasus")("ORGID")))
   set upl = Server.CreateObject("SoftArtisans.FileUp")
   prodId = trim(upl.Form("prodId"))
   groups_id = trim(upl.Form("groups_id"))	
   
   If (Not IsNumeric(prodId) Or trim(prodId) = "" Or IsNULL(prodId)) And isNumeric(OrgID) Then
		If Len(groups_id) > 0 Then
			sqlStr = "Select Count(PEOPLE_ID) FROM PEOPLES where ORGANIZATION_ID="& OrgID &_
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
		badCount = 0
		
	 ElseIf IsNumeric(OrgID) And IsNumeric(prodID) And Len(groups_id) > 0 Then
		sqlStr = "Select Count(PEOPLE_ID) FROM PRODUCT_CLIENT where ORGANIZATION_ID="& OrgID &_
		" AND GROUPID IN ("& groups_id &") AND PRODUCT_ID = " & prodID & " And DATE_SEND is NULL"
		set rsc = con.getRecordSet(sqlStr)
		If not rsc.eof Then
			totalCount = trim(rsc(0))
		Else
			totalCount = 0	
		End If
		set rsc = Nothing		
		
		sqlStr = "Select Count(PEOPLE_ID) FROM PRODUCT_CLIENT WHERE  ORGANIZATION_ID="& OrgID &_ 
		" AND PRODUCT_ID = " & prodID & " AND DATE_SEND is NULL AND IsNULL(IsEmailValid, 1) = 0"
		set rsc = con.getRecordSet(sqlStr)
		If not rsc.eof Then
			badCount = trim(rsc(0))
		Else
			badCount = 0	
		End If
		set rsc = Nothing			
   End If	
	  
	If IsNumeric(OrgID) And Len(groups_id) > 0 Then
			sqlStr = "Select Count(DISTINCT blocked_email) "&_ 
			" FROM blocked_emails where ORGANIZATION_ID="& OrgID &_ 
			" AND blocked_email IN (Select PEOPLE_EMAIL From PRODUCT_CLIENT WHERE GROUPID IN ("& groups_id &"))"
			set rsc = con.getRecordSet(sqlStr)
			If not rsc.eof Then
				blockedCount = trim(rsc(0))
			Else
				blockedCount = 0	
			End If
			set rsc = Nothing		 
			
			sql = "Select Email_Limit From Organizations WHERE ORGANIZATION_ID = " & OrgID
			set rs_org = con.getRecordSet(sql)	
			If not rs_org.eof Then
				emailLimit = trim(rs_org(0))
			End If
			set rs_org = Nothing	
				
			If isNumeric(emailLimit) Then
				emailLimit = cLng(emailLimit)
			End If	
	Else
			blockedCount = 0 : badCount = 0
    End If	
	
  If ((totalCount - blockedCount - badCount)>0) And ((emailLimit - (totalCount - blockedCount - badCount))=>0) Then	
   If Not IsNumeric(prodId) Or Len(prodId) = 0 Or IsNULL(prodId) then 	
	If trim(upl.Form("editFlag")) <> "" Then    
	         
			If (upl.TotalBytes > 102400) Then
			%>
			<div style="background-color:#D4D0C8;width:100%;height:100%;color:red">
			<br><br>
			<p align=center>You can not attach a file larger then 100 KB</p> 
			<br>
			<p align=center><input type=button value="Ok" onclick="window.close();" style="width:60"></p> 
			</div>			
			<%
			Set upl = Nothing
			Response.End
			End If 
	  
   			If trim(upl.Form("attachment")) <> "" Then   		
   				attachment=trim(upl.Form("attachment"))    		
   			Else
   				attachment=""	
   			End if   	   	
	 
			dateend = "NULL"		
			datestart = " getDate()"
			
			fromMail = sFix(trim(upl.Form("FROM_MAIL")))
			fromName = sFix(trim(upl.Form("FROM_NAME")))			
			emailSubject = sFix(trim(upl.Form("EMAIL_SUBJECT")))
			
			If trim(upl.Form("page_id")) <> "" And IsNumeric(upl.Form("page_id")) Then
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
			
			If IsNumeric(page_id) Then	
			
			strPassword = RandomPW(5)
			
			sql="SET NOCOUNT ON; INSERT INTO PRODUCTS (ORGANIZATION_ID,USER_ID,PRODUCT_NUMBER,DATE_START," & _
			"DATE_END,PAGE_ID,QUESTIONS_ID,FROM_MAIL,FROM_NAME,EMAIL_SUBJECT,FILE_ATTACHMENT,Send_Password) " & _
			" Values (" & OrgID & "," & UserId & ",'111'," & datestart & "," & dateend & "," & page_id &_
			"," & quest_id & ",'" & fromMail & "','" & fromName & "','" & emailSubject & "','" &_
			sFix(attachment) & "','" & sFix(strPassword) & "');SELECT @@IDENTITY AS NewID"
			set rs_tmp = con.getRecordSet(sql)
				prodId = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing
			
			sqlstr = "Exec dbo.Insert_Peoples '" & prodId & "','" & OrgID & "','" & UserId & "','" & groups_id & "'"
			con.ExecuteQuery(sqlstr)		
			
			Else
			%>
			<div style="background-color:#D4D0C8;width:100%;height:100%;color:red">
			<br><br>
			<p align=center>לא נמצא דף להפצה</p> 
			<br>
			<p align=center><input type=button value="Ok" onclick="window.close();" style="width:60"></p> 
			</div>			
			<%
			Set upl = Nothing
			Response.End			
			End If				
					
		End If					
	End If

	If isNumeric(prodID) And trim(strPassword) = "" Then
		sqlstr = "Select FROM_MAIL, Send_Password, Page_Id, EMAIL_SUBJECT From Products Where PRODUCT_ID = " & prodID
		set rs_p = con.getRecordSet(sqlstr)
		If not rs_p.eof Then
			strPassword = trim(rs_p("Send_Password"))
			fromMail = trim(rs_p("FROM_MAIL"))
			page_id = trim(rs_p("Page_Id"))
			emailSubject = trim(rs_p("EMAIL_SUBJECT"))
		Else
			strPassword = ""	
		End if
		set rs_p = Nothing
	End If

	If isNumeric(page_id) And Len(page_id) > 0 And IsNULL(page_id) = false Then
		sqlstr = "Select Page_Title From Pages Where Page_Id = " & page_id
		set rs_p = con.getRecordSet(sqlstr)
		If not rs_p.eof Then
			pageTitle = trim(rs_p(0))
		Else
			pageTitle = ""	
		End if
		set rs_p = Nothing
	End If					

	SendPassword(strPassword)
End If 
Sub SendPassword(strPassword)
			strBody = ""
			strBody = strBody & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//HE" & vbCrlf
            strBody = strBody & "<html><HEAD><title>Bizpower</title><meta charset=windows-1255>" & vbCrlf
            strBody = strBody & "<style>td{	font-size:13px;	font-family:Arial; color:Black } .title{font-size:15px;FONT-WEIGHT: bold;font-family:Arial; color:Black} a { font-size=13px; font-family=Arial; color=#FF0066 }</style></HEAD>" & vbCrlf
            strBody = strBody & "<body  marginheight=""10"" topmargin=""10"" leftmargin=""10"" bottommargin=""10""" & vbCrlf
            strBody = strBody & " rightmargin=""10"" >" & vbCrlf
            strBody = strBody & "<table width=600 border=0 cellpadding=2 cellspacing=2 align=center>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 height=10 valign=top></td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top class='title'>גולש/ת יקר/ה,</td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top>סיסמת אישור הפצה שלך היא : <b>" & strPassword & "</b></td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 height=10 valign=top></td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top>פרטי ההפצה:</td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top>נושא: " & emailSubject & "</td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top>דף מופץ: " & pageTitle & "</td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 height=10 valign=top></td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top>בברכה, צוות האתר</td></tr>" & vbCrlf
            strBody = strBody & "<tr><td colspan=2 align=right dir=rtl valign=top><a href='http://pegasus.bizpower.co.il'>http://pegasus.bizpower.co.il</a></td></tr>" & vbCrlf
            strBody = strBody & "</table>" & vbCrlf
            strBody = strBody & "</Body>" & vbCrlf
            strBody = strBody & "</html>"
			
			Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
				Msg.From = "support@cyberserve.co.il"
				Msg.MimeFormatted = true
				Msg.Subject = "BIZPOWER - סיסמת אישור הפצה"
				Msg.To = fromMail		
				Msg.HTMLBody = strBody	
				Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"		
				Msg.Send()						
			Set Msg = Nothing			
End Sub
  
%>
<script language=javascript>
<!--
	 function GoSend(btn)
	 {	
		var strPassword = document.getElementById("strPassword").value;
		if(document.getElementById("mail_password"))
		{
			var valPassword = document.getElementById("mail_password").value;
			if(valPassword == "")
			{
				window.alert("אנא הזן סיסמת אישור הפצה");
				document.getElementById("mail_password").focus();
				return false;
			}
			if(strPassword != valPassword)
			{
				window.alert("סיסמת ההפצה אשר הזנת אינה נכונה");
				document.getElementById("mail_password").focus();
				return false;	   
			}
		}	
		btn.disabled = true;
		strUrl = "sendMailQ.aspx?send_groupId=<%=groups_id%>&prodId=<%=prodId%>&pageId=<%=page_id%>";		
		document.location.href = strUrl;
		window.opener.document.location.href = "products.asp";
	 }
//-->
</script>
<table cellpadding=2 cellspacing=2 width="100%" border=0>
<tr><td height="10" nowrap></td></tr>
<tr><td align="center" class="td_subj" dir=ltr><font color="#023296"><b><%=totalCount%></b></font> מיילים לשליחה בקבוצות שנבחרו </td></tr>
<tr><td align="center" class="td_subj" dir=ltr>מתוכם נמצאו <font color="#023296"><b><%=blockedCount%></b></font> מיילים חסומים</td></tr>
<%if false then%>
<tr><td align="center" class="td_subj" dir=ltr>מתוכם נמצאו <font color="#023296"><b><%=badCount%></b></font> מיילים לא תקניים</td></tr>
<%end if%>
<tr><td align="center" class="td_subj" dir=ltr>סה"כ מיילים לשליחה <font color="#023296"><b><%=totalCount-blockedCount - badCount%></b></font></td></tr>
<tr><td align="center" class="td_subj" dir=ltr><font color="#023296" dir=ltr><b><%=emailLimit - (totalCount - blockedCount - badCount)%></b></font> יתרה לאחר השליחה</td></tr>
<tr><td height="10" nowrap></td></tr>
<%If ((totalCount - blockedCount - badCount)>0) And ((emailLimit - (totalCount - blockedCount - badCount))=>0) Then%>
<tr><td dir="rtl" align="center">ברגעים אלו נשלחה אליך סיסמת הפצה לכתובת האימייל אשר ציינת.</td></tr>
<tr><td dir="rtl" align="center">לאישור ההפצה אנא הקלד את הסיסמה בתיבה זו: <input type="text" maxlength="10" id="mail_password" id="mail_password" NAME="mail_password"></td></tr>
<tr><td height="5" nowrap><input type="hidden" id="strPassword" name="strPassword" value="<%=vFix(strPassword)%>"></td></tr>
<tr><td nowrap align="center">
<INPUT class="but_menu" style="width:90" type="button" onclick="GoSend(this)" value="שלח הפצה">
</td></tr>
<%Else%>
<tr><td align="center" class="td_subj" dir=ltr><font color=red><b>לא ניתן לבצע את השליחה</b></font></td></tr>
<%If ((emailLimit - totalCount - blockedCount - badCount)<=0) Then%>
<tr><td align="center" class="td_subj" dir=ltr><font color=red><b>04 - לקנית יתרה נוספת ניתן להתקשר ל 8770282</font></td></tr>
<%End if%>
<%End if%>
</table>
</body>
</html>