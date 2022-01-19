<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--INCLUDE FILE="../checkWorker.asp"-->
<%OrgID=264
lang_id = 1
dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"

   contactsCount=0 ' total count of selected appeals
   sentCount=0 ' count of appeals for which the mail was sent

 apealsIsMarketingEmailSend="" ' appaels ID for which the mail was sent 

  'appID=request.querystring("appID")
  appealsId=	request.querystring("appealsId")
  quest_id=Request.QueryString("quest_id")
'  response.Write "quest_id="&quest_id

	
	sqlstr = "Select appeal_id,appeals.Company_Id,appeals.Contact_Id,email,Reason_Id from appeals left join Contacts on appeals.Contact_id=contacts.Contact_Id WHERE  len(email)>0 and appeal_id IN (" & appealsId & ")"
'response.Write sqlstr
'response.end
set rs_Contacts = con.getRecordSet(sqlstr)	
	do while not rs_Contacts.eof

    contactsCount=contactsCount+1

	  companyID=rs_Contacts("Company_Id")
	contactID=rs_Contacts("contact_id")
	email=rs_Contacts("email")
	appID=rs_Contacts("appeal_id")
    Reason_Id=rs_Contacts("Reason_Id")
	  ' שליחת מייל ללקוח 
	   strBody=""
    
                strBody=strBody & (" <!DOCTYPE html>")
                strBody=strBody & ("<html><head><title></title>" & vbCrLf)
                strBody=strBody & ("<style>")
                strBody=strBody & ("td.content_cl")
                strBody=strBody & ("{")
                strBody=strBody & (vbTab & "FONT-WEIGHT: normal;")
                strBody=strBody & (vbTab & "FONT-SIZE: 11pt;")
                strBody=strBody & (vbTab & "COLOR: #193B6E;")
                strBody=strBody & (vbTab & "FONT-FAMILY: Arial;")
                strBody=strBody & ("}")
                strBody=strBody & ("td.content_b")
                strBody=strBody & ("{")
                strBody=strBody & (vbTab & "FONT-WEIGHT: bold;")
                strBody=strBody & (vbTab & "FONT-SIZE: 11pt;")
                strBody=strBody & (vbTab & "COLOR: #193B6E;")
                strBody=strBody & (vbTab & "FONT-FAMILY: Arial;")
                strBody=strBody & ("}")
                strBody=strBody & ("td.content_Title")
                strBody=strBody & ("{")
                strBody=strBody & (vbTab & "FONT-WEIGHT: bold;")
                strBody=strBody & (vbTab & "FONT-SIZE: 18pt;")
                strBody=strBody & (vbTab & "COLOR: #193B6E;")
                strBody=strBody & (vbTab & "FONT-FAMILY: Arial;")
                strBody=strBody & ("}")
                strBody=strBody & (".button")
                strBody=strBody & ("{")
                strBody=strBody & (vbTab & "display:block;")
                strBody=strBody & (vbTab & "width:auto;")
                strBody=strBody & (vbTab & "height: auto;")
                strBody=strBody & (vbTab & "margin: 0px;")
                strBody=strBody & (vbTab & "padding: 10px;")
                strBody=strBody & (vbTab & "border-radius:   8px;")
                strBody=strBody & (vbTab & "moz-border - radius:   8px;")
                strBody=strBody & (vbTab & "khtml-border - radius:   8px;")
                strBody=strBody & (vbTab & "o-border - radius:   8px;")
                strBody=strBody & (vbTab & "webkit-border - radius:   8px;")
                strBody=strBody & (vbTab & "ms-border - radius:   8px;  ")
                strBody=strBody & (vbTab & "background-color:  #93bd2a;")
                strBody=strBody & (vbTab & "color: #ffffff;")
                strBody=strBody & (vbTab & "Font-Size:   20pt;")
                strBody=strBody & (vbTab & "Font-weight:  700;")
                strBody=strBody & (vbTab & "text-align:   center;")
                strBody=strBody & (vbTab & "text-decoration:none;")
                strBody=strBody & ("}")
                strBody=strBody & ("</style></head><body>" & vbCrLf)
                strBody=strBody & ("<table cellpadding=0 cellspacing=5 align='center' border='0' align=right>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_b' nowrap width='100%'></td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>מטיילים יקרים</td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>תודה על התעניינותכם,<bR>בהמשך לשיחתנו תמצאו מצורפים במייל זה את המסמכים הבאים:</td>")
                strBody=strBody & ("</tr>")
     strBody=strBody & ("<tr>")
                strBody=strBody & ("<td bgcolor='#f5f5f5' align='right'>")
                strBody=strBody & ("<table cellspacing='0' cellpadding='10'>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_Title' nowrap >")
                strBody=strBody & ("<table cellspacing='0' cellpadding='10'>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td bgcolor='#93bd2a'>")
                strBody=strBody & ("<a href='" & Application("SiteUrl") & "/tours/tourRegistrationGalorContact.aspx?appId=" & appID & "'  class='button'>לחץ כאן</a>")
                strBody=strBody & ("</td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("</table>")
                strBody=strBody & ("</td>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_Title' nowrap > קישור להרשמה אינטרנטית </td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("</table>")
                strBody=strBody & ("</td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'><span style='font-size:12pt;font-weight:bold;color:red'>הרישום לטיול זה דרך האינטרנט הוא אישי: ""אין להעבירו לאדם אחר כיוון שעלולות להיגרם תקלות טכניות ולפגום ברישום""</span><br></td>")
                strBody=strBody & ("</tr>")

      strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>אנו ממליצים בחום להירשם ולהבטיח מקום בהקדם.</td>")
                strBody=strBody & ("</tr>")

                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>שימו לב כי המחיר נשמר רק לאחר אישור ההרשמה המלווה בתשלום. <br>הירשמו מוקדם כיוון שבמידה ומחיר הטיול ירד לאחר רישום לטיול, מדיניות פגסוס היא תמיד לשמור על המטייל <BR>שנרשם מוקדם ונוריד את המחיר למחיר המעודכן.<br></td>")
                strBody=strBody & ("</tr>")

                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>אנו ממליצים להקיף בעיגול כי הנכם מעוניינים לשמוע הצעת ביטוח רפואי המכסה דמי ביטול מיד עם הרשמתכם.<br></td>")
                strBody=strBody & ("</tr>")
                 

                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>לפקס 03-5436060 או למייל <a href='mailto:pegasus@pegasusisrael.co.il'>pegasus@pegasusisrael.co.il</a></td>")
                strBody=strBody & ("</tr>")


                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_b' nowrap width='100%'></td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>להתחלת תהליך רישום   <a href='" & Application("SiteUrl") & "/tours/tourRegistrationGalorContact.aspx?appId=" & appID & "' >לחץ כאן</a><br><br></td>")
                strBody=strBody & ("</tr>")

                strBody=strBody & ("<tr>")
                strBody=strBody & ("<td align='right' dir='rtl' class='content_b' nowrap width='100%'></td>")
                strBody=strBody & ("</tr>")
                strBody=strBody & ("<tr><td align='right' class='content_cl'  valign=top>טיול מהנה ומפלא</td></tr>")


                    strBody=strBody & ("<tr><td align='right' class='content_cl'  valign=top>צוות מכירות פגסוס</td></tr>")

                strBody=strBody & ("</table></body></html>")



   	
			Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
				Msg.From ="pegasus@pegasusisrael.co.il"
				Msg.MimeFormatted = true
			
				Msg.Subject="אתר פגסוס - טופס רישום לטיול"
					
				
				Msg.To = email '"faina@cyberserve.co.il"
				Msg.Bcc = "mila@cyberserve.co.il"
				Msg.HTMLBody = strBody
				Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"		
    
				Msg.Send()	
    
    sentCount=sentCount+1
			Set Msg = Nothing
    if apealsIsMarketingEmailSend<>"" then
        apealsIsMarketingEmailSend=apealsIsMarketingEmailSend & ","
    end if
    apealsIsMarketingEmailSend=apealsIsMarketingEmailSend & appID

'===========================
'Add to marketing mailing 
sqlIsMailing="INSERT INTO dbo.MarketingMailing " &_
             "  (CONTACT_ID, USER_ID, APPEAL_ID, Country_ID, Tour_ID, MailingType_ID, Recepient_Email, Subject_Email, Content_Email, DATE_SEND, DATE_OPENED, IS_OPENED)"
             "  (" & CONTACTID, USER_ID, APPEAL_ID, Country_ID, Tour_ID, MailingType_ID, Recepient_Email, Subject_Email, Content_Email, DATE_SEND, DATE_OPENED, IS_OPENED)"
'===========================

			rs_Contacts.MoveNext
			loop
    rs_Contacts.close

if apealsIsMarketingEmailSend<>""
	'update IsMarketingEmailSend in appeals  for which the mail was sent 
    sqlstr = "update appeals set IsMarketingEmailSend=IsMarketingEmailSend+1 where appeal_Id in ("& apealsIsMarketingEmailSend &")"
	con.executeQuery(sqlstr) 
end if 
%>
<html>
	<head>
		<!-- include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</head>
<body style="margin: 0px; background-color: #E6E6E6" onload="window.focus();" onbeforeunload="window.opener.location.reload(true)">
    <table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>" id="Table1">
        <tr>
            <td align="left" valign="middle" nowrap>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" id="Table2">
                    <tr>
                        <td class="page_title" dir="<%=dir_obj_var%>">&nbsp;שליחת מייל רישום</td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td width="100%">
                <table align="center" border="0" cellpadding="3" cellspacing="1" width="98%" align="center"
                    id="Table3">
                    <tr>
                        <td class="form_title">נבחרו <%=contactsCount %> משתמשים</td>
                    </tr>
                    <tr>
                        <td  height="5"></td>
                    </tr>
                    <tr>
                        <td  class="form_title">נשלחו מיילים ל <%=contactsCount %> משתמשים</td>
                    </tr>
                    
                    <tr>
                        <td  height="5"></td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table cellpadding="0" cellspacing="0" align="center" width="90%" align="left" id="Table4">
                    <tr valign="top">
                        <td  align="center"><a class="but_menu" href="#" style="width: 100px" onclick="window.close();"><span id="word4" name="word4">סגור חלון</span></a></td>

                    </tr>
                    
                    <tr>
                        <td  height="5"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
<%
set con = nothing
%>
</html>
