<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--INCLUDE FILE="../checkWorker.asp"-->

<%
UserId = trim(Request.Cookies("bizpegasus")("UserId"))
UserName=trim(Request.Cookies("bizpegasus")("UserName"))
sqlStr = "UPDATE USERS set  Add_Insurance = 0 WHERE USER_ID=" & USERID
'Response.Write sqlStr
					'Response.End
					con.GetRecordSet (sqlStr)	
	sqlstring = "SELECT isNULL(Add_Insurance, 0) as Add_Insurance FROM USERS WHERE USER_ID=" & USERID
	'Response.Write sqlstring
	'Response.End
	set worker=con.GetRecordSet(sqlstring)
	if not worker.EOF  then	
'response.Write trim(worker("Add_Insurance"))
'response.end
Response.Cookies("bizpegasus")("AddInsurance") = trim(worker("Add_Insurance"))
end if
'
		strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
			"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=right bgcolor=""#e6e6e6"">" & vbCrLf &_
	"<tr><td align=right class=page_title style=""background-color:#FF0000"" dir=rtl>מנהל יקר,</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_
     "<td align=right class=page_title style=""background-color:#FF0000"" dir=ltr>"& UserName &" שים לב שהנציג</td>" & vbCrLf &_
            "</tr><tr><td align=right class=page_title style=""background-color:#FF0000"" dir=ltr>אינו מוכן לעבוד לפי נוהל העבודה של הביטוחים ולכן הרשאת הביטוח נסגרה לו</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_   
   "<td align=right class=page_title style=""background-color:#FF0000"" dir=rtl> שים לב שעניין זה חמור ויש צורך להסביר לנציג שהוא חייב לעבוד לפי נהלי פגסוס</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_   
  "<td  align=right class=page_title style=""background-color:#FF0000"" dir=rtl>הקפידו בבקשה על רענון מול הנציג ואם צריך מול כל המחלקה</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_   

   "<td align=right class=page_title style=""background-color:#FF0000"" dir=rtl><BR>,תודה<BR> סמנכ``ל פגסוס <BR></td>" & vbCrLf &_
            "</tr></table></body></html>"
 '  response.Write strBody
  ' response.end
 

				Dim Msg
		Set Msg = Server.CreateObject("CDO.Message")
			Msg.BodyPart.Charset = "windows-1255"
			Msg.From ="doNotReplay@pegasusisrael.co.il" 'fromEmail
			Msg.MimeFormatted = true
			Msg.MimeFormatted = true
			strSub = "הרשאת הביטוח"
				
		
			Msg.Subject = strSub				
			Msg.To ="sales@pegasusisrael.co.il"			
			Msg.HTMLBody = strBody			
			Msg.Send()						
		Set Msg = Nothing
	 

					%>
					
	<SCRIPT LANGUAGE=javascript>
	<!--	
	opener.focus();
		opener.window.location.reload(true);
		self.close();
	-->
	</SCRIPT>				