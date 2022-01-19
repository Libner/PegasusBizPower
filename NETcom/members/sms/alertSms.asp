<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
	msg=trim(Request.QueryString("msg"))
	tid=trim(Request.QueryString("tid"))
	info=trim(Request.QueryString("info"))
	sms_status = 1
	
	Select Case UCase(trim(msg))
		Case "ERROR" : error = msg :  sms_status = 1 	'- אירעה שגיאה בתהליך שליחת ההודעות.
		Case "START" : error = "" :  sms_status = 2  '- התחילה שליחת ההודעות.
		Case "END" : error = "" :  sms_status = 3  '- הסתיימה שליחת ההודעות, הפרמטר info יכיל את מספר ההודעות שנשלחו.
		Case "STOP" : error = "" :  sms_status = 3  '- בוצעה עצירה יזומה של שליחת ההודעות, הפרמטר info יכיל את מספר ההודעות שנשלחו.
		Case "NO_SMS_LEFT" : error = msg :  sms_status = 1  '- השליחה הופסקה היות ולא נותרו SMS בחבילה שנקנתה, התראה זו תופיע בשליחת הודעות רגילות בלבד.
		Case "NO_USERS_FOUND" : error = msg :  sms_status = 1  '- לא נמצאו משתמשים במאגר שניתן לשלוח להם הודעה (המאגר ריק, או כל המשתמשים הגיעו למקסימום ההודעות שניתן לקבל).
		Case "UNKNOWN_ERROR" : error = msg :  sms_status = 1   '- אירעה שגיאה לא מוכרת.
		Case "CP_SEND_ERROR" : error = msg :  sms_status = 1  '- אירעה שגיאה בשליחה לחלק או כל המספרים
	End Select
	
	If IsNumeric(tid) And trim(tid) <> "" Then	
	 sqlstr = "UPDATE sms_sends SET error = '" & vFix(error) & "', status = '" & vFix(sms_status) & "', info = '" & vFix(info) & "' WHERE tid =  " & tid
	 con.ExecuteQuery(sqlstr)
   End If	 
	 
   Set con = Nothing %>