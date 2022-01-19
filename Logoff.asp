<!--#include file="include/connect.asp"-->
<!--#include file="include/reverse.asp"-->
<%
' Flushes authentication information for the user and ends the session
lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))		
If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
	lang_id = 1 
End If
If lang_id = 2 Then
	dir_var = "rtl"
	align_var = "left"
	dir_obj_var = "ltr"
Else
	dir_var = "ltr"
	align_var = "right"
	dir_obj_var = "rtl"
End If	
	
For Each SessionItem in Session.Contents
   If InStr(SessionItem,"admin") = 0 Then
		Session.Contents.Remove(SessionItem)
   End If
Next

Response.Cookies("bizpegasus").Expires=Now() - 1
Response.Cookies("bizpegasus") = ""
%>
<html>
<head>
<title>Bizpower - כלים חזקים אונליין</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="10" marginheight="0" hspace="10" vspace="0" topmargin="0" leftmargin="10" rightmargin="10" bgcolor="white">
<!--#include file="include/top.asp"-->
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td height=50 nowrap></td></tr>
<tr>
<td class="titleB" align=center dir="<%=dir_obj_var%>">
<%If trim(lang_id) = "1" Then%>
המערכת מבצעת יציאה לאחר 4.5 שעות של הפסקת פעילות, להמשך הפעילות אנא בצע כניסה מחדש בעזרת שם משתמש וסיסמא
<%Else%>
Bizpower logged out due to 4.5 hours with no activities in your account, Refill your User Info.
<%End If%>
</td>
</tr>
</table>
</body>
</html>
