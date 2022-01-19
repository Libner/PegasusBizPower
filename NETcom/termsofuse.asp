<!-- #include file="reverse.asp" -->
<%If Request.Form.Count > 0 Then
		'For each obj In  Request.Form
	 	'	Response.Write obj & " - " & Request.Form(obj) & "<br>"
		'Next	
		username=Request.Form("username")
		password=Request.Form("password")
		 
		If trim(Request.Form("chk_agree")) <> "" Then
			sqlstr = "UPDATE USERS.IsApproved = 1"
			Server.Transfer "/NETcom/default.asp"
		End if
	 
	 End If%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="title_meta_inc.asp" -->
<meta name=vs_defaultClientScript content="JavaScript">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name=ProgId content=VisualStudio.HTML>
<meta name=Originator content="Microsoft Visual Studio .NET 7.1">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="IE4.css" rel="STYLESHEET" type="text/css">
<script type="text/javascript">
function checkFields(objForm)
{
	if(document.getElementById("chk_agree").checked == false)
	{
		window.alert("עליך להסכים לתנאי השימוש באתר");
		return false;
	}
}	
</script>	
</head>
<body onload="self.focus();">
<form id="form1" name="form1" action="termsofuse.asp" method="post">
<table cellpadding="0" cellspacing="0" width="100%" border="0">
<tr><td align="center">
<input type="hidden" id="username" name="username" value="<%=vFix(username)%>">
<input type="hidden" id="password" name="password" value="<%=vFix(password)%>">
<TABLE id=TermsOfUse borderColor=#0000ff cellSpacing=0 cellPadding=0 width=580 
border=0>
<TBODY>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right>
<P><FONT face=arial,helvetica,sans-serif size=2><STRONG>תנאי 
שימוש</STRONG></FONT></P>
<P><FONT face=arial,helvetica,sans-serif size=2>בהצטרפותך למקבלי השירותים באתר 
זה (להלן "המזמין") הנך מצהיר/ה בזאת שאתה/את מקבל/ת על עצמך את תנאי השימוש 
בשירותים אלה כפי שמפורט להלן ומתחייב לשלם לסייברסרב את תמורת השירותים הללו, 
בהתאם לפרטי נספח הההזמנה, הכוללים את המחירים, דרך החישוב ותנאי התשלום : 
</FONT></P></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>1.&nbsp;הנהלת האתר אינה 
אחראית לתכנים המועברים באמצעותו או/ו לתוצאות שעשויות לנבוע לצד שלישי כל שהוא 
בגין תכנים אלה. </FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>2. המשתמשים במנגנוני 
האתר מתחייבים להימנע מהפצת דואר זבל (spam)&nbsp;שמשמעותו כי &nbsp;הנמענים למשלוח 
הדואר האלקטרוני לא הביעו הסכמתם לקבלתו, ולכל נמען הזכות לבקש להסירו מרשימת 
התפוצה<BR>(unsubscribe)</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>3.המזמין מצהיר ומאשר כי 
הוא מודע ומכיר את הוראות חוק התקשורת (בזק ושידורים) (תיקון מס' 40) התשס"ח - 2008 
(להלן" "החוק") אשר כולל בין השאר איסור על שיגור דבר פרסומת בהודעה אלקטרונית ללא 
הסכמת הנמען, ומתחייב לפעול במשלוח הדואר האלקטרוני שישלח לנמענים באמצעות שירותי 
סייברסרב בהתאם ובכפוף להוראות החוק.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>4.יודגש ויובהר כי החוק 
נכנס לתוקף מלא ומחייב מיום 1.12.08 והמזמין ער ומודע להוראותיו והכללים שנקבעו בו 
לרבות האחריות האישית של נושאי משרה במזמין, והיות הפרת הוראות החוק עוולה אזרחית 
ועבירה פלילית גם יחד על כל המשתמע מכך .</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>5. התכנים בדואר לא 
יחרגו מהוראות החוקים הרלוונטיים במדינת ישראל ובכלל זה הוצאת דיבה וגרימת נזק לצד 
שלישי כלשהו או לפגוע בזכויות קנייניות של מאן דהו. </FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>6. אין להשתמש בדואר 
אלקטרוני הכולל קישורים לאתרים לא חוקיים , לאתרים הכוללים חומר פורנוגראפי או כאלה 
הפוגעים בזכויות ילדים כולל אתרים המוגדרים "למבוגרים בלבד". </FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right>
<P><FONT face=arial,helvetica,sans-serif size=2>7. המזמין מתחייב לשפות ולפצות את 
סייברסרב באורח מלא ומיידי על כל נזק ו/או הוצאה ו/או עלות ו/או הפסדים שיגרמו 
לסייברסרב (אם יגרמו) בגין הפרת הוראות החוק בדואר האלקטרוני שישלח לנמענים באמצעות 
שירותי סייברסרב. <BR></FONT><FONT face=arial,helvetica,sans-serif size=2>מבלי 
לגרוע מהתוקף המלא ומכלליות התחייבות זו, המזמין מתחייב לשפות ולפצות את סייברסרב 
בנסיבות אלה, גם על כל ההוצאות הנלוות לרבות הוצאות הטיפול המשפטי בו תשתמש 
סייברסרב במקרה שינקטו נגדה הליכים בגין דואר אלקטרוני שישלח לנמענים כאמור. 
</FONT></P></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>8. הנהלת האתר עובדיו 
וכל הפועלים מטעמו לא ישאו באחריות כלשהיא לשיבושים או תקלות במשלוח הדואר 
האלקטרוני עקב נסיבות שאינן תלויות בהם כגון כוח עליון, אש , מלחמה, פעולות איבה או 
כל הקשור בספקי התקשורת ו/או האינטרנט.</FONT></TD></TR>
<TR height=15>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>9. אחריות הנהלת האתר 
לכל תקלה או עיכוב במשלוח הדואר האלקטרוני תהיה מוגבלת לסכום שאותו שילם הלקוח עבור 
משלוח הדואר האלקטרוני האמור.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>10. מקום השיפוט יהיה 
בית המשפט השלום בקרית-ביאליק.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>11. בכל מקרה של הפרת 
סעיף מהסעיפים הרשומים לעיל, הזכות בידי הנהלת האתר להפסיק את השרות לאלתר מבלי 
שתחול עליה החובה להחזרת סכום כלשהו ששולם עבור שירותים אלה.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>12.אין באמור בתנאי 
שימוש אלה, כדי לגרוע מחובת המזמין לקיים במסגרת פעילותו במשלוח דואר אלקטרוני 
באמצעות שירותי סייברסרב כאמור, את כל הוראות הדין ו/או הוראות תקנון השימוש באתר 
סייברסרב (ו/או בכל אתר שתספק סייברסרב לצרכי משלוח הדואר 
האלקטרוני).</FONT></TD></TR></TBODY>
</td></tr>
<tr><td height="10" nowrap></td></tr>
<tr><td align="right" dir="rtl"><input type="submit" id="chk_agree" name="chk_agree"
value="אני מסכים לתנאי שימוש באתר" class="add">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" value="ביטול" onclick="document.location.href='../default.asp';" class="add">
</td></tr>
<tr><td height="10" nowrap></td></tr>
</table>
</form>
</body>
</html>