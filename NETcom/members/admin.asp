<html>
<!--#include file="../connect.asp"-->
<!--#include file="../reverse.asp"-->
<!--#include file="checkWorker.asp"-->
<%set con = nothing%>
<head>
<title>BizPower</title>
<meta charset="windows-1255">
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bgcolor="#E5E5E5">
<div align="center"><center>
<table border="0" width="100%" bgcolor="#FFFFFF" height="70" cellspacing="0" cellpadding="0" ID="Table1">
  <tr>
    <td><img src="../../images/leumi_top.jpg"></td>
  </tr>
</table>
</center></div><div align="center"><center>

<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#060165">
  <%numOftab = 6%>
  <tr><td width=100% align="right"><!--#include file="../top.asp"--></td></tr>
</table>
</center></div>
<table border="0" width="100%" bgcolor="#FFFFFF" height="300" cellspacing="0" cellpadding="0" background="../images/bgr_list.jpg">
   <tr>
    <td width="118" valign="top"><img src="../../images/pict1_netcom.jpg" width="118" height="301"></td>
    <td valign="top" width="202"><img src="../../images/pict2_netcom.jpg" height="301"></td>
    <td width="100%" valign="top" align="right">
     <table border="0" width="100%" cellspacing="0" cellpadding="0">
<!-- MAIN -->    
      <tr><td width="100%" height="15"></td></tr>
      <tr><td width="100%">
        <table width=93% align="center" border="0" cellspacing="0" cellpadding="2">
           <tr>
              <td align="right" width="100%">.<font style="color:#003497;"><b><u>1000</u></b></font>&nbsp;-&nbsp;<b>יתרת מיילים לשליחה</b></td>      
           </tr>
           <tr><td width="100%" height="5"></td></tr>
           <tr>
               <td align="right" bgcolor="#EBEBEB" dir="rtl"><font style="color:#000000;font-size:11pt;"><b>הוראות לבניה ושליחת דואר מעוצב</b>(כולל משוב).</font>&nbsp;</td>      
           </tr>
           <tr><td width="100%" height="5"></td></tr>
			<tr>
			<td width="100%" valign="top" align="right">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">שלב 1. </span><a href="pages/default.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">דואר מעוצב</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">בנו מסמך שווקי</span>, באופן עצמאי או
מתוך מאגר תבניות מוכן מראש כולל שילוב לוגו החברה, תמונות וטקסטים שלכם.</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">שלב 2. </span><a href="products/questions.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">טפסי משוב</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">בנו טופס משוב</span> תוך הגדרת השדות
הרלוונטים (רישום למבצע, פרטי מידע, סקר וכיו&quot;ב).</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">שלב 3. </span><a href="companies/choose.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">מועדון לקוחות</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">הגדירו את לקוחות</span> שלכם
והזינו כתובות דוא&quot;ל למשלוח באופן ידני או בטעינה מקבצי <span dir=ltr>EXCEL</span></td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">שלב 4. </span><a href="products/products.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">הפצה לנמענים</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">בצעו הפצת</span> מסמך שווקי עם או בלי
טופס משוב ללקוחות מוגדרות.</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">שלב 5. </span><a href="admin_DB/default.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">משובים</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">תצפו במשובים</span> (תשובות) של נמעני ההפצה
שנקלטו במערכת.</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">שלב 6. </span><a href="reports/reports.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">דוחות</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">דוחות התפלגות המשובים</span> של נמעני
ההפצה, דוחות סטטיסטיים מוכנים מראש.</td></tr>
           <tr><td width="100%" height="5"></td></tr>
           <tr><td width="100%" align="right" dir="rtl">
נשמח לעמוד לרשותכם ולסייע לכם
בשימוש ב- <span dir=ltr>Netcommunicator</span></td></tr>
           <tr><td width="100%" align="right" dir="rtl">
<b>מרכז שירות הלקוחות</b> שלנו לרשותכם
בשעות העבודה בטל' <span dir=ltr><b>04-8770282</b></span>. בעת הפנייה, אנא הכינו פרטיכם, פרטי ארגונכם ובמידה
שיש לכם שאלות טכניות &#8211; העלו את המסך הרלוונטי של המערכת על מחשבכם.
           </td></tr>
           <tr><td width="100%" height="10"></td></tr>
			</table>
			</td>
			</tr>  
        </table>
       </td>
      </tr>  
<!-- END MAIN -->        
     </table>
    </td>
    <td width="17" valign="top" align="right"></td>
  </tr>
</table>
</center></div>
</body>
</html>
