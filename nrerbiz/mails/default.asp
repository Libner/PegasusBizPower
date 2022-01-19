<!--#include file="../../netcom/connect.asp"-->
<!--#include file="../../netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<%  Session.LCID = "1037"
    Response.CharSet = "windows-1255"%>
<body bgcolor="#FFFFFF">
<div align="right">
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="4" class="page_title" width=100%>רשימת הפצות</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" href="../../nrerbiz/choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="5%" align="center"></td> 
     <td width="*%" align="center"></td>      
  </tr>
</table>
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1">
	<tr bgcolor="#D7D7D7" height="22">
		<td class="11normalB" dir="rtl">מיילים<br>שנותרו</td>
		<td class="11normalB" dir="rtl">מיילים<br>שנשלחו</td>
		<td class="11normalB" dir="rtl">סיום שליחה</td>
		<td class="11normalB" dir="rtl">התחלת שליחה</td>
		<td class="11normalB" dir="rtl">נושא ההפצה</td>
		<td class="11normalB" dir="rtl">שם ארגון</td>
	</tr>
	<%sqlstr = "SET DATEFORMAT DMY;" & _
    " SELECT TOP 50 dbo.ORGANIZATIONS.ORGANIZATION_ID, dbo.ORGANIZATIONS.ORGANIZATION_NAME, " & _
    " dbo.PRODUCTS.PRODUCT_ID, dbo.PRODUCTS.EMAIL_SUBJECT, dbo.PRODUCTS.DATE_SEND_END, DATE_START, " & _
    " (Select Count(PEOPLE_ID) From PRODUCT_CLIENT WHERE PRODUCT_CLIENT.PRODUCT_ID = dbo.PRODUCTS.PRODUCT_ID AND DATE_SEND IS NULL) as count_not_send, " & _
    " (Select Count(PEOPLE_ID) From PRODUCT_CLIENT WHERE PRODUCT_CLIENT.PRODUCT_ID = dbo.PRODUCTS.PRODUCT_ID AND DATE_SEND IS NOT NULL) as count_send" & _
    " FROM dbo.ORGANIZATIONS INNER JOIN " & _
    " dbo.PRODUCTS ON dbo.ORGANIZATIONS.ORGANIZATION_ID = dbo.PRODUCTS.ORGANIZATION_ID " & _
	" WHERE (dbo.PRODUCTS.PRODUCT_NUMBER = '111') And (DATE_START Is not NULL) " & _
	" Order By DATE_START DESC"
	Set rs_pr = con.getRecordSet(sqlstr)
	While not rs_pr.Eof%>
	<tr bgcolor="#F5F5F5">
           <td class="10normal" dir="ltr" align="right"><%=rs_pr("count_not_send")%></td>
           <td class="10normal" dir="ltr" align="right"><%=rs_pr("count_send")%></td>
           <td dir="rtl" class="10normal"><%=rs_pr("DATE_SEND_END")%></td>
           <td dir="rtl" class="10normal"><%=rs_pr("DATE_START")%></td>
           <td dir="rtl" class="10normal"><%=rs_pr("EMAIL_SUBJECT")%></td>
           <td dir="rtl" class="10normal"><%=rs_pr("ORGANIZATION_NAME")%></td>
    </tr>		   
	<%rs_pr.moveNext
	Wend
	Set rs_pr = Nothing%>
</table>
</div>
</body>
</html>
<%set con = nothing%>