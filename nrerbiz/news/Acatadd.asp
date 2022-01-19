<!--#INCLUDE FILE="../../include/reverse.asp"-->
<!--#include file="../../include/connect.asp"-->
<!--#INCLUDE FILE="../checkAuWorker.asp"-->

<html>
<head>
<title>Administration Site</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link rel="stylesheet" type="text/css" href="../admin_style.css">

<script language="JavaScript">
     function CheckFields()
		{
	var doc=window.document;
	  if (	doc.forms[0].elements[0].value=='')
		   
		    	{
				alert('חובה למלא את השדה');
				return false;
				}
				else
				return true;
		}
</script>
</head>

<body class="body_admin">
<div align="right">

<%
catid=request("catid")
if catid<>nil and catid<>"" then
	set cc=con.execute("SELECT Category_Name FROM News_Categories WHERE Category_ID=" &catid)
	catname=cc("Category_Name")
	cc.close
    set cc=Nothing
end if
%>

<table border="0" width="100%" align="center"  cellspacing="0" cellpadding="0">
  <tr><td class="a_title_big" width="100%" dir=rtl><font size="4" color="#ffffff"><strong>ניהול אירועים</strong></font></td></tr>
  <tr>
    <td>
       <table class="table_admin_1" border="0" width="100%" align="center"  cellspacing="0" cellpadding="0">
         <tr>
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="adminCat.asp">חזרה לרשימת קטגוריות</a></td>
             <td class="td_admin_5" align="right" valign="middle" width="100%"><font style="font-size:12pt;"><b></b></font></td>
         </tr>
       </table>
    </td>
  </tr>
  <tr><td height="5"></td></tr>
  <tr>
    <td class="td_admin_4">
       <table border="0" width="100%" align="center"  cellspacing="1" cellpadding="2">
         <tr>
             <td class="td_title_2" align="right" width="100%"><font style="font-size:11pt;"><b><%if catid<>"" then%>עדכון קטגוריה<%else%>הוספת קטגוריה חדשה<%end if%>&nbsp;</b></font></td>
         </tr>
       </table>
    </td>
  </tr>
</table>

<table align=center border="0" cellpadding="2" cellspacing="1" width="100%">
<form method="post" action="adminCat.asp?catId=<%=catId%>">
<tr>
   <td align=right class="td_admin_4"><input dir="rtl" type="text" name="catname" value="<%=vFix(catname)%>" style="font-family:Arial" size="90">&nbsp;</td>
   <td align=center class="td_admin_4" nowrap><strong>&nbsp;&nbsp;שם הקטגוריה&nbsp;</strong></td>
</tr>
<tr><td colspan="2" class="td_admin_4" height="15">&nbsp;</td></tr>
<tr><td class="td_admin_5" colspan="2" align="center"><input type=submit value=" עדכון " onClick="return CheckFields()" style="font-family:Arial"></td></tr>
<tr><td colspan="2" class="td_line_between"></td></tr>
<tr><td colspan="2" height="15"></td></tr>
</form>
</table>
</div>
</body> 
</html>
<%
con.close 
set con=Nothing
%>
