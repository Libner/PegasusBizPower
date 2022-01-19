<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head> 
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr><td width=100%>
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 6%>
<%numOfLink = 0%>
<tr><td width=100%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title" dir=rtl>&nbsp;&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>         
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center>
<tr>    
    <td width="100%" height=100> 
    </td>
</tr>    
   <tr>    
    <td width="100%" valign="top" align="center">    
	<a class="linkfaq" style="font-size:14px" href="../../../desktop/setup/setupBP.exe" target=_blank>
	<%If trim(lang_id) = "1" Then%>
	Bizpower לחץ להורדת תוכנת
	<%Else%>
	Download Bizpower Reminder 
	<%End If%>
	</a>
   </td>
</tr>
<tr><td height=10 nowrap></td></tr>
</table>
</td></tr></table>
</body>
</html>
<%set con = nothing%>