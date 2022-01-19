<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	sqlstr = "Select Ip From Ip_Address"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	Ip_Address=rstitle("Ip")
	End If
	set rstitle = Nothing %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

</head>
<body>
<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>" ID="Table1">
<tr><td width="100%">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 4%>
<%numOfLink = 11%>
<%topLevel2 = 70 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width="100%" border="0" ID="Table2">
<tr>
<td width="100%" valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="center" border="0" bgcolor="#ffffff" ID="Table3">
<tr><td height=50 nowrap></td></tr>

<tr>
	<td align="center"nowrap><span style="font-size:20px;"><%=Ip_Address%></span></td>	
</tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>

 <tr><td align="center" colspan=2><a class="button_edit_1" style="width: 105px;"   href="javascript:void window.open('update_IP.asp', 'winUpdIP' , 'scrollbars=1,toolbar=0,top=150,left=50,width=600,height=200,align=center,resizable=1');">IP-עדכון כתובת ה</a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>


</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu" height="550">
<table cellpadding=1 cellspacing=1 width="100%" ID="Table4" border=0>
<tr><td align="<%=align_var%>" colspan=2 height="550" nowrap></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</td></tr></table>
</BODY>
</HTML>
