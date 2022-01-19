<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 56 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing	
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0"  ID="Table3">
<%numOftab = 3%>
<%numOfLink = 7%>
<%topLevel2 = 30 'current bar ID in top submenu - added 03/10/2019%>
  <tr><td width=100% align="right"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</center></div>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" background="../images/bgr_list.jpg">
<tr><td class="page_title" dir=rtl>&nbsp;</td></tr>

<tr><td height=20 nowrap></td></tr>
<tr>
   <td width="100%" align=center valign=top colspan=2>      
	<table border="0" cellspacing="1" width=450 align=center>	
	<TR>
		<td width=20%>&nbsp;</td>
		<TD align=center width=60%><A class="but_menu" href="default_user.asp" target=_self><span id=word1 name=word1><!--דו''ח אישי לתקופה--><%=arrTitles(1)%></span></a></TD>
		<td width=20%>&nbsp;</td>
	</TR>
	<TR>
		<td width=20%>&nbsp;</td>
		<TD align=center width=60%><A class="but_menu" href="default_client_user.asp" target=_self><span id="word2" name=word2><!--דו''ח  עלות לפי עובדים לתקופה--><%=arrTitles(2)%></span></a></TD>
		<td width=20%>&nbsp;</td>
	</TR>
	<TR>
		<td width=20%>&nbsp;</td>
		<TD align=center width=60%><A class="but_menu" href="default_client_job.asp" target=_self><span id="word3" name=word3><!--דו''ח  עלות לפי תפקידים לתקופה--><%=arrTitles(3)%></span></a></TD>
		<td width=20%>&nbsp;</td>
	</TR>																						
	</table>
  </td></tr>	
</td></tr></table>
</center></div>
</body>
<%set con=nothing%>
</html>
