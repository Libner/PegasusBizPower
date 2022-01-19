<!--#include file="../../../include/connect.inc"-->
<!--#include file="../../connect.inc"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="VISUAL">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>

<%if check_worker() then%>

<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0>
<table border="0" width="780" bgcolor="#c0c7c7" cellspacing="0" cellpadding="0" align=center>
<tr>
<td width=100% bgcolor="#0F212B">
<!--#INCLUDE FILE="../logo/logo_top.asp"-->
</td></tr>
<tr><td bgcolor="#0F212B">
<% numOftab = 2 %>
<!--#INCLUDE FILE="../../top_in.asp"-->
</td></tr></table> 
<table border="0" width="100%"  cellspacing="1" cellpadding="0">
	<tr>
		<td align="center"colspan=2 width=100%>&nbsp;</td>
			</tr>
</table>

<table width="770" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff" >
	<tr>
		<td>
			<%'//start of %>
		</td>
		
		<td>
			
		</td>
	</tr>
	
</table>
</body>
<%end if 'if check_worker()
set con=nothing
%>
</html>
