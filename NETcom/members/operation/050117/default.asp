<%@ LANGUAGE=VBScript %>
<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table1">
<tr><td width="100%" align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width="100%" align="<%=align_var%>">
  <%numOftab = 82%>
  <%numOfLink =0%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align=center>
<table border=0 cellpadding=2 cellspacing=0 align=center>
<tr>
<td><a href="default.asp?tab=8" class="button_edit_1" style="width:105;" >���� ��� �����</a></td>
<td width=10></td>
<td><a href="default.asp?tab=7" class="button_edit_1" style="width:105;" >����� �����</a></td>
<td width=10></td>

<td><a href="default.asp?tab=10" class="button_edit_1" style="width:115;" >����� ������ �����</a></td>
<td width=10></td>

<td><a href="default.asp?tab=5" class="button_edit_1" style="width:135;" >������� ����� ���� �����</a></td>
<td width=10></td>

<td><a href="default.asp?tab=4" class="button_edit_1" style="width:105;" >�� ���� ���� �����</a></td>
<td width=10></td>
<%if false then%>
<td><a href="default.asp?tab=10" class="button_edit_1" style="width:145;" >����� ������ ����� �����</a></td>
<td width=10></td>
<%end if%>
<td><a href="default.asp?tab=9" class="button_edit_1" style="width:165;" >�� ���� ����</a></td>
<td width=10></td>
<td><a href="default.asp?tab=11" class="button_edit_1" style="width:105;" >����� ������ ����</a></td>
<td width=10></td>


<td><a href="default.asp?tab=2" class="button_edit_1" style="width:105;" >������ ���� </a></td>
<td width=10></td>
<td><a href="default.asp?tab=3" class="button_edit_1" style="width:105;" >���"� ��� ������</a></td>
<td width=10></td>

<td><a href="default.asp?tab=1" class="button_edit_1" style="width:105;" >���� ���"�</a></td>
</tr>
</table><!--��� ���� �� �� ����� ����� �� �� �� ��� ���� ������ ����� �������--><%'=arrTitles(25)%></td></tr>
<tr>    
    <td width="100%" valign="top" align="center">
    <table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table2">
 <tr>
    <td align="left" width="100%" valign=top >   
    <iframe name="frameScrreen" id="frameScrreen"  src="WorkScreen.aspx?tab=<%=request.querystring("tab")%>"	ALLOWTRANSPARENCY=true height=2700  width="100%" marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="0" frameborder="0"></iframe>
    </td>
</tr></table>
	</td></tr>

	</table>

</body>
</html>
<%set con=Nothing%>