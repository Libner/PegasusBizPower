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
		window.alert("���� ������ ����� ������ ����");
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
<P><FONT face=arial,helvetica,sans-serif size=2><STRONG>���� 
�����</STRONG></FONT></P>
<P><FONT face=arial,helvetica,sans-serif size=2>��������� ������ �������� ���� 
�� (���� "������") ��� �����/� ���� ����/�� ����/� �� ���� �� ���� ������ 
�������� ��� ��� ������ ���� ������� ���� ��������� �� ����� �������� ����, 
����� ����� ���� �������, ������� �� �������, ��� ������ ����� ������ : 
</FONT></P></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>1.&nbsp;����� ���� ���� 
������ ������ �������� �������� ��/� ������� ������� ����� ��� ����� �� ���� 
���� ����� ���. </FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>2. �������� �������� 
���� �������� ������ ����� ���� ��� (spam)&nbsp;�������� �� &nbsp;������� ������ 
����� ��������� �� ����� ������ ������, ���� ���� ����� ���� ������ ������ 
������<BR>(unsubscribe)</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>3.������ ����� ����� �� 
��� ���� ����� �� ������ ��� ������� (��� ��������) (����� ��' 40) ����"� - 2008 
(����" "����") ��� ���� ��� ���� ����� �� ����� ��� ������ ������ ��������� ��� 
����� �����, ������� ����� ������ ����� ��������� ����� ������� ������� ������ 
�������� ����� ������ ������� ����.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>4.����� ������ �� ���� 
���� ����� ��� ������ ���� 1.12.08 ������� �� ����� ��������� ������� ������ �� 
����� ������� ������ �� ����� ���� ������, ����� ���� ������ ���� ����� ������ 
������ ������ �� ��� �� �� ������ ��� .</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>5. ������ ����� �� 
����� ������� ������ ���������� ������ ����� ����� �� ����� ���� ������ ��� ��� 
����� ����� �� ����� ������� �������� �� ��� ���. </FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>6. ��� ������ ����� 
�������� ����� ������� ������ �� ������ , ������ ������� ���� ���������� �� ���� 
������� ������� ����� ���� ����� �������� "�������� ����". </FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right>
<P><FONT face=arial,helvetica,sans-serif size=2>7. ������ ������ ����� ������ �� 
�������� ����� ��� ������ �� �� ��� �/�� ����� �/�� ���� �/�� ������ ������ 
��������� (�� �����) ���� ���� ������ ���� ����� ��������� ����� ������� ������� 
������ ��������. <BR></FONT><FONT face=arial,helvetica,sans-serif size=2>���� 
����� ������ ���� �������� �������� ��, ������ ������ ����� ������ �� �������� 
������� ���, �� �� �� ������� ������ ����� ������ ������ ������ �� ����� 
�������� ����� ������ ���� ������ ���� ���� �������� ����� ������� �����. 
</FONT></P></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>8. ����� ���� ������ 
��� ������� ����� �� ���� ������� ������ �������� �� ����� ������ ����� 
��������� ��� ������ ����� ������ ��� ���� ��� �����, �� , �����, ������ ���� �� 
�� ����� ����� ������� �/�� ��������.</FONT></TD></TR>
<TR height=15>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>9. ������ ����� ���� 
��� ���� �� ����� ������ ����� ��������� ���� ������ ����� ����� ���� ����� ���� 
����� ����� ��������� �����.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>10. ���� ������ ���� 
��� ����� ����� �����-������.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>11. ��� ���� �� ���� 
���� �������� ������� ����, ����� ���� ����� ���� ������ �� ����� ����� ���� 
����� ���� ����� ������ ���� ����� ����� ���� ������� ���.</FONT></TD></TR>
<TR height=5>
<TD><FONT face=arial,helvetica,sans-serif size=2></FONT></TD></TR>
<TR>
<TD dir=rtl style="FONT-SIZE: 11pt; COLOR: #666666; FONT-FAMILY: Arial" 
align=right><FONT face=arial,helvetica,sans-serif size=2>12.��� ����� ����� 
����� ���, ��� ����� ����� ������ ����� ������ ������� ������ ���� �������� 
������� ������ �������� �����, �� �� ������ ���� �/�� ������ ����� ������ ���� 
�������� (�/�� ��� ��� ����� �������� ����� ����� ����� 
���������).</FONT></TD></TR></TBODY>
</td></tr>
<tr><td height="10" nowrap></td></tr>
<tr><td align="right" dir="rtl"><input type="submit" id="chk_agree" name="chk_agree"
value="��� ����� ����� ����� ����" class="add">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" value="�����" onclick="document.location.href='../default.asp';" class="add">
</td></tr>
<tr><td height="10" nowrap></td></tr>
</table>
</form>
</body>
</html>