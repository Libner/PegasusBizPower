<%@ Page Language="vb" AutoEventWireup="false" Codebehind="checkUser.aspx.vb" Inherits="bizpower_pegasus2018.checkUser"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>checkUser</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
			<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
			<script>
			function CheckFields()
			{
					if (document.getElementById("pass").value=='')
					{
					return false;
					}
						document.getElementById("Form1").submit();
		return true;
			}
			</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" name="Form1" method="post" runat="server">
		<table border=0 cellpadding=0 cellspacing=0 align =center>
						<TR><td height=10 align=center>אנא הכנס סיסמה</td></TR>
						<tr>
							<td width="100%"><input type=password id="pass" name="pass"></td>
						</tr>
						<TR><td height=10>&nbsp;</td></tr>
												<TR><td align=center><input type="button" class="but_menu" style="width:90px" onclick="javascript:CheckFields(this); return false;" value="שלח" id="Button4" name="Button1"></td></TR>
						</table>
						<br clear=all>
						<asp:PlaceHolder  id="Divresult" runat=server Visible=false>
						<table border=0 cellpadding=0 cellspacing=0 align =center>
						<TR><td height=10 align=center style="color:#ff0000">סיסמה לא נכונה</td></TR>
						</table>
						</asp:PlaceHolder>
					</form>
	</body>
</HTML>
