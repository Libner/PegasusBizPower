<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SelectStatus.aspx.vb" Inherits="bizpower_pegasus.SelectStatus" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title><%=fTitle%></title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body style="MARGIN:0px" onload="self.focus();">
		<form id="Form1" method="post" runat="server">
			<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
				<tr>
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%"><b>?���� ����<br></b></td>
				</tr>
<tr>
					<td align="center" class="td_admin_4" height=10></td></tr>
				<tr>
					<td align="center">
						<select id="fselect" name="fselect" class="searchList" dir=rtl runat=server multiple size=4>
						
						</select>
					</td>
				</tr>
				<tr>
					<td align="center" class="td_admin_4" height=10></td></tr>
				<tr>
					<td align="center" class="td_admin_4"><asp:Button id="btnSubmit" BackColor="#FF9100" Font-Bold="True" Font-Name="Arial" Font-Size="10pt"
							ForeColor="#FFFFFF" EnableViewState="false" runat="server" Text="�����"></asp:Button></td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
