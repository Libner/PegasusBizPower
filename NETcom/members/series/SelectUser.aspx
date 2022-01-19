<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SelectUser.aspx.vb" Inherits="bizpower_pegasus.SelectUser1" %>
<!DOCTYPE html>
<html>
	<head>
		<title><%=fTitle%></title>
	<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body style="MARGIN:0px" onload="self.focus();">
		<form id="Form1" method="post" runat="server">
			<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
				<tr>
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%"><b>אופרטור<br></b></td>
				</tr>
<tr>
					<td align="center" class="td_admin_4" height=10></td></tr>
				<tr>
					<td align="center">
						<select id="fselect" name="fselect"  dir=rtl runat=server multiple size=20 style="width:220px;font-size:9pt;font-family:arial;">
						
						</select>
					</td>
				</tr>
				<tr>
					<td align="center" class="td_admin_4" height=10></td></tr>
				<tr>
					<td align="center" class="td_admin_4"><asp:Button id="btnSubmit" BackColor="#FF9100" Font-Bold="True" Font-Name="Arial" Font-Size="10pt"
							ForeColor="#FFFFFF" EnableViewState="false" runat="server" Text="אישור"></asp:Button></td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
