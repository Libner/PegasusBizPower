<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editFinancial_Euro.aspx.vb" Inherits="bizpower_pegasus.editFinancial_Euro" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>הערכות כספית  € </title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">
  </head>
<body style="margin:0px" onload="self.focus();">
		<form id="Form1" method="post" runat="server">
			<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
				<tr>
					<td bgcolor="#e6e6e6" height="5" align="center" width=100%><b>הערכות כספית  € <br> <%=DepartureCode%></b></td>
				</tr>
				<tr>
					<td align="center" >
						<input dir="ltr" name="Financial_Euro" id="Financial_Euro" style="WIDTH:60px;HEIGHT:30px" value="<%=Financial_Euro%>">
					</td>
				</tr>
				<tr>
					<td align="center" class="td_admin_4"><asp:Button id="btnSubmit" BackColor="#FF9100" Font-Bold="True" Font-Name="Arial" Font-Size="10pt"
							ForeColor="#FFFFFF" EnableViewState="false" runat="server" Text="אישור"></asp:Button></td>
				</tr>
				
			</table>
		</form>
	</body>
</html>
