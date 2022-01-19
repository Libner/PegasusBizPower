<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editStatus.aspx.vb" Inherits="bizpower_pegasus.editStatus"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>editStatus</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
  </head>
  <body  style="background-color:#e6e6e6">

    <form id="Form1" method="post" runat="server">
    <table cellpadding=0 cellspacing=0 align=center  >
    <tr><td align=center>
<select ID="sStatus" Runat="server" Width="550" style="height:30px;font-size:12pt;direction:rtl"></select>
 </td></tr>
 <tr>
					<td height="5" ></td>
				</tr>
 	<tr>
					<td align="center" class="td_admin_4"><asp:Button id="btnSubmit" BackColor="#FF9100" Font-Bold="True" Font-Name="Arial" Font-Size="10pt"
							ForeColor="#FFFFFF" EnableViewState="false" runat="server" Text="אישור"></asp:Button></td>
				</tr>
				<tr>
					<td height="5" ></td>
				</tr>
 </table>
    </form>

  </body>
</html>
