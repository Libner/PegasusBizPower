<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editSelect.aspx.vb" Inherits="bizpower_pegasus.editSelect" %>
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
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%"><b><%=fTitle%><br> <%=DepartureCode%></b></td>
				</tr>
<tr>
					<td align="center" class="td_admin_4" height=10></td></tr>
				<tr>
					<td align="center">
						<select id="fselect" name="fselect" class="searchList" dir=rtl>
						<option value=" ">&nbsp;&nbsp;&nbsp;&nbsp;בחר&nbsp;&nbsp;&nbsp;&nbsp;</option>
						<option value="כן" <%if trim(fValue)="כן" then%> selected <%end if%>>כן</option>
						<%if fName="Vouchers_Provider" then%>
						<option value="מותאם">מותאם</option>
						<%end if%>
						<option value=" ">&nbsp;</option>
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
