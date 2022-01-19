<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editVVouchers_Provider_Status.aspx.vb" Inherits="bizpower_pegasus2018.editVouchers_Provider_Status" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title></title>
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
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%"><b>ספקים<br>
							<%=DepartureCode%>
						</b>
					</td>
				</tr>
				<tr>
					<td align="center" class="td_admin_4" height="10"></td>
				</tr>
				<tr>
					<td align="center">
						<asp:DataList ID="dtlSupplier" Runat="server" RepeatColumns="4" RepeatDirection="Vertical" CellSpacing="1"
							RepeatLayout="Table" CellPadding="1" BackColor="#f5f5f5" ItemStyle-CssClass="td_admin_4" ItemStyle-HorizontalAlign="Right">
							<ItemTemplate>
								<table border="0" cellpadding="2" cellspacing="2">
									<tr>
										<td align="right" dir="rtl">
											<a href="#" onclick="window.open('GetSupplierDetails.aspx?id=<%#Container.DataItem("supplier_Id")%>', 'suppDetails', 'toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=100,width=600,height=850');"><%#Container.DataItem("supplier_Name")%></a>
											<br>
											&nbsp;<span style="color:#aaaaaa"><%#Container.DataItem("Country_Name")%></span>
										</td>
										<td align="right" valign="top">
										<select id="sVouchers_Provider_<%#Container.DataItem("supplier_Id")%>" name="sVouchers_Provider_<%#Container.DataItem("supplier_Id")%>" dir="rtl" class="searchList">
										<option value="" <%#IIF (func.dbNullFix(Container.DataItem("Vouchers_Status"))="","selected","")%>>לא</option>
										<option value="כן" <%#IIF (func.dbNullFix(Container.DataItem("Vouchers_Status"))="כן","selected","")%>>כן</option>
										<option value="מותאם" <%#IIF (func.dbNullFix(Container.DataItem("Vouchers_Status"))="מותאם","selected","")%>>מותאם</option>
									</select></td>
									</tr>
									<tr>
										<td colspan="2" height="5"></td>
									</tr>
								</table>
							</ItemTemplate>
						</asp:DataList></td>
				<tr>
					<td align="center" class="td_admin_4" height="10"></td>
				</tr>
				<tr>
					<td align="center" class="td_admin_4">
						<table border="0" cellpadding="5" cellspacing="0">
							<tr>
						
									<td align="center" class="td_admin_4">
									<asp:Button id="btnSubmit" BackColor="#FF9100" Font-Bold="True" Font-Name="Arial" Font-Size="10pt"
										ForeColor="#FFFFFF" EnableViewState="false" runat="server" Text="אישור"></asp:Button></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
