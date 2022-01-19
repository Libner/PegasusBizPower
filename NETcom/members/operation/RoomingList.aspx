<%@ Page Language="vb" AutoEventWireup="false" Codebehind="RoomingList.aspx.vb" Inherits="bizpower_pegasus2018.Rooming_List" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Rooming List Upload</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body style="MARGIN:0px" onload="self.focus();">
		<form id="Form1" method="post" runat="server">
			<table  align="center" cellspacing="0" cellpadding="2" border="0">
			<tr><td align="center" class="td_admin_4" height=30></td></tr>
				<tr>
					<td  height="5" align="center" width="100%"><b>Rooming List – <%=func.GetDepatureCode(DepartureId)%> <br></b></td>
				</tr>
<tr><td align="center" class="td_admin_4" height=10></td></tr>
						<tr>
					<td align="center"><a href="#"  class="button_edit_1" style="width:200" onclick="window.open('RoomingListUpload.aspx?dID=<%=DepartureId%>','winUPD','top=20, left=10, width=1000, height=350, scrollbars=1');">Rooming List Upload</a>
				
					</td>
				</tr>
					<tr>
					<td align="center" class="td_admin_4" height=20></td></tr>
					<tr><td align="center" style="background-color:green;width:200;LINE-HEIGHT: 20px;color:#ffffff;font-weight:bold">כניסה לאזור האישי של הספק</td></tr>
				<tr>
					<td align="center" class="td_admin_4" height=3></td></tr>
						<tr>
					<td align="center"  style="background-color:#f5f5f5;width:200;LINE-HEIGHT: 20px;color:#ffffff;font-weight:bold">
						<asp:DataList ID="dtlSupplier" Runat="server" RepeatColumns="1" RepeatDirection="Vertical" CellSpacing="1"
							RepeatLayout="Table" CellPadding="1" BackColor="#f5f5f5" ItemStyle-CssClass="td_admin_4" ItemStyle-HorizontalAlign="Right">
							<ItemTemplate>
								<table border="0" cellpadding="2" cellspacing="2">
									<tr>
										<td align="right" dir="rtl">
											<a href="<%=SiteRootPath%>/suppliers/?pp=<%#Container.DataItem("GUID")%>&crmadm=1&UU=<%=trim(Request.Cookies("bizpegasus")("UserId"))%>" target="_blank"><%#Container.DataItem("supplier_Name")%>&nbsp;<%#Container.DataItem("Country_Name")%></a>
											</td>
									</tr>
									<tr>
										<td colspan="2" height="5"></td>
									</tr>
								</table>
							</ItemTemplate>
						</asp:DataList></td></tr>
		</form>
	</body>
</HTML>
