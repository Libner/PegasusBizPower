<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ConcactLog.aspx.vb" Inherits="bizpower_pegasus2018.ConcactLogToursScreen" %>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>
<html>
	<head>
		<title>screenSetting</title>
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
		<script src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
	</head>
	<body>
		<form id="Form1" method="post" runat="server" autocomplete="off">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td style="background-color: #FFD011">
						<table cellpadding="0" cellspacing="0" width="100%" align="left" style="border: solid 0px #d3d3d3">
							<tr>
								<td><a class="button_edit_1" style="width: 95px; font_family: ARIAL; font-size: 14px; font-weight: normal"
										href="javascript:void(0)" onclick="javascript:window.open('','_self').close();">חלון&nbsp;סגור</a></td>
								<td width="100%" class="page_title" dir="rtl">&nbsp;<font style="color: #000000;"><b>היסטוריית 
											תאריכי כניסות משתמש</b></font>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			<tr>
			<td >
			<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="rtl">
				<tr>
					<asp:Repeater ID="rptTitle" runat="server">
						<ItemTemplate>
							<td class="title_sort" align="center">
								<%#Container.DataItem("Column_Title")%>
							</td>
						</ItemTemplate>
					</asp:Repeater>
				</tr>
				<asp:Repeater ID="rptLogs" runat="server">
					<ItemTemplate>
						<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';"
							style="background-color: rgb(201, 201, 201);">
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("CONTACT_NAME"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Phone"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Email"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("SubCategory_Name"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Country_Name"))%>
							</td>
							<td align="center" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Series_Name"))%>
							</td>
							<td align="center" class="td_admin_4">
								<%#func.dbNullDate(Container.DataItem("LastEnter_Date"),"dd/MM/yyyy HH:mm")%>
							</td>
							<td align="center" class="td_admin_4">
								<%#func.fixnumeric(Container.DataItem("countEnters"))%>
							</td>
							<td align="center" class="td_admin_4">
								<asp:PlaceHolder ID="phAppleal16735" Runat="server" Visible="False">
									<img src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16735")),"v.png","notok_icon.gif")%>" border="0" height="20">
								</asp:PlaceHolder>
							</td>
							<td align="center" class="td_admin_4" id="blockExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" <%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"style=background-color:#ccffcc","")%>>
								<asp:PlaceHolder ID="phAppleal16504" Runat="server" Visible="False">
									<img id="imgExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"v.png","notok_icon.gif")%>" border="0" height="20">
								</asp:PlaceHolder>
							</td>
						</tr>
					</ItemTemplate>
					<AlternatingItemTemplate>
						<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
							style="background-color: rgb(230, 230, 230);">
								<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("CONTACT_NAME"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Phone"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Email"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("SubCategory_Name"))%>
							</td>
							<td align="right" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Country_Name"))%>
							</td>
							<td align="center" class="td_admin_4">
								<%#func.dbNullFix(Container.DataItem("Series_Name"))%>
							</td>
							<td align="center" class="td_admin_4">
								<%#func.dbNullDate(Container.DataItem("LastEnter_Date"),"dd/MM/yyyy HH:mm")%>
							</td>
							<td align="center" class="td_admin_4">
								<%#func.fixnumeric(Container.DataItem("countEnters"))%>
							</td>
							<td align="center" class="td_admin_4">
								<asp:PlaceHolder ID="phAppleal16735" Runat="server" Visible="False">
									<img src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16735")),"v.png","notok_icon.gif")%>" border="0" height="20">
								</asp:PlaceHolder>
							</td>
							<td align="center" class="td_admin_4" id="blockExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" <%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"style=background-color:#ccffcc","")%>>
								<asp:PlaceHolder ID="phAppleal16504" Runat="server" Visible="False">
									<img id="imgExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"v.png","notok_icon.gif")%>" border="0" height="20">
								</asp:PlaceHolder>
							</td>
							</td>
						</tr>
					</AlternatingItemTemplate>
				</asp:Repeater>
				<asp:PlaceHolder ID="pnlPages" runat="server">
					<tr>
						<td height="2" colspan="13">
						</td>
					</tr>
					<tr>
						<td class="plata_paging" valign="top" align="center" colspan="18" bgcolor="#D8D8D8">
							<table dir="ltr" cellspacing="0" cellpadding="2" width="100%" border="0">
								<tr>
									<td class="plata_paging" valign="baseline" nowrap align="left" width="150">
										&nbsp;הצג
										<asp:DropDownList ID="PageSize" CssClass="PageSelect" runat="server" AutoPostBack="true">
											<asp:ListItem Value="10">10</asp:ListItem>
											<asp:ListItem Value="30" Selected="True">30</asp:ListItem>
											<asp:ListItem Value="50">50</asp:ListItem>
										</asp:DropDownList>
										&nbsp;בדף&nbsp;
									</td>
									<td valign="baseline" nowrap align="right">
										<asp:LinkButton ID="cmdNext" runat="server" CssClass="page_link" Text="«עמוד הבא"></asp:LinkButton>
									</td>
									<td class="plata_paging" valign="baseline" nowrap align="center" width="160">
										<asp:Label ID="lblTotalPages" runat="server"></asp:Label>&nbsp;דף&nbsp;
										<asp:DropDownList ID="pageList" CssClass="PageSelect" runat="server" AutoPostBack="true"></asp:DropDownList>
										&nbsp;מתוך&nbsp;
									</td>
									<td valign="baseline" nowrap align="left">
										<asp:LinkButton ID="cmdPrev" runat="server" CssClass="page_link" Text="עמוד קודם»"></asp:LinkButton>
									</td>
									<td class="plata_paging" valign="baseline" align="right">
										&nbsp;נמצאו&nbsp;&nbsp;&nbsp;
										<asp:Label CssClass="plata_paging" ID="lblCount" runat="server"></asp:Label>&nbsp;יציאות
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</asp:PlaceHolder>
							</table>
						</td>
					</tr>
			</table>
		</form>
	</body>
</html>
