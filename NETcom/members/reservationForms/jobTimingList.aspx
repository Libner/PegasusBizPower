<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="jobTimingList.aspx.vb" Inherits="bizpower_pegasus2018.jobTimingList"%>
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
		<script>
	function openTimingUpd(TimingId)
	{
		h = 600;
		w = 500;
		S_Wind = window.open("updTiming.aspx?TimingId="+TimingId, "S_Wind" ,"scrollbars=1,toolbar=0,top=00,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	
		</script>
	</head>
	<body>
		<form id="Form1" method="post" runat="server" autocomplete="off">
			<DS:LOGOTOP id="logotop" runat="server"></DS:LOGOTOP>
			<DS:TOPIN id="topin" numOfLink="5" numOfTab="108" topLevel2="155" runat="server"></DS:TOPIN>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td height="30" align="left" style="WIDTH: 232px;">
						<a class="but_menu" style="WIDTH: 220px;padding:6px;BACKGROUND-COLOR: #98fc00;color:#000000"
							href="default.aspx">חזרה למסך שליחת טפסי רישום</a>
					</td>
					<td align="left" valign="middle" nowrap class="td_title_2" dir="rtl">
						<a class="but_menu" style="WIDTH: 220px;padding:6px" href="" onclick="openTimingUpd();return false">
							הוספת תזמון שליחה אוטומטית חדש</a>
					</td>
				</tr>
				<tr>
					<td height="10" colspan="2"></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="80%"align="center">
				<tr>
					<td valign="top"align="center">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center">
									<table border="0" cellpadding="1" cellspacing="1" width="100%">
										<tr bgcolor="#d8d8d8" id="blockSearchTerms">
											<td class="td_admin_5" width="70" align="center" nowrap>
												מחיקה
											</td>
											<td class="td_admin_5" width="70" align="center" nowrap>
												עדכון
											</td>
											<td class="td_admin_5" align="center" nowrap>
												פעיל/לא פעיל
											</td>
											<td class="td_admin_5" align="center" nowrap>
												מדינות
											</td>
											<td class="td_admin_5" align="center" nowrap>
												טווח תאריכי פתיחת טפסי הרישום
											</td>
											<td class="td_admin_5" align="center" nowrap>
												מספר יום
											</td>
											<td class="td_admin_5" align="center" nowrap>
												מועד השליחה
											</td>
											<td class="td_admin_5" align="center" nowrap>
												תאור
											</td>
										</tr>
										<asp:Repeater ID="rptTimings" runat="server">
											<ItemTemplate>
												<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';"
													style="background-color: rgb(201, 201, 201);">
													<td align="center" class="td_admin_1" valign="top"><a  href="jobTimingList.aspx?delId=<%#Container.DataItem("Timing_Id")%>" ONCLICK="return CheckDel()"><IMG SRC="../../images/delete_icon.gif" BORDER=0></a></td>
													<td align="center" class="td_admin_1" valign="top"><a  href="javascript:void(0)" onclick="openTimingUpd('<%#Container.DataItem("Timing_Id")%>');return false"><IMG SRC="../../images/edit_icon.gif" BORDER=0 ></a></td>
													<td align="right" class="td_admin_4">
														<%#getTimingVisiblity(func.dbnullfix(Container.DataItem("Sending_Frequancy")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#func.dbnullfix(Container.DataItem("Sending_countriesList"))%>
													</td>
													<td align="right" class="td_admin_4">
															<%#getTimingFormsDateInterval(func.dbnullfix(Container.DataItem("Sending_Forms_DateInterval")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#getTimingDay(func.dbnullfix(Container.DataItem("Sending_Forms_DayNumber")),func.dbnullfix(Container.DataItem("Sending_Frequancy")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#getTimingFrequancy(func.dbnullfix(Container.DataItem("Sending_Frequancy")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#func.dbnullfix(Container.DataItem("Timing_Title"))%>
													</td>
												</tr>
											</ItemTemplate>
											<AlternatingItemTemplate>
												<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
													style="background-color: rgb(230, 230, 230);">
													<td align="center" class="td_admin_1" valign="top"><a  href="jobTimingList.aspx?delId=<%#Container.DataItem("Timing_Id")%>" ONCLICK="return CheckDel()"><IMG SRC="../../images/delete_icon.gif" BORDER=0></a></td>
													<td align="center" class="td_admin_1" valign="top"><a  href="javascript:void(0)" onclick="openTimingUpd('<%#Container.DataItem("Timing_Id")%>');return false"><IMG SRC="../../images/edit_icon.gif" BORDER=0 ></a></td>
													<td align="right" class="td_admin_4">
														<%#getTimingVisiblity(func.dbnullfix(Container.DataItem("Sending_Frequancy")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#func.dbnullfix(Container.DataItem("Sending_countriesList"))%>
													</td>
													<td align="right" class="td_admin_4">
															<%#getTimingFormsDateInterval(func.dbnullfix(Container.DataItem("Sending_Forms_DateInterval")))%>
												</td>
													<td align="right" class="td_admin_4">
														<%#getTimingDay(func.dbnullfix(Container.DataItem("Sending_Forms_DayNumber")),func.dbnullfix(Container.DataItem("Sending_Frequancy")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#getTimingFrequancy(func.dbnullfix(Container.DataItem("Sending_Frequancy")))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#func.dbnullfix(Container.DataItem("Timing_Title"))%>
													</td>
												</tr>
											</AlternatingItemTemplate>
										</asp:Repeater>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
