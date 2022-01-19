<%@ Page Language="vb" AutoEventWireup="false" Codebehind="List_openDays.aspx.vb" Inherits="bizpower_pegasus2018.List_openDays"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>List_openDays</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table border="0" cellpadding="2" cellspacing="0" align="center" width="100%">
				<tr>
					<td width="10"></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
				<tr>
					<td align="center" colspan="2" style="COLOR:#000000">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td height="30" align="center"><span style="FONT-SIZE: 14pt;COLOR: #6f6da6"><%=func.GetSelectUserName(uid)%>
										- ימים שלא נכח נציג בעבודה </span>
								</td>
							</tr>
							<tr>
								<td height="30" align="center"><span style="FONT-SIZE: 14pt;COLOR: #6f6da6"><%=dateStart%>
										-
										<%=dateEnd%>
										:בין תאריכים </span>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td align="center">
						<table border="1" cellspacing="3" cellpadding="3" align="center">
							<asp:Repeater ID="rptData" Runat="server">
								<ItemTemplate>
									<tr>
										<td align="center" style="background-color:#e1e1e1;width:250px" dir="rtl"><%#DataBinder.Eval(Container.DataItem, "DateKey", "{0:d/M/yyyy}")%></td>
										<td align="center" style="background-color:#e1e1e1;width:250px" dir="rtl"><%#Container.DataItem("WeekdayNameHeb")%></td>
									</tr>
								</ItemTemplate>
							</asp:Repeater>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
