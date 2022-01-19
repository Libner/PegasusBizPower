<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AddMessage.aspx.vb" Inherits="bizpower_pegasus.AddMessage" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AddMessage</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<TBODY>
					<tr>
						<td align="left" valign="middle" nowrap>
							<table width="100%" border="0" cellpadding="0" cellspacing="0" ID="Table8">
								<tr>
									<td class="page_title">הוספת&nbsp;שיחה&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td width="100%" bgcolor="#e6e6e6">
							<table align="center" border="0" cellpadding="3" cellspacing="1" width="500">
								<tr>
									<td align="right"><select runat="server" id="sSeries" class="searchList" dir="rtl" name="sSeries" onchange="selectToursDate(this)"></select>
									</td>
									<td align="right" nowrap width="100" valign="top">סדרה&nbsp;</td>
								</tr>
								<tr>
									<td align="right"><select runat="server" id="sDepCode" class="searchList" dir="rtl" name="sDepCode"></select>
									</td>
									<td align="right" nowrap width="100" valign="top">קוד טיול&nbsp;</td>
								</tr>
								<tr>
									<td align="right"><textarea dir="rtl" name="Messages_Content" style="HEIGHT:200px;FONT-FAMILY:arial;WIDTH:300px"
											ID="Messages_Content">&lt;%=func.vfix(Messages_Content)%&gt;</textarea></td>
									<td align="right" nowrap width="100" valign="top">פרוט שיחה&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
		</form>
		</TBODY></TABLE>
	</body>
</HTML>
