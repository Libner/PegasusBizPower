<%@ Page Language="vb" EnableViewState="true" AutoEventWireup="false" validateRequest="false" Codebehind="sendMailQ.aspx.vb" Inherits="bizpower_pegasus.sendMailQ" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<%Response.Buffer = False%>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
			<style> .normalB { FONT-WEIGHT: bold; FONT-SIZE: 11pt; COLOR: #2e2380; FONT-FAMILY: Arial }
	</style>
	</HEAD>
	<BODY onload="window.focus();">
		<form id="form1" name="form1" method="post" runat="server">
			<input type="hidden" id="hdnemailLimit" runat="server"> <input type="hidden" id="hdnUseBizLogo" runat="server" NAME="hdnUseBizLogo">
			<input type="hidden" id="hdnPageLang" runat="server" NAME="hdnPageLang"> <input type="hidden" id="hdnfirst_slice" runat="server">
			<input type="hidden" id="hdnsecond_slice" runat="server"> <input type="hidden" id="hdnpagesource" runat="server">
			<table cellpadding="0" cellspacing="0" width="100%">
				<asp:Panel ID="pnlStartSend" Runat="server" Visible="True">
					<TBODY>
						<TR>
							<TD height="20"></TD>
						</TR>
						<TR>
							<TD class="normalB" dir="rtl" noWrap align="center"><FONT color="red">תהליך שליחת 
									מיילים מתבצע כעת, התהליך יארך מספר דקות</FONT></TD>
						</TR>
						<TR>
							<TD class="normalB" dir="rtl" noWrap align="center"><FONT color="red">אין לסגור חלון זה 
									ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</FONT></TD>
						</TR>
						<TR>
							<TD height="30" noWrap></TD>
						</TR>
						<TR>
							<TD class="td_subj" align="center">שליחת המיילים החלה</TD>
						</TR>
				</asp:Panel>
				<asp:Panel ID="pnlEndSend" Runat="server" Visible="False">
					<TR>
						<TD height="20" noWrap></TD>
					</TR>
					<%If (currEmailCount <= 0) Or (emailCount >= emailLimit) Then%>
					<TR>
						<TD class="td_subj" align="center">שליחת המיילים הסתיימה</TD>
					</TR>
					<%ElseIf currEmailCount >= emailLimit Then%>
					<TR>
						<TD class="td_subj" align="center">כיוון שיתרתך לשליחת מיילים נגמרה שליחת מיילים 
							הסתיימה</TD>
					</TR>
					<%End If%>
					<TR>
						<TD height="10" noWrap></TD>
					</TR>
				</asp:Panel>
				<%If isNumeric(trim(sendCount)) Then%>
				<tr>
					<td align="center" dir="rtl" class="td_subj">
						נשלחו &nbsp; <font color="#023296"><b>
								<%=ViewState("sendCount")%>
							</b></font>מיילים</td>
				</tr>
				<%End If%>
				<%If isNumeric(trim(notSendCount)) Then%>
				<tr>
					<td align="center" class="td_subj" dir="rtl">
						סה"כ מיילים לשליחה <font color="#023296"><b>
								<%=ViewState("notSendCount")%>
							</b></font>
					</td>
				</tr>
				<%End If%>
				<%if false then%>
				<tr>
					<td align="center" class="td_subj" dir="rtl">
						מתוכם נמצאו לא תקניים <font color="#023296"><b>
								<%=ViewState("badCount")%>
							</b></font>
					</td>
				</tr>
				<%end if%>
				<tr>
					<td height="20" nowrap></td>
				</tr>
				<tr>
					<td nowrap align="center">
						<INPUT class="but_menu" style="WIDTH:90px" type="button" onclick="javascript:window.close()"
							value="סגור חלון" id="button3" name="button3">
					</td>
				</tr>
				</TBODY>
			</table>
		</form>
	</BODY>
</HTML>
