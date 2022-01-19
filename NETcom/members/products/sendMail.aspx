<%@ Page Language="vb" EnableViewState="false" AutoEventWireup="false" validateRequest="false" Codebehind="sendMail.aspx.vb" Inherits="bizpower_pegasus.sendMail" %>
<%Response.Buffer = False%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<style>
.normalB
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11pt;
    COLOR: #2e2380;
    FONT-FAMILY: Arial
}
</style>
</HEAD>
<BODY onload="window.focus();">
<form id="form1" name="form1" method="post" runat="server">
<input type="hidden" id="hdnemailLimit" runat="server" value="">
<input type="hidden" id="hdnUseBizLogo" runat="server" value="" NAME="hdnUseBizLogo">
<input type="hidden" id="hdnPageLang" runat="server" value="" NAME="hdnPageLang">
<input type="hidden" id="hdnfirst_slice" runat="server" value="">
<input type="hidden" id="hdnsecond_slice" runat="server" value="">
<input type="hidden" id="hdnpagesource" runat="server" value="">
<table cellpadding=0 cellspacing=0 width="100%">
<asp:Panel ID="pnlStartSend" Runat="server" Visible="True">
<tr><td height="20"></td></tr>	
<tr>
	<td align="center" nowrap class="normalB" dir=rtl><font color="red">תהליך שליחת מיילים מתבצע כעת, התהליך יארך מספר דקות</font></td>
</tr>		
<tr>
	<td  align="center" nowrap class="normalB" dir=rtl><font color="red">אין לסגור חלון זה ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</font></td>
</tr>	
<tr><td height="30" nowrap></td></tr>
<tr><td align="center" class="td_subj">שליחת המיילים החלה</td></tr>
</asp:Panel>
<asp:Panel ID="pnlEndSend" Runat="server" Visible="False">
<tr><td height="20" nowrap></td></tr>
<%If (currEmailCount <= 0) Or (emailCount >= emailLimit) Then%>
<tr><td align="center" class="td_subj">שליחת המיילים הסתיימה</td></tr>
<%ElseIf currEmailCount >= emailLimit Then%>
<tr><td align="center" class="td_subj">כיוון שיתרתך לשליחת מיילים נגמרה שליחת מיילים הסתיימה</td></tr>
<%End If%>
<tr><td height="10" nowrap></td></tr>
</asp:Panel>
<%If isNumeric(trim(sendCount)) Then%>
<tr><td align="center" dir=rtl class="td_subj"> נשלחו &nbsp; <font color="#023296"><b><%=sendCount%></b></font> מיילים</td></tr>
<%End If%>
<%If isNumeric(trim(totalCount)) Then%>
<tr><td align="center" class="td_subj" dir=rtl> סה"כ מיילים לשליחה <font color="#023296"><b><%=totalCount%></b></font></td></tr>
<%End If%>
<tr><td height="20" nowrap></td></tr>
<tr><td nowrap align="center">
<INPUT class="but_menu" style="width:90" type="button" onclick="javascript:window.close()" value="סגור חלון" id=button3 name=button3>
</td></tr>
</table>
</form>
</body>
</html>