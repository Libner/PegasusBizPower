<%@ Control Language="vb" AutoEventWireup="false" Codebehind="top_in.ascx.vb" Inherits="bizpower_pegasus.top_in" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<table cellpadding="0" cellspacing="0" width="100%" border="0">
	<tr>
		<td align="right" bgcolor="#ffffff">
			<table border="0" bgcolor="#ffffff" cellspacing="0" cellpadding="0" align="<%=align_var%>" dir=<%=dir_obj_var%>>
				<tr>
					<asp:PlaceHolder id="BarTopPH" runat="server"></asp:PlaceHolder>
				</tr>
			</table>
		</td>
	</tr>
	<tr><td width="100%" height="20" nowrap bgcolor="#ffffff">&nbsp;</td></tr>
	<tr>
		<td width="100%">
			<TABLE width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff" align="<%=align_var%>" dir=<%=dir_obj_var%>>
				<TR>
					<asp:PlaceHolder id="BarBottomPH" runat="server"></asp:PlaceHolder>
					<TD width="100%" style="BORDER-BOTTOM: #6f6d6b 1px solid">&nbsp;</TD>
				</TR>
			</TABLE>
		</td>
	</tr>
</table>
