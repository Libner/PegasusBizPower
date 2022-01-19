<%@ Control Language="vb" AutoEventWireup="false" Codebehind="top_in.ascx.vb" Inherits="bizpower_pegasus2018.top_in_webflowCSS" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>

<div class="biz_nav_section">
    <div class="biz_buttons">
     <a href="<%=strLocal%>/netCom/" class="bizlink_level_1 <%If trim(numOfTab) = "-1" Then%>current <%End if%>w-inline-block">
        <div>ראשי</div>
      </a>
<asp:PlaceHolder id="BarTopPH" runat="server"></asp:PlaceHolder>
    </div>
    <div class="biz_links" >
    <asp:PlaceHolder id="BarBottomPH" runat="server"></asp:PlaceHolder>
    </div>
    </div>
    <table cellpading="0 cellspacing=0" width="100%" align="right"><tr><td class="page_title">&nbsp;<!--אנא הקלד את שם החברה מימין או את שם איש הקשר לאיתור במאגר הלקוחות--></td></tr>
    </table>
    <!--
<table cellpadding="0" cellspacing="0" width="100%" border="0">
	<tr>
		<td align="right" bgcolor="#ffffff">
			<table border="0" bgcolor="#ffffff" cellspacing="0" cellpadding="0" align="<%=align_var%>" dir=<%=dir_obj_var%>>
				<tr>
					<asp:PlaceHolder id="BarTopPH_old" runat="server"></asp:PlaceHolder>
				</tr>
			</table>
		</td>
	</tr>
	<tr><td width="100%" height="20" nowrap bgcolor="#ffffff">&nbsp;</td></tr>
	<tr>
		<td width="100%">
			<TABLE width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff" align="<%=align_var%>" dir=<%=dir_obj_var%>>
				<TR>
					<asp:PlaceHolder id="BarBottomPH_old" runat="server"></asp:PlaceHolder>
					<TD width="100%" style="BORDER-BOTTOM: #6f6d6b 1px solid">&nbsp;</TD>
				</TR>
			</TABLE>
		</td>
	</tr>
</table>
-->
