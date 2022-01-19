<%@ Control Language="vb" AutoEventWireup="false" Codebehind="logotop.ascx.vb" Inherits="bizpower_pegasus2018.logotop" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<%if ShowLogo = "1" then%>
<%Dim show_logo as Boolean
  Dim align_logo as String
  show_logo = true : align_logo = "center"
  If Not isNothing(Request.Cookies("bizpegasus")("UseBizLogo")) Then
	If trim(Request.Cookies("bizpegasus")("UseBizLogo")) = "0" Then 
		show_logo = false : align_logo = "left"
	End If	
  End If%>  
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
	<tr>
		<td width="100%" valign="bottom" height="79">
			<table border="0" width="100%" height="66"  cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top" align="center" width="100%" height="62"></td>
          <td valign="bottom" align="<%=align_var%>"><img src="<%=Application("VirDir")%>/images/topslog.gif" border="0" alt=""></td>
         <td width="201" valign=bottom>
          <table bgcolor="#6F6DA6" border="0" width="201" cellspacing="0" cellpadding="0" 
          style="background-image:URL('<%=Application("VirDir")%>/images/exit<%=img_%>.gif'); background-repeat:no-repeat;  background-position:top left;">
            <tr>
              <td width="201" nowrap height="21"><a href="<%=Application("VirDir")%>/exit.aspx" target="_self"><img 
              src="<%=Application("VirDir")%>/images/top_knisa<%=img_%>.gif" border="0"></a></td>
            </tr>
            <tr>            
              <td width="100%" height="45">
              <table border="0" width="201" height="45" cellspacing="1" cellpadding="0" dir="<%=dir_var%>">
                <tr>                                   
                  <td width=120 nowrap align="<%=align_var%>" style="color:#FFD011;font-weight:600;vertical-align:middle">&nbsp;<%=user_name%>&nbsp;</td>
                  <td width=73 nowrap align="<%=align_var%>" style="color:#FFFFFF;font-weight:600;vertical-align:middle">&nbsp;<span id=word1>משתמש</span>&nbsp;</td>
                  <td width=4 nowrap></td>
                </tr>
                <tr>
                 <td width=120 nowrap valign=top align="<%=align_var%>" style="color:#FFD011;font-weight:600;vertical-align:middle">&nbsp;<%=org_name%>&nbsp;</td>
                 <td width=73 nowrap align="<%=align_var%>" style="color:#FFFFFF;font-weight:600;vertical-align:middle" >&nbsp;<span id=word2>חברה</span>&nbsp;</td>                                 
                 <td width=4 nowrap></td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </td>
        </tr>
      </table>
		</td>
	</tr>
	<tr>
		<td width="100%" bgcolor="#060165" height="18" valign="top">
			<table border="0" width="100%" height="18" cellspacing="0" cellpadding="0">
				<tr>
					<td width="100%" height="1" nowrap></td>
				</tr>
				<tr>
					<td bgcolor="#060165" height="18" dir="rtl">
						<%''/strat of news%>
						<marquee direction="right" scrolldelay="120">
							<asp:PlaceHolder id="NewsPH" runat="server"></asp:PlaceHolder>
						</marquee>
						<%'//end of news%>
					</td>
				</tr>
				<tr>
					<td width="100%" height="1" nowrap></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end if ' show logo%>