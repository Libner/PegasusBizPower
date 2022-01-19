<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="logScreenMembersLastEnter.aspx.vb" Inherits="bizpower_pegasus2018.logScreenMembersLastEnter" %>
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
		<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
    <script>
        function scrollWin() {
            window.scrollTo(0, 0);
        }
        
        function FuncSort(query, par, srt) {

           
            if (query != "") {
                //alert(query)
                var r;
                r = "&" + par + "=ASC"
                query = query.replace(r, "");
                r = "&" + par + "=DESC"
                query = query.replace(r, "");
                //alert(query)

//alert("logScreen.aspx?" + query + "&" + par + "=" + srt)
                window.location = "logScreen.aspx?" + query + "&" + par + "=" + srt
            }
            else {
				var querySearch
				querySearch=""
				
				$("#blockSearchTerms").find("input").each(function () {
                if($(this).val()!="")
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + encodeURIComponent($(this).val())
					}
				});
				$("#blockSearchTerms").find("select").each(function () {
                if($(this).val()!="")
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + encodeURIComponent($(this).val())
					}
				});
//alert("logScreen.aspx?" + querySearch + "&" + par + "=" + srt)
                window.location = "logScreenMembersLastEnter.aspx.aspx?" + querySearch + "&" + par + "=" + srt
            }
        }
    </script>
    
	</head>
	<body>
		<form id="Form1" method="post" runat="server" autocomplete="off">
			<div class="biz_nav_section">
				<div class="bizpower_top">
					<div class="biz_pegasus_login">
						<div class="biz_login_name"><%=HttpContext.Current.Server.UrlDecode(Trim(Request.Cookies("bizpegasus")("UserName")))%></div>
						<div></div>
					</div>
					<div class="biz_logo_group w-clearfix">
						<div class="biz_logo"><img src="<%=Application("VirDir")%>/NETcom/webflowCSS/images/top_biz_logo.png" alt="" style="max-width: 100%;"></div>
						<div class="biz_logo_txt">pegasus</div>
						<div class="biz_logo_line">/</div>
					</div>
				</div>
			</div>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td height="30">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table border="0" cellpadding="1" cellspacing="1" width="100%" dir="rtl">
							 <tr bgcolor="#d8d8d8"  id="blockSearchTerms">
                    <td></td>
                        <td class="td_admin_5" nowrap>
                                <input runat="server" id="sContactName" dir="rtl" name="sContactName" class="search">
                        </td>
                        <td class="td_admin_5" nowrap>
                                <input runat="server" id="sPhone" dir="rtl" name="sPhone" class="search">
                        </td>
                        
                        <td class="td_admin_5" nowrap >
                        </td>
                       
                        <td class="td_admin_5" >
                        
                        </td>
                        
                        <td class="td_admin_5" nowrap  style="width: 250px">
                            
                        </td>
                        <td class="td_admin_5" nowrap width="100">
                            
                        </td>
                        
                        <td class="td_admin_5" align="center" nowrap dir="ltr">
							
                        </td>
                         <td class="td_admin_5" nowrap width="100">
                               
                        </td>
                            <td class="td_admin_5" align="center" width="50" nowrap>
                            <asp:LinkButton runat="server" ID="btnSearch" CssClass="button_small1"  BackColor=#ff9900>חפש</asp:LinkButton>
                        </td>
                      
                        <td class="td_admin_5" align="center" width="50" nowrap>
                            <a href="logScreenMembersLastEnter.aspx?<%=qrystring%>" class="button_small1">הצג הכל</a>
                         </td>
                      
                      </tr>
                      <tr>
								<td class="td_admin_5" align="center" nowrap>
									סדר
								</td>
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
                    <td class="td_admin_4" align="center">
                        <font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
                    </td>
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
						  <asp:PlaceHolder ID="phAppleal16735" Runat=server Visible=False>
                         <img src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16735")),"v.png","notok_icon.gif")%>" border="0" height="20">
					</asp:PlaceHolder>
                    </td>
                   <td align="center" class="td_admin_4" id="blockExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" <%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"style=background-color:#ccffcc","")%>>
                          <asp:PlaceHolder ID="phAppleal16504" Runat=server Visible=False>                        
                         <img id="imgExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"v.png","notok_icon.gif")%>" border="0" height="20">
						 </asp:PlaceHolder>
                    </td>
									</tr>
								</ItemTemplate>
								<AlternatingItemTemplate>
									<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
										style="background-color: rgb(230, 230, 230);">
										<td class="td_admin_4" align="center">
											<font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
										</td>
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
						  <asp:PlaceHolder ID="phAppleal16735" Runat=server Visible=False>
                         <img src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16735")),"v.png","notok_icon.gif")%>" border="0" height="20">
					</asp:PlaceHolder>
                    </td>
                   <td align="center" class="td_admin_4" id="blockExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" <%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"style=background-color:#ccffcc","")%>>
                          <asp:PlaceHolder ID="phAppleal16504" Runat=server Visible=False>                        
                         <img id="imgExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"v.png","notok_icon.gif")%>" border="0" height="20">
						 </asp:PlaceHolder>
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
														<asp:ListItem Value="30">30</asp:ListItem>
														<asp:ListItem Value="50">50</asp:ListItem>
														<asp:ListItem Value="100" Selected="True">100</asp:ListItem>
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
