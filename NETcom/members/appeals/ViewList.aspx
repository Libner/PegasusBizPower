<%@ Page Language="vb" enableViewState="false" AutoEventWireup="false" Codebehind="ViewList.aspx.vb" Inherits="bizpower_pegasus.ViewList" %>
<%@ Register TagPrefix="AtrafDating" TagName="bannersInclude" Src="../global/bannersInclude.ascx" %>
<%@ Register TagPrefix="AtrafDating" TagName="topInclude" Src="../global/topInclude.ascx" %>
<%@ Register TagPrefix="AtrafDating" TagName="rightInclude" Src="../global/rightInclude.ascx" %>
<%@ Register TagPrefix="AtrafDating" TagName="bottomInclude" Src="../global/bottomInclude.ascx" %>
<%@ OutputCache Duration="5" VaryByParam="none" Location="none"%>
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>אטרף דייטינג - תזכורות אישיות</title>
	<LINK href="/global/styles.css?v=<%=Application("sversion")%>" type="text/css" rel="STYLESHEET">
	<%If siteAudience <> "" Then%><LINK href="/global/skin-<%=siteAudience%>.css?v=<%=Application("sversion")%>" type="text/css" rel="STYLESHEET"><%else%><LINK href="/global/skin-master.css?v=<%=Application("sversion")%>" type="text/css" rel="STYLESHEET"><%end if%>		
	<script src="/javaScripts/jquery-1.11.2.min.js"></script>	
	<script language="javascript" src="/javaScripts/atrafdating.js?v=<%=Application("sversion")%>" type="text/javascript"></script>	
	<script language="javascript" src="/javaScripts/pushpop.js?v=<%=Application("sversion")%>" type="text/javascript"></script>	
	<script type="text/javascript" src="/javaScripts/fancybox/jquery.fancybox.js?v=<%=Application("sversion")%>"></script>
	<link rel="stylesheet" type="text/css" href="/javaScripts/fancybox/jquery.fancybox.css?v=<%=Application("sversion")%>" media="screen" />
	<script type="text/javascript" src="/javaScripts/jquery.bpopup.min.js?v=<%=Application("sversion")%>"></script>			
</HEAD>

<body stats="1" class="body_background">
<form method="post" runat="server" ID="Form1">

<atrafdating:topInclude id="topInclude" runat="server"></atrafdating:topInclude>

<div id="contentMain" class="contentMain" align="center">

<table id="siteMainTable" border="0" class="site_mainTable" cellspacing="0" cellpadding="0">

	<tr>
		<atrafdating:rightInclude id="BRM"  showPirsum="True" showBanner="True" runat="server"></atrafdating:rightInclude>
	
		<td width="610" align="left" valign="top">	
			<div class="content_mainTable">
				
				<div class="ttl-title-parent">
					<div class="ttl-title">תזכורות אישיות</div>
					<div style="float:left;position: relative;top: 15px;">
									<%'//start of info box%>
									<div dir="rtl">
									<table border="0" cellspacing="0" cellpadding="0">             
										<tr>	
											<td align="left" valign="middle">	
															<asp:TextBox ID="txtSearch" dir="rtl" value="חפש תזכורות" CssClass="right-style-search" style="width:170px; height: 31px;  margin: 0; padding: 0 5px; background-color: #FFFFFF;"  Runat="server" onclick="clearDefaultSearch('txtSearch')" onkeypress="return searchMemo();" />											
											</td>
										</tr>
									</table>																						
									</div>
									<%'//end of info box%>	
					</div>
				</div>
																
				<atrafdating:bannersInclude id="bannersInclude" runat="server"></atrafdating:bannersInclude>
								
				<div>
					<table border="0" width="100%" cellspacing="0" cellpadding="0">             
						<tr>	
							<td align="right" valign="top">	
								<%'//start of page content%>	
						<table cellSpacing="0" cellPadding="0" width="100%" align="center" bgColor="#ffffff" border="0">							
							<TR>
								<td align="center" width="100%">
									<%'//start of content%>
									<asp:Repeater ID="RLi" Runat="server" enableViewState="false">																		
										<ItemTemplate>
												<div class="lcrop-main">
														<div class="lcrop-pic">
															<div id="divPic" runat="server" class="lcrop-pic-in pphover">
																<div class="lcrop-pic-nw" id="divNW"  onclick="stopBubble(event);" runat="server"></div>
															</div>
														</div>
														<div class="lcrop-c">
															<div id="divMembername" runat="server"></div>	
															<div class="lcrop-ll">
																<div class="grayText normalText"><asp:Label id="lblMemo_membername" Runat="Server" /></div>
																<div class="grayText normalText"><asp:Label id="lblMemoPhone" Runat="Server" /></div>																																
																<div  class="grayText normalText lcrop-memo"><asp:Label id="lblMemoText" Runat="Server" /></div>																																
															</div>
														</div>
														<div class="lcrop-left">
															<asp:placeholder id="PHMessage" Runat="server"></asp:placeholder>
														</div>
												</div>
										</ItemTemplate>
										<HeaderTemplate>
											<div class="lcrop-line"></div>
										</HeaderTemplate>										
										<SeparatorTemplate>
											<div class="lcrop-h-sep"></div><div class="lcrop-line"></div>																								
										</SeparatorTemplate>
										<FooterTemplate>
											<%If IsDataBound = True Then%>
											<div class="lcrop-h-sep"></div><div class="lcrop-line"></div>
											<div class="gallery-banner" align="center">
												<iframe src="/banners/getNormalBanner.aspx" allowtransparency="true" width="468" height="70" frameborder="0" scrolling="no"></iframe>
											</div>		
											<div class="lcrop-h-sep"></div><div class="lcrop-line"></div>
											<%end if%>										
										</FooterTemplate>								
									</asp:Repeater>
									
									<%If IsDataBound = False Then%>
									<div align="center" class="largeText bold"><br>הרשימה ריקה</div>							
									<%end if%>	
									
									<asp:placeholder id="PHPagerList" Runat="server">
										<table align="center" border="0" width="95%" cellspacing="0" cellpadding="0">
										<tr><td height="20"></td></tr>	   
										<tr>
												<td align="center">
													<table border="0" cellspacing="0" cellpadding="2">
														<tr>
																<asp:placeholder id="PHPagerPage" Runat="server"></asp:placeholder>
														</tr>
													</table>
												</td>			
											</tr>
										</table>	
									</asp:placeholder>																		
									<%'//end of content%>
								</td>
							</TR>
						</table>
								<%'//end of page content%>	
							</td>
						</tr>
					</table>							
				</div>
			</div>		
		</td>		
	</tr>
</table>
</div>		

<atrafdating:bottomInclude id="bottomInclude" runat="server"></atrafdating:bottomInclude>

</form>	
</body>
</HTML>
