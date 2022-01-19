<%@ OutputCache Duration="1" VaryByParam="none" Location="none"%>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editPage.aspx.vb" Inherits="bizpower_pegasus.editTextE" validateRequest="false" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>תבניות תוכן</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
		<script language="javascript">
		<!-- // load htmlarea
		_editor_url = "../../htmlarea/";						// URL to htmlarea files
		<%if siteId = 2 then%>
		_editor_lang = "en";
		<%end if%>									// interface language
		_view_form = false;
		var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
		if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
		if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
		if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
		if (win_ie_ver >= 5.5) {
		document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
		document.write(' language="javascript"></scr' + 'ipt>');  
		} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
		
		function CheckFields(){
			if (document.Form1.content.value=='')
			{
				alert('חובה למלא את השדה טקסט');
				return false;
			}else{
				return true;
			}	
		}
		function SetHTMLArea() {
					//HTMLArea.replace("content");
					var config = new Object();    // create new config object
					config.bodyStyle = ' font-family: Arial; font-size: 10pt; PADDING-LEFT: 10px; PADDING-RIGHT: 10px';
					config.fontsize_default = "10";
					editor_generate("content", config);
					openPreview();
			}
			
		// -->
		</script>
	</HEAD>
	<body topmargin="0" leftmargin="0" rightmargin="0" onload="SetHTMLArea()">
		<form id="Form1" method="post" runat="server" autocomplete="off">
			<table width="100%" align="center" cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td class="a_title_big" width="100%"><center>"מנגנון ניהול אתר "<%=siteName%></center>
					</td>
				</tr>
				<tr>
					<td bgcolor="#e6e6e6" height="1"></td>
				</tr>
				<tr>
					<td class="title_admin_1" align="right" valign="middle" width="100%" nowrap>&nbsp;&nbsp;<%=ptitle%>&nbsp;&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#e6e6e6" height="1"><INPUT type="hidden" id="previewpage" name="previewpage" runat="server"></td>
				</tr>
				<tr>
					<td width="100%" nowrap class="td_title_2">
						<table border="0" cellpadding="2" cellspacing="0">
							<tr>
								<td width="50%" align="left">
									<a class = "button_admin_1" href="publications.asp?pageId=<%=pageId%>&amp;innerparent=<%=pageId%>&amp;catId=<%=catida%>">
										חזרה</a>
								</td>
								<td width="50%" align="right">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td nowrap>
												<asp:Button id="Button3" runat="server" Text="" TabIndex="1" BorderStyle="None" BorderWidth="0"
													Height="1" Width="1"></asp:Button>
												<asp:Button id="Button2" runat="server" Text="תצוגה מקדימה" CssClass="button_admin_1" TabIndex="10"></asp:Button>
												<!--									
											<a class="button_admin_1" target="_blank" href="../../template/default.asp?maincat=<%=maincat%>&catId=<%=catida%>&PageId=<%=pageId%>&id_site=1&show=1">תצוגה מקדימה</a>
-->
											</td>
											<td width="5"></td>
											<td>
												<a class="button_admin_e" href="" onclick="return false;">עדכון הדף</a>
											</td>
											<td width="5"></td>
											<td>
												<a class="button_admin_1" href="editPage.asp?maincat=<%=maincat%>&amp;pageId=<%=pageId%>&amp;catId=<%=catida%>">
													עדכון מאפייני הדף</a>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td class="td_admin_4_left_align" height="10"></td>
				</tr>
				<tr>
					<td align="center" bgcolor="#e6e6e6" width="100%">
						<table align="center" cellspacing="0" cellpadding="0" border="0" width="100%">
							<tr>
								<td bgcolor="#e6e6e6" width="170" nowrap>&nbsp;</td>
								<td width="100%" align="center">
									<textarea ID="content" runat="server"></textarea>
								</td>
								<td bgcolor="#e6e6e6" width="170" nowrap>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgcolor="#e6e6e6" height="10"></td>
				</tr>
				<tr>
					<td bgcolor="#e6e6e6" height="10" align="center">
						<asp:Button id="Button1" BackColor="#FF9100" Font-Bold="True" Font-Name="Arial" Font-Size="10pt"
							ForeColor="#ffffff" runat="server" Text="  אישור  " TabIndex="2"></asp:Button>
						<input type="hidden" name="pageId" value="<%=pageId%>"> <input type="hidden" name="homePage" value="0">
						<input type="hidden" name="catId" value="<%=catida%>">
					</td>
				</tr>
				<tr>
					<td bgcolor="#e6e6e6" height="10"></td>
				</tr>
			</table>
			<asp:PlaceHolder id="PHAlert" runat="server"></asp:PlaceHolder>
		</form>
		<script language='javascript'>
				function openPreview()
				{
				<%	dim strpreview as String
						strpreview = "../../template/default.asp?maincat=" & maincat & "&catId=" & catida & "&PageId=" & pageId & "&id_site=" & siteId & "&show=1" %>
					if (document.Form1.previewpage.value == 'yes')
					{
					var newwin = window.open("<%=strpreview%>","_blank");
					document.Form1.previewpage.value = '';
					newwin.focus();
					}
				}
		</script>
	</body>
</HTML>
