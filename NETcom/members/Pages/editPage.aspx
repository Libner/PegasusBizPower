<%@ OutputCache Duration="1" VaryByParam="none" Location="none"%>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editPage.aspx.vb" validateRequest="false" Inherits="bizpower_pegasus.editText" %>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta charset="windows-1255">
		<LINK href="../../IE4.css" type="text/css" rel="STYLESHEET">
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript">
		<!-- // load htmlarea
		_editor_url = "../../../htmlarea/";                     // URL to htmlarea files
		var page_id = '<%=pageID%>';
		_view_form = false;
		license = getCookie("SURVEYS");
		license = new String(license);
		license = license.slice(0,1)
		//window.alert(license);
		if(license == "1")
			_view_form = true;
		//window.alert(_view_form);
		
		var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
		if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
		if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
		if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
		if (win_ie_ver >= 5.5) {
		document.write('<script src="' +_editor_url+ 'editor.js"');
		document.write(' language="javascript"></script>');  
		} 
		else 
		{ 		   
		  document.write('<script>');
		  document.write(' function editor_generate() {');
		  document.write(' window.document.all("title1").style.display = "none";');
		  document.write(' window.document.all("td_textarea").className = "title_wizard";');		 
		  document.write(' window.document.all("td_textarea").innerHTML = "עורך הדפים עובד בדפדפן אקספלורר 5.5 ומעלה בלבד";');
		  document.write(' window.document.all("td_textarea").innerHTML = window.document.all("td_textarea").innerHTML + "<br>" + "במידה שאינך רואה את העורך עליך להתקין גרסת דפדפן עדכנית";');		  
		  document.write(' window.document.all("td_textarea").innerHTML = window.document.all("td_textarea").innerHTML + "<br>" + "<a href=http://www.microsoft.com/windows/ie_intl/he/default.mspx style=color:red target=_blank>לחץ כאן</a> להורדת גרסת דפדפן עדכנית<img src=../../images/icon_explor.gif vspace=0 style=' + 'padding-left:10px;position:relative;top:3px;' + '>";'); 		  
		  document.write(' return false; }');
		  document.write('</script>'); 
		}
		
		function CheckFields(){
			if (document.Form1.title1.value=='')
			{
				<%
					If trim(lang_id) = "1" Then
						str_alert = "חובה למלא את השדה טקסט"
					Else
						str_alert = "Please insert the page content"
					End If   
				%>					
				window.alert("<%=str_alert%>");
				return false;
			}else{
				return true;
			}	
		}
		function openPreview(pageId)
		{
			page = window.open("result.asp?pageId="+pageId,"Page","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=580, left=2, top=2");	
		}
	    function getCookie(name)
	    {
			var prefix = name + "="
			var cookieStartIndex = document.cookie.indexOf(prefix)
			if (cookieStartIndex == -1)
				return null
			var cookieEndIndex = document.cookie.indexOf("&", cookieStartIndex +
				prefix.length)
			if (cookieEndIndex == -1)
				cookieEndIndex = document.cookie.length
			return unescape(document.cookie.substring(cookieStartIndex +
				prefix.length, cookieEndIndex))
        }
        
        function openTemplate()
        {
			if(window.Form1.title1.value != "")
			{
			   <%
					If trim(lang_id) = "1" Then
						str_confirm = "			            ! שים לב\n\n.בחירת תבנית תמחוק את התוכן הקיים בדף\n\n                                    ? האם ברצונך להמשיך"
					Else
						str_confirm = "Pay attention!  \n\nSelection of the template will delete the current page content. \n\nAre you sure want to procced?"
					End If   
				%>					
				if (window.confirm("<%=str_confirm%>"))
				{
					window.open('templates_list.asp?template_id='+document.all('template_id').value,'','menubar=0,toolbar=0,width=550,height=480,scrollbars=1,left=50,top=10,resizable=1');
				}
			}
			else
			window.open('templates_list.asp?template_id='+document.all('template_id').value,'','menubar=0,toolbar=0,width=550,height=480,scrollbars=1,left=50,top=10,resizable=1');
				
			return false;
        }
        function openHelp()
        {
			window.open("help_editor.html",'','menubar=0,toolbar=0,width=560,height=530,scrollbars=1,left=50,top=5,resizable=1');
			return false;
        }
        function SetHTMLArea() {
					var config = new Object();    // create new config object
					config.bodyStyle = 'font-family: Arial; font-size: 10pt; MARGIN: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; SCROLLBAR-FACE-COLOR: #DDECFF;SCROLLBAR-HIGHLIGHT-COLOR: #DEEBFE;SCROLLBAR-SHADOW-COLOR: #719BED;SCROLLBAR-3DLIGHT-COLOR: #DDECFF;SCROLLBAR-ARROW-COLOR: #11449D;SCROLLBAR-TRACK-COLOR: #DDECFF;SCROLLBAR-DARKSHADOW-COLOR: #ffffff';
					config.fontsize_default = "10";
					config.width = "642";
					editor_generate("title1", config);
			}
        // -->
		</script>			
	</HEAD>
<body onload="SetHTMLArea();">
<form id="Form1"  method="post" runat="server" autocomplete=off>	
	<DS:LOGOTOP id="logotop"  runat="server"></DS:LOGOTOP>
	<DS:TOPIN id="topin" numOfLink="0"  numOfTab="2" runat="server"></DS:TOPIN>
	<table cellSpacing="0" cellPadding="0" width="100%">
	<tr><td class="page_title" width="100%">&nbsp;</td></tr>									
    <tr><td width="100%">
	<table cellpadding="1" cellspacing="1" width="100%" border="0" dir="<%=dir_var%>">
	<tr>
		<td width="100%" bgcolor="#E6E6E6">															
		<table cellSpacing="0" cellPadding="0" align="center" border="0">
		<tr><td height="10" nowrap></td></tr>
		<tr>
			<td align="<%=align_var%>">
			<table cellpadding=0 cellspacing=0 width=100% border=0>
			<tr>
			<td width=100% align="<%=align_var%>">
			<input type="hidden" name="template_id" id="template_id" runat="server">
				<input class="texts" id="pageTitle1" style="WIDTH: 350px;" type="text" maxLength="150"
					name="pageTitle1" runat="server"></td>
			<td width="100" nowrap align="center">&nbsp;<b><!--שם הדף--><%=arrTitles(3)%></b>&nbsp;</td>
			</tr>
			<tr>
			<td align="<%=align_var%>">
			<select name="PageLang" id="PageLang" class="norm" dir="<%=dir_obj_var%>">
			<option value="1" <%if trim(PageLang) = "1" then%> selected<%end if%>>&nbsp;&nbsp;<!--עברית--><%=arrTitles(14)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</OPTION>
			<option value="2" <%if trim(PageLang) = "2" then%> selected<%end if%>>&nbsp;&nbsp;<!--אנגלית--><%=arrTitles(15)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</OPTION>		  
			</select>
			</td>
			<td width="100" nowrap align="center"><b>&nbsp;<!--שפת הדף--><%=arrTitles(13)%>&nbsp;</b></td>
			</tr>			
			</table>
			</td>
		</tr>
		<tr><td height="10"></td></tr>
		<tr><td dir=rtl style="color:red" align="<%=align_var%>">&nbsp;<!--בעת עריכת הדף יש להשתמש בSHIFT ENTER על מנת לרדת לשורה חדשה.--><%=arrTitles(12)%>&nbsp;
		<br>{$var_rcptname} - מתחלף לשם המלא של נמען
		</td></tr>
		<tr><td height="10"></td></tr>
		<tr><td align="<%=align_var%>" valign=top id="td_textarea" name="td_textarea">
		<textarea id="title1"  style="WIDTH: 642px;height:400px; margin-left: 0px;padding: 0px" runat="server" ><p dir=rtl>&nbsp;</p></textarea></td></tr>
		<tr>
			<td height="10"></td>
		</tr>
		<tr>
			<td dir="ltr" align="center" height="10">
			<asp:Button id="Button2" BackColor="#060165" Font-Bold="True" Font-Name="Arial" Font-Size="12px"
					ForeColor="#ffffff" runat="server" Text="  ביטול  " Width=85px></asp:Button>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<asp:button id="Button1" runat="server" Text="  אישור  " ForeColor="#ffffff" Font-Size="12px"
					Font-Name="Arial" Font-Bold="True" BackColor="#060165" Width=85px></asp:button>	
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
				<asp:button id="Button6" runat="server" Text="  שמור שינויים  " ForeColor="#ffffff" Font-Size="12px"
					Font-Name="Arial" Font-Bold="True" BackColor="#060165" Width=85px></asp:button>													
				<input type="hidden" name="pageId1" id="pageId1" runat=server value="<%=pageId%>">
			</td>
		</tr>
		<tr><td height=10 nowrap></td></tr>										
		<tr><td><asp:PlaceHolder id="PHAlert" runat="server"></asp:PlaceHolder></td></tr>
		</table>								
		</td>
		<td width="110" nowrap align="<%=align_var%>" valign="top" style="border:1px solid #808080;" class="td_menu">
			<table cellpadding="1" cellspacing="1" width="100%" border=0>				
				<tr>
					<td align="<%=align_var%>" colspan="2" height="5" nowrap></td>
				</tr>
				<tr>
					<td nowrap colspan="2" align="center"><b><!--שם התבנית--><%=arrTitles(4)%></b></td>
				</tr>									
				<tr>
					<td nowrap colspan="2" align="center">
					<asp:Label style="border: 0px; background: #d3d3d3;" name="template_title" id="template_title" runat="server" CssClass="texts">ללא תבנית</asp:Label></td>
				</tr>
				<tr>
					<td align="<%=align_var%>" colspan="2" height="5" nowrap></td>
				</tr>
				<tr>
					<td nowrap colspan="2" align="center"><a class="button_edit_1" style="WIDTH:100px; line-height:22px" href='#' onclick="return openTemplate();">
                    <!--בחר	תבנית--><%=arrTitles(5)%></a></td>
				</tr>
				<%If isNumeric(pageId) Then%>
				<tr>
					<td nowrap colspan="2" align="center">
					<asp:Button id="Button7" CssClass="button_edit_1"  runat="server" Text="תצוגה מקדימה" Width=100px></asp:Button></td>
				</tr>
				<%End If%>
			</table></td></tr>
		</table></td></tr>			
	</table>
	</form>
  </body> 
</HTML>