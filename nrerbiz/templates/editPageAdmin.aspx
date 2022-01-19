<%@ Page Language="vb" AutoEventWireup="false" Codebehind="editPageAdmin.aspx.vb" validateRequest="false" Inherits="bizpower_pegasus.editTextAdmin" %>
<%@ OutputCache Duration="1" VaryByParam="none" Location="none"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Bizpower Administration</title>
	<meta charset="windows-1255">
	<link href="../../admin.css" rel="STYLESHEET" type="text/css">
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript">
		<!-- // load htmlarea
		_editor_url = "../../../htmlarea/";                     // URL to htmlarea files
		var page_id = '<%=templateID%>';
		_view_form = false;		
		
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
				alert('חובה למלא את השדה טקסט');
				return false;
			}else{
				return true;
			}	
		}
		function openPreview(templateID)
		{
			page = window.open("result.asp?templateID="+templateID,"Page","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");	
		}	
		function openHelp()
        {
			window.open("help_editor.html",'','menubar=0,toolbar=0,width=560,height=530,scrollbars=1,left=50,top=5,resizable=1');
			return false;
        }
      function SetHTMLArea() {
					var config = new Object();    // create new config object
					config.bodyStyle = 'font-family: Arial; font-size: 12pt; MARGIN: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; SCROLLBAR-FACE-COLOR: #DDECFF;SCROLLBAR-HIGHLIGHT-COLOR: #DEEBFE;SCROLLBAR-SHADOW-COLOR: #719BED;SCROLLBAR-3DLIGHT-COLOR: #DDECFF;SCROLLBAR-ARROW-COLOR: #11449D;SCROLLBAR-TRACK-COLOR: #DDECFF;SCROLLBAR-DARKSHADOW-COLOR: #ffffff';
					config.debug = 0;
					config.fontsize_default = "10";
					config.width = "641";
					editor_generate("title1", config);
			}

		// -->
		</script>
	</HEAD>
	<body  onload="SetHTMLArea()">
	
		<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td colspan="5" class="page_title">&nbsp;<%if trim(templateId)<>"" then%>עדכון<%else%>הוספת<%end if%>&nbsp;תבנית</td>
		</tr>
		<tr>			
			<td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדף ניהול תבניות</a></td>     
			<td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>  
			<td width="5%" align="center"><a class="button_admin_1" href="#" onclick="openPreview('<%=templateId%>')">תצוגה מקדימה</a></td>  
			<td width="*%" align="center"></td> 
		</tr>
		</table>
		<table width="630" align=center border="0" cellpadding="2" cellspacing="1" bgcolor="#ffffff">		
		<tr><td height="10" nowrap></td></tr>
		<tr>
			<td width="100%" colspan=2>
			<form id="Form1" method="post" runat="server">
			<table cellSpacing="0" cellPadding="0" width="630" align="center" border="0">
			<tr>
			<td width="530" nowrap bgcolor="#E6E6E6" align="right"><input 
            class=texts id=templateTitle1 dir=rtl style="WIDTH: 350px;text-align:right" type=text 
            maxLength=150 value="" name=templateTitle1  runat="server"
            ></td>
			<td  width="100" nowrap class="10normalB" bgcolor="#DDDDDD" align="center">&nbsp;<b>שם התבנית</b>&nbsp;</td>
							</tr>
							<tr>
								<td colSpan="2" height="10"></td>
							</tr>
							<tr><td colSpan="2" align="right" valign=top id="td_textarea" name="td_textarea">
							<textarea id="title1" style="WIDTH: 637px;height:400px; margin-left: 0px;padding: 0px" runat="server"></textarea></td></tr>

							<tr>
								<td colSpan="2" height="10"></td>
							</tr>
							<tr>
								<td dir="rtl" align="center" colSpan="2" height="10"><asp:button id="Button1" runat="server" Text="  אישור  " ForeColor="#ffffff" Font-Size="12px"
										Font-Name="Arial" Font-Bold="True" BackColor="#060165"></asp:button>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<asp:Button id="Button2" BackColor="#060165" Font-Bold="True" Font-Name="Arial" Font-Size="12px"
										ForeColor="#ffffff" runat="server" Text="  ביטול  "></asp:Button>
									<input type="hidden" name="templateId1" id="templateId1" value="" runat="server"> <input type="hidden" name="homePage" value="0">
								</td>
							</tr>
							<tr>
								<td height="10" colspan="2"></td>
							</tr>						
						</table>
						<asp:PlaceHolder id="PHAlert" runat="server"></asp:PlaceHolder></form>
				</td>
			</tr>
		</table>
	</body>
</HTML>
