<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SendMail.aspx.vb" Inherits="bizpower_pegasus2018.SendMail1" validateRequest="false"%>
<%@ Register TagPrefix="UC" TagName="logotop" Src="~/include/logotop.ascx" %>
<%@ Register TagPrefix="UC" TagName="topInclude" Src="~/include/topInclude.ascx"   %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SendMail</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script type="text/javascript" src="../../../../tinymce/tinymce.min.js"></script>
			<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
		<script>
		function ClearData()
		{
		$('#SelSearchUsers option').remove()
	$('#SelSearchDep option:selected').prop("selected", false);
//$('#SelSearchDep').find('option:selected').remove().end();
}
 function SendData()
{
$( "#SelSearchDep" )
    var str = "";
    $( "#SelSearchDep option:selected" ).each(function() {
      str += $( this ).val() + ", ";
  })
 if (str!='')
 {
 //alert(begin --str)
//alert(str)
   var xmlHTTP = new XMLHttpRequest();
            xmlHTTP.open("POST", 'getWorkersByDepId.aspx', false);
            xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><depId>' + str + '</depId></request>');
            ExpResult = new String(xmlHTTP.responseText);
   //  alert(ExpResult)
        ExpList = ExpResult.split("#")
    // alert(ExpList)     
          if (window.document.getElementById("SelSearchUsers").options) {
                window.document.getElementById("SelSearchUsers").options.length = 0;
        // window.document.getElementById("SelSearchUsers").options[0] = new Option("כל התחומים", 0);
        //       window.document.getElementById("SelSearchUsers").options.selectedIndex = 0;
         //     window.document.getElementById("SelSearchUsers").options[0].setAttribute("disabled", "disabled");
            }
        //  alert(ExpList.length)
        for (i = 0; i < ExpList.length; i++) {
                exp = new String(ExpList[i]);
                exp = exp.split(":")
                window.document.getElementById("SelSearchUsers").options[window.document.getElementById("SelSearchUsers").options.length] = new Option(exp[0], exp[1]);
                  window.document.getElementById("SelSearchUsers").options[window.document.getElementById("SelSearchUsers").options[i].setAttribute("selected", "selected")]
            }
           
 //alert(end  --str)
 }
 }
		function CheckForm()
		{
	//	alert("CheckForm")
		   if (document.getElementById("SelSearchUsers").value == '') //&& document.getElementById("sTo").value == '')
            {
                alert("אנא בחר איש קשר");
                document.getElementById("SelSearchUsers").focus()
                return false;
            }
               if (document.getElementById("pSubject").value == '') //&& document.getElementById("sTo").value == '')
            {
                alert("אנא מלא שדה נושא");
                document.getElementById("pSubject").focus()
                return false;
            }
            
		    document.getElementById("Form1").submit();
            return true;

	
		}
		</script>
		<script language="javascript" type="text/javascript">
		<!--
		tinymce.init({
			selector: "textarea#sContent",
			theme: "modern",
			width : <%=pwidth%>,
			height: 300,
			plugins: [
                "advlist autolink autosave lists link image charmap print preview hr anchor pagebreak spellchecker",
                "searchreplace wordcount visualblocks visualchars code fullscreen insertdatetime media nonbreaking",
                "save table contextmenu directionality emoticons template textcolor colorpicker paste"
			],
			toolbar1: "formatselect fontselect fontsizeselect | bold italic underline strikethrough | forecolor backcolor",
			toolbar2: "undo redo | outdent indent | bullist numlist | ltr rtl | alignleft aligncenter alignright alignjustify | link unlink anchor image media | visualchars visualblocks template",
			toolbar3: "pastetext searchreplace | hr removeformat | subscript superscript | charmap emoticons",
			templates: [ 
				{"title": "תבנית עברית", "description": "", "url": "../templates/TemplateHeb.htm"},
             	{"title": "תבנית אנגלית", "description": "", "url": "../templates/TemplateEng.htm"},
                  	{"title": "תבליט ריבוע", "description": "תבליט ריבוע", "url": "../templates/CusotmUl1.htm"},
				{"title": "תבליט חץ", "description": "תבליט חץ", "url": "../templates/CusotmUl2.htm"}
		
			
	
			],
			menubar: true,
			toolbar_items_size: 'small',
			content_css: "<%=Application("VirDir")%>/include/tinymce.css",
			button_tile_map : true,
			theme_advanced_resizing : true,
			nonbreaking_force_tab : false,
			extended_valid_elements : "*[*]",
			valid_elements : "*[*]",
			apply_source_formatting : true,
			convert_fonts_to_styles : true,
			relative_urls : false,
			remove_script_host : false,
			debug : false,
			document_base_url : "<%=BaseUrl%>",
			directionality : "rtl",
		    paste_auto_cleanup_on_paste: false,
   			paste_strip_class_attributes : "all",   			
   			paste_remove_styles : false,
			cleanup : true,
			fix_list_elements : true,
			fix_table_elements : true,
			inline_styles : true,
			cleanup_on_startup : false,
			image_advtab: true,
			table_grid: true,
			paste_data_images: false,
			paste_as_text: false,
			paste_webkit_styles: "all",
			paste_retain_style_properties: "all",
			language : 'he_IL',
			file_browser_callback: function(field_name, url, type, win) { 
				if (type == "media")
					{
						win.open('<%=Application("VirDir")%>/tinymce/plugins/media/AMediaAdd.aspx?field_name=' + field_name + '&type=' + type, 'AMediaAdd', 'width=300,height=150,left=200,top=200,scrollbars=yes,status=yes,location=no,resizable=yes,dependent');
					}	
					else if(type == "file")	
					{
						win.open('<%=Application("VirDir")%>/tinymce/plugins/advlink/AFileAdd.aspx?field_name=' + field_name + '&type=' + type, 'AFileAdd', 'width=300,height=150,left=200,top=200,scrollbars=yes,status=yes,location=no,resizable=yes,dependent');
					}	
					else
					{			
						win.open('<%=Application("VirDir")%>/tinymce/themes/advanced/AImgAdd.aspx?field_name=' + field_name + '&type=' + type, 'AImgAdd', 'width=300,height=150,left=200,top=200,scrollbars=yes,status=yes,location=no,resizable=yes,dependent');
					}	
			}
		}); 
		//-->
		</script>
	</HEAD>
	<body style="BACKGROUND-COLOR:#e6e6e6">
		<UC:logotop id="logotop" runat="server"></UC:logotop>
		<UC:topInclude ID="topInclude" runat="server"></UC:topInclude>
		<div align="center"><center>
		<asp:PlaceHolder ID="plhResultSuccess" Runat=server Visible=False>
		<table cellpadding="0" cellspacing="2" align="center" border="0">
			<tr>
							<td height="20" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2" align="center"><table border="0" cellpadding="0" cellspacing="0">
								
									<tr>
										<td align=center><span style="COLOR: #6F6DA6;font-size:16pt">המייל נשלח בהצלחה</span></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td height="20" colspan="2"></td>
						</tr>
						</table>
		</asp:PlaceHolder>
		<asp:PlaceHolder ID="plhResultError" Runat=server Visible=False>
		<table cellpadding="0" cellspacing="2" align="center" border="0">
			<tr>
							<td height="20" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2" align="center"><table border="0" cellpadding="0" cellspacing="0">
									<tr>
										<td align=center><span style="COLOR: #ff0000;font-size:16pt">אירעה שגיאה בשליחת מייל </span></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td height="20" colspan="2"></td>
						</tr>
						</table>
		</asp:PlaceHolder>
		<asp:PlaceHolder ID=plhForm Runat=server>
				<form id="Form1" method="post" runat="server" name="Form1">
					<table cellpadding="0" cellspacing="2" align="center" border="0">
						<tr>
							<td colspan="2" align="center"><table border="0" cellpadding="0" cellspacing="0">
									<tr>
										<td><input type="button" name="Button1" value="שלח  מייל" id="Button1" class="but_menu" onClick="return CheckForm()"></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td height="20" colspan="2"></td>
						</tr>
						<TR>
							<td align="left">From:</td>
							<td align="left"><input id="pFrom" name="pFrom" style="WIDTH:300px" readonly value="<%=FromEmail%>"></td>
						</TR>
						<tr><td colspan=2 height=15></td></tr>
						<TR>
							<td align="left" valign="top">To:</td>
							<td align="left"><table border=0 cellpadding=0 cellspacing=0>
							<tr><td align=center>עובד</td><td width=50>&nbsp;&nbsp;&nbsp;</td><td align=center>מחלקה</td></tr>
							<tr>
						
								<td align="left"  valign=top>
								<table border=0 cellpadding=0 cellspacing=0>
							<tr><td valign=top>	
								<select ID="SelSearchUsers" name="SelSearchUsers" multiple runat="server" style="HEIGHT: 120px;WIDTH: 300px;DIRECTION: rtl"
									AutoPostBack="false">
								</select></td></tr>
								<tr><td align=center height=10></td></tr></table>
							</td>
							<td width=50>&nbsp;&nbsp;&nbsp;</td>
							<td valign=top>
							<table border=0 cellpadding=2 cellspacing=2>
							<tr><td>		<select ID="SelSearchDep" name="SelSearchDep" multiple runat="server" style="HEIGHT: 120px;WIDTH: 300px;DIRECTION: rtl"
									AutoPostBack="false">
								</select>
							</td></tr>
							<tr><td align=center>
							<table border=0 cellpadding=0 cellspacing=0>
							<tr>
							<td><a class="button_edit_1" style="width:120px; line-height:110%; padding:3px" href="javascript:void(0)" onclick="javascript:ClearData()">אפס רשימת עובדים</a></td>
							<td width=20>&nbsp;&nbsp;&nbsp;</td>
							<td>	<a class="button_edit_1" style="width:90px; line-height:110%; padding:3px" href="javascript:void(0)" onclick="javascript:SendData()">עדכן עובדים</a></td></tr>
							
											
							</table>
							</td></TR>
							</table>
						</td>
							</tr>
							
							</table>
						</td>
						</TR>
					
						<tr>
							<td align="left">CC:</td>
							<td><input id="pCC" name="pCC" style="WIDTH:300px"></td>
						</tr>
						<tr>
							<td align="left">BCC:</td>
							<td><input id="pBCC" name="pBCC" style="WIDTH:300px"> <span style="COLOR:#ff0000">*(להפריד 
									את כתובות המייל באמצעות פסיק)</span></td>
						</tr>
						<TR>
							<td align="left">Subject:</td>
							<td align="left"><input id="pSubject" name="pSubject" style="WIDTH:300px"></td>
						</TR>
						<tr>
							<td align="left">Upload File</td>
							<td align="left"><input name="fileupload1" type="file" id="fileupload1" dir="rtl" runat="server"></td>
						</tr>
						<tr>
							<td align="right" colspan="2"><div style="WIDTH: 950px; TEXT-ALIGN: right"><textarea ID="sContent" name="sContent" class="input_text mceEditor" runat="server" style="PADDING-BOTTOM: 0px; PADDING-TOP: 0px; PADDING-LEFT: 0px; MARGIN: 0px; PADDING-RIGHT: 0px"></textarea></div>
							</td>
						</tr>
						<tr>
							<td height="20" colspan="2"></td>
						</tr>
						<tr>
							<td colspan="2" align="center"><table border="0" cellpadding="0" cellspacing="0">
									<tr>
										<td><input type="button" name="Button2" value="שלח  מייל" id="Button2" class="but_menu" onClick="return CheckForm()"></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td height="20" colspan="2"></td>
						</tr>
					</table>
				</form>
				</asp:PlaceHolder>
			</center>
		</div>
	</body>
</HTML>
