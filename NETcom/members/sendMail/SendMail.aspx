<%@ Register TagPrefix="UC" TagName="topInclude" Src="~/include/topInclude.ascx"   %>
<%@ Register TagPrefix="UC" TagName="logotop" Src="~/include/logotop.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SendMail.aspx.vb" Inherits="bizpower_pegasus2018.SendMail1" validateRequest="false"%>
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
		<%if false then%>
		<UC:logotop id="logotop" runat="server"></UC:logotop>
		<UC:topInclude ID="topInclude" runat="server"></UC:topInclude>
		<%end if%>
		<div align="center"><center>
				<asp:PlaceHolder ID="plhResultSuccess" Runat="server" Visible="False">
					<TABLE cellSpacing="2" cellPadding="0" align="center" border="0">
						<TR>
							<TD height="20" colSpan="2"></TD>
						</TR>
						<TR>
							<TD colSpan="2" align="center">
								<TABLE cellSpacing="0" cellPadding="0" border="0">
									<TR>
										<TD align="center"><SPAN style="FONT-SIZE: 16pt; COLOR: #6f6da6">המייל נשלח בהצלחה</SPAN></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR>
							<TD height="20" colSpan="2"></TD>
						</TR>
					</TABLE>
				</asp:PlaceHolder>
				<asp:PlaceHolder ID="plhResultError" Runat="server" Visible="False">
					<TABLE cellSpacing="2" cellPadding="0" align="center" border="0">
						<TR>
							<TD height="20" colSpan="2"></TD>
						</TR>
						<TR>
							<TD colSpan="2" align="center">
								<TABLE cellSpacing="0" cellPadding="0" border="0">
									<TR>
										<TD align="center"><SPAN style="FONT-SIZE: 16pt; COLOR: #ff0000">אירעה שגיאה בשליחת 
												מייל </SPAN>
										</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR>
							<TD height="20" colSpan="2"></TD>
						</TR>
					</TABLE>
				</asp:PlaceHolder>
				<asp:PlaceHolder ID="plhForm" Runat="server">
					<FORM id="Form1" method="post" name="Form1" runat="server">
						<TABLE cellSpacing="2" cellPadding="0" align="center" border="0">
							<TR>
								<TD colSpan="2" align="center">
									<TABLE cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><INPUT onclick="return CheckForm()" id="Button1" class="but_menu" type="button" value="שלח  מייל"
													name="Button1"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD height="20" colSpan="2"></TD>
							</TR>
							<TR>
								<TD align="left">From:</TD>
								<TD align="left"><INPUT id=pFrom style="WIDTH: 300px" readOnly 
      value="<%=FromEmail%>" name=pFrom></TD>
							</TR>
							<TR>
								<TD height="15" colSpan="2"></TD>
							</TR>
							<TR>
								<TD vAlign="top" align="left">To:</TD>
								<TD align="left">
									<TABLE cellSpacing="0" cellPadding="0" border="0">
										<%if false then%>
										<TR>
											<TD align="center">עובד</TD>
											<TD width="50">&nbsp;&nbsp;&nbsp;</TD>
											<TD align="center">מחלקה</TD>
										</TR>
										<%end if%>
										<TR>
											<TD vAlign="top" align="left">
												<TABLE cellSpacing="0" cellPadding="0" border="0">
													<TR>
														<TD vAlign="top"><SELECT id="SelSearchUsers" style="HEIGHT: 120px; WIDTH: 300px; DIRECTION: rtl" multiple
																name="SelSearchUsers" runat="server" AutoPostBack="false"></SELECT></TD>
													</TR>
													<TR>
														<TD height="10" align="center"></TD>
													</TR>
												</TABLE>
											</TD>
											<%if false then%>
											<TD width="50">&nbsp;&nbsp;&nbsp;</TD>
											<TD vAlign="top">
												<TABLE cellSpacing="2" cellPadding="2" border="0">
													<TR>
														<TD><SELECT id="SelSearchDep" style="HEIGHT: 120px; WIDTH: 300px; DIRECTION: rtl" multiple name="SelSearchDep"
																runat="server" AutoPostBack="false"></SELECT>
														</TD>
													</TR>
													<TR>
														<TD align="center">
															<TABLE cellSpacing="0" cellPadding="0" border="0">
																<TR>
																	<TD><A onclick="javascript:ClearData()" class="button_edit_1" style="WIDTH: 120px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px; PADDING-LEFT: 3px; LINE-HEIGHT: 110%; PADDING-RIGHT: 3px"
																			href="javascript:void(0)">אפס רשימת עובדים</A></TD>
																	<TD width="20">&nbsp;&nbsp;&nbsp;</TD>
																	<TD><A onclick="javascript:SendData()" class="button_edit_1" style="WIDTH: 90px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px; PADDING-LEFT: 3px; LINE-HEIGHT: 110%; PADDING-RIGHT: 3px"
																			href="javascript:void(0)">עדכן עובדים</A></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>
												<%end if%>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD align="left">CC:</TD>
								<TD><INPUT id="pCC" style="WIDTH: 300px" name="pCC"></TD>
							</TR>
							<TR>
								<TD align="left">BCC:</TD>
								<TD><INPUT id="pBCC" style="WIDTH: 300px" name="pBCC"> <SPAN style="COLOR: #ff0000">*(להפריד 
										את כתובות המייל באמצעות פסיק)</SPAN></TD>
							</TR>
							<TR>
								<TD align="left">Subject:</TD>
								<TD align="left"><INPUT id="pSubject" style="WIDTH: 300px" name="pSubject"></TD>
							</TR>
							<TR>
								<TD align="left">Upload File</TD>
								<TD align="left"><INPUT id="fileupload1" dir="rtl" type="file" name="fileupload1" runat="server"></TD>
							</TR>
							<TR>
								<TD colSpan="2" align="right">
									<DIV style="WIDTH: 950px; TEXT-ALIGN: right"><TEXTAREA id="sContent" class="input_text mceEditor" style="PADDING-BOTTOM: 0px; PADDING-TOP: 0px; PADDING-LEFT: 0px; MARGIN: 0px; PADDING-RIGHT: 0px"
											name="sContent" runat="server"></TEXTAREA></DIV>
								</TD>
							</TR>
							<TR>
								<TD height="20" colSpan="2"></TD>
							</TR>
							<TR>
								<TD colSpan="2" align="center">
									<TABLE cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><INPUT onclick="return CheckForm()" id="Button2" class="but_menu" type="button" value="שלח  מייל"
													name="Button2"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD height="20" colSpan="2"></TD>
							</TR>
						</TABLE>
					</FORM>
				</asp:PlaceHolder>
			</center>
		</div>
	</body>
</HTML>
