<%@ Page Language="vb" EnableViewState="true" AutoEventWireup="false" validateRequest="false" Codebehind="addResponse.aspx.vb" Inherits="bizpower_pegasus2018.addResponse" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>טופס מתעניין בטיול</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
		<script language="javascript" type="text/javascript" charset="windows-1255">

		function checkForm()
		{
			if(window.document.getElementById("Response_Title").value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס מענה "
				Else
					str_alert = "Please insert the reply !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.getElementById("Response_Title").focus();
				return false;
			}
			if(window.document.getElementById("Response_Content").value.length > 5000)
			{
				window.alert("התוכן שהזנת הינו גדול ממספר התוים המקסימלי");
				return false;
			}
			if(!checkFile("fileuploadF1"))
			{
					return false;
			}
			if(!checkFile("fileuploadF2"))
			{
					return false;
			}
			if(!checkFile("fileuploadF3"))
			{
					return false;
			}
			
			return true;			
		}

		function checkFile(fileId)
		{
		if (window.document.getElementById(fileId).value  != '')
			{
				var filename = new String(window.document.getElementById(fileId).value);
				filename = filename.toLowerCase();
				if (
				(filename.lastIndexOf(".doc") == -1) && (filename.lastIndexOf(".pdf") == -1)
				&& (filename.lastIndexOf(".rtf") == -1)  && (filename.lastIndexOf(".ppt") == -1)
				&& (filename.lastIndexOf(".docx") == -1) && (filename.lastIndexOf(".pptx") == -1)
				&& (filename.lastIndexOf(".gif") == -1) && (filename.lastIndexOf(".jpg") == -1)
				&& (filename.lastIndexOf(".jpeg") == -1)  && (filename.lastIndexOf(".bmp") == -1)
				&& (filename.lastIndexOf(".png") == -1)
				)
				{
					window.alert("סיומת קובץ אינה חוקית, אנא בחר קובץ חוקי");
					window.document.getElementById(fileId).focus();
					return false;
				}
			}
			return true				
		}
		</script>
	</HEAD>
	<body style="MARGIN:0px; BACKGROUND-COLOR:#e6e6e6">
		<div align="center">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td align="left" valign="middle" nowrap>
						<table width="100%" border="0" cellpadding="1" cellspacing="0">
							<tr>
								<td class="page_title" valign="middle" width="100%"><%If Len(responseID) > 0 Then%><span id="word1" name="word1"><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name="word2"><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id="word3" name="word3"><!--מענה--><%=arrTitles(3)%></span>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td height="15" nowrap></td>
				</tr>
				<tr>
					<td width="100%">
						<table width="100%" cellspacing="1" cellpadding="2" align="center" border="0">
							<form id="Form1" method="post" runat="server">
								<input type=hidden name=responseID id=responseID value="<%=responseID%>"> <input type=hidden name=prodID id="prodID" value="<%=prodID%>">
								<TBODY>
									<tr>
										<td width="100%" valign="top">
											<input class="texts" name="Response_Title" id="Response_Title" value="<%=func.vFix(Response_Title)%>" style="WIDTH:450px" maxLength=100>
										</td>
										<td width="70" nowrap valign="top">&nbsp;<b><span id="word4" name="word4"><!--כותרת--><%=arrTitles(4)%></span></b>&nbsp;</td>
									</tr>
									<tr>
										<td width="100%" valign="top">
											<textarea class="texts" name="Response_Content" id="Response_Content" style="WIDTH:450px"
												rows="9"><%=(Response_Content)%></textarea>
										</td>
										<td width="70" nowrap valign="top">&nbsp;<b><span id="word5" name="word5"><!--תוכן--><%=arrTitles(5)%></span></b>&nbsp;</td>
									</tr>
									<tr>
										<td height="5" colspan="2" nowrap></td>
									</tr>
									<tr>
										<td width="100%" valign="top">
											<a ID="aFile1" name="aFile1" Runat="server" target="_blank"></a>
											<asp:Button CausesValidation="False" OnClick="Delete_File" CssClass="button_small" ID="btnDeleteFile1"
												Width="85" Runat="server" Text="מחק קובץ" Visible="False" EnableViewState="false" CommandArgument="1"></asp:Button>
											<input id="fileuploadF1" style="WIDTH:280px" type="file" name="fileuploadF1" runat="server"
												EnableViewState="false" class="input_form">
										</td>
										<td width="70" nowrap valign="top">&nbsp;קובץ מצורף&nbsp;</td>
									</tr>
									<tr>
										<td height="5" colspan="2" nowrap></td>
									</tr>
									<tr>
										<td width="100%" valign="top">
											<a ID="aFile2" name="aFile2" Runat="server" target="_blank"></a>
											<asp:Button CausesValidation="False" OnClick="Delete_File" CssClass="button_small" ID="btnDeleteFile2"
												Width="85" Runat="server" Text="מחק קובץ" Visible="False" EnableViewState="false" CommandArgument="2"></asp:Button>
											<input id="fileuploadF2" style="WIDTH:280px" type="file" name="fileuploadF2" runat="server"
												EnableViewState="false" class="input_form">
										</td>
										<td width="70" nowrap valign="top">&nbsp;קובץ מצורף&nbsp;</td>
									</tr>
									<tr>
										<td height="5" colspan="2" nowrap></td>
									</tr>
									<tr>
										<td width="100%" valign="top">
											<a ID="aFile3" name="aFile3" Runat="server" target="_blank"></a>
											<asp:Button CausesValidation="False" OnClick="Delete_File" CssClass="button_small" ID="btnDeleteFile3"
												Width="85" Runat="server" Text="מחק קובץ" Visible="False" EnableViewState="false" CommandArgument="3"></asp:Button>
											<input id="fileuploadF3" style="WIDTH:280px" type="file" name="fileuploadF3" runat="server"
												EnableViewState="false" class="input_form">
										</td>
										<td width="70" nowrap valign="top">&nbsp;קובץ מצורף&nbsp;</td>
									</tr>
									<tr>
										<td align="center" colspan="2">
											<input type=button value="<%=arrButtons(8)%>" class="but_menu" style="WIDTH:90px" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											<input runat="server" type="submit" id="btnSubmit" name="btnSubmit" class="but_menu" style="WIDTH:90px"
												onclick="return checkForm()" value="Submit Query">
										</td>
									</tr>
							</form>
						</table>
					</td>
				</tr>
				</TBODY></table>
		</div>
	</body>
</HTML>
