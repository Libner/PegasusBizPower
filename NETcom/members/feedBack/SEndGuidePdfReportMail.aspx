<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SEndGuidePdfReportMail.aspx.vb" Inherits="bizpower_pegasus2018.SEndGuidePdfReportMail"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>SEndGuidePdfReportMail</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

    	<script LANGUAGE="JavaScript">
<!--
function CheckFields(action)
{	
alert("CheckFields")
	if (!checkEmail(window.document.getElementById("sendermail").value) || window.document.getElementById("sendermail").value == "")
	{	str_alert = "כתובת דואר  אלקטרוני לא חוקית"
		
		window.alert(str_alert);
		document.frmMain.sendermail.focus();
		return false;
	}
	
	if (window.document.getElementById("ReplyText").value == '') {
              window.alert("! נא להכניס תוכן ");
              window.document.getElementById("ReplyText").focus();
              return false;
          }
      	
				if(window.document.getElementById("ReplyText").value.length > 5000)
				{
					window.alert("התוכן שהזנת הינו גדול ממספר התוים המקסימלי");
					return false;
				}
					
			
	return true;	
}


function checkEmail(addr)
{
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}

//-->
	</script>
  </head>
<body style="margin:0px; background-color:#E6E6E6" onload="window.focus();">
				<FORM name="frmMain" runat=server onSubmit="return CheckFields()" ID="frmMain" method="post">
	
					    <input type=hidden id="gId" name="gId" value="<%=gId%>">
						<input type=hidden id="gyear" name="gyear" value="<%=gyear%>">
								<table border="0" cellpadding="0" cellspacing="1" width="100%" dir="right" ID="Table1">
			<tr>
				<td align="left" valign="middle" nowrap>
					<table width="100%" border="0" cellpadding="0" cellspacing="0" ID="Table2">
						<tr>
							<td class="page_title" dir="rtl">&nbsp;שליחת מייל למדריך&nbsp;<%=Guide_Name%> </td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="100%">
					<table align="center" border="0" cellpadding="3" cellspacing="1" width="98%" align="center"
						ID="Table3">
						<tr>
							<td height="10"></td>
						</tr>
		
				
							<tr>
								<td align="right" width=100% valign=top>
									<span id="word7" name="word7"><!--כתובת אימייל למשלוח תגובה-->שם שולח אימייל</span>&nbsp;
									<input dir="rtl" type="text" class="texts" style="width:285" id="sender_name" name="sender_name" value="<%=selUserName%>" maxlength=50>
								<input type=hidden id="senderId" name="senderId" value="<%=Request.Cookies("bizpegasus")("UserId")%>">
								</td>
								<td align="right" nowrap width="70">&nbsp;<span id="word8" name="word8">מאת</span>&nbsp;</td>
							</tr>
							<tr>
								<td align="right" width=100% valign=top>כתובת שממנה יוצאת אימייל&nbsp; <input id="sendermail" name="sendermail" dir="ltr" class="texts" style="width:285;" value="<%=EMAIL%>">
								</td>
								<td align="right" nowrap width="70">&nbsp;<span id="word9" name="word9"><!--מאת-->מאת</span>&nbsp;</td>
							</tr>
							<tr>
								<td align="right" width=100% valign=top>&nbsp; <input id="Guide_Email" name="Guide_Email" dir="ltr" class="texts" style="width:285;" value="<%=Guide_Email%>">
								</td>
								<td align="right" nowrap width="70">&nbsp;<span id="Span1" name="word9"><!--מאת-->אל</span>&nbsp;</td>
							</tr>
							
							<TR>
								<td align="right" width=100%>
									<textarea dir="rtl" type="text" class="texts" style="width:450" rows=1 id="SubjectText" name="SubjectText">דוח שנתי למדרך <%=Guide_Name%></textarea>
								</td>
								<td align="right" nowrap width="70">&nbsp;<!--תוכן תגובה--> נושא&nbsp;</td>
							</TR>
						
							<tr valign="top">
								<td align="right" width=100%>
									<textarea dir="rtl" type="text" class="texts" style="width:450" rows=9 id="ReplyText" name="ReplyText"></textarea>
								</td>
								<td align="right" nowrap width="70">&nbsp;תוכן אימייל&nbsp;</td>
							</tr>
						<tr>
								<td  dir="rtl"><input type="file" name="attachment_file" ID="attachment_file" size="33" value=""></td>
								<td align="right" nowrap width="70">&nbsp;מסמך מצורף</td>
							</tr>
							<tr>
								<td colspan="2" height="10"></td>
							</tr>	
							<tr>
								<td align=right>
									<table cellpadding="0" cellspacing="0"  width="50%" align="right" ID="Table4" border=0>
										<tr valign="top">
											<td width="28%" align="right"><A class="but_menu" href="#" style="width:100px;Height:20" onclick="window.close();"><span id="word4" name="word4">ביטול</span></A></td>
											<td width="10" nowrap></td><td width="28%" align="left" nowrap><asp:Button Runat=server cssclass="but_menu" ID=btnSend Height=20  Text="שלח אימייל"></asp:Button></td>
										</tr>
									</table>
								</td>
								<td align="right" nowrap width="70">&nbsp;</td>
							</tr>
					
						<tr>
							<td colspan="2" height="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
			</FORM>
  </body>
</html>
