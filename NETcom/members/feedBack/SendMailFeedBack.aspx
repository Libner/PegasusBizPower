<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SendMailFeedBack.aspx.vb" Inherits="bizpower_pegasus2018.SendMailFeedBack"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SendEmail</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<link href="http://pegasus.bizpower.co.il/Netcom/IE4.css" rel="STYLESHEET" type="text/css">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
<!--
      function textLimit(field, maxlen) {
          if (field.value.length > maxlen + 1)
              alert('הקלט נחתך!');
          if (field.value.length > maxlen)
              field.value = field.value.substring(0, maxlen);
      }
      function closeWin() {
          opener.focus();
          self.close();
      }

      function CheckFields() {

          /*   if(window.document.formSendEmail.Email_type_id.value == '')
          {
          window.alert("! נא בחר סוג ההודעה");
          window.document.formSendEmail.Email_type_id.focus();
          return false;
          }*/

          if (window.document.getElementById("FromEmail").value == '') {
              window.alert("! נא להכניס טלפון שולח ");
              window.document.getElementById("FromEmail").focus();
              return false;
          }
          /*if (document.getElementById("FromEmail").value.length!=10)
          {
          window.alert('!נא למלא טלפון 10 ספרות');
          document.getElementById("FromEmail").focus();
          return false;

          }
          var phoneTxt=document.getElementById("FromEmail").value
          if(!/^(0[23489]([- ]{0,1})[^0\D]{1}\d{6})$|^(05[0247]{1}([- ]{0,1})[0-9]{7})$/.test(phoneTxt))
          {
          window.alert("מספר טלפון לא חוקי");
          window.document.getElementById("FromEmail").focus()
          return false;
          }
          else
          {
          if(/^(050)|(052)|(054)|(057)/.test(phoneTxt))
          isMobile=true
          else
          return false;
				
          }
          */
        /* if (window.document.getElementById("Email_content").value == '') {
              window.alert("! נא להכניס  טקסט ההודעה");
              window.document.getElementById("Email_content").focus();
              return false;
          }*/
          document.getElementById("formSendEmail").submit();
          return true;
      }



      function checkTel(cell) {
      }
//-->
		</script>
	</HEAD>
	<body style="BACKGROUND:#e6e6e6;MARGIN:0px">
		<form id="formSendEmail" name="formSendEmail" method="post" runat="server">
			<input type=hidden id=DepartureId name=DepartureId value="<%=DepartureId%>">
			<table border="0" cellpadding="1" cellspacing="0" width="100%" align="center" dir="ltr"
				ID="Table2">
				<TBODY>
					<tr>
						<td align="center" colspan="2" dir="rtl"><b>סה''כ מייל לשליחה
								<%=CountMail%>
							</b>
						</td>
					</tr>
					<tr>
						<td align="right" width="100%">
							שליחת קישור לטופס משוב
						</td>
						<td align="right" width="140" nowrap><b><!--תאריך--> סוג ההודעה
								<br>
								(לצורך תיעוד בלבד)</b></td>
					</tr>
					<tr>
						<td align="right" width="100%" valign="top"><input dir="rtl" class="texts" 
	style="WIDTH:150px;COLOR:#000000;BACKGROUND-COLOR:#c0c0c0" id="FromEmail" name="FromEmail" value="<%=FromEmail%>"  readonly maxlength="100"></td>
						<td align="right" width="140" nowrap dir="rtl"><b>מייל שולח</b></td>
					</tr>
					<%If False Then%>
					<tr>
						<td align="right" nowrap dir="rtl">
							<textarea id="Email_content" name="Email_content" dir="rtl" onkeyup="textLimit(this, <%=EmailText%>);" style="HEIGHT:40px;WIDTH:250px;LINE-HEIGHT:120%" class="Form" >&lt;%=Email_content%&gt;</textarea>
						</td>
						<td align="right" width="140" nowrap valign="top"><b>תוכן ההודעה<BR>
								(<%=EmailText%>
								תווים )</b></td>
					</tr>
					<%end if%>
					<tr>
						<td align="center" colspan="2" height="20" nowrap></td>
					</tr>
					<tr>
						<td align="center" colspan="2">
							<table cellpadding="0" cellspacing="0" width="100%" ID="Table7">
								<tr>
									<td align="center" width="50%">
										<input type="button" value="ביטול" class="but_menu" style="WIDTH:120px" onclick="closeWin();"
											ID="Button2" NAME="Button2">
									</td>
									<td align="center" width="50%">
										<A class="but_menu" href="#" style="WIDTH:120px;PADDING-BOTTOM:4px;PADDING-TOP:4px;PADDING-LEFT:4px;LINE-HEIGHT:120%;PADDING-RIGHT:4px"
											onClick="return CheckFields()"><!--הוסף--> שלח</A>
									</td>
								</tr>
							</table>
		</form>
		</TD></TR></TBODY></TABLE>
	</body>
</HTML>
