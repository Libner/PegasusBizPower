<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SendSMSFeedBack.aspx.vb" Inherits="bizpower_pegasus2018.SendSMSFeedBack" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>SendSMS</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <link href="http://pegasus.bizpower.co.il/Netcom/IE4.css" rel="STYLESHEET" type="text/css">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
  <script language="javascript" type="text/javascript">
<!--
      function textLimit(field, maxlen) {
          if (field.value.length > maxlen + 1)
              alert('���� ����!');
          if (field.value.length > maxlen)
              field.value = field.value.substring(0, maxlen);
      }
      function closeWin() {
          opener.focus();
          self.close();
      }

      function CheckFields() {

          /*   if(window.document.formSendSMS.SMS_type_id.value == '')
          {
          window.alert("! �� ��� ��� ������");
          window.document.formSendSMS.SMS_type_id.focus();
          return false;
          }*/

          if (window.document.getElementById("sms_phone").value == '') {
              window.alert("! �� ������ ����� ���� ");
              window.document.getElementById("sms_phone").focus();
              return false;
          }
          /*if (document.getElementById("sms_phone").value.length!=10)
          {
          window.alert('!�� ���� ����� 10 �����');
          document.getElementById("sms_phone").focus();
          return false;

          }
          var phoneTxt=document.getElementById("sms_phone").value
          if(!/^(0[23489]([- ]{0,1})[^0\D]{1}\d{6})$|^(05[0247]{1}([- ]{0,1})[0-9]{7})$/.test(phoneTxt))
          {
          window.alert("���� ����� �� ����");
          window.document.getElementById("sms_phone").focus()
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
        /* if (window.document.getElementById("SMS_content").value == '') {
              window.alert("! �� ������  ���� ������");
              window.document.getElementById("SMS_content").focus();
              return false;
          }*/
          document.getElementById("formSendSMS").submit();
          return true;
      }



      function checkTel(cell) {
      }
//-->
</script>  
  </head>
  <body style="margin:0px;background:#e6e6e6">

    <form id="formSendSMS" name="formSendSMS" method="post" runat="server">
<input type=hidden id=DepartureId name=DepartureId value="<%=DepartureId%>">
<table border="0" cellpadding="1" cellspacing="0" width="100%" align=center dir="ltr" ID="Table2">
  <tr><td align=center colspan=2  dir=rtl><b>��''� SMS ������ <%=CountSms%></b></td></tr>
   <tr>
	<td align="right" width=100%>
	����� ����� ����� ����
</td>
	<td align="right" width=140 nowrap><b><!--�����-->��� ������ <br>(����� ����� ����)</b></td>
   </tr>
<tr>
	<td align="right" width="100%" valign="top"><input dir="rtl" type="text" class="texts" 
	style="width:100;background-color:#c0c0c0;color:#808080" id="sms_phone" name="sms_phone" value="<%=sms_phone%>"  readonly maxlength="10"></td>
	<td align="right" width="140" nowrap dir="rtl"><b>����� ����</b></td>
</tr>
<%If False Then%>	  <tr>             
	<td align="right" nowrap  dir="rtl">
	<textarea id="SMS_content" name="SMS_content" dir="rtl" onkeyup="textLimit(this, <%=smsText%>);" style="height:40;width:250;line-height:120%;" class="Form" ><%=SMS_content%></textarea>
	</td>	
	<td align="right" width=140 nowrap valign=top><b>���� ������<BR>(<%=smsText%> ����� )</b></td>
	</tr> 
    <%end if%>
	<tr><td align=center colspan="2" height=20 nowrap></td></tr>
<tr><td align=center colspan="2">
<table cellpadding=0 cellspacing=0 width=100% ID="Table7">
<tr>
<td align=center width=50%>
<input type=button value="�����" class="but_menu" style="width:120" onclick="closeWin();" ID="Button2" NAME="Button2">
</td>
<td align=center width=50%>
<A class="but_menu" href="#" style="width:120;line-height:120%;padding:4px" onClick="return CheckFields()"><!--����-->���</a>
</td>
</tr>
</table>
    </form>

  </body>
</html>
