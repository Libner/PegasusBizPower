<%@ Page Language="vb" AutoEventWireup="false" Codebehind="VerificationPage.aspx.vb" Inherits="bizpower_pegasus2018.VerificationPage"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>VerificationPage</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="/dynamic_style.css" rel="STYLESHEET" type="text/css"><script>
function checkForm()
{
       if (document.getElementById("VCode").value == '') //&& document.getElementById("sTo").value == '')
            {
                alert("אנא מלא שדה ");
                document.getElementById("VCode").focus()
                return false;
            }
            
		    document.getElementById("Form1").submit();
            return true;

}
</script>
  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tbody><tr>
          <td valign="middle"><a href="/default.asp"><img src="/images/top_logo.gif" width="223" height="59" border="0"></a></td>
          <td width="100%" align="right" dir="rtl" valign="bottom">&nbsp;</td>
          <td valign="middle" align="right">
          <table border="0" width="195" cellspacing="0" cellpadding="0">
            <tbody><tr>
              <td width="100%" height="10"></td>
            </tr>
            <tr>
              <td width="100%" valign="middle">
			
         </td></tr>
    </tbody></table></td></tr>
    </tbody></table>
    <table border=0 align=center>
    <tr><td class='titleB' align=center>'אנא הקלד את הקוד האימות שנשלח לסלולרי שלך ולחץ 'אישור</td></tr>
    <tr><td height=50></td></tr>
    <tr><td align=center  height="21"><input type="text" name="VCode" id="VCode" class="form_text_home" style="height:21px;" onkeypress='return event.charCode >= 48 && event.charCode <= 57' ></td></tr>
        <tr><td height=20></td></tr>  
         <tr>
              <td align=center height="21"><input type=submit value="אישור" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</td>
            </tr></td></tr>
            <asp:PlaceHolder ID=plhError Visible=False Runat=server>
            <tr><td height=25></td></tr>
            <tr><td class='titleB'  style="color:#ff0000" align=center> אנא הקלד את הקוד האימות הנכון</td></tr>
            </asp:PlaceHolder>
    </table>
    </form>

  </body>
</html>
