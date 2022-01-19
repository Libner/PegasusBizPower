<script language="javascript" type="text/javascript">
<!--
function checkFields(objForm)
{
	if(document.form1.username.value == '')
	{
		window.alert("!!נא להכניס שם משתמש");
		document.form1.username.select();
		return false;
	}
	if(document.form1.password.value == '')
	{
		window.alert("!!נא להכניס סיסמה");
		document.form1.password.select();
		return false;
	}
	return true;
}
//-->
</script>
<%maincat=1%>
<!--#include file="../PopUpMenus/java_layer.asp"-->
<table border="0" width="780" cellspacing="0" cellpadding="0" align=center>
   <tr>
      <td width="100%" valign="middle">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="middle"><a href="<%=Application("VirDir")%>/default.asp"><img src="<%=Application("VirDir")%>/images/top_logo.gif" width="223" height="59" border="0"></a></td>
          <td width="100%" align="right" dir="rtl" valign="bottom">&nbsp;</td>
          <td valign="middle" align="right">
          <table border="0" width="195" cellspacing="0" cellpadding="0">
            <tr>
              <td width="100%" height="10"></td>
            </tr>
            <tr>
              <td width="100%" valign=middle>
			  <FORM style="margin-middle: 0" method="POST" action="<%=Application("VirDir")%>/netCom/default.asp" name=form1 id=form1 onSubmit="return checkFields(this)">									
              <table border="0" width="185" cellspacing="0" cellpadding="0" align=right>				
                <tr>
                  <td align="right" colspan=2><input type=hidden name="wizard_id" id="wizard_id" value="<%=Request("wizard_id")%>">
                  <input type="text" name="username" id="username" class="form_text_home" >
                  </td>
                  <td width="60" class="lhome" align="right"><b>משתמש</b></td>
                 </tr>
                 <tr> 
                  <td valign="bottom" nowrap width="40" align=right>
                  <input type="image" src="<%=Application("VirDir")%>/images/send2.gif" id="btnLogin" name="btnLogin" tabindex="-1">                  
                  </td>                  
                  <td align="right" nowrap width="90" >
                  <input type="password" name="password" id="password" class="form_text_home">
                  </td>
                  <td nowrap class="lhome" align="right"><b>סיסמא</b></td>
                </tr>				         
              </table></form>       
         </td></tr>
    </table></td></tr>
    </table></td></tr></table>
    <!--#include file="../include/top_bar_links.asp"-->