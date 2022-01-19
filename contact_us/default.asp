<!--#include file="../include/connect.asp"-->
<!--#include file="../include/reverse.asp"-->
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<!--#include file="../include/title_meta_inc.asp"-->
<title>Bizpower - טופס הצטרפות</title>
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function check_fields(formObj) 
{
	
	if (emptyField(formObj.name))
	{
		alert('אנא מלא שם פרטי ומשפחה');
		formObj.name.focus();
		return false;	
	}	
	if (emptyField(formObj.email))
	{
		alert('אנא מלא דואר אלקטרוני');
		formObj.email.focus();
		return false;	
	}			
	if (emptyField(formObj.phone))
	{
		alert('אנא מלא מספר טלפון');
		formObj.phone.focus();
		return false;	
	}					

}  

function emptyField(textObj)
{
	if (textObj.value.length == 0) return true;
	for (var i=0; i<textObj.value.length; i++) {
		var ch = textObj.value.charAt(i);
		if (ch != ' ' && ch != '\t') return false;	
	}
	return true;	
} 

//-->
</SCRIPT>
</head>
<body marginwidth="0" marginheight="0" topmargin="5" leftmargin="0" rightmargin="0" bgcolor="white">
<!--#include file="../include/top.asp"-->
<table dir="rtl" border="1" style="border-collapse:collapse" bordercolor="#999999" cellspacing="0" cellpadding="0" width="780" align="center" ID="Table5">
	<tr><td colspan=2 bgcolor="#FFFFFF" align="center" width="780">
	<div align="center">
			<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
				ID="Shockwaveflash1" VIEWASTEXT width=780 height="175">
				<PARAM NAME="movie" VALUE="../flash/BP-flash-test.swf">
				<PARAM NAME="quality" VALUE="best">
				<PARAM NAME="wmode" VALUE="transparent">
				<PARAM NAME="bgcolor" VALUE="#EBEBEB">					
				<EMBED src="../flash/BP-flash-test.swf" quality="best" wmode="transparent" bgcolor="#FFFFFF"
					TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
					width="780" height="175">
				</EMBED>
			</OBJECT>
	</div>				
	</td></tr>
	<tr>
			<td  align="right" class="title_page" dir="rtl"	style="padding:5px;padding-right:10px">
			טופס הצטרפות													
			</td>	
			<td rowspan=2 align="right" valign="top" width="180" nowrap bgcolor="#FFFFFF">
		    <!--#include file="../include/ashafim_inc.asp"-->
		    </td>											
		</tr>			
	<tr>
		<td width="600" nowrap bgcolor="#F9F9F9" valign="top" align="right" style="padding-right:20px; padding-left:20px;">
        <table border="0" cellspacing="0" cellpadding="0" width="480" dir="ltr"> 
			<%'//start of centent%>
				<tr>
					<td width="100%" valign="top" align="right" height="15"></td>
				</tr>							
				<tr>
					<td align="right" width="100%">
					<table width="100%" align="right" border="0" cellspacing="0" cellpadding="0">
					<tr>
					<td width="100%">
					<%'//start of content%>
					<form style="margin:0px" name="myform" action="thanks.asp" method="post" onsubmit="return check_fields(this)" ID="myform">					
					<table width=100% align="center" border="0">
					<tr>
					<td width="70%" nowrap align="right"><input type="text" class="form_text" dir="rtl" name="name" id="name" style="width:255px"></td>
					<td width="30%" nowrap align="right" dir="rtl" nowrap class="forms"><font color="red">*</font>&nbsp;שם פרטי ומשפחה:</td>
					</tr>
					<tr>
					<td align="right"><input type="text" class="form_text" dir="ltr" name="email" id="email" style="width:255px"></td>
					<td align="right" width="30%" dir="rtl" class="forms"><font color="red">*</font>&nbsp;דואר אלקטרוני:</td>
					</tr>
					<tr>
					<td align="right" ><input type="text" class="form_text" dir="ltr" name="phone" id="phone" style="width:255px"></td>
					<td align="right" width="30%" dir="rtl" class="forms"><font color="red">*</font>&nbsp;מספר טלפון:</td>
					</tr>
					<tr>
					<td align="right"><input type="text" class="form_text" dir="rtl" name="company" id="company" style="width:255px"></td>
					<td  align="right" width="30%" dir="rtl" class="forms">&nbsp;&nbsp;שם חברה:&nbsp;</td>
					</tr>
					<tr>
					<td  align="right" ><input type="text" class="form_text" dir="rtl" name="tafkid" id="tafkid" style="width:255px"></td>
					<td width="30%" align="right" dir="rtl" class="forms">&nbsp;&nbsp;תפקיד:&nbsp;</td>											
					</tr>
					<tr>
					<td align=right><input dir="rtl" name="address" id="address" class="form_text" style="width:255px"></td>
					<td  align="right" width="30%"  dir="rtl" class="forms">&nbsp;&nbsp;כתובת:&nbsp;</td>
					</tr>
  					<tr>
						<td  align="right" ><textarea dir="rtl" class="form_text" name="comments" id="comments" rows="7"  style="width:255px"></textarea></td>
						<td align="right" width="30%" dir="rtl" valign="top" class="forms">&nbsp;&nbsp;תוכן הפניה:&nbsp;</td>
					</tr>		
					<tr><td colspan="2" align="right" dir="rtl"><font color="red">*</font> - שדה חובה</td></tr>
					<tr>					
					<td  align="right">
					<INPUT TYPE="submit" class="button" NAME="send" Value="&nbsp;שלח&nbsp;" ID="Submit1">
					</td>
				    <td></td>
				    </tr>
				    <tr><td height="10" nowrap></td></tr>
				    </table> 
			</form>						
		<%'//end of content%>																			
	    </td></tr></table></td></tr></table></td>									
	</tr>					
</table>

<!--#include file="../include/bottom.asp"-->

</body>
</html>
