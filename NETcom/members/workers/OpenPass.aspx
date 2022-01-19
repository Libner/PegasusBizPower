<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OpenPass.aspx.vb" Inherits="bizpower_pegasus2018.OpenPass"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OpenPass</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script>
		function CheckFields(p)
		{
	
		 if (!checkPassword(document.getElementById("pass").value) && document.all("pass").value != "")
		{
        
				window.alert("סיסמה לא חוקית\n על כל סיסמה להיות בנויה מ 8 תוים המכילים בתוכם לפחות אות גדולה אחת, לפחות מספר אחד ולפחות תו אחד (!, @, #, וכו)");
				document.getElementById("pass").focus();
				return false;
		}	
		document.getElementById("Form1").submit();
		return true;
		}
		function checkPassword(pass)
{
 var str = pass;
    if (str.length < 8) {
        alert("על כל סיסמה להיות בנויה ממינימום 8 תוים");
        return false;

}
var validChars=/^(?=.*?[A-Z])(?=.*?[a-z])(?=.*?[0-9])(?=.*?[^\w\s]).{8,}$/
return (validChars.test(pass))
}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<table border="0" width="480" cellspacing="0" cellpadding="0" align="center" dir="rtl">
			<tr>
				<td align="left" valign="middle" nowrap>
					<table width="100%" border="0" cellpadding="1" cellspacing="0">
						<tr>
							<td class="page_title" align="center" valign="middle" width="100%"><!--עדכון--> 
								&nbsp;עדכון סיסמה לחוזי עבודה &nbsp;</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td height="15" nowrap></td>
			</tr>
			<tr>
				<td width="100%">
					<form id="Form1" name="Form1" method="post" runat="server">
						<tr>
							<td width="100%"><input type=password id="pass" name="pass"  value="<%=pass%>"></td>
						</tr>
						<TR><td height=10>&nbsp;</td></TR>
						<TR><td><input type="button" class="but_menu" style="width:90px" onclick="javascript:CheckFields(this); return false;" value="עדכן סיסמה" id="Button4" name="Button1"></td></TR>
					</form>
				</td>
			</tr>
		</table>
	</body>
</HTML>
