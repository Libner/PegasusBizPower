<%@ Page EnableViewState="false" CodeBehind="upload.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="bizpower_pegasus.Upload" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>Administration</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <LINK href="../../IE4.css" type="text/css" rel="STYLESHEET">
	<script language="JavaScript" type="text/javascript">
	<!--
	function ifFieldEmpty(btnSubmit)
	{
	  if (document.form1.elements("UploadFile").value=='')
		{
			alert("נא לבחור קובץ חוקי Excel !");
			return false;
		}
	  else
	  {
		var fname=new String(document.form1.elements("UploadFile").value)
		ext=fname.substring(fname.lastIndexOf('.')+1,fname.length)
		if ((ext != 'xls') && (ext != 'txt'))
		{
			alert("נא לבחור קובץ חוקי Excel !");
			return false;
		}
		else
		{
			btnSubmit.style.color = "red";  btnSubmit.value = "אנא המתן ..";  btnSubmit.disabled = true;
			document.form1.submit();
			return false;
		}
	  }
	}
	//-->
	</script>
</head>      
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus();" style="margin: 0px;">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td class="form_header" dir="rtl" colspan="2">&nbsp;<%=object_name%> יבוא&nbsp;</td></tr>
<tr>
<td width="100%" align="right" dir="rtl" style="padding: 10px;"><a href="../../../download/import/gilboa.xls" target="_blank" class="link_news">הורד קובץ קיים</a></td>
</tr>
<tr>
<td width="100%" align="right" dir="rtl" style="padding: 10px;">לחץ על כפתור [עיון] ובחר Excel הקובץ מהמחשב שלך.</td>
</tr>
<tr><td width="100%">
<form name="form1" id="form1" target="_self" runat="server" METHOD="post" ENCTYPE="multipart/form-data" style="margin: 0px;">
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="2">
<tr>
	<td align="right">
	<table border="0" width="100%" cellpadding="0" cellspacing="0">		
		<tr>																							
			<td align="right" width="100%" dir="rtl"><INPUT style="font-size:11pt" TYPE="FILE" 
			NAME="UploadFile" style="width: 300px" ID="UploadFile"></td>														
			<td width="10" nowrap align="center"></td>
		</tr>		
	</table>					
	</td>
</tr>
<tr><td width="100%" style="padding: 10px;">
<table cellpadding="0" cellspacing="0" width="100%">
<tr>
	<td width="100%" align="right"><INPUT type="submit" class="add" value="טעינת הקובץ" id="upload_excel" 
	name="upload_excel" style="width: 100px" onClick="return ifFieldEmpty(this);"></td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td></tr>	
<tr><td height="10" nowrap></td></tr>	
</table>
</body>		
</html>