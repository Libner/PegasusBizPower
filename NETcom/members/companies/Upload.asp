<%Response.Buffer = False%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
	function ifFieldEmpty(){
	
	if (document.form1.UploadFile.value=='')
	{
		alert('בחר קובץ');
		document.form1.UploadFile.focus();		
		return false;
	}
	else if (document.form1.UploadFile.value !='')
	{
			var fname=new String();
			var fext=new String();
			var extfiles=new String();
			fname=document.form1.UploadFile.value;
			indexOfDot = fname.lastIndexOf('.')
			fext=fname.slice(indexOfDot+1,-1)		
			fext=fext.toUpperCase();
			extfiles='CSV';		
			if ((extfiles.indexOf(fext)>-1) == false)
			{	
				window.alert(extfiles + ' סיומת הקובץ חייבת להיות ');
				return false;
			}    
	}    		  
	return true;		
}
//-->
</script>
</head>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus();">
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr><td>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir=rtl>&nbsp;ייבוא <%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%>&nbsp;</td></tr>
<table cellpadding=5 cellspacing=0 width=450 align=center>
<tr><td width="100%" colspan=2 height=30></td></tr>
<tr>
<td width=100% align=right dir=rtl> בצע ייצוא אנשי קשר מתוכנת הדואר שלך בפורמט ערכים מופרדים בפסיקים (קובץ CSV)</td>
<td width=20 nowrap align=center><b>1</b></td></tr>
<tr>
<td width=100% align=right dir=rtl>שמור את הקובץ בכונן המקומי שלך.</td>
<td width=20 nowrap align=center><b>2</b></td></tr>
<tr>
<td width=100% align=right>.מהכונן שלך CSVובחר את הקובץ ה Browse לחץ על לחצן</td>
<td width=20 nowrap align=center><b>3</b></td></tr>
<tr><td width="100%" colspan=2>
<table border="0" width="100%" align=center cellspacing="1" cellpadding="2">
<form name="form1" id="form1" target="_self" ACTION="Fileadd.asp" METHOD="post" ENCTYPE="multipart/form-data">
<tr>
	<td align=right>
	<table border=0 width=100% cellpadding=0 cellspacing=0>		
		<tr height=25>																							
			<td align="right" width=100%>&nbsp;<INPUT style="font-size:11pt" TYPE="FILE" NAME="UploadFile" style="width:300" ID="File1"></td>														
		</tr>		
	</table>					
	</td>
</tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right><INPUT type="submit" class="but_menu" value="טען רשימה" id="upload_excel" name="upload_excel" style="width:100" onClick="return ifFieldEmpty();"></td>
<td width=150 nowrap align=right>לחץ על לחצן טען רשימה</td>
<td width=20 nowrap align=center><b>4</b></td></tr>
</table>
</td></tr>
</form>	
<tr><td height=10 nowrap></td></tr>	
</table></td></tr></table>
</body>		
</html>		