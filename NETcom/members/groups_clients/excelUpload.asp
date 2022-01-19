<%groupId = trim(Request.QueryString("groupId"))%>
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
		return true;		
}
//-->
</script>

</head>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus();">
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td width="100%" class="page_title" dir=rtl>&nbsp;Excel&nbsp;טעינת נמענים מקובץ&nbsp;<font color="#6E6DA6"><%=groupName%></font></td></tr>         
<tr><td width="100%" colspan=2 height=10></td></tr>
<tr><td width="100%" colspan=2>
<table border="0" width="95%" align=center cellspacing="1" cellpadding="2">
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right><INPUT type=button class="but_menu" onclick="window.open('../../../download/list.xls')" value="הורד" style="width:70"></td>
<td width=200 nowrap align=right>Excel הורד תבנית נמענים בפורמט</td>
<td width=20 nowrap align=center><b>1</b></td></tr>
</table>
</td></tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%">
<tr>
<td width=100% align=right>.ושמור אותה בכונן המקומי שלך ,Excelעדכן את רשימת הנמענים לתבנית ה</td>
<td width=20 nowrap align=center><b>2</b></td></tr>
</table>
</td></tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right>.מהכונן שלך Excelובחר את הקובץ ה Browse לחץ על לחצן</td>
<td width=20 nowrap align=center><b>3</b></td></tr>
</table>
</td></tr>
<form name="form1" id="form1" target="_self" ACTION="Fileadd.asp?C=1&groupId=<%=groupId%>" METHOD="post" ENCTYPE="multipart/form-data">
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
<tr><td dir="rtl"><font color="red">שימו לב, לא ניתן להעלות קבוצות של יותר מ-60 אלף נמענים</font></td></tr>	
</table></td></tr></table>
</body>		
</html>		