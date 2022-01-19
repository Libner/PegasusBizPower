<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--

 var preview, frame;
 
function CheckField() {
  if (window.document.all("pageTitle").value=='')
  {
	alert('! חובה להכניס שם חדש של דף מעוצב');
	window.document.all("pageTitle").select();
	window.document.all("pageTitle").scrollIntoView();
	return false;
  }
  
  document.dataForm.submit();
  return true;
}
<!--End-->
</script>  
<%
    UserID=trim(Request.Cookies("bizpegasus")("UserID"))
    OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))

    copypageId=trim(Request("copyPageId")) 'מספר דף לשכפול    
	If trim(copypageId) <> "" Then
		sqlstr = "Select Page_Title From Pages Where Page_Id = " & copypageId
		set rs_p = con.getRecordSet(sqlstr)
		If not rs_p.eof Then
			copypageTitle = trim(rs_p(0))
		Else
			copypageTitle = ""	
		End if
		set rs_p = Nothing
	End If
	
	if Request("editFlag")<>nil then			
		
		if trim(copypageId) <> "" then 	
			'הוספת דף חדש עבור הדף מבצע
			sqlstring="SET NOCOUNT ON; INSERT INTO Pages (Page_Source,ORGANIZATION_ID, USER_ID, Product_Id, " & _
			" FORM_LINK, FORM_LINK_IMAGE, LINK_IMAGE,LINK_IMAGE_ALIGN,LINK_TEXT,LINK_TEXT_ALIGN,LINK_BGCOLOR,"&_
			" LINK_FONT_TYPE,LINK_FONT_SIZE,LINK_FONT_COLOR,FORM_BGCOLOR,"&_
			" FORM_FONT_TYPE,FORM_FONT_SIZE,FORM_FONT_COLOR,Page_Lang) "&_
			" Select Page_Source, ORGANIZATION_ID, USER_ID, Product_Id, FORM_LINK, FORM_LINK_IMAGE,"&_
			" LINK_IMAGE,LINK_IMAGE_ALIGN,LINK_TEXT,LINK_TEXT_ALIGN,LINK_BGCOLOR,"&_
			" LINK_FONT_TYPE,LINK_FONT_SIZE,LINK_FONT_COLOR,FORM_BGCOLOR,"&_
			" FORM_FONT_TYPE,FORM_FONT_SIZE,FORM_FONT_COLOR,Page_Lang "&_
			" From Pages Where Page_ID = " & copypageId & "; SELECT @@IDENTITY AS NewID"			
			'Response.Write sqlstring
			'Response.End
			set rs_tmp = con.getRecordSet(sqlstring)
				pageId = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing	  					
			
			if trim(pageId) <> "" Then
				sqlstr="Update Pages Set Page_Title = '" & sFix(Request.Form("pageTitle")) & "'" &_
				" Where Page_Id = " & pageId
				con.executeQuery(sqlstr)
			end if			
			
		end if	
			
		%>
		<script language=javascript>
		<!--
			window.opener.document.location.href = window.opener.document.location.href;
			window.close();	
		//-->
		</script>
		<%
		
	End if		
		
%>	
<body style="margin:0px;background:#E5E5E5" onload="window.focus();">
<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td>
<table width=100% border=0 cellspacing="0" cellpadding="0">
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td class="page_title" align="right" dir=rtl>&nbsp;שכפול דף "<%=copypageTitle%>" &nbsp;</td></tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
</table>
</td></tr>
<tr><td height=15 nowrap></td></tr>
  <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" align=center border="0" cellpadding="2" cellspacing="1">
	<FORM name="dataForm" ACTION="copyPage.asp?editFlag=1" METHOD="post">
	<input type="hidden" name="copypageId" id="copypageId" value="<%=copypageId%>">
	<tr>
		<td align=right><input dir=rtl type="text" class="texts" style="width:350" id="pageTitle" name="pageTitle" value="<%=vfix(pageTitle)%>"></td>
		<td align="center"><b>&nbsp;שם דף חדש&nbsp;</b></td>
	</tr>		 
	<tr><td colspan="2" height="25" nowrap></td></tr>
	<tr><td colspan="2" align=center>
	<table width=100% border=0>
	<tr><td width=45% align="right">
	<INPUT class="but_menu" style="width:90" type="button" value="ביטול" id=button2 name=button2 onclick="window.close();"></td>
	<td width=5% nowrap></td>
	<td width=45% align="left">
	<input class="but_menu" style="width:90" type="button" value="אישור" onClick="return CheckField()">
	</td>
	<td height="15" align="right"></td></tr>
	</table></td></tr>
	</table>
	</td></tr>
	<tr><td height=10 nowrap></td></tr>
	</table>
</center></div>
<%set con=nothing%>
</body>
</html>

