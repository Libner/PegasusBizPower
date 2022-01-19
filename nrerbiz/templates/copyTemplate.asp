<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--

 var preview, frame;
 
function CheckField() {
  if (window.document.all("TemplateTitle").value=='')
  {
	alert('! חובה להכניס שם חדש של תבנית מעוצב');
	window.document.all("TemplateTitle").select();
	window.document.all("TemplateTitle").scrollIntoView();
	return false;
  }
  
  document.dataForm.submit();
  return true;
}
<!--End-->
</script>  
<%
    copytemplateId=trim(Request("copytemplateId")) 'מספר תבנית לשכפול    
	If trim(copytemplateId) <> "" Then
		sqlstr = "Select template_Title From templates Where template_Id = " & copytemplateId
		set rs_p = con.getRecordSet(sqlstr)
		If not rs_p.eof Then
			copytemplateTitle = trim(rs_p(0))
		Else
			copytemplateTitle = ""	
		End if
		set rs_p = Nothing
	End If
	
	if Request("editFlag")<>nil then			
		
		if trim(copytemplateId) <> "" then 	
			sqlstr="Select Template_Source,Template_Screenshot"&_
			" FROM Templates WHERE Template_ID =" & copytemplateId				
			set rs_p = con.getRecordSet(sqlstr)
			'Response.Write sFix(rs_p(1))
			'Response.End
			If not rs_p.eof Then
				set con_copy=server.createobject("adodb.connection")					
				con_copy.open "FILE NAME="& server.MapPath("..\..\netcom/netReply.udl")									
				set Template_copy = Server.CreateObject("ADODB.RecordSet")	
				sqltext = "SELECT * FROM Templates "
				Template_copy.Open sqltext,con_copy,3,3				
				Template_copy.AddNew
				
				Template_copy("Template_Title") = trim(Request.Form("TemplateTitle"))
				Template_copy("Template_Source") = rs_p("Template_Source")
				Template_copy("Template_Screenshot") = rs_p("Template_Screenshot")			
				
				Template_copy.Update
				Template_copy.Close
				set Template_copy = Nothing
			 End If	
			
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
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">
<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td>
<table width=100% border=0 cellspacing="0" cellpadding="0">
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td class="page_title" align="right" dir=rtl>&nbsp;שכפול תבנית "<%=copyTemplateTitle%>" &nbsp;</td></tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
</table>
</td></tr>
<tr><td height=15 nowrap></td></tr>
  <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" align=center border="0" cellpadding="2" cellspacing="1">
	<FORM name="dataForm" ACTION="copyTemplate.asp?editFlag=1" METHOD="post">
	<input type="hidden" name="copyTemplateId" id="copyTemplateId" value="<%=copyTemplateId%>">
	<tr>
		<td align=right><input dir=rtl type="text" class="texts" style="width:250" id="TemplateTitle" name="TemplateTitle" value="<%=vfix(TemplateTitle)%>"></td>
		<td align="center" class="11normalB"><b>&nbsp;שם תבנית חדשה&nbsp;</b></td>
	</tr>		 
	<tr><td colspan="2" height="25" nowrap></td></tr>
	<tr><td colspan="2" align=center>
	<table width=100% border=0>
	<tr><td width=45% align="right">
	<INPUT class="but_browse" style="width:90" type="button" value="ביטול" id=button2 name=button2 onclick="window.close();"></td>
	<td width=5% nowrap></td>
	<td width=45% align="left">
	<input class="but_browse" style="width:90" type="button" value="אישור" onClick="return CheckField()">
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

