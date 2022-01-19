<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--

var frame;

function CheckField() {
  if (window.document.all("pageTitle").value=='')
  {
	alert('! חובה למלא את השם של דף מבצע');
	window.document.all("pageTitle").select();
	window.document.all("pageTitle").scrollIntoView();
	return false;
  }
  
 
  return true;
}
<!--End-->
</script>  
<body topmargin=0 rightmargin=0 leftmargin=0 bottommargin=0>
<%
    copytemplateId=trim(Request("copytemplateId")) 'מספר דף לשכפול		
	If trim(copytemplateId) <> "" Then
	sqlpg="SELECT Template_Source FROM Templates WHERE Template_Id="&copytemplateId&""
	set pg=con.getRecordSet(sqlpg)
	If not pg.eof Then 
		tempSource=pg(0)
	End if
	set pg = Nothing	
				
	if Request("editFlag")<>nil then	
						
		sqlstr = " updateTemplateContent (0,'" & sFix(Request.Form("pageTitle")) & "','" & tempSource & "')"						
		con.ExecuteQuery(sqlstr)
					
		sql="SELECT Max(Template_Id) FROM Templates"			
		set rs_max=con.getRecordSet(sql)
		If not rs_max.eof Then
			templateId = trim(rs_max(0))	
		End If
		set rs_max = Nothing	
		
	    Response.Redirect  "editPageAdmin.aspx?templateId=" & templateId
	
    Else   	
	If isNumeric(trim(copytemplateId)) Then
	sqlpg="SELECT * FROM Templates WHERE Template_Id="&copytemplateId&""
	'Response.Write sqlpg
	'Response.End
	set pg=con.getRecordSet(sqlpg)
	If not pg.eof Then
		pageTitle=pg("Template_Title")
	End If
	set pg = Nothing
	End If
%>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5" class="page_title" dir=rtl>&nbsp;שכפול&nbsp;תבנית&nbsp;"<%=pageTitle%>"</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" href="#" onClick="document.dataForm.submit();return false;">עדכון</a></td>  
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדף ניהול תבניות</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="*%" align="center"></td> 
  </tr>
</table>

<table border="0" width="100%"cellspacing="1" cellpadding="0">
	<tr>
		
		<td align="center">&nbsp;</td>
	</tr>
	
</table>
<table width="600" align=center border="0" cellpadding="2" cellspacing="1" bgcolor="#ffffff">
<FORM name="dataForm" ACTION="addTemplate.asp?editFlag=1" METHOD="post" onSubmit="return CheckField();">
<tr><td colspan="2" height="10" class="title_table_admin">
<table width=100% border=0>
</table></td></tr>
<tr>
	<td align=right bgcolor="#dbdbdb"><input dir=rtl type="text" class="texts" style="width:350" id="pageTitle" name="pageTitle" value=""></td>
	<td align="center" bgcolor="#b0b0b0"><font style="font-size:10pt;color:#ffffff;"><b>&nbsp;שם התבנית&nbsp;</b></font></td>
</tr>

   <input type="hidden" name="copytemplateId" value="<%=copytemplateId%>">	
   

	</form>	
	
</table>
<br>
<%End If%>
<%End If%>
<%set con=nothing%>
</body>
</html>

