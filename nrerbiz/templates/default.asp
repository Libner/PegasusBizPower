<!--#include file="../../netcom/connect.asp"-->
<!--#include file="../..//netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
function CheckDelProd() {
  return (confirm("?האם ברצונך למחוק את הדף מבצע"))    
}
function openCopy(TemplateId)
 {
	h = 200;
	w = 500;
	S_Wind = window.open("copyTemplate.asp?copyTemplateId=" + TemplateId, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
	S_Wind.focus();	
	return false;
 }
//-->
</script>  
</head>
<%
	templateId = Request.QueryString("templateId")
	If request("deltemplate")="1" and isNumeric(templateId) Then		
		con.ExecuteQuery("Delete from Templates WHERE template_id="&templateId&"")		
		Response.Redirect "default.asp"
	End If
	
	If Request.QueryString("idactive") <> nil Then
		idactive = trim(Request.QueryString("idactive"))
		sqlstr = "Update Templates Set Template_Active = 1 - Template_Active WHERE Template_ID = " & idactive
		con.executeQuery(sqlstr)
		Response.Redirect "default.asp"
	End If
  %>
<body>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="4" class="page_title">דף ניהול תבניות דפי מבצע</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" href="editPageAdmin.aspx">הוספת תבנית חדשה</a></td>    
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="5%" align="center"></td> 
     <td width="*%" align="center"></td>      
  </tr>
</table>
<br>
<table width="90%" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff" >

<tr>
	<th width="100" nowrap bgcolor="#999999">&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>מחיקה</b></font>&nbsp;</th>	
	<th width="100" nowrap bgcolor="#999999">&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>פעיל</b></font>&nbsp;</th>	
	<th width="100" nowrap bgcolor="#999999">&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>הרשאות</b></font>&nbsp;</th>	
	<th width="100" nowrap bgcolor="#999999">&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>שכפול</b></font>&nbsp;</th>
	<th width="100" nowrap bgcolor="#999999">&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>פרטי התבנית</b></font>&nbsp;</th>
	<th width="100" nowrap bgcolor="#999999">&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>עריכת תוכן</b></font>&nbsp;</th>
	<th width="100%" bgcolor="#999999">&nbsp;&nbsp;<font style="font-size:10pt;color:#ffffff;"><b>שם התבנית</b></font>&nbsp;&nbsp;</th>
</tr>
<%	set rs_Templates = con.GetRecordSet("Select template_id, template_title, IsNull(template_active,0) from Templates Order by template_id")
    if not rs_Templates.eof then 
		do while not rs_Templates.eof
    	templateId = trim(rs_Templates(0))
    	template_name = rs_Templates(1)
    	template_active = rs_Templates(2)
%>
<tr>
	<td align="center" bgcolor="#dddddd" nowrap>&nbsp;	
	<a href="default.asp?templateId=<%=templateId%>&deltemplate=1" ONCLICK="return CheckDelProd()"><IMG SRC="<%=Application("VirDir")%>/netcom/images/delete_icon.gif" BORDER=0 alt="מחיקת דף מבצע"></a>&nbsp;	
	</td>
	<td align="center" bgcolor="#dddddd" nowrap><a href="default.asp?idactive=<%=templateId%>"><%if template_active = "0" then%><img src="../images/lamp_off.gif" alt="לא מופיע באתר" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../images/lamp_on.gif" alt="מופיע באתר" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>
	<td align="center" bgcolor="#dddddd" nowrap>&nbsp;<a href="permitions.asp?templateId=<%=templateId%>"><IMG SRC="<%=Application("VirDir")%>/netcom/images/signIn_icon.gif" BORDER=0 alt="הרשאות"></a>&nbsp;</td>		
	<td align="center" bgcolor="#dddddd" nowrap>&nbsp;<INPUT type=image onclick="return openCopy('<%=templateId%>')" SRC="<%=Application("VirDir")%>/netcom/images/copy_icon.gif" BORDER=0 alt="שכפול הדף מבצע">&nbsp;</td>
	<td align="center" bgcolor="#dddddd" nowrap>&nbsp;<a href="editScreen.asp?templateId=<%=templateId%>"><IMG SRC="<%=Application("VirDir")%>/netcom/images/form_icon.gif" BORDER=0 alt="שכפול הדף מבצע"></a>&nbsp;</td>
	<td align="center" bgcolor="#dddddd" nowrap>&nbsp;	
	<a href="editPageAdmin.aspx?templateId=<%=templateId%>"><IMG SRC="<%=Application("VirDir")%>/netcom/images/write.gif" BORDER=0 alt="עריכת דף מבצע"></a>&nbsp;	
	</td>
	<td class="td_subj" align=right bgcolor="#dddddd"><a class="linkFaq" href="result.asp?templateId=<%=templateId%>" target="_new"><%=template_name%></a>&nbsp;&nbsp;</td>
</tr>
<%		rs_Templates.MoveNext
		loop
	end if
	set rs_Templates=nothing%>
</table> 
</body>
<%set con=nothing%>
</html>
