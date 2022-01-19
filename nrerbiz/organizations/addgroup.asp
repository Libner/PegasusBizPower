<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  OrgId = request("ORGANIZATION_ID")
    If Request.QueryString("add") <> nil Then
        group_id = trim(Request.Form("group_id"))	
        group_name = sFix(Request.Form("group_name"))       
     
		If trim(Request.Form("group_id")) = "" Then ' add type
			sqlstr = "Insert into Users_Groups (Organization_ID,group_name) values (" &_
			trim(OrgId) & ",'" & trim(group_name) & "')"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update
					
			sqlstr="Update Users_Groups set group_name = '" & trim(group_name) &_
			"' Where group_id = " & group_id
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>	
	<%	End If
	
 End If	
%>
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{
			if(window.document.form1.group_name.value == "")
			{
				window.alert("!!נא להכניס קבוצה");
				window.document.form1.group_name.focus();
				return false;
			}			
			return true;
			   
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("group_id") <> nil Then
		group_id = trim(Request.QueryString("group_id"))
		If Len(group_id) > 0 Then
			sqlstr="Select group_name From Users_Groups Where group_id = " & group_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				group_name = trim(rssub("group_name"))					
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table class="table_admin_2" border="0" width="480" cellspacing="0" cellpadding="0" align="center">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table width="100%" border="0"  cellpadding="1" cellspacing="0">
	   <tr>		 
		 <td class="page_title" align="right" valign="middle" width="100%"><%If Len(group_id) > 0 Then%>עדכון קבוצה<%Else%>הוספת קבוצה<%End If%>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="460" cellspacing="1" cellpadding="2" align=center border="0">
<form name=form1 id=form1 action="addgroup.asp?add=1" target="_self" method="post">
<input type=hidden name=group_id id=group_id value="<%=group_id%>">
<input type=hidden name=ORGANIZATION_ID id="ORGANIZATION_ID" value="<%=OrgId%>">
<tr>
	<td align=right width=330 nowrap>
	<input type="text" name="group_name" id="group_name" value="<%=vFix(group_name)%>" dir=rtl style="width:300" maxLength=50>	
	</td>
	<td width="100" nowrap align="right">&nbsp;<b>שם קבוצה</b>&nbsp;</td>	
</tr>
<tr><td height=35 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="ביטול" class="but_browse" style="width:90" onclick="window.close();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="אישור" class="but_browse" style="width:90" onclick="return checkForm()"></td></tr>
</form>
</table>
</td></tr> 
</BODY>
</HTML>
