<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
 	templateId = Request.QueryString("templateId")
 	If isNumeric(trim(templateId)) Then
	sqlpg="SELECT * FROM Templates WHERE Template_Id="&templateId&""
	'Response.Write sqlpg
	'Response.End
	set pg=con.getRecordSet(sqlpg)
	If not pg.eof Then
		pageTitle=pg("Template_Title")
	End If
	set pg = Nothing
	End If
	
  If trim(Request.QueryString("add")) <> nil And trim(templateId) <> "" Then 
      
	organizations_list = Request.Form("organizations")
	If Len(organizations_list) > 0 Then
		organizations_arr = Split(organizations_list,",")		
	End If		
	
	sqlstr = "Delete FROM Templates_To_Organizations WHERE Template_Id = " & templateId
	con.executeQuery(sqlstr)
	
    If IsArray(organizations_arr) And IsNull(organizations_arr)=false Then
		For i=0 To Ubound(organizations_arr)
			sqlstr = "Insert Into Templates_To_Organizations values (" & organizations_arr(i) & "," & templateId & ")"
			con.executeQuery(sqlstr)
		Next
    End If
    
    Response.Redirect "default.asp"
  End if
%>
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
function check_all_orgs(objChk)
{
	input_arr = document.getElementsByName("organizations");	
	for(i=0;i<input_arr.length;i++)
	{
		input_arr(i).checked = objChk.checked;		
	}
	return true;
}
//-->
</script>
</head>
<body>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="4" class="page_title" dir=rtl>&nbsp;ניהול הרשאות לתבנית&nbsp;"<%=pageTitle%>"</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" href="#" onClick="document.form1.submit();return false;">עדכון</a></td>  
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדף ניהול תבניות</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="100%" align="center"></td>      
  </tr>
</table>
<br>
<table width="90%" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff">
<form name="form1" id="form1" target="_self" ACTION="permitions.asp?templateId=<%=templateId%>&add=1" METHOD="post">
<tr>
	<td align="right">
	<input type=hidden name="orgs_id" id="orgs_id" value="">
	<table border=0 width=500 cellpadding=0 cellspacing=0 align=center>		
		<tr height=25>																							
			<td align="right" width=100%>
			<table cellpadding=2 cellspacing=1 width=100% dir=rtl>			
			<%
				sqlStr = "Select Organization_ID, Organization_NAME from Organizations Order by 2"
				'Response.Write sqlStr
				set rs_orgs = con.GetRecordSet(sqlStr)
				If not 	rs_orgs.Eof Then
					OrgsCount = rs_orgs.recordCount
					sqlstr = "Select Count(*) From Templates_To_Organizations WHERE Template_ID = " & templateId
					set rs_tmp = con.getRecordSet(sqlstr)
					If not rs_tmp.eof Then
						OrgTempCount = rs_tmp(0)
					Else
						OrgTempCount = 0
					End If
					If trim(OrgsCount) = trim(OrgTempCount) Then
						is_checked_all = true
					Else
						is_checked_all = false
					End If			
					
			%>
			<tr><td colspan=2 height=1 bgcolor=#808080></td></tr>
			<tr>
			<td width=22 align="right" nowrap class="12normalB" bgcolor="#b0b0b0"><input type=checkbox name="all_orgs_id" value="all_orgs_id" ID="all_orgs_id" onclick="return check_all_orgs(this)" <%If is_checked_all Then%> checked <%End If%>></td>
			<td align="right" width=100% class=12normalB bgcolor="#dbdbdb">&nbsp;כל הארגונים&nbsp;</td>
			</tr>	
			<%	
				Do while not rs_orgs.eof
					Organization_ID = trim(rs_orgs(0))
					Organization_NAME = trim(rs_orgs(1))
					is_checked = false
					If Len(templateId) > 0 Then
						sqlstr = "Select Top 1 Template_ID From Templates_To_Organizations WHERE Template_ID = " & templateId &_
						" And Organization_ID = " & Organization_ID
						set rs_check = con.getRecordSet(sqlstr)
						If not rs_check.eof Then
							is_checked = true
						Else
							is_checked = false
						End if
						set rs_check = Nothing
					End if
			%>
			<tr>
			<td width=22 align="right" nowrap class="10normal" bgcolor="#b0b0b0"><input type=checkbox ID="<%=Organization_ID%>" value="<%=Organization_ID%>" name="organizations" <%If is_checked Then%> checked <%End If%>></td>
			<td align="right" width=100% class="11normal" bgcolor="#dbdbdb">&nbsp;<%=Organization_NAME%>&nbsp;</td>			
			</tr>
			<%
				
				rs_orgs.movenext
				loop
				set rs_orgs = nothing
				End If										
			%>
			<tr><td colspan=2 height=1 bgcolor=#808080></td></tr>
			</table>		
			</td>
			<td width=20 nowrap align=center>&nbsp;</td>														
		</tr>		
	</table>					
	</td>
</tr>
<tr><td colspan=2 height=20 nowrap></td></tr>
</form>	
</table></td></tr>
</table>
</body>		
</html>		