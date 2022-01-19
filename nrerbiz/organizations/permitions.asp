<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
 	OrgID = Request.QueryString("ORGANIZATION_ID")
 	If isNumeric(trim(OrgID)) Then
	sqlpg="SELECT * FROM Organizations WHERE Organization_Id="&OrgID&""
	'Response.Write sqlpg
	'Response.End
	set rs_org=con.getRecordSet(sqlpg)
	If not rs_org.eof Then
		OrgName=rs_org("Organization_Name")
	End If
	set rs_org = Nothing
	End If
	
  If trim(Request.QueryString("add")) <> nil And trim(OrgID) <> "" Then 
      
	templates_list = Request.Form("templates")
	If Len(templates_list) > 0 Then
		templates_arr = Split(templates_list,",")		
	End If		
	
	sqlstr = "Delete FROM Templates_To_Organizations WHERE Organization_Id = " & OrgID
	con.executeQuery(sqlstr)
	
    If IsArray(templates_arr) And IsNull(templates_arr)=false Then
		For i=0 To Ubound(templates_arr)
			sqlstr = "Insert Into Templates_To_Organizations values (" & OrgID & "," & templates_arr(i) & ")"		
			con.executeQuery(sqlstr)
		Next
    End If
    
    Response.Redirect "default.asp"
  End if
%>
<html>
<head>
<Title>Bizpower Administration</Title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
function check_all_templates(objChk)
{
	input_arr = document.getElementsByName("templates");	
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
    <td colspan="4" class="page_Title" dir=rtl>&nbsp;ניהול הרשאות לתבניות עבור הארגון&nbsp;"<%=OrgName%>"</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" href="#" onClick="document.form1.submit();return false;">עדכון</a></td>  
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp" target=_self>חזרה לדף ארגונים</a></td>
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="100%" align="center"></td>      
  </tr>
</table>
<br>
<table width="90%" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff">
<form name="form1" id="form1" target="_self" ACTION="permitions.asp?ORGANIZATION_ID=<%=OrgID%>&add=1" METHOD="post">
<tr>
	<td align="right">
	<input type=hidden name="orgs_id" id="orgs_id" value="">
	<table border=0 width=500 cellpadding=0 cellspacing=0 align=center>		
		<tr height=25>																							
			<td align="right" width=100%>
			<table cellpadding=2 cellspacing=1 width=100% dir=rtl>			
			<%
				sqlStr = "Select Template_ID, Template_Title from Templates Where IsNull(Template_Active,0) = '1' Order by 1"
				'Response.Write sqlStr
				set rs_templates = con.GetRecordSet(sqlStr)
				If not 	rs_templates.Eof Then
					TempsCount = rs_templates.recordCount
					sqlstr = "Select Count(*) From Templates_To_Organizations WHERE Organization_ID = " & OrgID
					set rs_tmp = con.getRecordSet(sqlstr)
					If not rs_tmp.eof Then
						OrgTempCount = rs_tmp(0)
					Else
						OrgTempCount = 0
					End If
					If cInt(TempsCount) <= Cint(OrgTempCount) Then
						is_checked_all = true
					Else
						is_checked_all = false
					End If			
					
			%>
			<tr><td colspan=2 height=1 bgcolor=#808080></td></tr>
			<tr>
			<td width=22 align="right" nowrap class="12normalB" bgcolor="#b0b0b0"><input type=checkbox name="all_templates_id" value="all_templates_id" ID="all_templates_id" onclick="return check_all_templates(this)" <%If is_checked_all Then%> checked <%End If%>></td>
			<td align="right" width=100% class=12normalB bgcolor="#dbdbdb">&nbsp;כל התבניות&nbsp;</td>
			</tr>	
			<%	
				Do while not rs_templates.eof
					Template_ID = trim(rs_templates(0))
					Template_Title = trim(rs_templates(1))
					is_checked = false
					If Len(OrgID) > 0 Then
						sqlstr = "Select Top 1 Template_ID From Templates_To_Organizations WHERE Organization_ID = " & OrgID &_
						" And Template_ID = " & Template_ID
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
			<td width=22 align="right" nowrap class="10normal" bgcolor="#b0b0b0"><input type=checkbox ID="<%=Template_ID%>" value="<%=Template_ID%>" name="templates" <%If is_checked Then%> checked <%End If%>></td>
			<td align="right" width=100% bgcolor="#dbdbdb">&nbsp;<a class="11normal" href="../templates/result.asp?templateId=<%=Template_ID%>" target=_blank><%=Template_Title%></a>&nbsp;</td>			
			</tr>
			<%
				
				rs_templates.movenext
				loop
				set rs_templates = nothing
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