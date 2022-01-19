<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    OrgId = trim(Request("ORGANIZATION_ID"))
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("job_id")) = "" Then ' add type
			sqlstr = "Insert into jobs (Organization_ID,job_name,hour_pay) values (" &_
			trim(OrgId) & ",'" & sFix(Request.Form("job_name")) & "','" & sFix(Request.Form("hour_pay")) & "')"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			job_id = trim(Request.Form("job_id"))
			sqlstr="Update jobs set job_name = '" & sFix(Request.Form("job_name")) &_
			"', hour_pay = '" & sFix(Request.Form("hour_pay")) & "' Where job_id = " & job_id
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
			if(window.document.form1.job_name.value == "")
			{
				window.alert("!!נא להכניס תפקיד");
				window.document.form1.job_name.focus();
				return false;
			}
			if(window.document.form1.hour_pay.value == "")
			{
				window.alert("!!נא להכניס עלות שעה");
				window.document.form1.hour_pay.focus();
				return false;
			}
				
			return true;
			   
		}
	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("job_id") <> nil Then
		job_id = trim(Request.QueryString("job_id"))
		If Len(job_id) > 0 Then
			sqlstr="Select job_name, hour_pay From jobs Where job_id = " & job_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				job_name = trim(rssub("job_name"))	
				hour_pay = trim(rssub("hour_pay"))	
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
		 <td class="page_title" align="right" valign="middle" width="100%"><%If Len(job_id) > 0 Then%>עדכון תפקיד<%Else%>הוספת תפקיד<%End If%>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0">
<form name=form1 id=form1 action="addjob.asp?add=1" target="_self" method="post">
<input type=hidden name=job_id id=job_id value="<%=job_id%>">
<input type=hidden name=ORGANIZATION_ID id="ORGANIZATION_ID" value="<%=OrgId%>">
<tr>
	<td align=right width=330 nowrap>
	<input type="text" name="job_name" id="job_name" value="<%=vFix(job_name)%>" dir=rtl size=40 maxLength=50>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b>שם תפקיד</b>&nbsp;</td>	
</tr>
<tr>
	<td align=right width=330 nowrap>
	<input type="text" name="hour_pay" id="hour_pay" value="<%=vFix(hour_pay)%>" dir=rtl size=4 maxLength=4 onkeypress="GetNumbers()">	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b>עלות שעה</b>&nbsp;</td>	
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
