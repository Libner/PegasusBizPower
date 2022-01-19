<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<%
	OrgId = trim(Request("ORGANIZATION_ID"))
	If trim(OrgId) <> "" Then
	sqlStr = "select ORGANIZATION_NAME from ORGANIZATIONS where ORGANIZATION_ID=" & OrgId  
	''Response.Write sqlStr
	set rs_org = con.GetRecordSet(sqlStr)
	if not rs_org.eof then
		orgName = trim(rs_org(0))
	end if
	set rs_org = nothing
	End If  
	
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then
		    con.executeQuery("Delete From JOBS Where job_id = " & delId)		    
			con.executeQuery("Delete From USERS Where job_id = " & delId)
		End If
		Response.Redirect "jobs.asp?ORGANIZATION_ID=" & OrgId
	End If
%>
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function checkDelete(countUsers)
	{
		if(countUsers > 0)
			return window.confirm("?האם ברצונך למחוק את התפקיד עם כל העובדים");
		else return window.confirm("?האם ברצונך למחוק את התפקיד");	
	}
	
	function openjobs(jobID,OrgID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addjob.asp?job_id=" + jobID + "&ORGANIZATION_ID="+OrgID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
	function openUpload(jobID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("uploadExcel.asp?job_id=" + jobID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
//-->
</SCRIPT>

</HEAD>

<body bgcolor="#FFFFFF">
<div align="right">

<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table3">
  <tr>
    <td colspan="4" class="page_title" width=100% dir=rtl><%=orgName%> - רשימת תפקידים</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1"  href="" onClick="return openjobs('','<%=OrgID%>')">הוספת תפקיד חדש</a></td>    
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדף ארגונים</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>       
     <td width="*%" align="center"></td>      
  </tr>
</table>
<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr><td height=20 nowrap></td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table1">
<tr>
<td width=100% valign=top>
<table width="650" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff">
<tr bgcolor="#B4B4B4" height="22">
	<th width="80" nowrap align="center" class="11normalB">&nbsp;מחיקה&nbsp;</th>	
	<th width="80" nowrap align="center" class="11normalB">&nbsp;עריכה&nbsp;</th>	
	<th width="20%" align="center" class="11normalB">&nbsp;עלות שעה&nbsp;</th>
	<th width="80%" align="right" class="11normalB">&nbsp;תפקיד&nbsp;</th>	
</tr>
<%	set rs_jobs = con.GetRecordSet("Select job_id, job_name, hour_pay from jobs WHERE Organization_ID = " & trim(OrgID) & " order by job_id")
    if not rs_jobs.eof then 
		do while not rs_jobs.eof
    	job_Id = trim(rs_jobs("job_id"))
    	job_name = trim(rs_jobs("job_name"))
    	hour_pay = trim(rs_jobs("hour_pay"))
    	If Len(job_id) > 0 Then
    		sqlstr = "Select Count(USER_ID) From Users Where job_id = " & job_id
    		set rscount = con.getRecordSet(sqlstr)
    		If not rscount.eof Then
    			countUsers = rscount(0)
    		Else countUsers = 0	
    		End If
    		set rscount = Nothing
    	Else countUsers = 0	
    	End If
    	%>
<tr>
	<td align="center" class="normalB" bgcolor="#e6e6e6" nowrap><a href="jobs.asp?deleteId=<%=job_id%>&ORGANIZATION_ID=<%=OrgId%>" ONCLICK="return checkDelete(<%=countUsers %>)"><IMG SRC="../images/delete_icon.gif" BORDER=0 alt="מחיקת קבוצת עובדים שיחה"></a></td>
	<td align="center" class="normalB" bgcolor="#e6e6e6" nowrap><a href="" onClick="return openjobs('<%=job_Id%>','<%=OrgID%>')" target="_blank"><IMG SRC="../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align=center class="normalB" bgcolor="#e6e6e6"><%=hour_pay%>&nbsp;</td>	
	<td align=right class="normalB" bgcolor="#e6e6e6"><b><%=job_name%></b>&nbsp;</td>	
</tr>
<%		rs_jobs.MoveNext
		loop
	end if
	set rs_jobs=nothing%>
</table>
</td>
</tr>
<tr><td height=10 nowrap></td></tr>
</table>
</BODY>
</HTML>
