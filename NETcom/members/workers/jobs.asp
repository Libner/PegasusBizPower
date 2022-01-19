<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%If Request.QueryString("deleteId") <> nil Then
			delId = trim(Request.QueryString("deleteId"))
			If Len(delId) > 0 Then
				con.executeQuery("Delete From JOBS Where job_id = " & delId)			
			End If
			Response.Redirect "jobs.asp"
	End If

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 36 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
		dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	End If
	set rstitle = Nothing %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
	function checkDelete()
	{
		<%If trim(lang_id) = "1" Then%>
			str_confirm = "?האם ברצונך למחוק את הסוג"
		<%Else%>
			str_confirm = "Are you sure want to delete the position?"		
		<%End If%>		
		return window.confirm(str_confirm);		
	}
	
	function openjobs(jobID)
	{
		h = 500;
		w = 500;
		S_Wind = window.open("addjob.asp?job_id=" + jobID, "S_Wind" ,"scrollbars=1,toolbar=0,top=50,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
	function openUpload(jobID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("uploadExcel.asp?job_id=" + jobID, "S_Wind" ,"scrollbars=1,toolbar=0,top=50,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
//-->
</script>
</head>
<body>
<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr><td width="100%">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 4%>
<%numOfLink = 4%>
<%topLevel2 = 35 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width="100%" border="0" >
<tr>
<td width="100%" valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff">
<tr>
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<!--מחיקה--><%=arrTitles(3)%>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<!--עדכון--><%=arrTitles(4)%>&nbsp;</td>	
	<td width="80%" align="<%=align_var%>" class="title_sort">&nbsp;<!--סוג עובד--><%=arrTitles(5)%>&nbsp;</td>	
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
	<td align="center" bgcolor="#e6e6e6" nowrap>
	<%If countUsers = 0 Then%>
	<a href="jobs.asp?deleteId=<%=job_id%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" border="0" title="מחיקת קבוצת עובדים"></a>
	<%Else%>
	<%If trim(lang_id) = "1" Then
	    str_alert = "קיימים עובדים עם תפקיד זה במערכת\n\n" & Space(16) & "על מנת למחוק את התפקיד\n\nאליך למחוק קודם את העובדים הנ\'\'ל"
	Else
	    str_alert = "There are employees with this type in the system,\n\n therefore you can\'t delete this employee type"
	End If%>	
	<input type=image src="../../images/delete_icon.gif" onclick="window.alert('<%=str_alert%>');return false;">
	<%End If%>	
	</td>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="javascript:void(0)" onClick="return openjobs(<%=job_Id%>)" target="_blank"><IMG SRC="../../images/edit_icon.gif" border="0"></a></td>	
	<td align="<%=align_var%>" bgcolor="#e6e6e6"><b><%=job_name%></b>&nbsp;</td>	
</tr>
<%		rs_jobs.MoveNext
		loop
	end if
	set rs_jobs=nothing%>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width="100%" >
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openjobs('')"><span id=word6 name=word6><!--הוספת סוג עובד--><%=arrTitles(6)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href="excel_jobs.asp" target=_blank><!--הצג דוח--><%=arrTitles(9)%></a></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</td></tr></table>
</BODY>
</HTML>
