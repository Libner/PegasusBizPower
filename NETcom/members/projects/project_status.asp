<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
project_id=request.querystring("project_id")
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 1%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title">&nbsp;סטטוס תכנון מול ביצוע&nbsp;</td></tr>		   
	  		       	
   </table></td></tr> 
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="2" cellspacing="2">
<tr><td height=10></td></tr>  
<tr>
<td>
<table cellpadding=2 cellpadding=2 align=center>
<%
if project_id<>nil and project_id<>"" then
  sqlStr = "select company_id, project_code,project_name, project_description, start_date, end_date, profit_percent from projects where project_ID=" & project_id  
  ''Response.Write sqlStr
  set rs_projects = con.GetRecordSet(sqlStr)
	if not rs_projects.eof then
		company_id = rs_projects("company_id")
		If trim(company_id) <> "" Then
			sqlstr = "Select company_name FROM companies WHERE company_id = " & company_id
			set rs_compName = con.getRecordSet(sqlstr)
			if not rs_compName.eof then
				company_name = rs_compName("company_name")
			end if
			set rs_compName = nothing	
		End If
		If trim(company_name) = "" Then
			company_name = "כללי"
		End If
		project_code = rs_projects("project_code")
		project_name = rs_projects("project_name")
		project_description = rs_projects("project_description")
		dateStart = rs_projects("start_date")
		dateEnd = rs_projects("end_date")	
		profit_percent = rs_projects("profit_percent")	
		If IsNull(profit_percent) Then
			profit_percent = 0
		End if	
		'Response.Write profit_percent
	end if
	set rs_projects = nothing
	
	sqlstr = "Select TOP 1 project_id FROM hours WHERE organization_id = " & trim(OrgID)
	set rs_h = con.getRecordSet(sqlstr)
	if not rs_h.eof then
		is_worked = true
	else
		is_worked = false
	end if
	set rs_h = nothing
end if
'Response.Write company_id
%>
<tr>
	<td align=right nowrap width="80%" dir=rtl><%=project_code%></td>
	<td align="right" nowrap width="20%" dir=rtl><b>קוד פרויקט</b>&nbsp;</td>
</tr>
<tr>
	<td align=right nowrap width="80%" dir=rtl><%=project_name%></td>
	<td align="right" nowrap width="20%" dir=rtl><b>שם פרויקט</b>&nbsp;</td>
</tr>
<tr>
	<td align=right dir=rtl><%=project_description%></td>
	<td align="right" nowrap width="120" dir=rtl><b>תיאור פרויקט</b>&nbsp;</td>
</tr>

<tr>
	<td align=right dir=rtl><%=company_name%></td>
	<td align="right" nowrap width="120" dir=rtl><b>לקוח</b>&nbsp;</td>
</tr>
</table></td></tr>
<%if project_id<>nil and project_id<>"" then%>
<tr>
<td>
<table cellpadding=3 cellspacing=1 border=0 bgcolor="#E6E6E6" align=right>
<%
	sum_all = 0
	sqlstr = "Select DISTINCT project_id,company_id From pricing_to_projects WHERE pricing_id = " & project_id & " Order By project_id"
	'Response.Write sqlstr
	set rs = con.getRecordSet(sqlstr)
	if not rs.eof then
		projNumber = rs.recordCount
%>
<tr>
<td class="title_sort" align=center colspan=2>&nbsp;עלות עבודה בפועל&nbsp;</td>
<td class="title_sort" align=center colspan=3>&nbsp;תמחור הפרויקט&nbsp;</td></tr>
<tr style="line-height:22px">
<td class="title_sort2" width=67 nowrap align=center>עלות</td>
<td class="title_sort2" width=67 nowrap align=center>שעות</td>
<td class="title_sort2" width=67 nowrap align=center>עלות</td>
<td align=center class="title_sort2">שעות מתוכננות</td>
<td align=right class="title_sort2">&nbsp;תפקידים&nbsp;</td></tr>
<%
 	
  sqlstr="Select DISTINCT job_id, job_name From jobs WHERE ORGANIZATION_ID = " & OrgId & " Order By job_id"   
  set rs_pr = con.getRecordSet(sqlstr)
  while not rs_pr.eof
	 job_id = trim(rs_pr(0))
%>
	<tr>
	
	<%
	 sum_hour = 0
	 sqlstr = "Select SUM(minuts / 60.0) From hours_view "&_
	 " WHERE project_id = " & project_id  & " AND job_id = " & job_id
	 'Response.Write sqlstr
	 set rs_sum = con.getRecordSet(sqlstr)		
	 if not rs_sum.eof then	   
		sum_hour = rs_sum(0)		
	 end if  
	 set rs_sum = nothing
	 
	 If IsNumeric(sum_hour) Then
		sum_hour_ = FormatNumber(Round(trim(sum_hour),1),1,-1,0,0)		
	 Else
		sum_hour_ = "0.0"
	 End If	
	 
	 sum_job = 0	
	 sqlstr = "Select  SUM(minuts / 60.0 * HOUR_PAY)  From hours_view "&_
	 " WHERE project_id = " & project_id & " AND job_id = " & job_id
	'Response.Write sqlstr
	set rs_sum = con.getRecordSet(sqlstr)		
	if not rs_sum.eof then	   
		sum_job = rs_sum(0)		
	end if  
	set rs_sum = nothing	
	
	If IsNumeric(sum_job) Then
		sum_job_ = FormatNumber(Round(trim(sum_job),1),1,-1,0,0)
		sum_all = sum_all + sum_job
	Else
		sum_job_ = "0.0"
		sum_all = sum_all + 0
	End If
	%>
	<td class="card4" align=center><%=sum_job_%></td>
    <td class="card4" align=center><%=sum_hour_%></td>
	<%
	sqlstr = "Select AVG(hour_pay) From JOBS "&_
	" WHERE ORGANIZATION_ID = " & OrgID & " AND job_id = " & job_id
	'Response.Write sqlstr
	set rs_sum = con.getRecordSet(sqlstr)
	if not rs_sum.eof Then
	
	If IsNumeric(rs_sum(0)) Then
			job_pay = cDbl(rs_sum(0))
			job_pay = FormatNumber(Round(trim(job_pay),1),1,-1,0,0)
	Else
			job_pay = 0
	End If	
		
	set rs_sum = nothing
	end if	
    %>	
	<%
		sqlstr = "Select hours FROM  pricing_to_jobs WHERE pricing_id = " & project_id & " AND job_id = " & job_id
		set rs_h = con.getRecordSet(sqlstr)
		if not rs_h.eof Then
			hours_job = trim(rs_h(0))
		else
			hours_job = 0
		end if	
		set rs_h = nothing			
		
	%>
	<td class="card5" align=center><%=job_pay*hours_job%></td>
	<td class="card5" align=center><%=hours_job%></td>
	<td class="title_sort3" align=right nowrap>&nbsp;<%=trim(rs_pr("job_name"))%>&nbsp;</td>
	</tr>
<%
	rs_pr.moveNext
	Wend
	set rs_pr = Nothing		
%>
<tr>
<%
	sum_all = FormatNumber(Round(sum_all,1),1,-1,0,0)
%>
<td style="font-size:11pt" align=center><font color="#736BA6"><b><%=sum_all%></b></font></td>
<td style="font-size:11pt" colspan=2 align=left>&nbsp;<b>סה"כ לפני רווח</b>&nbsp;</td>
</tr>
<tr>
<td align=center><B>%&nbsp;<%=profit_percent%></B>&nbsp;</td>
<td colspan=2 align=left>&nbsp;<b>אחוז רווח</b>&nbsp;</td>
</tr>
<%If isNumeric(trim(profit_percent)) Then
	   sum_after_profit = cDbl(sum_all) + (cDbl(profit_percent) * cDbl(sum_all)) / 100
  Else
		sum_after_profit = sum_all
  End If
  sum_after_profit = FormatNumber(Round(sum_after_profit,1),1,-1,0,0) 	
%>
<tr>
<td style="font-size:11pt" align=center><font color="#736BA6"><b><%=sum_after_profit%></b></font></td>
<td style="font-size:11pt"  colspan=2 align=left>&nbsp;<b>עלות סופית</b>&nbsp;</td>
</tr>
<%End if%>
</table>
<%End if%>
</td></tr></table>
</td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%
set con = nothing
%>
</html>
