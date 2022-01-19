<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	projectID = trim(Request.QueryString("project_id"))
	project_title = trim(Request.Cookies("bizpegasus")("Projectone"))
%>
<html>
<head>
<!--#include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOfTab = 3%>
<%numOfLink = 1%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table width="100%" border="0" cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;שעות עבודה על&nbsp;<%=project_title%>&nbsp;</td></tr>		   	  		       	
     </table>
   </td>
</tr> 
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="2" cellspacing="2">
<tr><td height=10></td></tr>  
<tr>
<td>
<table cellpadding=2 cellpadding=2 align=right width=360>
<%
found_project = false
if projectID<>nil and projectID<>"" then
  sqlStr = "Select company_id, project_code,project_name, project_description, start_date, end_date,"&_
  " profit_percent From projects Where project_ID = " & projectID & " And ORGANIZATION_ID = " & OrgID 
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
		found_project = true
	else
		found_project = false	
	end if
	set rs_projects = nothing
end if
'Response.Write company_id
%>
<%If found_project Then%>
<tr>
	<td align=right nowrap width="200px" dir=rtl class="Form_R"><%=project_code%></td>
	<td align="right" nowrap width="120px" dir=rtl>קוד <%=project_title%></td>
</tr>
<tr>
	<td align=right nowrap dir=rtl class="Form_R"><%=project_name%></td>
	<td align="right" nowrap dir=rtl>שם <%=project_title%></td>
</tr>
<tr>
	<td align=right dir=rtl class="Form_R" style="line-height:120%;"><%=project_description%></td>
	<td align="right" nowrap dir=rtl>תיאור <%=project_title%></td>
</tr>
     <!-- start fields dynamics -->		
	 <!--#INCLUDE FILE="project_fields.asp"-->	
	 <!-- end fields dynamics -->	
<tr>
	<td align=right dir=rtl class="Form_R"><%=company_name%></td>
	<td align="right" nowrap dir=rtl>לקוח</td>
</tr>
</table></td></tr>
<%if projectID<>nil and projectID<>"" then%>
<tr>
<td>
<%
  sum_hours = 0 : sum_all = 0
  sqlstr="Select DISTINCT job_id, job_name From HOURS_VIEW WHERE project_id = " & projectID &" And ORGANIZATION_ID = " & OrgId & " Order By job_id"   
  set rs_pr = con.getRecordSet(sqlstr)
  if not rs_pr.eof then
%>
 <table cellpadding=3 cellspacing=1 border=0 bgcolor="#FFFFFF" align=right width=360>
	<tr style="line-height:22px">
	<td class="title_sort3" width=70 nowrap align=center>עלות</td>
	<td class="title_sort3" width=70 nowrap align=center>שעות</td>
	<td align=right class="title_sort3" width=100 nowrap>&nbsp;עובד&nbsp;</td>
	<td align=right class="title_sort3" width=120 nowrap>&nbsp;תפקיד&nbsp;</td></tr>
<%  
  while not rs_pr.eof
	 job_id = trim(rs_pr(0))	 
	 
	 sum_hour = 0
	 sum_hours = 0
	 sqlstr = "Select SUM(minuts) From hours_view "&_
	 " WHERE project_id = " & projectID  & " AND job_id = " & job_id
	 'Response.Write sqlstr
	 set rs_sum = con.getRecordSet(sqlstr)		
	 if not rs_sum.eof then	   
		sum_hour = cDbl(rs_sum(0)) / 60
		sum_hours = sum_hours + sum_hour	
		sum_hours_all=sum_hours_all+ sum_hour	
	 end if  
	 set rs_sum = nothing
	 
	 If IsNumeric(sum_hour) Then
		sum_hour_ = FormatNumber(Round(trim(sum_hour),1),1,-1,0,0)		
	 Else
		sum_hour_ = "0.0"
	 End If	
	 
	 sum_job = 0	
	 sqlstr = "Select HOUR_PAY From Jobs Where job_id = " & job_id
	'Response.Write sqlstr
	set rs_sum = con.getRecordSet(sqlstr)		
	if not rs_sum.eof then	   
		sum_job = cDbl(rs_sum(0)) * sum_hours
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
	<tr>
	<td class="title_sort1" align=center><b><%=sum_job_%></b></td>
    <td class="title_sort1" align=center><b><%=sum_hour_%></b></td>
    <td class="title_sort1" align=center>&nbsp;</td>		
	<td class="title_sort1" align=right nowrap>&nbsp;<b><%=trim(rs_pr("job_name"))%></b>&nbsp;</td>
	</tr>
	<%
	' הצג רשימת על העובדים בתפקיד הנ''ל
	sqlstr = "Select FIRSTNAME +  ' ' + LASTNAME, HOUR_PAY, USER_ID FROM USERS WHERE ORGANIZATION_ID = " & OrgID & " AND job_id = " & job_id &_
	" AND USER_ID IN (Select Distinct USER_ID FROM HOURS WHERE project_id = " & projectID & " AND job_id = " & job_id	& ")"
	set rs_users = con.getRecordSet(sqlstr)
	while not rs_users.eof 		
		sum_hour = 0
		hour_pay = rs_users(1)
		
		sqlstr = "Select SUM(minuts) From hours_view "&_
		" WHERE project_id = " & projectID & " AND USER_ID = " & trim(rs_users(2))
		'Response.Write sqlstr
		set rs_sum = con.getRecordSet(sqlstr)		
		if not rs_sum.eof then	   
			sum_hour = cDbl(rs_sum(0)) / 60
		end if  
		set rs_sum = nothing	
		
		If IsNumeric(sum_hour) Then
			sum_hour_ = FormatNumber(Round(trim(sum_hour),1),1,-1,0,0)		
		Else
			sum_hour_ = "0.0"
		End If		
		
		sum_user = sum_hour * cDbl(hour_pay)		
		
		If IsNumeric(sum_user) Then
			sum_user_ = FormatNumber(Round(trim(sum_user),1),1,-1,0,0)		
		Else
			sum_user_ = "0.0"
		End If		
	%>
	<tr>
	<td class="card1" align=center><%=sum_user_%></td>
    <td class="card1" align=center><%=sum_hour_%></td>	
	<td class="card1" align=right nowrap>&nbsp;<%=trim(rs_users(0))%>&nbsp;</td>
	<td class="card1" align=right>&nbsp;</td>
	</tr>	
	
	<%
		rs_users.moveNext
		Wend
		set rs_users = Nothing
	%>
	<tr><td colspan=4 height=1 bgcolor="#736BA6"></td></tr>
<%
	rs_pr.moveNext
	Wend
	set rs_pr = Nothing	
	
%>
<tr>
<%
	sum_all = FormatNumber(Round(sum_all,1),1,-1,0,0)
	sum_hours = FormatNumber(Round(sum_hours,1),1,-1,0,0)
	sum_hours_all =  FormatNumber(Round(sum_hours_all,1),1,-1,0,0)
%>
<td style="font-size:11pt" align=center><font color="#736BA6"><b><%=sum_all%></b></font></td>
<td style="font-size:11pt" align=center><font color="#736BA6"><b><%=sum_hours_all%></b></font></td>
<td style="font-size:11pt" align=left colspan=2>&nbsp;<b>סה"כ</b>&nbsp;</td>
</tr>
<%End if%>
</table>
<%end if %>
</td></tr>
<%sqlstr = "Select mechanism_id, mechanism_name, budget_hours, proposal_hours From mechanism Where Project_id = " & projectID
  set rs_mech = con.getRecordSet(sqlstr)
  while not rs_mech.eof
	mechID = trim(rs_mech(0))
	mechName = trim(rs_mech(1))
	budget_hours = trim(rs_mech(2))
	proposal_hours = trim(rs_mech(3))
	
	If IsNull(budget_hours) Or IsNumeric(budget_hours)=false Then
		budget_hours = 0
	End If
	If IsNull(proposal_hours) Or IsNumeric(proposal_hours)=false Then
		proposal_hours = 0
	End If	
	
	sum_hours = 0 : sum_all = 0 : sum_hours_all = 0 : sum_days_all = 0
	sqlstr="Select DISTINCT job_id, job_name From HOURS_VIEW WHERE project_id = " & projectID &_
	" And mechanism_id = " & mechID & " And ORGANIZATION_ID = " & OrgId & " Order By job_id"   
	set rs_pr = con.getRecordSet(sqlstr)
	if not rs_pr.eof then
%>
<tr><td><table cellpadding=3 cellspacing=1 border=0 bgcolor="#FEFCF1" align=right width=360>
    <tr><td colspan=10 align=center class="card3"><%=mechName%></td></tr>
	<tr style="line-height:22px">
	<td class="title_sort4" width=70 nowrap align=center>עלות</td>
	<td class="title_sort4" width=70 nowrap align=center>הצעה</td>
	<td class="title_sort4" width=70 nowrap align=center>הפרש</td>
	<td class="title_sort4" width=70 nowrap align=center>שעות ביצוע</td>
	<td class="title_sort4" width=70 nowrap align=center>שעות תכנון</td>
	<td class="title_sort4" width=70 nowrap align=center>הפרש</td>
	<td class="title_sort4" width=70 nowrap align=center>ימי ביצוע</td>
	<td class="title_sort4" width=70 nowrap align=center>ימי תכנון</td>
	<td align=right class="title_sort4" width=100 nowrap>&nbsp;עובד&nbsp;</td>
	<td align=right class="title_sort4" width=120 nowrap>&nbsp;תפקיד&nbsp;</td></tr>
<%  
  while not rs_pr.eof
	 job_id = trim(rs_pr(0))	 
	 
	 sum_hour = 0 : sum_hours = 0
	 sqlstr = "Select SUM(minuts) From hours_view Where project_id = " & projectID &_
	 " And mechanism_id = " & mechID & " AND job_id = " & job_id
	 'Response.Write sqlstr
	 set rs_sum = con.getRecordSet(sqlstr)		
	 if not rs_sum.eof then	   
		sum_hour = cDbl(rs_sum(0)) / 60
		sum_hours = sum_hours + sum_hour	
		sum_hours_all=sum_hours_all+ sum_hour	
	 end if  
	 set rs_sum = nothing
	 
	 If IsNumeric(sum_hour) Then
		sum_hour_ = FormatNumber(Round(trim(sum_hour),1),1,-1,0,0)		
	 Else
		sum_hour_ = "0.0"
	 End If
	 
	 sum_day = 0 : sum_days = 0
	 sqlstr = "Select COUNT(DISTINCT DATE) From hours_view Where project_id = " & projectID &_
	 " And mechanism_id = " & mechID & " AND job_id = " & job_id
	 'Response.Write sqlstr
	 set rs_sum = con.getRecordSet(sqlstr)		
	 if not rs_sum.eof then	   
		sum_day = cInt(rs_sum(0))
		sum_days = sum_days + sum_day	
		sum_days_all = sum_days_all + sum_days	
	 end if  
	 set rs_sum = nothing
	 
	 If IsNumeric(sum_hour) Then
		sum_day_ = FormatNumber(Round(trim(sum_day),1),0,0,0)
	 Else
		sum_day_ = "0"
	 End If	 
	 
	 sum_job = 0	
	 sqlstr = "Select HOUR_PAY From Jobs Where job_id = " & job_id
	'Response.Write sqlstr
	set rs_sum = con.getRecordSet(sqlstr)		
	if not rs_sum.eof then	   
		sum_job = cDbl(rs_sum(0)) * sum_hours
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
	<tr>
	<td class="title_sort5" align=center><b><%=sum_job_%></b></td>
	<td class="title_sort5" align=center>&nbsp;</td>
	<td class="title_sort5" align=center>&nbsp;</td>
    <td class="title_sort5" align=center><b><%=sum_hour_%></b></td>
    <td class="title_sort5" align=center>&nbsp;</td>
    <td class="title_sort5" align=center>&nbsp;</td>
    <td class="title_sort5" align=center><b><%=sum_day_%></b></td>
    <td class="title_sort5" align=center>&nbsp;</td>
    <td class="title_sort5" align=center>&nbsp;</td>    	
	<td class="title_sort5" align=right nowrap>&nbsp;<b><%=trim(rs_pr("job_name"))%></b>&nbsp;</td>
	</tr>
	<%
	' הצג רשימת על העובדים בתפקיד הנ''ל
	sqlstr = "Select FIRSTNAME +  ' ' + LASTNAME, HOUR_PAY, USER_ID FROM USERS WHERE ORGANIZATION_ID = " & OrgID & " AND job_id = " & job_id &_
	" AND USER_ID IN (Select Distinct USER_ID FROM HOURS WHERE project_id = " & projectID & " And mechanism_id = " & mechID & " AND job_id = " & job_id	& ")"
	set rs_users = con.getRecordSet(sqlstr)
	while not rs_users.eof 		
		sum_hour = 0
		hour_pay = rs_users(1)
		sqlstr = "Select SUM(minuts) From hours_view WHERE project_id = " & projectID &_
		" And mechanism_id = " & mechID & " AND USER_ID = " & trim(rs_users(2))
		'Response.Write sqlstr
		set rs_sum = con.getRecordSet(sqlstr)		
		if not rs_sum.eof then	   
			sum_hour = cDbl(rs_sum(0)) / 60
		end if  
		set rs_sum = nothing	
		
		If IsNumeric(sum_hour) Then
			sum_hour_ = FormatNumber(Round(trim(sum_hour),1),1,-1,0,0)		
		Else
			sum_hour_ = "0.0"
		End If
		
		sum_day = 0 : sum_days = 0
		sqlstr = "Select COUNT(DISTINCT DATE) From hours_view Where project_id = " & projectID &_
		" And mechanism_id = " & mechID & " AND USER_ID = " & trim(rs_users(2))
		'Response.Write sqlstr
		set rs_sum = con.getRecordSet(sqlstr)		
		if not rs_sum.eof then	   
			sum_day = cInt(rs_sum(0))
		end if  
		set rs_sum = nothing
		 
		If IsNumeric(sum_hour) Then
			sum_day_ = FormatNumber(Round(trim(sum_day),1),0,0,0)
		Else
			sum_day_ = "0"
		End If
			
		sum_user = sum_hour * cDbl(hour_pay)		
		
		If IsNumeric(sum_user) Then
			sum_user_ = FormatNumber(Round(trim(sum_user),1),1,-1,0,0)		
		Else
			sum_user_ = "0.0"
		End If
		
	%>
	<tr>
	<td class="card4" align=center><%=sum_user_%></td>
	<td class="card4" align=center>&nbsp;</td>
	<td class="card4" align=center>&nbsp;</td>
    <td class="card4" align=center><%=sum_hour_%></td>
    <td class="card4" align=center>&nbsp;</td>
    <td class="card4" align=center>&nbsp;</td>
    <td class="card4" align=center><%=sum_day_%></td>
    <td class="card4" align=center>&nbsp;</td>
	<td class="card4" align=right nowrap>&nbsp;<%=trim(rs_users(0))%>&nbsp;</td>
	<td class="card4" align=right>&nbsp;</td>
	</tr>	
	
	<%
		rs_users.moveNext
		Wend
		set rs_users = Nothing
	%>
	<tr><td colspan=10 height=1 bgcolor="#D5AA00"></td></tr>
<%
	rs_pr.moveNext
	Wend
	set rs_pr = Nothing	
	
%>
<tr>
<%
	difference = budget_hours - sum_hours_all
	If difference < 0 Then
		dif_color = "Red"
	Else
		dif_color = "Green"
	End If	
	difference = FormatNumber(Round(difference,1),1,-1,0,0)
	sum_all = FormatNumber(Round(sum_all,1),1,-1,0,0)
	sum_hours = FormatNumber(Round(sum_hours,1),1,-1,0,0)
	sum_hours_all = FormatNumber(Round(sum_hours_all,1),1,-1,0,0)	
	
	sqlstr = "Select COUNT(DISTINCT DATE) From hours_view Where project_id = " & projectID &_
	" And mechanism_id = " & mechID 
	'Response.Write sqlstr
	set rs_sum = con.getRecordSet(sqlstr)		
	if not rs_sum.eof then	   
		sum_day_pr = cInt(rs_sum(0))
	end if  
	set rs_sum = nothing
	
	If IsNumeric(sum_day_pr) = false Then
		sum_day_pr = 0
	End If
	
	planed_days = 0 : planed_day = 0
	sqlstr = "Select DateDiff(d,Start_date,End_Date) From mechanism Where project_id = " & projectID &_
	" And mechanism_id = " & mechID
	'Response.Write sqlstr
	set rs_dates = con.getRecordSet(sqlstr)		
	while not rs_dates.eof
		planed_day = trim(rs_dates(0))
		If IsNumeric(planed_day) And Len(planed_day) > 0 Then
			planed_days = planed_days + (planed_day + 1)
		End If
		rs_dates.moveNext
	wend
	set rs_dates = nothing
	
	If IsNumeric(planed_days) = false Then
		planed_days = 0
	End If	
	
	difference_days = planed_days - sum_day_pr
	If difference < 0 Then
		dif_color_pr = "Red"
	Else
		dif_color_pr = "Green"
	End If	
	
%>
<td style="font-size:11pt" align=center><font color="#D5AA00"><b><%=sum_all%></b></font></td>
<td style="font-size:11pt" align=center><font color="#D5AA00"><b><%=proposal_hours%></b></font></td>
<td style="font-size:11pt" align=center><font color="<%=dif_color%>"><b><%=difference%></b></font></td>
<td style="font-size:11pt" align=center><font color="#D5AA00"><b><%=sum_hours_all%></b></font></td>
<td style="font-size:11pt" align=center><font color="#D5AA00"><b><%=budget_hours%></b></font></td>
<td style="font-size:11pt" align=center><font color="<%=dif_color_pr%>"><b><%=difference_days%></b></font></td>
<td style="font-size:11pt" align=center><font color="#D5AA00"><b><%=sum_day_pr%></b></font></td>
<td style="font-size:11pt" align=center><font color="#D5AA00"><b><%=planed_days%></b></font></td>
<td style="font-size:11pt" align=left>&nbsp;<b>סה"כ</b>&nbsp;</td>
</tr>
</table></td></tr>
<%  End If
	rs_mech.moveNext
	Wend
	set rs_mech = Nothing
%>
</table>
</td></tr>
<%End If%>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%
set con = nothing
%>
</html>
