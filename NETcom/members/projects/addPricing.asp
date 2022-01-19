<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
pricing_id=request.querystring("pricing_id")
errName = false
project_title = trim(Request.Cookies("bizpegasus")("Projectone"))

if request("addPricing")<>nil then 'after form filling
	project_id = pricing_id	 	
	profit_percent = trim(Request.Form("profit_percent"))	 
	
	sqlStr = "Update projects set profit_percent = '" & profit_percent &_					
	"' where project_ID=" & project_ID
	con.GetRecordSet (sqlStr)
	
	sqlstr = "Delete FROM pricing_to_jobs WHERE pricing_id=" & pricing_id
	con.executeQuery(sqlstr)
	
	sqlstr="Select DISTINCT job_id From jobs WHERE ORGANIZATION_ID = " & OrgId & " Order By job_id"   
	set rs_pr = con.getRecordSet(sqlstr)
	while not rs_pr.eof
		job_id = trim(rs_pr(0))
		If Request.Form("hours"&job_id) <> nil Then
			hour_job = trim(Request.Form("hours"&job_id))
		Else
			hour_job = 0
		End If
		sqlstr = "Insert INTO pricing_to_jobs (pricing_id, job_id, hours) values "&_
		"(" & pricing_id & "," & job_id & "," & hour_job & ")"
		'Response.Write sqlstr
		'Response.End
		con.executeQuery(sqlstr)
		rs_pr.moveNext
	wend
	set rs_pr = Nothing					
	Response.Redirect ("addPricing.asp?pricing_id="&pricing_id) 
  
end if
%>
<%
	If Request.QueryString("delProg") <> nil And Request.QueryString("delComp") <> nil AND pricing_id <> nil Then
		
		delProg = Request.QueryString("delProg")
		delComp = Request.QueryString("delComp")
		delMech = Request.QueryString("delMech")
		
		sqlstr = "Delete FROM pricing_to_projects WHERE pricing_id = " & pricing_id &_
		" AND project_id = " & delProg &_
		" AND company_id = " & delComp &_
		" AND mechanism_id = " & delMech
		
		con.executeQuery(sqlstr)		
		Response.Redirect "addPricing.asp?pricing_id=" & pricing_id
		
	End if
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

	function deleteProject(prog_id,comp_id,mech_id)
	{
		if(window.confirm("? האם ברצונך למחוק את המנגנון הנ''ל מרשימת מנגנונים להשוואה"))
		{
			document.location.href = "addPricing.asp?pricing_id=<%=pricing_id%>&delProg=" + prog_id + "&delComp=" + comp_id + "&delMech=" + mech_id;
		}
		return false;
	}
	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 	
<!--End-->
</script>  
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 2%>
<%If trim(wizard_id) = "" Then%>
<!--#include file="../../top_in.asp"-->
<%End If%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title"><%If IsNumeric(wizard_id) Then%> <%=page_title_%> <%Else%>&nbsp;<%if pricing_id<>nil then%>תמחור<%else%>הוספת<%end if%>&nbsp;<%=project_title%> עתידי&nbsp;<%End If%></td></tr>		   
	  		       	
   </table>
   </td>
</tr> 
<% wizard_page_id = 4 %>
<tr><td width=100% align="right">
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<%If trim(wizard_id) <> "" Then%>
<tr>
<td width=100% align="right" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width=100% ID="Table3">
<tr><td class="explain_title">צור טופס רישום חדש!</td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td class=explain>
בשדה <b>שפת הטופס</b> , בחר את השפה בה ייכתב טופס ההרשמה.
</td></tr>
<tr><td class=explain>
בשדה <b>שם הטופס</b> הזן שם לטופס הרישום. שם זה לא יופיע בדיוור אלא ישמש לזיהוי הטופס במערכת ה-Bizpower.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט מקדים</b> , הקלד את הטקסט שיופיע בתחילת הטופס לפני שדות הרישום.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט תודה לאחר הרישום</b> , הקלד את הטקסט שברצונך שהנרשם יראה לאחר שליחת הטופס. <br>לדוגמה: "תודה! נציגינו יחזרו אליך ב-24 השעות הקרובות"
</td></tr>
</table>
</td>
</tr>
<tr>
    <td width="100%" height="2" nowrap bgcolor="#FFFFFF"></td>
</tr>
<%End If%>       
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="3" cellspacing="1"  bgcolor="#E6E6E6">
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="2" cellspacing="2">
<tr><td height=10></td></tr>  
<tr>
<td>
<table cellpadding=0 cellpadding=0 align=center border=0>         
<%
found_pricing = false
if pricing_id<>nil and pricing_id<>"" then
  sqlStr = "select company_id, project_code,project_name, project_description, start_date, end_date, "&_
  " profit_percent from projects where project_ID=" & pricing_id  & " And Organization_ID = " & OrgID
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
		found_pricing = true
	else
		found_pricing = false	
	end if
	set rs_projects = nothing	

end if
'Response.Write company_id
%>
<%If found_pricing Then%>
<tr>
	<td align=right nowrap width="200px" dir=rtl class="Form_R"><%=project_code%></td>
	<td align="right" nowrap width="120px" dir=rtl>קוד <%=project_title%>&nbsp;</td>
</tr>
<tr>
	<td align=right nowrap dir=rtl class="Form_R"><%=project_name%></td>
	<td align="right" nowrap dir=rtl>שם <%=project_title%>&nbsp;</td>
</tr>
<tr>
	<td align=right dir=rtl class="Form_R"><%=project_description%></td>
	<td align="right" nowrap dir=rtl>תיאור <%=project_title%>&nbsp;</td>
</tr>
     <!-- start fields dynamics -->
	 <!--#INCLUDE FILE="project_fields.asp"-->	
	 <!-- end fields dynamics -->	
<tr>
	<td align=right dir=rtl class="Form_R"><%=company_name%></td>
	<td align="right" nowrap dir=rtl>לקוח&nbsp;</td>
</tr>
</table></td></tr>
<FORM name="frmMain" ACTION="addPricing.asp?pricing_id=<%=pricing_id%>&addPricing=1" METHOD="post">
<%if pricing_id<>nil and pricing_id<>"" then%>
<tr><td align=right colspan=2 height=10></td></tr>
<tr><td align=right colspan=2><input type=button value="הוסף מנגנון להשוואה" class="but_menu" style="width:200" onclick="window.open('addpricing_project.asp?pricing_id=<%=pricing_id%>','','top=90,left=90,height=200,width=400,status=yes,toolbar=no,menubar=no,location=no')"></td></tr><tr>
<tr><td align=right colspan=2 height=5></td></tr>
<tr>
<td colspan=2 bgcolor="#E6E6E6">
<table cellpadding=2 cellspacing=1 width=100% bgcolor="#F0F0F0">
<%
	sum_all = 0
	sqlstr = "Select DISTINCT project_id,company_id,mechanism_id,project_name,company_name,mechanism_name From pricing_to_jobs_view WHERE pricing_id = " & pricing_id & " Order By project_id"
	set rs = con.getRecordSet(sqlstr)
	if not rs.eof then
		projNumber = rs.recordCount
%>
<tr><td class="title_sort" align=center colspan=2>&nbsp;תמחור <%=project_title%> עתידי&nbsp;</td>
<td class="title_sort" align=center colspan=2>&nbsp;ממוצעים&nbsp;</td>
<td class="title_sort" align=center colspan="<%=projNumber%>">&nbsp;מנגנונים להשוואה&nbsp;</td>
<td class="title_sort">&nbsp;</td></tr>
<tr><td class="title_sort" align=center colspan=2>&nbsp;&nbsp;<%=company_name%>&nbsp;<br>&nbsp;<font color="#555290"><%=project_name%></font>&nbsp;</td>
<td class="title_sort" colspan=2>&nbsp;</td>
<% while not rs.eof	%>
<td class="title_sort" align=right valign=top>
<table cellpadding=1 cellspacing=1 align=right border=0>
<tr><td align=right nowrap><%=rs(4)%><br><%=rs(3)%><br><%=rs(5)%></td>
<td align=right valign=top><img src="../../images/delete_icon.gif" class="hand" hspace=0 vspace=0 border=0 style="vertical-align:top" onclick="return deleteProject('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>')" title="מחק את המנגנון מרשימת <%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))	%> להשוואה"></td>
 </tr>
</table>
</td>
<%	
   rs.moveNext
   wend
   set rs = nothing	
%>
<td class="title_sort" align=right>&nbsp;</td></tr>
<tr style="line-height:20px">
<td class="title_sort2" width=67 nowrap align=center>עלות</td>
<td class="title_sort2" width=67 nowrap align=center>שעות</td>
<td class="title_sort2" width=67 nowrap align=center>עלות בש"ח</td>
<td class="title_sort2" width=67 nowrap align=center>שעות</td>
<%For i=1 To projNumber%>
<td align=center class="title_sort2">סה"כ שעות</td>
<%Next%>
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
	 sqlstr = "Select SUM(minuts / 60.0) From pricing_to_jobs_view "&_
	 " WHERE pricing_id = " & pricing_id  & " AND job_id = " & job_id
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
	
	avg_pay = 0
	
	sqlstr = "Select AVG(hour_pay) From USERS "&_
	" WHERE ORGANIZATION_ID = " & OrgID & " AND job_id = " & job_id
	'Response.Write sqlstr
	set rs_sum = con.getRecordSet(sqlstr)
	if not rs_sum.eof Then
	
	   If IsNumeric(rs_sum(0)) Then
			avg_pay = cDbl(rs_sum(0))
			avg_pay = FormatNumber(Round(trim(avg_pay),1),1,-1,0,0)
	   Else
			avg_pay = 0
	   End If	
	   
	set rs_sum = nothing
	end if	
		
    %>	
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
	
		sqlstr = "Select hours FROM  pricing_to_jobs WHERE pricing_id = " & pricing_id & " AND job_id = " & job_id
		set rs_h = con.getRecordSet(sqlstr)
		if not rs_h.eof Then
			hours_job = trim(rs_h(0))
		else
			hours_job = 0
		end if	
		set rs_h = nothing	
		
		sum_all = sum_all + job_pay*hours_job
	%>	
	<td class="card4" align=center><%=job_pay*hours_job%></td>
	<td class="card4" align=center><input type=text name="hours<%=job_id%>" value="<%=hours_job%>" class="texts" style="width:40" onkeypress="GetNumbers()"></td>
	<td class="card5" align=center><%=sum_hour_*avg_pay%></td>
    <td class="card5" align=center><%=sum_hour_%></td>
	<%
	sqlstr = "Select DISTINCT company_id, project_id FROM pricing_to_projects WHERE pricing_id = " & pricing_id & " Order BY company_id, project_id"
	set rs_comp = con.getRecordSet(sqlstr)
	while not rs_comp.eof
		company_id = rs_comp(0)
		project_id = rs_comp(1)		
		
		sqlstr = "Select Sum(minuts / 60) From pricing_to_jobs_view " &_
		" WHERE pricing_id = " & pricing_id & " AND job_id = " & job_id &_
		" AND company_id = " & company_id & " AND project_id = " & project_id
		'Response.Write sqlstr
		'Response.End
		set rs = con.getRecordSet(sqlstr)
		if not rs.eof then
			If IsNumeric(rs(0)) Then
				hours = FormatNumber(Round(trim(rs(0)),1),1,-1,0,0)
			else
				hours = 0
			end if
		%>	
		<td class="card5" align=center><%=hours%></td>		
		<%
		else
		%>
		<td class="card5" align=center>0</td>
		<%		    
		end if
		set rs = nothing
	
	rs_comp.moveNext
	Wend
	set rs_comp = Nothing		
					
	%>
	<td class="title_sort3" align=right>&nbsp;<%=rs_pr("job_name")%>&nbsp;</td>
	</tr>
<%
	rs_pr.moveNext
	Wend
	set rs_pr = Nothing		
%>
</table>
</td>
</tr>
<tr>
<td>
<table cellpadding=0 cellspacing=0 ID="Table2">
<tr>
<%
	sum_all = FormatNumber(Round(sum_all,1),1,-1,0,0)
%>
<td style="font-size:11pt" width=80 nowrap align=center><font color="#736BA6"><b><%=sum_all%></b></font></td>
<td style="font-size:11pt" width=100% align=left>&nbsp;<b>סה"כ לפני רווח</b>&nbsp;</td>
</tr>
</table>
</td>
</tr>
<tr>
<td>
<table cellpadding=0 cellspacing=0 >
<tr>
<td width=80 nowrap align=center>%<input type=text name="profit_percent" id="profit_percent" value="<%=profit_percent%>" class=texts style="width:40" onkeypress="GetNumbers()">&nbsp;</td>
<td width=100% align=left>&nbsp;<b>אחוז רווח</b>&nbsp;</td>
</tr>
</table>
</td>
</tr>
<%If isNumeric(trim(profit_percent)) Then
	   sum_after_profit = cDbl(sum_all) + (cDbl(profit_percent) * cDbl(sum_all)) / 100
  Else
		sum_after_profit = sum_all
  End If
  sum_after_profit = FormatNumber(Round(sum_after_profit,1),1,-1,0,0) 	
%>
<tr>
<td>
<table cellpadding=0 cellspacing=0 ID="Table1">
<tr>
<td style="font-size:11pt" width=80 nowrap align=center><font color="#736BA6"><b><%=sum_after_profit%></b></font></td>
<td style="font-size:11pt" width=100% align=left>&nbsp;<b>עלות סופית</b>&nbsp;</td>
</tr>
</table>
</td>
</tr>
<%End if%>
<%End if%>
<tr bgcolor="#E6E6E6"><td colspan="2" height=10></td></tr>
<%If trim(wizard_id) <> "" Then%>	
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center bgcolor="#E6E6E6" ID="Table4">
<tr>
<td width=50% align=right><input type=button class="but_menu" style="width:90px" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id%>.asp'" value="ביטול" ID="Button1" NAME="Button1"></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=submit class="but_menu" style="width:90px" value="אישור" ID="Button2" NAME="Button2"></td></tr>
</table></td></tr>		
<%Else%>	
<tr bgcolor="#E6E6E6">
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center bgcolor="#E6E6E6" >
<tr>
<td width=50% align=right><input type=button class="but_menu" style="width:90px" onclick="document.location.href='prices.asp';" value="ביטול"></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=submit class="but_menu" style="width:90px" value="אישור"></td></tr>
</table></td></tr>
<%End If%>
</form>
<tr><td colspan="2" height=10></td></tr></table>
</td></tr>
<%End If%>
<tr><td colspan="2" height=10></td></tr>
</table>
</body>
<%
set con = nothing
%>
</html>
