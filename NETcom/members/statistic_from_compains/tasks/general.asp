<%@ Language=VBScript%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
function SendData()
{
form1.submit()
}
//-->
</script> 
</head>
<body>
<FORM action="general.asp" method=POST id="form1" name="form1" target="_self">   

<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="ltr" ID="Table1">
<tr><td width="100%">
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%appDepId=Request("app_DepId")
'	response.Write "appDepId="&appDepId 
	'dim arrayCountryID as Array
	if appDepId<>"" then
	arrayappDepId=Split(appDepId,",")
	else
	appDepId=0

	end if%>
<%numOfTab = 5%>
<%numOfLink = 0%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" colspan=2>&nbsp;</td></tr>
<tr><td height="20" nowrap colspan=2></td></tr>
<tr>		
    <td width="100%" align="center" valign="top" >
	<table border="0" cellspacing="1" width="100%" align="center" cellpadding="0" dir="ltr" ID="Table2">	
	<tr><td >
	<table border="0" cellspacing="1" width="60%" align="center" cellpadding="0" dir="rtl">
	<tr>	
    <th width="70%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>" rowspan=2>
   	&nbsp;ריכוז משימות:&nbsp;</th>	
   	<th class=textp align="center" rowspan=2 width="30%">מחלקה</th>
	<th class=textp colspan="2" align="center">
	משימות קישורית
	</th>	
	<th class=textp colspan="2" align="center">
	משימות מעקב 
	</th>	<th class=textp colspan="2" align="center">
	רישום ופתיחת טופס או עדכון כרטיס
	</th>	
	</tr>	
	<tr>    

	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">חדש</th>
		<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">בטיפול</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">חדש</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">בטיפול</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">חדש</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">בטיפול</th>
	
	
    </tr>    
	<%'sqlstr = "Exec dbo.GetTasksCount '"  & OrgID  & "'" 	
	if len(appDepId)>0 and appDepId<>"0" then
sqlQuery =" where U.Department_Id in("& appDepId &")"
	arrayDepId=Split(appDepId,",")
else
sqlQuery=""
end if
		sqlstr = "select distinct * from dbo.[TBL_taskCount](264) U left join jobs J ON U.job_Id=J.job_Id left join Departments on U.Department_Id=Departments.DepartmentId	 "& sqlQuery & " order by  DepartmentName,U.FIRSTNAME,U.LASTNAME"
''	Response.Write sqlstr
		'Response.End 
		set rs_products = con.GetRecordSet(sqlstr)
		If not rs_products.eof Then
		While not rs_products.eof 
			prod_Id = trim(rs_products(0))
			product_name = trim(rs_products("FIRSTNAME"))&" "&  trim(rs_products("LASTNAME"))%>
	
	<%KNew=rs_products("KNew")
	KTipul=rs_products("KTipul")
	MaakavNew=rs_products("MaakavNew")
	MaakavTipul=rs_products("MaakavTipul")
	KartisNew=rs_products("KartisNew")
	KartisTipul=rs_products("KartisTipul")
	DepartmentName=rs_products("DepartmentName")
		'sql_Knew="SELECT COUNT(T.task_id) FROM tasks T WHERE  CHARINDEX(CONVERT(varchar, '2026'),T.task_types)>0	AND (reciver_id  = " & prod_Id &") AND (T.task_status = 1)	"
	'set rs_KNew=con.GetRecordSet(sql_Knew)
	'if not rs_KNew.eof then
	'KNew=rs_KNew(0)
	'else
	'KNew=0
	'end if
	'set rs_KNew=Nothing
	'sql_KTipul="SELECT COUNT(T.task_id) FROM tasks T WHERE  CHARINDEX(CONVERT(varchar, '2026'),T.task_types)>0	AND (reciver_id  = " & prod_Id &") AND (T.task_status = 2)	"
	'set rs_KTipul=con.GetRecordSet(sql_KTipul)
	'if not rs_KTipul.eof then
	'KTipul=rs_KTipul(0)
	'else
	'KTipul=0
	'end if
	'set rs_KTipul=Nothing
	
	'sql_MaakavNew="SELECT COUNT(T.task_id) FROM tasks T WHERE  CHARINDEX(CONVERT(varchar, '2005'),T.task_types)>0	AND (reciver_id  = " & prod_Id &") AND (T.task_status = 1)	"
	'set rs_MaakavNew=con.GetRecordSet(sql_MaakavNew)
	'if not rs_MaakavNew.eof then
	'MaakavNew=rs_MaakavNew(0)
	'else
	'MaakavNew=0
	'end if
	'set rs_MaakavNew=Nothing
	
	'sql_MaakavTipul="SELECT COUNT(T.task_id) FROM tasks T WHERE  CHARINDEX(CONVERT(varchar, '2005'),T.task_types)>0	AND (reciver_id  = " & prod_Id &") AND (T.task_status = 2)	"
	'set rs_MaakavTipul=con.GetRecordSet(sql_MaakavTipul)
	'if not rs_MaakavTipul.eof then
	'MaakavTipul=rs_MaakavTipul(0)
	'else
	'MaakavTipul=0
	'end if
	'set rs_MaakavTipul=Nothing

    'sql_KartisNew="SELECT COUNT(T.task_id) FROM tasks T WHERE  CHARINDEX(CONVERT(varchar, '2027'),T.task_types)>0	AND (reciver_id  = " & prod_Id &") AND (T.task_status = 1)	"
	'set rs_KartisNew=con.GetRecordSet(sql_KartisNew)
	'if not rs_KartisNew.eof then
	'KartisNew=rs_KartisNew(0)
	'else
	'KartisNew=0
	'end if
	'set rs_KartisNew=Nothing
	
	'sql_KartisTipul="SELECT COUNT(T.task_id) FROM tasks T WHERE  CHARINDEX(CONVERT(varchar, '2027'),T.task_types)>0	AND (reciver_id  = " & prod_Id &") AND (T.task_status = 2)	"
	'set rs_KartisTipul=con.GetRecordSet(sql_KartisTipul)
	'if not rs_KartisTipul.eof then
	'KartisTipul=rs_KartisTipul(0)
	'else
	'KartisTipul=0
	'end if
	'set rs_KartisTipul=Nothing
	
	
	%>
	<tr height=24 bgcolor="#e6e6e6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';" style="background-color: rgb(230, 230, 230);">
	<td width="70%" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<a class="link_categ" href="javascript:void(0)" onclick="window.open('taskDetailsAll.asp?pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=800, scrollbars=1');"><b><%=product_name%></b></a></td>	
	<td  width="30%" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap>&nbsp;<a class="link_categ" href="javascript:void(0)" onclick="window.open('taskDetailsAll.asp?pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=800, scrollbars=1');"><b><%=DepartmentName%></b></a></td>	
	
	<%'For jj=1 To Ubound(arr_Status)%>
   <td valign="middle" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num1" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=1&Ttype=2026&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=KNew%><%'=rs_products("KNew")%></a></td>	  
   <td  valign="middle" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num2" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=2&Ttype=2026&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=KTipul%><%'=rs_products("KTipul")%></a></td>	  
   <td  valign="middle" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num1" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=1&Ttype=2005&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=MaakavNew%><%'=rs_products("MaakavNew")%></a></td>	  
   <td  valign="middle" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num2" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=2&Ttype=2005&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=MaakavTipul%><%'=rs_products("MaakavTipul")%></a></td>	  
   <td  valign="middle" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num1" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=1&Ttype=2027&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=KartisNew%><%'=rs_products("KartisNew")%></a></td>	  
   <td  valign="middle" align="center" dir="<%=dir_obj_var%>"><a href="javascript:void(0)" class="task_status_num2" onclick="window.open('taskDetails.asp?status=2&Ttype=2027&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=KartisTipul%><%'=rs_products("KartisTipul")%></a></td>	  
	<%'Next%>		
	</tr>								
<%	rs_products.moveNext
	Wend
	Set rs_products = Nothing
    Else	%>
	<tr>		
    <th  colspan=9 width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>">
    <%If trim(lang_id) = "1" Then%>	&nbsp;לא קיימים נתונים במערכת.&nbsp;
	<%Else%>	&nbsp;Not found forms&nbsp;<%End If%>
	</th>
    </tr>
<%		
 End If	
%>	
</table></td><td width=110 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width="100%" ID="Table3">
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td colspan=2>
	<select name="app_DepId" id="app_DepId" dir="rtl" class="norm" style="width: 120px;height:200px" multiple>
    <OPTION value="0" >כל המחלקות</OPTION>
<%sqlstrDep = "SELECT departmentId, departmentName  FROM Departments   ORDER BY departmentName"
   set rs_Dep= con.getRecordSet(sqlstrDep)
 
   do while not rs_Dep.EOF 
		Dep_Id =rs_Dep("departmentId")
		Dep_Name =rs_Dep("departmentName")
		
		if appDepId<>"0" then
							for i=0 to UBound(arrayappDepId)
							if cint(Dep_Id)=cint(arrayappDepId(i)) then
							 s=1 
							  Exit For
							 else
							 s=0
							 end if%>
							<%next%>
								
			<option value="<%=Dep_Id%>" <%if s="1" then%> selected <%end if%>><%=Dep_Name%></option>
			<%else%>
		<option value="<%=Dep_Id%>"><%=Dep_Name%></option>
			<%end if%>
			<%rs_Dep.moveNext
		loop
		rs_Dep.close
		set rs_Dep=Nothing	
%>
															</select></td></tr>
																<tr>
				<td colspan="2" height="10">&nbsp;</td>
			</tr>

		<TR><td colspan=2 align=center>
		<a class="button_edit_1" style="width:90px; line-height:110%; padding:3px" href="javascript:void(0)" onclick="javascript:SendData()">הצג</a>
		</td></tr>
</table>
</td></tr>
</table></td></tR>

</tr>
<tr><td height="20" nowrap></td>
</tr>
</table>
</form>
</body>
</html>
<%  set con = nothing %>