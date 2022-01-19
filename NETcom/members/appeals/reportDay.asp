
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%end_date=request("end_date")
  selDep_id=request("selDep_id")
  if not IsDate(end_date) then
  end_date=day(Now())&"/"& month(Now()) &"/"& year(Now())
  
  end if
%>

	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="stylesheet" type="text/css">
		<script src="http://code.jquery.com/jquery-1.9.1.js"></script>
	
	<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>
		<script>
	
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}
	</script>
	</head>
<body style="margin: 0px;" onload="self.focus();"  >
  <form id="form1" name="form1" action="reportDay.asp" method="post">
	  <table cellpadding="0" cellspacing="0" dir="<%=dir_var%>" width="100%" border="0" ID="Table1">  
  <tr><td width="100%"><A name="table_appeals"></A>  
  <table cellpadding=0 cellspacing=0 width="100%" dir="<%=dir_var%>" border="0" ID="Table2">
  <tr>
    <td valign="top" class="title_form" width="100%" align="<%=align_var%>" dir=<%=dir_obj_var%>>&nbsp;סיכום יומי&nbsp;<font color="#E6E6E6"></font>&nbsp;</td>
  </tr>
   <tr>  
  <td class="title_form"  align=center>
	<table cellpadding=2 cellspacing=0 align=right width=50%>
		<tr>
		<td><input type="submit" value="הצג נתונים" class="button_edit_2" 
	style="width: 90px; display: inline;" ID="btnData" NAME="btnData"></td>
		<td align="center"><a href='' onclick='callCalendar(document.form1.end_date,"Asend_date");return false;' id='Asend_date'><image src='../../images/calend.gif' border=0 vspace=0 valign=middle></a>&nbsp;<input dir='ltr' type='text' class="Form_R" style="width:75" id="end_date" name='end_date' value='<%=end_date%>' maxlength=10 readonly></td>	
	
<td align="right">
<select name="selDep_id" dir="rtl" class="search" style="width: 150px; display: inline;" id="selDep_id" onchange=javascript:ChangeSelect(this);>
		<OPTION value="">בחר מחלקה</OPTION>
<%
sqlstrDep = "SELECT departmentId, departmentName  FROM Departments   ORDER BY PriorityLevel"
    

	'Response.Write sqlstr
	'Response.End
	set rs_Dep= con.getRecordSet(sqlstrDep)
   do while not rs_Dep.EOF 
		Dep_Id =rs_Dep("departmentId")
		Dep_Name =rs_Dep("departmentName")%>
																<OPTION value="<%=Dep_Id%>" <%If trim(Dep_Id) = trim(selDep_id) Then%> selected <%End If%>><%=Dep_Name%>
																</OPTION>
																<%rs_Dep.moveNext
		loop
		rs_Dep.close
		set rs_Dep=Nothing	%>
</select>
</td></tr></table></td></tr>
<tr>
	<td width="100%" align="center" valign="top" height=15 bgcolor="#808080">&nbsp;</td></tr>
<tr>
	<td width="100%" align="center" valign="top">
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>" ID="Table3" width=100%>	
    <tr>
	    <td nowrap class="title_sort" dir="rtl" align="center"><b>כמות הביטוחים שנשלו/לא נשלחו</b></td>
		<td nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><b>כמות הרישומים שלו</b></td>
		<td  nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><b>המוכר הטוב ביותר</b></td>
		<td   nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><b>מכירות ליום<br><%=end_date%></b></td>
		<td   nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><b>מכירות לחודש<br><%=month(end_date)%>/<%=year(end_date)%></b></td>
		<td   nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><b>שם מחלקה</b></td>
		
			</tr>
	<%
		j=0
		if selDep_id="" then  selDep_id=0
	if selDep_id>0 then
	 sqlstr = "SELECT departmentId, departmentName  FROM Departments where departmentId="& selDep_id &"  ORDER BY PriorityLevel "
	else
	 sqlstr = "SELECT departmentId, departmentName  FROM Departments ORDER BY PriorityLevel"
    end if
   'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)

	Do while (Not app.eof)
	departmentId=app("departmentId")
	sqlMonth="SET DATEFORMAT DMY;select sum(cast(FIELD_VALUE as int)) as C1 from FORM_VALUE fv left join  Appeals App on App.Appeal_Id=fv.APPEAL_ID where "& _
    " App.Department_Id="& departmentId &" and  month(APPEAL_DATE)=month('"& end_date &"') and  year(APPEAL_DATE)=year('"& end_date &"') and Questions_Id = 16735 and "& _  
    " (fv.Field_Id = 40660) and   (Docket_Status IN (2,3))"
	'and (FIELD_ID = 40622 and len(FIELD_VALUE)=0) 
set AppMonth=con.GetRecordSet(sqlMonth)
if not AppMonth.eof then
C1=AppMonth("C1")
else
C1=0
end if
if DateDiff("d",end_date,now())=0 then
	sqlday="SET DATEFORMAT DMY;select sum(cast(FIELD_VALUE as int)) as CDay from FORM_VALUE fv left join  Appeals App on App.Appeal_Id=fv.APPEAL_ID where "& _
    " App.Department_Id="& departmentId &" and   month(APPEAL_DATE)='"& month(end_date) &"' and day(appeal_date)='"& day(end_date)& "' and year(appeal_date)='"& year(end_date)&"' and Questions_Id = 16735 and "& _  
    " (fv.Field_Id = 40660) "
else
	sqlday="SET DATEFORMAT DMY;select sum(cast(FIELD_VALUE as int)) as CDay from FORM_VALUE fv left join  Appeals App on App.Appeal_Id=fv.APPEAL_ID where "& _
    " App.Department_Id="& departmentId &" and   month(APPEAL_DATE)='"& month(end_date) &"' and day(appeal_date)='"& day(end_date)& "' and year(appeal_date)='"& year(end_date)&"' and Questions_Id = 16735 and "& _  
    " (fv.Field_Id = 40660) and   (Docket_Status IN (2,3))"
end if

	'and (FIELD_ID = 40622 and len(FIELD_VALUE)=0) 
set Appday=con.GetRecordSet(sqlday)
if not Appday.eof then
CDay=Appday("CDay")
else
CDay=0
end if
'---------------------
if DateDiff("d",end_date,now())=0 then
 sqlOwner="SET DATEFORMAT DMY;select top 1 sum(cast(FIELD_VALUE as int)) as SumOwner,User_Id_Order_Owner,(U.FIRSTNAME + Char(32) + U.LASTNAME) as USER_NAME from FORM_VALUE fv left join  Appeals App on App.Appeal_Id=fv.APPEAL_ID "& _
 " left join Users U on U.User_Id=App.User_Id_Order_Owner where App.Department_Id="& departmentId &" and  " & _
 " month(APPEAL_DATE)=month('"& end_date &"')  and  year(APPEAL_DATE)=year('" & end_date &"') and  day(APPEAL_DATE)=day('" &end_date &"')  and Questions_Id = 16735 and   "& _
 "(fv.Field_Id = 40660)  group by User_Id_Order_Owner,U.FIRSTNAME,U.LASTNAME order by SumOwner desc"

else
 sqlOwner="SET DATEFORMAT DMY;select top 1 sum(cast(FIELD_VALUE as int)) as SumOwner,User_Id_Order_Owner,(U.FIRSTNAME + Char(32) + U.LASTNAME) as USER_NAME from FORM_VALUE fv left join  Appeals App on App.Appeal_Id=fv.APPEAL_ID "& _
 " left join Users U on U.User_Id=App.User_Id_Order_Owner where App.Department_Id="& departmentId &" and  " & _
 " month(APPEAL_DATE)=month('"& end_date &"')  and  year(APPEAL_DATE)=year('" & end_date &"') and  day(APPEAL_DATE)=day('" &end_date &"')  and Questions_Id = 16735 and   "& _
 "(fv.Field_Id = 40660) and   (Docket_Status IN (2,3)) group by User_Id_Order_Owner,U.FIRSTNAME,U.LASTNAME order by SumOwner desc"
 end if
 set appOwner=con.GetRecordSet(sqlOwner)
 if not appOwner.eof then
 User_Owner=appOwner("USER_NAME")
 SumOwner=appOwner("SumOwner")
 else
 User_Owner=""
 SumOwner=0
 end if
 
 
 if DateDiff("d",end_date,now())=0 then
 sqlInsurance="SET DATEFORMAT DMY;select count(Insurance_Status) as InsuranceCount from   Appeals App  where" & _
    " App.Department_Id="& departmentId &" and   month(APPEAL_DATE)=month('" & end_date &"') and day(appeal_date)=day('" & end_date &"') and year(appeal_date)=year('" & end_date &"') " & _
    " and Questions_Id = 16735  and Insurance_Status=1 "

else
 sqlInsurance="SET DATEFORMAT DMY;select count(Insurance_Status) as InsuranceCount from   Appeals App  where" & _
    " App.Department_Id="& departmentId &" and   month(APPEAL_DATE)=month('" & end_date &"') and day(appeal_date)=day('" & end_date &"') and year(appeal_date)=year('" & end_date &"') " & _
    " and Questions_Id = 16735 and   (Docket_Status IN (2,3)) and Insurance_Status=1 "
end if
set appInsurance=con.GetRecordSet(sqlInsurance)
if not appInsurance.eof then
 InsuranceCount=appInsurance("InsuranceCount")
end if
 if DateDiff("d",end_date,now())=0 then
 sqlInsuranceNo="SET DATEFORMAT DMY;select count(Insurance_Status) as InsuranceCountNo from   Appeals App  where" & _
    " App.Department_Id="& departmentId &" and   month(APPEAL_DATE)=month('" & end_date &"') and day(appeal_date)=day('" & end_date &"') and year(appeal_date)=year('" & end_date &"') " & _
    " and Questions_Id = 16735  and Insurance_Status=0 "

else
 sqlInsuranceNo="SET DATEFORMAT DMY;select count(Insurance_Status) as InsuranceCountNo from   Appeals App  where" & _
    " App.Department_Id="& departmentId &" and   month(APPEAL_DATE)=month('" & end_date &"') and day(appeal_date)=day('" & end_date &"') and year(appeal_date)=year('" & end_date &"') " & _
    " and Questions_Id = 16735 and   (Docket_Status IN (2,3)) and Insurance_Status=0 "
end if
set appInsuranceNo=con.GetRecordSet(sqlInsuranceNo)
if not appInsuranceNo.eof then
 InsuranceCountNo=appInsuranceNo("InsuranceCountNo")
end if



'---------------------

	If j Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If		
	%>

	<tr height=30 bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
	    <td nowrap  dir="rtl" align="center"><%=InsuranceCountNo%>/<%=InsuranceCount%></td>
		<td nowrap  dir="<%=dir_obj_var%>" align="center"><%=SumOwner%></td>
		<td nowrap  dir="<%=dir_obj_var%>" align="center"><%=User_Owner%></td>
		<td nowrap  dir="<%=dir_obj_var%>" align="center"><%IF len(CDay)>0 then%><%=CDay%><%else%>0<%end if%></td>
		<td nowrap  dir="<%=dir_obj_var%>" align="center"><%IF len(C1)>0 then%><%=C1%><%else%>0<%end if%></td>
		<td nowrap  dir="<%=dir_obj_var%>" align="right"><%=app("departmentName")%></td>
		
			</tr>
	<%		app.movenext
	j=j+1
	%>
	<%loop%>
<%set app = nothing	%>
<%Set con = Nothing%>	

</table>
	</form>



			 	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
</body>
</html>