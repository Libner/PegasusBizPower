
<!--#include file="../../connect.asp"-->

 <%PermisionAdmin=trim(Request.Cookies("bizpegasus")("PermisionAdmin"))%>
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->

<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
function CheckDel(str) {
	 <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את העובד"
     Else
		str_confirm = "Are you sure want to delete the employee ?"
     End If   
     %>		
	 return window.confirm("<%=str_confirm%>");	 
}
	function openPass()
	{
		h = 200;
		w = 500;
		S_Wind = window.open("OpenPass.aspx", "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	

function SendData()
{
form1.submit()
}
//-->
</script> 
<script>
function cball_onclick() {
	var strid = new String(document.form1.ids.value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.form1.elements['cbMail'+ arrid[i]].checked = document.form1.cb_all.checked ;
	
}
function cdall_onclick()
{
	var strid = new String(document.form1.idsCode.value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.form1.elements['cdCode'+ arrid[i]].checked = document.form1.cd_all.checked ;
	
}
</script>
</head> 

<% 'On Error Resume Next

	if trim(Request.Form("trappCode")) <> "" and trim(request.Form("change_code_flag"))="1" then
	trappCode=trim(Request.Form("trappCode"))
	'response.Write trappCode
	'response.end
	strCode=split(trappCode,",") 
	'response.Write ubound(strCode)
	for i=0 to ubound(strCode)
	'response.Write ubound(strCode)
	if IsNumeric( strCode(i)) then
		sqlstr = "UPDATE Users SET VerificationCode_Confirmed=1 WHERE User_id= "& strCode(i)
		'Response.Write pSqlString
		con.ExecuteQuery(sqlstr)
		end if
	'	 If conn.Errors.Count > 0 Then
	'	Response.Write pSqlString
     '   ExecuteQuery = False
        'Response.End 
   ' Else
      '  ExecuteQuery = True
    'End If
	'	response.Write sqlstr
	next
	'sqlstr = "UPDATE Users SET SendVerificationCode=null WHERE User_id IN (" & Request.Form("trappCode") & ")"
	'response.Write sqlstr
'	response.end
	'con.ExecuteQuery(sqlstr)
'response.end
Response.Redirect "default.asp"
	end if
	'response.end
%>

<% 

 workerId = trim(Request.Cookies("bizpegasus")("UserId"))
 PermisionAdmin=trim(Request.Cookies("bizpegasus")("PermisionAdmin"))
 'response.Write workerId &"<BR>"
PByDepartmentId=PermissionsByDepartmentId(workerId)
 'response.Write "PByDepartmentId="& PByDepartmentId

urlSort="default.asp?1=1"
arch=0
if request.QueryString ("arch")<>"" then
if IsNumeric(request.QueryString("arch")) then
     arch=request.QueryString("arch")
     end if
     else
     arch=0
 end if 
   %>
<%
appDepId=Request("app_DepId")
	'response.Write "AppCountryID="&AppCountryID 
	'dim arrayCountryID as Array
	if appDepId<>"" then
	arrayappDepId=Split(appDepId,",")
	else
	appDepId=0

	end if
 sort = Request.QueryString("sort")	
 if trim(sort)="" then  sort=1 end if  
 dim sortby(4)	
sortby(1) = "FIRSTNAME, LASTNAME"
sortby(2) = "FIRSTNAME desc, LASTNAME DESC"
sortby(3) = "birthday"
sortby(4) = "birthday DESC"

if Request.QueryString("delUserID")<>nil then					
		
			'--insert into changes table
			sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
			" SELECT 'חשבון משתמש', IsNULL(U.FIRSTNAME, '') + ' ' + IsNULL(LASTNAME, ''), [USER_ID], 'מחיקה', getDate(), " & UserID & _
			" FROM dbo.USERS U WHERE ([USER_ID] = " & User_ID & ")"
			con.executeQuery(sqlstr)	
		
			con.ExecuteQuery "DELETE from USERS where USER_ID=" & Request.QueryString("delUserID")
			Response.Redirect "default.asp"
	end if
%>
<%	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 38 Order By word_id"				
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
	  set rstitle = Nothing
%>



	<%    if trim(Request.Form("trapp")) <> "" And trim(Request.Form("change_status_flag")) = "1" then
	If isNumeric(Request.Form("cmb_change_status")) And trim(Request.Form("cmb_change_status")) <> "" Then
		cmb_change_status = cInt(Request.Form("cmb_change_status"))
	Else
		cmb_change_status = 1
	End If
	'response.Write "app_DepId="& appDepId &"<BR>"
		if  appDepId>0 then
		sqlDelete= "UPDATE Users SET User_STATUS=Null WHERE Department_Id IN (" & appDepId & ")"
		else
		
	sqlDelete= "UPDATE Users SET User_STATUS=Null"
	end if
	'response.Write sqlDelete
	
	'response.end
	con.ExecuteQuery(sqlDelete)
	
	sqlstr = "UPDATE Users SET User_STATUS=" & cmb_change_status & " WHERE User_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	con.ExecuteQuery(sqlstr)
	Response.Redirect "default.asp?arch=" & arch
end if%>
<%	if trim(Request.Form("trappEmail")) <> "" And trim(Request.Form("change_status_flagEmail")) = "1" then
	If isNumeric(Request.Form("change_status_flagEmail")) And trim(Request.Form("change_status_flagEmail")) <> "" Then
		cmb_chanchange_status_flagEmailge_status = cInt(Request.Form("change_status_flagEmail"))
	Else
		cmb_change_status = 1
	End If
	'''	sqlDelete= "UPDATE Users SET User_STATUS=Null"
	'''con.ExecuteQuery(sqlDelete)
	'''sqlstr = "UPDATE bar_users SET User_STATUS=" & cmb_change_status & " WHERE User_id IN (" & Request.Form("trappEmail") & ")"
	''''Response.Write(sqlstr)
	''''Response.End
	'''con.ExecuteQuery(sqlstr)
	'''Response.Redirect "default.asp?arch=" & arch
end if%>


<body>
<FORM action="default.asp?arch=<%=arch%>" method=POST id="form1" name="form1" target="_self">   

	<input type="hidden" name="trapp" value="" ID="trapp">
	<input type="hidden" name="trappEmail" value="" ID="trappEmail">
	<input type="hidden" name="trappCode" value="" ID="trappCode">
	<input type="hidden" name="change_status_flag" value="0" ID="change_status_flag">		
		<input type="hidden" name="change_code_flag" value="0" ID="change_code_flag">		
	<input type="hidden" name="change_status_flagEmail" value="0" ID="change_status_flagEmail">		

<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width="100%"><!--#include file="../../logo_top.asp"--></td></tr>
<%numOftab = 4%>
<%numOfLink = 5%>
<%topLevel2 = 36 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width="100%"><!--#include file="../../top_in.asp"--></td></tr>
<tr>
<td width="100%" class="page_title_left" dir=rtl align=left style=padding-left:20>&nbsp;&nbsp;<%if arch<>1 then%><a class=Link href="default.asp?arch=1">עובדים לא פעילים</a><%else%><a class=Link href="default.asp">עובדים פעילים</a><%end if%></td></tr>
<tr><td width="100%">
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
   <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1">
	<tr style="height:20px">
	<%if trim(Request.Cookies("bizpegasus")("DeleteVerificationCode"))=1 then%>
	<td width="7%" align="center" class="title_sort">תאריך התחלה <BR>קוד אימות<BR>אוטומטי</td>
	<td width="7%" align="center" class="title_sort">ביטול<br> אימות קוד<br><INPUT type="checkbox" LANGUAGE="javascript" onclick="return cdall_onclick()" title="" id="cd_all" name="cd_all"></td>
	<%end if%>
	<td width="7%" align="center" class="title_sort"><%if arch<>1 then%>חסום<%End If%> <%If 0 Then%><!--מחיקה--><%=arrTitles(3)%><%End If%></td>
	<td width="7%" align="center" class="title_sort"><!--עדכון--><%=arrTitles(4)%></td>
	<td width="7%" align="center" class="title_sort"><!--פעיל--><%=arrTitles(8)%></td>
	<td width="7%" align="center" class="title_sort" nowrap>דרגת חשיבות ידיעת המשוב</td>
	
	<td width="10%" align="<%=align_var%>" class="title_sort">שווי כספי של הזמנה</td>
	<td width="10%" align="<%=align_var%>" class="title_sort">יעד הזמנות מינימלי</td>
		<td width="10%" align="<%=align_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>">
		<%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="">
<%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=3"  title=""><%end if%>
	<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a>&nbsp;<span id="word5" name=word5>תאריך לידה</span>&nbsp;
	</td>	

	<td width="15%" align="<%=align_var%>" class="title_sort"><!--סוג עובד--><%=arrTitles(5)%></td>
	<td width="10%" align="<%=align_var%>" class="title_sort">שם המחלקה</td>
	<td width="15%" align="<%=align_var%>" nowrap class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>">
			<%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="">
<%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=1"  title=""><%end if%>
	<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a>&nbsp;<span id="Span1" name=word5><%=arrTitles(9)%></span>&nbsp;
	</td>
	<%if arch=0 then%>
	 <td align="center" class="title_sort" nowrap>הצג בריכוז <BR>משימות</td>
	<td align="center" class="title_sort"><%if trim(Request.Cookies("bizpegasus")("EmailGroupSend"))=1 then%>שליחת מייל<br>קבוצתי<INPUT type="checkbox" LANGUAGE="javascript" onclick="return cball_onclick()" title="" id="cb_all" name="cb_all"><%end if%></td>
	<%end if%>
	</tr>
<% 
 UpdateFieldUserBlocked = trim(Request.Cookies("bizpegasus")("UpdateFieldUserBlocked"))
 InsertNewUser = trim(Request.Cookies("bizpegasus")("InsertNewUser"))
 'response.Write "UpdateFieldUserBlocked="&UpdateFieldUserBlocked

if len(appDepId)>0 and appDepId<>"0" then
sqlQuery =" and U.Department_Id in("& appDepId &")"
	arrayDepId=Split(appDepId,",")
else
sqlQuery="0"
end if
'response.Write "11=" & PByDepartmentId &":"

if len(PByDepartmentId)>0 then
sqlStr = "EXEC dbo.users_site_listPByDepartmentId " & OrgId & "," &arch & ",'" & sqlQuery &"',' order by "& sortby(sort) &"' ,'"& PByDepartmentId  &"'"
'Response.Write sqlStr
'response.end

end if
if PermisionAdmin=1 then
sqlStr = "EXEC dbo.users_site_list " & OrgId & "," &arch & ",'" & sqlQuery &"',' order by "& sortby(sort) &"'"
end if
'where departmentId in ("& PByDepartmentId &") 
'Response.Write sqlStr
'response.end
 j=0
'Response.Write sqlStr
'response.end

set rs_USERS = con.GetRecordSet(sqlStr)
if not rs_USERS.EOF then
do while not rs_USERS.EOF
j=j+1

	USER_ID = rs_USERS("USER_ID")
	ids = ids & USER_ID 	
	
	user_name = rs_USERS(1)
	job_name = rs_USERS("job_name")
	departmentName=rs_USERS("departmentName")
	
	active = rs_USERS("ACTIVE")
	block=rs_Users("User_Bloked")
	Month_Min_Order = nFix(rs_USERS("Month_Min_Order"))
	Order_Price = nFix(rs_USERS("Order_Price"))
	count_tasks = nFix(rs_USERS("count_tasks"))
	count_appeals = nFix(rs_USERS("count_appeals"))
	user_status= rs_USERS("user_status")
	ImportanceId=rs_USERS("ImportanceId")
	birthday=rs_USERS("birthday")
	SendVerificationCode=rs_USERS("SendVerificationCode")
	VerificationCode_Confirmed=rs_USERS("VerificationCode_Confirmed") 
		if (IsDate(SendVerificationCode) and  VerificationCode_Confirmed=0) then
		idsCode=idsCode & USER_ID 	
		end if
select case ImportanceId
case "1"
ImportanceName="שוטף"
case "2"
ImportanceName="חשיבות נמוכה"
case "3"
ImportanceName="חשוב"
case "4"
ImportanceName="חשוב מאוד"
case else
ImportanceName=""
end select


	 %>
	<tr>
	<%if trim(Request.Cookies("bizpegasus")("DeleteVerificationCode"))=1 then%>
	<td dir="<%=dir_obj_var%>" align="center" class="card"><%if IsDate(SendVerificationCode) then%><%=FormatDateTime(SendVerificationCode,2)%><%end if%></td>
	<td dir="<%=dir_obj_var%>" align="center" class="card"><%if IsDate(SendVerificationCode) and  VerificationCode_Confirmed=0 then%><INPUT type="checkbox" name="cdCode" Id="cdCode<%=USER_ID%>" ><%end if%></td>
	<%end if%>
		<td dir="<%=dir_obj_var%>" align="center" class="card">
		<%If 0 Then%>
		<%If count_tasks = 0 And count_appeals = 0 Then%>
		<a href="default.asp?delUserID=<%=USER_ID%>" ONCLICK="return CheckDel()"><img src="../../images/delete_icon.gif" border="0" alt="מחיקה"></a>
		<%Else%>
		<%If trim(lang_id) = "1" Then
			str_alert = "שים לב, קיימת מידע במערכת עבור עובד זה\n\n" & Space(3) & "לפי כך לא ניתן למחוק את העובד ממערכת\n\n" & Space(4) & "אלא להעביר את העובד לסטטוס לא פעיל"
		Else
			str_alert = "Pay attention,\n\n you can\'t delete this employee \n\n however you can deactivate him"
		End If%>		
		<input type=image src="../../images/delete_icon.gif" border=0 Onclick="window.alert('<%=str_alert%>');return false;">
		<%End If%>
		<%End If%>	
		<%if arch<>1  and cint(UpdateFieldUserBlocked)=1 then%><a href="vsbBlock_worker.asp?idsite=<%=USER_ID%>"><%if block = "1" then%><img src="../../images/lamp_off.gif" alt="משתמש פעיל" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../../images/lamp_on.gif" alt="משתמש חסום" border="0" WIDTH="13" HEIGHT="18"><%end if%></a>
		<%end if%>
		</td>
			<td dir="<%=dir_obj_var%>" align="center" class="card"><a href="addWorker.asp?USER_ID=<%=USER_ID%>" target=_self><img src="../../images/edit_icon.gif" border="0"></a></td>
	    <td dir="<%=dir_obj_var%>" align="center" class="card"><a href="vsbPress_worker.asp?idsite=<%=USER_ID%>"><%if active = "0" then%><img src="../../images/lamp_off.gif" alt="<%=arrTitles(13)%>" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../../images/lamp_on.gif" alt="<%=arrTitles(14)%>" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>
	       <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=ImportanceName%>&nbsp;</td>
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=Order_Price%>&nbsp;</td>
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=Month_Min_Order%>&nbsp;</td>
  <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=birthday%>&nbsp;</td>	   	    
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=job_name%>&nbsp;</td>
	     <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=departmentName%>&nbsp;</td>
		<td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card" nowrap ><a href="addWorker.asp?USER_ID=<%=USER_ID%>" class="link_categ" target=_self>&nbsp;<strong><%=user_name%></strong>&nbsp;</a></td>
	<%if arch=0 then%>
		 <td align="center" class="card"><INPUT type="checkbox" id="cb<%=USER_ID%>" name="cb<%=USER_ID%>" <%if user_status=1 then%> checked <%end if%>></td>
		 <td align="center" class="card"><%if trim(Request.Cookies("bizpegasus")("EmailGroupSend"))=1 then%><%'if false then%><INPUT type="checkbox" id="cbMail<%=USER_ID%>" name="cbMail<%=USER_ID%>" ><%end if%></td>
<%if false then%>
		<td width="65" align="center" valign="middle" class="card">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table3">             
			<tr>	
				<td width="20" nowrap align="center" valign="middle"><a href="" onclick="return moveTo('<%=elmOrd%>','<%=i%>','<%=id%>')"><img src="../../../images/arrow_top_bot.gif" align="top" border=0 alt="העבר למקום הרצוי"></a></td>
			    <td width="25" nowrap align="right" valign="middle"><input type="text" name="toPlace-<%=j%>" value="" onKeyPress="return getNumbers(this)" class="Form" style="width:80%" ID="toPlace-<%=j%>"></td>		    			
			    <td width="20" nowrap align="right" class="td_admin_4"><font color="#060165"><B>&nbsp;<%=j%>&nbsp;</b></font></td>
			</tr>
		</table>		
	</td><%end if%>
	<%end if%>
	</tr>
<%
rs_USERS.movenext
if not rs_USERS.eof then
		ids = ids & ","
		end if
		if not rs_USERS.eof and (IsDate(SendVerificationCode) and  VerificationCode_Confirmed=0) then
		idsCode = idsCode & ","
		end if
loop%>
<input type="hidden" name="ids" value="<%=ids%>" ID="ids">
<input type="hidden" name="idsCode" value="<%=idsCode%>" ID="idsCode">


<%set rs_USERS = nothing
else
'Response.Redirect "addWorker.asp" %>
<tr><td colspan="11" class="title_sort1" align="center"><!--לא נמצאו עובדים--><%=arrTitles(10)%></td></tr>
<%end if%>
</table>
</td>
<td width=110 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width="100%">
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href='WorkersTmp.aspx' target=_blank>מסך הגדרות עובדים</a></td></tr>

<%IF InsertNewUser=1 THEN%><tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href='addWorker.asp'><!--הוספת עובד--><%=arrTitles(6)%></a></td></tr><%END IF%>
<tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href="javascript:void window.open('update_order_param.asp', 'winUpd' , 'scrollbars=1,toolbar=0,top=150,left=50,width=400,height=200,align=center,resizable=1');">יעד/עלות הזמנות</a></td></tr>
<tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href="excel_workers.asp" target=_blank><!--הצג דוח--><%=arrTitles(15)%></a></td></tr>
<%if arch=0 then%>
<tr><td colspan="2" align="right" bgcolor="#B2B2B2"><input type="button" value="הצג במשימות" class="button_edit_2"
 onclick="if (checkChangeStatus()) {document.form1.submit()} "></td></tr>
 <tr><td colspan="2" align="right" bgcolor="#B2B2B2"><%if trim(Request.Cookies("bizpegasus")("EmailGroupSend"))=1 then%><%'if false then%><input type="button" value="שלח מייל" class="button_edit_2"
 onclick="if (checkChangeStatusEMail()) {document.form1.submit()} " ID="Button1" NAME="Button1"><%end if%></td></tr>
 <%end if%>
 <%if trim(Request.Cookies("bizpegasus")("DeleteVerificationCode"))=1 then%>
  <tr><td colspan="2" align="right" bgcolor="#B2B2B2">
 <input type="button" value="בטל קוד אימות" class="button_edit_2" onclick="if (checkChangeCode()) {document.form1.submit()} " ID="Button1" NAME="Button1">
 </td></tr>
 <%end if%>
 
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
<%if false then%> <tr><td align="center" colspan=2><a class="button_edit_1" style="width: 105px;"   href="javascript:void window.open('update_IP.asp', 'winUpdIP' , 'scrollbars=1,toolbar=0,top=150,left=50,width=600,height=200,align=center,resizable=1');">IP-עדכון כתובת ה</a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr><%end if%>
<tr><td colspan=2>
	<select name="app_DepId" id="app_DepId" dir="rtl" class="norm" style="width: 120px;height:100px" multiple>
 <%if false then%>   <OPTION value="0" <%if trim(app_DepId)="0" or trim(app_DepId)="" then%>selected  <%end if%>>כל המחלקות</OPTION><%end if%>
<%'sqlstrDep = "SELECT departmentId, departmentName  FROM Departments   ORDER BY departmentName"
if PermisionAdmin=1  then 'workerid=893
sqlstrDep = "SELECT departmentId,departmentName FROM Departments ORDER BY departmentName "

else
sqlstrDep = "SELECT departmentId,departmentName FROM Departments where departmentId in ("& PByDepartmentId &") ORDER BY departmentName "
 end if
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
		<tr>
				<td colspan="2" height="10">&nbsp;</td>
			</tr>
		<tr><td colspan=2 align=center><%if workerId=893  then%><a class="button_edit_1" href="" onclick="return openPass()" target="_blank">סיסמה לחוזי עבודה</a><%end if%></td></tr>

</table>
</td></tr></table>
</td></tr>

<tr><td height="15"</td></tr>

</table>
</form>
</body>
</html>
<%set con = nothing%>
<script>
function checkChangeCode()
{
	
		var fl = 0;
		document.form1.trappCode.value = '';
		//alert(document.form1.idsCode.value)
		if(document.form1.idsCode)
		{
		var strid = new String(document.form1.idsCode.value);
		if(strid != "")
		{
		//alert("ddddd")
					var arrid = strid.split(',');
				//	alert(arrid.length)
					for (i=0;i<arrid.length;i++)
					{
						if (document.form1.elements['cdCode'+ arrid[i]].checked==true)
						{
						
							document.form1.trappCode.value = document.form1.trappCode.value + arrid[i] + ',';
							fl = 1;
						}	
					}
						<%
						If trim(lang_id) = "1" Then
							str_confirm = "? האם ברצונך לבטל קוד אימות להמשתמשים המסומנים"
						Else
							str_confirm = "Are you sure want to move the selected users to the selected status?"
						End If   
						%>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trappCode.value);
				//alert(txtnew.length)
				document.form1.trappCode.value = txtnew.substr(0,txtnew.length - 1);
				document.form1.action = "<%=urlSort%>";
				//window.alert(document.form1.action);
				window.document.all("change_code_flag").value = "1";
			//alert(document.form1.trappCode.value)
			document.form1.submit()
				return false;
			}
			else if (fl) return false;
		}
						<%
							If trim(lang_id) = "1" Then
								str_confirm = "! נא לסמן שם עובד "
							Else
								str_confirm = "Please select user !"
							End If   
						%>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}

function checkChangeStatusEMail()
	{
	
		var fl = 0;
		document.form1.trapp.value = '';
		
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
			//alert(document.form1.getElementsByTagName['cb'+ arrid[i]].value)
				if (document.form1.elements['cbMail'+ arrid[i]].checked==true)
				{	document.form1.trappEmail.value = document.form1.trappEmail.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "האם ברצונך לשלוח מייל  להמשתמשים המסומנים?"
			Else
				str_confirm = "Are you sure want to move the selected users to the selected status?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trappEmail.value);
				document.form1.trappEmail.value = txtnew.substr(0,txtnew.length - 1);
			//	alert(document.form1.trappEmail.value)
			//	document.form1.action = "../sendMail/SendMail.aspx";
			 window.open('', 'formpopup', 'width=1000,height=1000,resizeable,scrollbars');
         	 document.form1.target = 'formpopup';
   		     document.form1.action = "../sendMail/SendMail.aspx";

				//window.alert(document.form1.action);
				window.document.all("change_status_flagEmail").value = "1";
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "נא לסמן שם עובד! "
			Else
				str_confirm = "Please select user !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}
function checkChangeStatus()
	{
	
		var fl = 0;
		document.form1.trapp.value = '';
		
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
			//alert(document.form1.getElementsByTagName['cb'+ arrid[i]].value)
				if (document.form1.elements['cb'+ arrid[i]].checked==true)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך להעביר את המשתמשים המסומנים לסטאטוס הנבחר"
			Else
				str_confirm = "Are you sure want to move the selected users to the selected status?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				document.form1.action = "<%=urlSort%>";
				//window.alert(document.form1.action);
				window.document.all("change_status_flag").value = "1";
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן שם עובד "
			Else
				str_confirm = "Please select user !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}

</script>