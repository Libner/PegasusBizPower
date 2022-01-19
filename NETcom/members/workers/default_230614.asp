<!--#include file="../../connect.asp"-->
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
//-->
</script> 
</head> 
<% arch=0
if request.QueryString ("arch")<>"" then
if IsNumeric(request.QueryString("arch")) then
     arch=request.QueryString("arch")
     end if
     else
     arch=0
 end if 
   %>
<%
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
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 38 Order By word_id"				
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
	  set rstitle = Nothing  %>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width="100%"><!--#include file="../../logo_top.asp"--></td></tr>
<%numOftab = 4%>
<%numOfLink = 5%>
<tr><td width="100%"><!--#include file="../../top_in.asp"--></td></tr>
<tr>
<td width="100%" class="page_title_left" dir=rtl align=left style=padding-left:20>&nbsp;&nbsp;<%if arch<>1 then%><a class=Link href="default.asp?arch=1">עובדים לא פעילים</a><%else%><a class=Link href="default.asp">עובדים פעילים</a><%end if%></td></tr>
<tr><td width="100%">
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
   <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1">
	<tr style="height:20px">
	<td width="7%" align="center" class="title_sort"><%If 0 Then%><!--מחיקה--><%=arrTitles(3)%><%End If%></td>
	<td width="7%" align="center" class="title_sort"><!--עדכון--><%=arrTitles(4)%></td>
	<td width="7%" align="center" class="title_sort"><!--פעיל--><%=arrTitles(8)%></td>
	<td width="20%" align="<%=align_var%>" class="title_sort">שווי כספי של הזמנה</td>
	<td width="20%" align="<%=align_var%>" class="title_sort">יעד הזמנות מינימלי</td>
	<td width="20%" align="<%=align_var%>" class="title_sort"><!--סוג עובד--><%=arrTitles(5)%></td>
	<td width="20%" align="<%=align_var%>" class="title_sort"><!--שם עובד--><%=arrTitles(9)%></td>
	</tr>
<%
sqlStr = "EXEC dbo.users_site_list " & OrgId & "," &arch
'Response.Write sqlStr
'response.end
set rs_USERS = con.GetRecordSet(sqlStr)
if not rs_USERS.EOF then
do while not rs_USERS.EOF
	USER_ID = rs_USERS("USER_ID")
	user_name = rs_USERS(1)
	job_name = rs_USERS("job_name")
	active = rs_USERS("ACTIVE")
	Month_Min_Order = nFix(rs_USERS("Month_Min_Order"))
	Order_Price = nFix(rs_USERS("Order_Price"))
	count_tasks = nFix(rs_USERS("count_tasks"))
	count_appeals = nFix(rs_USERS("count_appeals")) %>
	<tr>
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
		</td>
		<td dir="<%=dir_obj_var%>" align="center" class="card"><a href="addWorker.asp?USER_ID=<%=USER_ID%>" target=_self><img src="../../images/edit_icon.gif" border="0"></a></td>
	    <td dir="<%=dir_obj_var%>" align="center" class="card"><a href="vsbPress_worker.asp?idsite=<%=USER_ID%>"><%if active = "0" then%><img src="../../images/lamp_off.gif" alt="<%=arrTitles(13)%>" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../../images/lamp_on.gif" alt="<%=arrTitles(14)%>" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=Order_Price%>&nbsp;</td>
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=Month_Min_Order%>&nbsp;</td>	   	    
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=job_name%>&nbsp;</td>
		<td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card"><a href="addWorker.asp?USER_ID=<%=USER_ID%>" class="link_categ" target=_self>&nbsp;<strong><%=user_name%></strong>&nbsp;</a></td>
	</tr>
<%
rs_USERS.movenext
loop
set rs_USERS = nothing
else
Response.Redirect "addWorker.asp" %>
<tr><td colspan="4" class="title_sort1" align="center"><!--לא נמצאו עובדים--><%=arrTitles(10)%></td></tr>
<%end if%>
</table>
</td>
<td width=110 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width="100%">
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href='addWorker.asp'><!--הוספת עובד--><%=arrTitles(6)%></a></td></tr>
<tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href="javascript:void window.open('update_order_param.asp', 'winUpd' , 'scrollbars=1,toolbar=0,top=150,left=50,width=400,height=200,align=center,resizable=1');">יעד/עלות הזמנות</a></td></tr>
<tr><td nowrap colspan="2" align="center"><a class="button_edit_1" style="width: 105px;"  href="excel_workers.asp" target=_blank><!--הצג דוח--><%=arrTitles(15)%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table>
</td></tr></table>
</td></tr>
<tr><td height="15"</td></tr>
</table>
</body>
</html>
<%set con = nothing%>