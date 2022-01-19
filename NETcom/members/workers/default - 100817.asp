<%@ Language=VBScript%>


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
	  set rstitle = Nothing
	  
	    %>
	<%    if trim(Request.Form("trapp")) <> "" And trim(Request.Form("change_status_flag")) = "1" then
	If isNumeric(Request.Form("cmb_change_status")) And trim(Request.Form("cmb_change_status")) <> "" Then
		cmb_change_status = cInt(Request.Form("cmb_change_status"))
	Else
		cmb_change_status = 1
	End If	
	sqlDelete= "UPDATE Users SET User_STATUS=Null"
	con.ExecuteQuery(sqlDelete)
	sqlstr = "UPDATE Users SET User_STATUS=" & cmb_change_status & " WHERE User_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	con.ExecuteQuery(sqlstr)
	Response.Redirect "default.asp?arch=" & arch
end if%>
<body>
<FORM action="default.asp?arch=<%=arch%>" method=POST id="form1" name="form1" target="_self">   

	<input type="hidden" name="trapp" value="" ID="trapp">
	<input type="hidden" name="change_status_flag" value="0" ID="change_status_flag">		

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
	<td width="7%" align="center" class="title_sort" nowrap>דרגת חשיבות ידיעת המשוב</td>
	
	<td width="20%" align="<%=align_var%>" class="title_sort">שווי כספי של הזמנה</td>
	<td width="20%" align="<%=align_var%>" class="title_sort">יעד הזמנות מינימלי</td>
	<td width="20%" align="<%=align_var%>" class="title_sort"><!--סוג עובד--><%=arrTitles(5)%></td>
	<td width="20%" align="<%=align_var%>" class="title_sort"><!--שם עובד--><%=arrTitles(9)%></td>
	<%if arch=0 then%>
	 <td align="center" class="title_sort" nowrap>הצג בריכוז <BR>משימות</td>
	<%if false then%><td align="center" class="title_sort">&nbsp;</td><%end if%>
	<%end if%>
	</tr>
<%
sqlStr = "EXEC dbo.users_site_list " & OrgId & "," &arch
'Response.Write sqlStr
'response.end
 j=0
set rs_USERS = con.GetRecordSet(sqlStr)
if not rs_USERS.EOF then
do while not rs_USERS.EOF
j=j+1

	USER_ID = rs_USERS("USER_ID")
	ids = ids & USER_ID 		
	user_name = rs_USERS(1)
	job_name = rs_USERS("job_name")
	active = rs_USERS("ACTIVE")
	Month_Min_Order = nFix(rs_USERS("Month_Min_Order"))
	Order_Price = nFix(rs_USERS("Order_Price"))
	count_tasks = nFix(rs_USERS("count_tasks"))
	count_appeals = nFix(rs_USERS("count_appeals"))
	user_status= rs_USERS("user_status")
	ImportanceId=rs_USERS("ImportanceId")
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

	'response.Write user_status
	 %>
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
	       <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=ImportanceName%>&nbsp;</td>
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=Order_Price%>&nbsp;</td>
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=Month_Min_Order%>&nbsp;</td>	   	    
	    <td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card">&nbsp;<%=job_name%>&nbsp;</td>
		<td dir="<%=dir_obj_var%>" align="<%=align_var%>" class="card"><a href="addWorker.asp?USER_ID=<%=USER_ID%>" class="link_categ" target=_self>&nbsp;<strong><%=user_name%></strong>&nbsp;</a></td>
	<%if arch=0 then%>
		 <td align="center" class="card"><INPUT type="checkbox" id="cb<%=USER_ID%>" name="cb<%=USER_ID%>" <%if user_status=1 then%> checked <%end if%>></td>
<%if false then%>
		<td width="65" align="center" valign="middle" class="card">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table3">             
			<tr>	
				<td width="20" nowrap align="center" valign="middle"><a href="" onclick="return moveTo('<%=elmOrd%>','<%=i%>','<%=id%>')"><img src="../../../images/arrow_top_bot.gif" align="top" border=0 alt="העבר למקום הרצוי"></a></td>
			    <td width="25" nowrap align="right" valign="middle"><input type="text" name="toPlace-<%=j%>" value="" onKeyPress="return getNumbers(this)" class="Form" style="width:80%" ID="Text1"></td>		    			
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
loop%>
<input type="hidden" name="ids" value="<%=ids%>" ID="ids">


<%set rs_USERS = nothing
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
<%if arch=0 then%>
<tr><td colspan="2" align="right" bgcolor="#B2B2B2"><input type="button" value="הצג במשימות" class="button_edit_2"
 onclick="if (checkChangeStatus()) {document.form1.submit()} "></td></tr><%end if%>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
<%if false then%> <tr><td align="center" colspan=2><a class="button_edit_1" style="width: 105px;"   href="javascript:void window.open('update_IP.asp', 'winUpdIP' , 'scrollbars=1,toolbar=0,top=150,left=50,width=600,height=200,align=center,resizable=1');">IP-עדכון כתובת ה</a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr><%end if%>
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
				if (document.forms('form1').elements('cb'+ arrid[i]).checked)
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