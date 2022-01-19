<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%  
    If Request.QueryString("OrgId") = nil Then
		OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	Else
		OrgId = trim(Request.QueryString("OrgId"))
	End If	
	UserId = trim(Request.Cookies("bizpegasus")("UserId"))
    taskID = trim(Request("taskID"))
    task_addressees = trim(Request("task_addressees"))
    lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))	 
	If lang_id = "2" Then
			dir_var = "rtl"
			align_var = "left"
			dir_obj_var = "ltr"
			self_name = "Self"
	Else
			dir_var = "ltr"
			align_var = "right"
			dir_obj_var = "rtl"
			self_name = "עצמי"
	End If           
   	
	If trim(Request.QueryString("add")) <> nil Then		
		If Request.Form("user") <> nil Or Request.Form("all_users") <> nil Then			
			users_id = trim(Request.Form("user"))	
			sqlstr = "SELECT UserName = CASE User_ID WHEN "&UserID&" THEN '"&self_name&"' ELSE FIRSTNAME + ' ' + LASTNAME END "&_
			" FROM Users WHERE ORGANIZATION_ID = " & OrgID & " And User_ID IN (" & users_id & ") "&_
			" ORDER BY Case When User_ID = "&UserID&" Then 0 Else 1 End, FIRSTNAME + ' ' + LASTNAME"
			'Response.Write sqlstr
			'Response.End
			set rs_n = con.getRecordSet(sqlstr)
			If not rs_n.eof Then
				users_name = trim(rs_n.getString(,,",",","))
				users_name = Left(users_name, Len(users_name)-1)
			Else
				users_name = ""	
			End If
			set rs_n = nothing							
	End If
	%>
		<script language=javascript>
		<!--
			window.opener.document.all("task_addressees").value = "<%=users_id%>";
			window.opener.document.all("users_names").innerHTML = "<%=Replace(users_name,Chr(34),Chr(39)&Chr(39))%>";			
			window.close();
		//-->
		</script>

	<%
   End If
   
	sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	set rsbuttons = con.getRecordSet(sqlstr)
	If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	End If
	set rsbuttons=nothing	   				
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<SCRIPT LANGUAGE=javascript>
function check_all_users(objChk)
{
	input_arr = document.getElementsByName("user");	
	for(i=0;i<input_arr.length;i++)
	{
		//input_arr(i).disabled = objChk.checked;
		input_arr(i).checked = objChk.checked;
		if(objChk.checked == true)
			window.document.all("tr"+input_arr(i).value).style.background = "#B1B0CF";
		else
			window.document.all("tr"+input_arr(i).value).style.background = "#E6E6E6";
	}
	return true;
}
function selectUser(objChk,userID)
{
	if(objChk.checked == true)
		window.document.all("tr"+userID).style.background = "#B1B0CF";
	else
		window.document.all("tr"+userID).style.background = "#E6E6E6";
}
</SCRIPT>
</head>
<body style="margin:0px;background-color:#e6e6e6" onload="window.focus()">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF ID="Table1">
<tr>
<td width="100%" class="page_title" dir=rtl>&nbsp;
<%If trim(lang_id) = "1" Then%>
רשימת עובדים
<%Else%>
Employees list
<%End If%>
&nbsp;</td></tr>         
<tr><td width="100%">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" valign="top" align="right">    
<%'//start of contact_user%>
	<table border=0 width=100% cellpadding=0 cellspacing=0 bgcolor="#e6e6e6">								
		<tr><td align=right valign=top>
		<form name="form1" id="form1" action="users_list.asp?add=1" method=post>		
			<table border=0 width=100% cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">
			<%
				sqlStr = "SELECT User_ID, UserName = CASE User_ID WHEN "&UserID&" THEN '"&self_name&"' ELSE FIRSTNAME + ' ' + LASTNAME END FROM Users WHERE ORGANIZATION_ID = " & OrgID & " AND ACTIVE = 1 ORDER BY Case When User_ID = "&UserID&" Then 0 Else 1 End, FIRSTNAME + ' ' + LASTNAME"
				set rs_users = con.GetRecordSet(sqlStr)
				If not rs_users.eof Then												
			%>
				<tr>								
					<td width=100% style="line-height:18px" align="<%=align_var%>" class="title_sort">
					<%If trim(lang_id) = "1" Then%>
					כל העובדים
					<%Else%>
					All employees
					<%End if%>
					&nbsp;</td>
					<td width=25 class="title_sort" align=center nowrap><INPUT type=checkbox name="all_users" id="all_users" onclick="return check_all_users(this)"></td>
				</tr>					
			<%
				do while not rs_users.eof
				userID = trim(rs_users(0))				
				userName = trim(rs_users(1))						
				If trim(taskID) <> "" And trim(userID) <> "" Then					
				sqlstr = "Select * From task_to_users WHERE User_ID = " & userID &_				
				" AND task_id = " & taskID				
				set rs_chk = con.getRecordSet(sqlstr)
				If not rs_chk.eof Then
					checked = "checked"
				Else
					checked = ""					
				End If	
				End If
				If trim(task_addressees) <> "" Then
				sqlstr = "Select * From Users WHERE User_id = " & userID &_
				" AND User_id IN (" & task_addressees & ")"
				'Response.Write sqlstr
				'Response.End
				set rs_chk = con.getRecordSet(sqlstr)
				If not rs_chk.eof Then
					checked = "checked"
				Else
					checked = ""
				End If	
				End If					
			%>		
			<tr id="tr<%=userID%>" name="tr<%=userID%>" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align="<%=align_var%>" dir=rtl>&nbsp;<%=userName%>&nbsp;</td>
				<td align=center><INPUT type=checkbox name="user" <%=checked%> id="<%=userID%>" value="<%=userID%>" onclick=" selectUser(this,<%=userID%>)"></td>
			</tr>					
			<%
		rs_users.movenext
		loop
		set rs_users = nothing
		If Len(ids) > 1 Then
			ids = Left(ids, Len(ids)-1)
		End If
		If Len(names) > 1 Then
			names = Left(names, Len(names)-1)
		End If
		End If
		%>					
		</table>						
	  </td></tr>
	<tr><td colspan=2 height="15" nowrap>
	<input type=hidden name="all_ids" id="all_ids" value="<%=ids%>">
	<input type=hidden name="all_names" id="all_names" value="<%=vFix(names)%>">
	</td></tr>
	<tr>
		<td colspan=2 align="center" nowrap dir="<%=dir_var%>">
			<input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button1" NAME="Button1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type=submit class="but_menu" value="<%=arrButtons(1)%>" style="width:80" ID="Button2" NAME="Button2">
		</td>		 
	</tr>
	<tr><td height=10 nowrap colspan=2></td></tr>
	</table></td></tr>
	</form>
	</table></td></tr></table>
	
</body>
<%set con=nothing%>
</html>
