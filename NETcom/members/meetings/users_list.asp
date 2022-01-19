<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<% 	OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	UserId = trim(Request.Cookies("bizpegasus")("UserId")) 	
	
    meetingID = trim(Request("meetingID"))
    meeting_participants = trim(Request("meeting_participants"))
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
	
	'משתמש מורשה להוסיף פגישות לאחרים
	sqlstr = "Select IsNULL(Add_Meetings,0) From Users Where User_ID = " & UserID
	set rs_check = con.getRecordSet(sqlstr)
	if not rs_check.eof Then
		AddMeetings = trim(rs_check.Fields(0))
		participant_id = trim(Request("participant_id"))	
	else
		AddMeetings = 0
		participant_id = ""
	end if				           
   	
	If trim(Request.QueryString("add")) <> nil Then		
				
			users_id = trim(Request.Form("user"))	
			If trim(AddMeetings) = "0" Then
			    If Len(users_id) > 0 Then
					users_id = users_id & "," & participant_id
				Else
					users_id = participant_id
				End If
			End If		
		
			sqlstr = "SELECT FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " And User_ID IN (" & users_id & ") "&_
			" AND ACTIVE = 1 ORDER BY FIRSTNAME + ' ' + LASTNAME"
			'Response.Write sqlstr
			'Response.End
			set rs_n = con.getRecordSet(sqlstr)
			If not rs_n.eof Then
				users_name = trim(rs_n.getString(,,"<br>","<br>"))
				users_name = Left(users_name, Len(users_name)-1)
			Else
				users_name = ""	
			End If
			set rs_n = nothing							
			
	%>
		<script language=javascript>
		<!--
			window.opener.document.all("meeting_participants").value = "<%=users_id%>";
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
				If trim(AddMeetings) = "0" Then
				sqlStr = "SELECT User_ID, FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " And ACTIVE = 1 ORDER BY Case When User_ID = "&participant_id&" Then 0 Else 1 End, FIRSTNAME + ' ' + LASTNAME"
				Else
				sqlStr = "SELECT User_ID, FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " And ACTIVE = 1 ORDER BY FIRSTNAME + ' ' + LASTNAME"
				End If
				set rs_users = con.GetRecordSet(sqlStr)
				If not rs_users.eof Then												
			
				do while not rs_users.eof
				currUserID = trim(rs_users(0))				
				curruserName = trim(rs_users(1))						
				If trim(meetingID) <> "" And trim(currUserID) <> "" Then					
				sqlstr = "Select * From meeting_to_users WHERE User_ID = " & currUserID &_				
				" AND meeting_id = " & meetingID				
				set rs_chk = con.getRecordSet(sqlstr)
				If not rs_chk.eof Then
					checked = "checked"
				Else
					checked = ""					
				End If	
				End If
				If trim(meeting_participants) <> "" Then
				sqlstr = "Select * From Users WHERE User_id = " & currUserID &_
				" AND User_id IN (" & meeting_participants & ")"
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
			<tr id="tr<%=currUserID%>" name="tr<%=currUserID%>" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align="<%=align_var%>" dir=rtl>&nbsp;<%=curruserName%>&nbsp;</td>
				<td align=center><INPUT type=checkbox name="user" <%=checked%> id="<%=currUserID%>" value="<%=currUserID%>" onclick="selectUser(this,<%=currUserID%>)" <%If (trim(currUserID) = trim(participant_id)) And trim(AddMeetings) = "0" Then%> disabled <%End if%>></td>
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
	<input type=hidden name="participant_id" id="participant_id" value="<%=vFix(participant_id)%>">
	</td></tr>
	<tr>
		<td colspan=2 align="center" nowrap dir="<%=dir_var%>">
			<input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type=submit class="but_menu" value="<%=arrButtons(1)%>" style="width:80">
		</td>		 
	</tr>
	<tr><td height=10 nowrap colspan=2></td></tr>
	</table></td></tr>
	</form>
	</table></td></tr></table>
	
</body>
<%set con=nothing%>
</html>
