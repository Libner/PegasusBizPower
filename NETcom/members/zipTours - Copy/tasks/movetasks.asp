<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%Response.CharSet = "windows-1255"

  'העברה גורפת של המשימות
  tasksId = trim(Request("tasksId"))
  
  SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
  COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
  UserID=trim(Request.Cookies("bizpegasus")("UserID"))
  OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))

  lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
  if lang_id = "1" then
    arr_Status = Array("","חדש","בטיפול","סגור")	
  else
    arr_Status = Array("","new","active","close")	
  end if
  If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
 End If
  If lang_id = 2 Then
		dir_var = "rtl"  : 	align_var = "left"  :  dir_obj_var = "ltr"  :  	self_name = "Self"
  Else
		dir_var = "ltr"  :  	align_var = "right"   :   dir_obj_var = "rtl"   :   self_name = "עצמי"
  End If		  
    
  task_sender_name = trim(Request.Cookies("bizpegasus")("UserName")) %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
	function closeWin()
	{
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	}
		
	function CheckFields(act)
	{
		if(document.getElementById("task_reciver_id"))
		{
				if(document.getElementById("task_reciver_id").value=='')
				{
 					window.alert("! נא לבחור את מקבל המשימה");     
					document.getElementById("task_reciver_id").focus();
					return false; 
				}
		}
	
		if (act=='submit')
		{ 
			document.forms[0].submit();             
		}	 
		
		return true;
	}
//-->
</script>  
</head>
<%If Request.Form("add") <> nil Then    
  	
    'הוספת משימות לטפסים מרובים
    tasksId = trim(Request.Form("tasksId")) 
   	arrTasks = Split(tasksId, ",")   		

	If trim(Request.Form("add")) = "1" And isArray(arrTasks) Then
		task_reciver_id = trim(Request.Form("task_reciver_id"))
		
		For aa=0 To Ubound(arrTasks)
				
			taskId	= cLng(arrTasks(aa))
	
			sqlstr="UPDATE tasks SET reciver_id = " & task_reciver_id & ", task_status = 1 WHERE Task_Id = " & taskId
			con.ExecuteQuery(sqlstr)				
				
			sqlstr = "EXEC dbo.get_task '" & task_reciver_id & "','" & OrgID & "','" & lang_id & "','" & taskId & "'"
			set rs_task = con.getRecordSet(sqlstr)
			if not rs_task.eof then
					task_content = trim(rs_task("task_content"))
					task_date = trim(rs_task("task_date"))
					task_open_date = trim(rs_task("task_open_date"))
					task_status = trim(rs_task("task_status"))	
					activityId = trim(rs_task("parent_id"))
					task_sender_name = trim(rs_task("sender_name"))
					task_reciver_name = trim(rs_task("reciver_name"))
					task_contact_name = trim(rs_task("contact_name"))
					task_company_name = trim(rs_task("company_name"))
					task_project_name = trim(rs_task("project_name"))
					task_replay = trim(rs_task("task_replay"))
					task_close_date = trim(rs_task("task_close_date"))
					task_sender_id = trim(rs_task("User_ID"))
					task_reciver_id = trim(rs_task("reciver_id"))
					task_contact_id = trim(rs_task("contact_id"))
					task_company_id = trim(rs_task("company_id"))
					task_appeal_id = trim(rs_task("appeal_id"))
					task_project_id = trim(rs_task("project_id"))
					parentID = trim(rs_task("parent_ID"))
					private_flag = trim(rs_task("private_flag"))
						
					If IsDate(task_date) Then
						task_date = Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)
					End If	
					
					If IsDate(task_open_date) Then
						task_open_date = FormatDateTime(task_open_date,2) & " " & FormatDateTime(task_open_date,4)
					End If		
					
					sqlstr = "Select EMAIL From USERS Where USER_ID = " & task_sender_id
					set rswrk = con.getRecordSet(sqlstr)
					If not rswrk.eof Then
						fromEmail = trim(rswrk("EMAIL"))
					End If
					set rswrk = Nothing													
							
					If task_reciver_id <> "" Then
						sqlstr = "Select EMAIL From USERS Where USER_ID = " & task_reciver_id
						set rswrk = con.getRecordSet(sqlstr)
						If not rswrk.eof Then
						toMail = trim(rswrk("EMAIL"))
						End If
						set rswrk = Nothing					
					End If			
						
					mail_recivers = ""
					sqlstr = "Select FirstName + Char(32) + LastName From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
					"Where Task_ID = " & taskID
					set rs_names = con.getRecordSet(sqlstr)
					if not rs_names.eof then
						mail_recivers = rs_names.getString(,,",",",")
					end if
					set rs_names = nothing
					If Len(mail_recivers) > 1 Then
						mail_recivers = Left(mail_recivers,Len(mail_recivers)-1)
					End If					
				
					If trim(task_appeal_id) <> "" Then
						sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & task_appeal_id & "'"
						set rs_app = con.getRecordSet(sqlstr)
						if not rs_app.eof then			
							productName = trim(rs_app("product_Name"))	
						end if
						set rs_app = nothing
					End If
					attachment = trim(rs_task("attachment"))
					attachment_closing = trim(rs_task("attachment_closing"))
						
					sqlstr = "Exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
					set rs_task_types = con.getRecordSet(sqlstr)
					If not rs_task_types.eof Then
						task_types_names = rs_task_types.getString(,,",",",")
					Else
						task_types_names = ""
					End If		
					
					If Len(task_types_names) > 0 Then
						task_types_names = Left(task_types_names,(Len(task_types_names)-1))
					End If
			
				end if
				set rs_task = Nothing	
				   
			If trim(lang_id) = "1" Then
					'<!-- start sending mail -->		
					strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
					"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
					"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=""#e6e6e6"">" & vbCrLf &_
					"<tr><td class=page_title style=""background-color:#FF0000"" dir=" & dir_obj_var & ">" & trim(Request.Cookies("bizpegasus")("TasksOne")) & "&nbsp;חדשה&nbsp;-&nbsp;BIZPOWER&nbsp;</td>" & vbCrLf &_
					"</tr><tr>" & vbCrLf &_
					"<td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
					"<table width=100% border=""0"" cellpadding=""0"" align=center cellspacing=""1"">" & vbCrLf &_
					"<tr><td colspan=4 height=""5"" nowrap></td></tr>" & vbCrLf &_
					"<tr><td align=""" & align_var & """ nowrap><span  class=""Form_R"" style=""width:75;"">" & task_open_date & "</span></td>" & vbCrLf &_
					"<td align=""" & align_var & """ width=80 nowrap><b>תאריך פתיחה</b></td></tr>" & vbCrLf &_			
					"<tr><td align=" & align_var & "><span style=""width:75;text-align:center"" class=""task_status_num" & task_status & """>" & vbCrLf &_
					arr_Status(task_status) & "</span></td><td align=" & align_var & " width=80 style=""padding-right:5px;padding-left:5px"" nowrap><b>סטטוס</b></td>" & vbCrLf &_
					"</tr><tr><td align=" & align_var & " dir=" & dir_obj_var & "><span  class=Form_R style=""width:75;text-align:right"">" & vbCrLf &_
					task_date & "</span></td><td align=" & align_var & " width=80 style=""padding-right:5px;padding-left:5px"" nowrap><b>תאריך יעד</b></td>" & vbCrLf &_
					"</tr><tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_sender_name & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>מאת</b></td></tr><tr>" & vbCrLf &_
					"<td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_reciver_name & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>אל</b></td></tr>" & vbCrLf &_
					"<tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=Form_R style=""width:280;"">" & mail_recivers & "</span></td>" & vbCrLf &_
					"<td align="& align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>מכותבים</b></td></tr>" & vbCrLf &_
					"<tr><td align=" & align_var & " dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & task_types_names & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>סוג " & trim(Request.Cookies("bizpegasus")("TasksOne")) & "</b></td>" & vbCrLf &_
					"</tr><tr><td align=" & align_var & " valign=top dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(task_content) & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap valign=top><b>תוכן</b></td></tr>" & vbCrLf &_
					"<tr><td height=5 colspan=2 nowrap></td></tr>"
					If not (IsNull(attachment) Or trim(attachment) = "") Then
					strBody = strBody & "<tr><td dir=" & dir_obj_var & " valign=top>" & vbCrLf &_
					"<a class=""file_link"" href=""" & strLocal & "download/tasks_attachments/" & attachment & """ target=_blank>" & attachment & "</a>" & vbCrLf &_
					"</td><td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap valign=top><b>מסמך</b></td></tr>" & vbCrLf
					End if
					strBody = strBody & "</table></td></tr>" & vbCrLf &_
					"<tr><td align=" & align_var & " bgcolor=""#C9C9C9""  style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
					"<table cellpadding=0 cellspacing=1 border=0 width=100% align=center><tr><td height=5 colspan=2 nowrap></td></tr>"
					If trim(task_company_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_company_name &	"</td>" & vbCrLf &_
					"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>קישור ל" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</b></td>" & vbCrLf &_
					"</tr>"
					End If
					If trim(task_contact_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_contact_name & "</td>" & vbCrLf &_
					"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>קישור ל" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</b></td>" & vbCrLf &_
					"</tr>"
					End If
					If trim(task_project_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_project_name & "</td>" & vbCrLf &_
				"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>קישור ל" &_
					trim(Request.Cookies("bizpegasus")("Projectone")) & "</b></td></tr>"
					End If
					If trim(task_appeal_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_appeal_id & " - " & productName & "</td>" & vbCrLf &_
				"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>קישור לטופס</b></td></tr>"
					End If
			Else
		  			'<!-- start sending mail -->		
					strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
					"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
					"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=""#e6e6e6"" dir=" & dir_var & ">" & vbCrLf &_
					"<tr><td class=page_title style=""background-color:#FF0000"" dir=" & dir_obj_var & ">New&nbsp;" & trim(Request.Cookies("bizpegasus")("TasksOne")) & "&nbsp;-&nbsp;BIZPOWER&nbsp;</td>" & vbCrLf &_
					"</tr><tr>" & vbCrLf &_
					"<td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
					"<table width=100% border=""0"" cellpadding=""0"" align=center cellspacing=""1"" dir=" & dir_var & ">" & vbCrLf &_
					"<tr><td colspan=4 height=""5"" nowrap></td></tr>" & vbCrLf &_
					"<tr><td align=""" & align_var & """ nowrap><span  class=""Form_R"" style=""width:75;"">" & task_open_date & "</span></td>" & vbCrLf &_
					"<td align=""" & align_var & """ width=80 nowrap><b>Date</b></td></tr>"	 & vbCrLf &_		
					"<tr><td align=" & align_var & "><span style=""width:75;text-align:center"" class=""task_status_num" & task_status & """>" & vbCrLf &_
					arr_Status(task_status) & "</span></td><td align=" & align_var & " width=80 style=""padding-right:5px;padding-left:5px"" nowrap><b>Status</b></td>" & vbCrLf &_
					"</tr><tr><td align=" & align_var & " dir=" & dir_obj_var & "><span  class=Form_R style=""width:75;text-align:right"">" & vbCrLf &_
					task_date & "</span></td><td align=" & align_var & " width=80 style=""padding-right:5px;padding-left:5px"" nowrap><b>Target date</b></td>" & vbCrLf &_
					"</tr><tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_sender_name & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>From</b></td></tr><tr>" & vbCrLf &_
					"<td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_reciver_name & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>To</b></td></tr>" & vbCrLf &_
					"<tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=Form_R style=""width:280;"">" & mail_recivers & "</span></td>" & vbCrLf &_
					"<td align="& align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>Addressees</b></td></tr>" & vbCrLf &_			
					"<tr><td align=" & align_var & " dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & task_types_names & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap><b>Type</b></td>" & vbCrLf &_
					"</tr><tr><td align=" & align_var & " valign=top dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(task_content) & "</span></td>" & vbCrLf &_
					"<td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap valign=top><b>Content</b></td></tr>" & vbCrLf &_
					"<tr><td height=5 colspan=2 nowrap></td></tr>"
					If not (IsNull(attachment) Or trim(attachment) = "") Then
					strBody = strBody & "<tr><td dir=" & dir_obj_var & " valign=top>" & vbCrLf &_
					"<a class=""file_link"" href=""" & strLocal & "download/tasks_attachments/" & attachment & """ target=_blank>" & attachment & "</a>" & vbCrLf &_
					"</td><td align=" & align_var & " style=""padding-right:5px;padding-left:5px"" nowrap valign=top><b>Document</b></td></tr>" & vbCrLf
					End if
					strBody = strBody & "</table></td></tr>" & vbCrLf &_
					"<tr><td align=" & align_var & " bgcolor=""#C9C9C9""  style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
					"<table cellpadding=0 cellspacing=1 border=0 width=100% align=center><tr><td height=5 colspan=2 nowrap></td></tr>"
					If trim(task_company_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_company_name &	"</td>" & vbCrLf &_
					"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>Link to " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</b></td>" & vbCrLf &_
					"</tr>"
					End If
					If trim(task_contact_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_contact_name & "</td>" & vbCrLf &_
					"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>Link to " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</b></td>" & vbCrLf &_
					"</tr>"
					End If
					If trim(task_project_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_project_name & "</td>" & vbCrLf &_
				"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>Link to " &_
					trim(Request.Cookies("bizpegasus")("Projectone")) & "</b></td></tr>"
					End If
					If trim(task_appeal_id) <> "" Then
					strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_appeal_id & " - " & productName & "</td>" & vbCrLf &_
				"<td align=" & align_var & " width=120 nowrap style=""padding-right:5px;padding-left:5px""><b>Link to form</b></td></tr>"
					End If

			End If
			'Response.Write strBody
			'Response.End	
			  
			If Len(toMail) > 0 And Len(fromEmail) > 0 Then
				If trim(lang_id) = "1" Then
				strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
				"</td></tr><tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/tasks/default.asp?T=IN&task_status=1&UserID="&task_reciver_id&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF;font-weight:600'>לחץ כאן לפרטים</a><br></td></tr>" & vbCrLf  &_
				"</table></body></html>"
				Else
				strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
				"</td></tr><tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/tasks/default.asp?T=IN&task_status=1&UserID="&task_reciver_id&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF;font-weight:600'>Click here for details</a><br></td></tr>" & vbCrLf  &_
				"</table></body></html>" 	  
				End If
				SendTask strBodyTo,toMail,fromEmail,0 		 
			End If			
		Next
			
		Sub SendTask(strBody,toMail,fromEmail,flag)
			Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
				Msg.From = fromEmail
				Msg.MimeFormatted = true
				Msg.MimeFormatted = true
				If flag = 0 Then 'מקבל המשימה
					If trim(lang_id) = "1" Then
						strSub = "" & trim(Request.Cookies("bizpegasus")("TasksOne")) & " חדשה - BIZPOWER"
					Else
						strSub = "New " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER"
					End If	
				Else
					If trim(lang_id) = "1" Then
						strSub = "" & trim(Request.Cookies("bizpegasus")("TasksOne")) & " חדשה - BIZPOWER - לידיעה"
					Else
						strSub = "New " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER - FYI"
					End If				
				End if
				Msg.Subject = strSub				
				Msg.To = toMail			
				Msg.HTMLBody = strBody			
				Msg.Send()						
			Set Msg = Nothing
		End Sub	
		End If		
			
		Set upl = Nothing %>
		<script language="javascript" type="text/javascript">
		<!--	
			opener.focus();
			opener.window.location.reload(true);
			self.close();
		//-->
			</script>
	<%	End If 			
			
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 28 Order By word_id"				
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
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  	  
%>
<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr>	     
	<td class="page_title"  dir="<%=dir_obj_var%>">&nbsp;העברה גורפת של משימות&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap dir="<%=dir_var%>">
<FORM name="Form1" ID="Form1" ACTION="movetasks.asp" METHOD="post" onSubmit="return CheckFields();">
<table border="0" cellpadding="1" cellspacing="0" width="90%" align=center dir="<%=dir_var%>">
<tr>
   <td align="<%=align_var%>" width=100% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
  	<tr><td colspan="2" height="20" nowrap></td></tr>
 	<tr><td align="<%=align_var%>" width="100%"> 
    <select name="task_reciver_id"  dir="<%=dir_obj_var%>" class="norm" style="width: 250px" ID="task_reciver_id">
    <option value="" ><%=arrTitles(21)%></option>
    <%set UserList=con.GetRecordSet("SELECT User_ID, UserName = CASE User_ID WHEN "&UserID&" THEN '"&self_name&"' ELSE FIRSTNAME + ' ' + LASTNAME END FROM Users WHERE ORGANIZATION_ID = " & OrgID & " AND ACTIVE = 1 ORDER BY Case When User_ID = "&UserID&" Then 0 Else 1 End, FIRSTNAME + ' ' + LASTNAME")
    do while not UserList.EOF
    selUserID=UserList(0)
    selUserName=UserList(1)%>
    <option value="<%=selUserID%>" <%if trim(task_reciver_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
     </select>
    </td>
    <td align="<%=align_var%>" width="80" nowrap><b><!--אל--><%=arrTitles(5)%></b></td>
    </tr>     
    <tr><td colspan="2" height="10" nowrap></td></tr>
	 <%If trim(tasksId) <> "" Then%>
	<tr>		
		<td align="<%=align_var%>" width="100%"><%=tasksId%><input type="hidden" id="tasksId" name="tasksId" value="<%=tasksId%>" ></td>
		<td align="<%=align_var%>" width="80" nowrap><b>משימות</b></td>  		
	</tr>
	<%End If%>      
	<tr><td height="5" colspan="2" nowrap><input type=hidden name="add" id="add" value="1"></td></tr>
</table></form></td></tr>
<tr><td align=center colspan="2" height="20" nowrap></td></tr>
<tr><td align=center colspan="2">
<table cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="center" width="50%"><input type="button" value="<%=arrButtons(2)%>" class="but_menu" style="width:100" onclick="closeWin();"></td>
<td align="center" width="50%"><A class="but_menu" href="javascript:void(0)" style="width:100px" onClick="return CheckFields('submit')">העבר משימות</a></td>
</tr>
</table>
</td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
</table>
</body>
</html>
<%set con=Nothing%>