<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->

	messages_type_id=trim(Request("messages_type_id"))
	messageID=trim(Request("messageID"))
	'הוספת משימות לטפסים מרובים
  
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
		dir_var = "ltr"  :  	align_var = "right"  :  	dir_obj_var = "rtl"  :  self_name = "עצמי"
	End If		  
    
	task_sender_name = trim(Request.Cookies("bizpegasus")("UserName"))
    
   %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
//start of pop-ups related functions//
	

	function ContentOnchange(pId)
	{
	//alert(pId.value)
	//document.all("message_content").value="uuuuuu"
		if(window.XMLHttpRequest && !(window.ActiveXObject))
	 var xmlHTTP = new XMLHttpRequest();
    else
		if(window.ActiveXObject)
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	if (!xmlHTTP)
		var xmlHTTP = new ActiveXObject("Msxml2.XMLHTTP");
			if(pId.length > 0)
			{	 
				xmlHTTP.open("POST", 'GetType.asp', false);
				xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><typeId>' + pId.value + '</typeId></request>');	 
				
				result = new String(xmlHTTP.responseText);
	document.getElementById("message_content").value=result

		}	
	}
	


//end of pop-ups related functions//

function closeWin()
{
	opener.focus();
	opener.window.location.reload(true);
	self.close();
}	
function CheckFields(act)
{
if(window.document.all("task_addressees").value=='')
{
window.alert("בחר מתוך רשימה של משתמשים אשר שולחים הודעה")
window.document.all("task_addressees").focus()
}
   if(window.document.all("messages_type_id"))
   {
		if(window.document.all("messages_type_id").value=='')
		{
   			<%
				If trim(lang_id) = "1" Then
					str_alert = "! נא לבחור את נושא ההודעה"
				Else
					str_alert = "Please, choose the message subject !"
				End If	
			%>
 			window.alert("<%=str_alert%>");     
			window.document.all("messages_type_id").focus();
			return false; 
		}
   }
   /*
   if(window.document.all("contactID").value=='')
   {
      window.alert("! נא לבחור <%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>"); 
      window.document.all("contactID").focus();    
      return false; 
   }   
   */
  if (document.all("message_content").value=='')
   {
      <%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור תוכן " & trim(Request.Cookies("bizpegasus")("TasksOne"))
			Else
				str_alert = "Please, choose the " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " content !"
			End If	
	  %>
 	  window.alert("<%=str_alert%>");        
      document.all("message_content").focus();
      return false;
   }
 
   if (act=='submit')
   { 
       document.formtask.submit();             
   }	 
   return true;

}



//-->
</script>  
</head>
<%  
set upl=Server.CreateObject("SoftArtisans.FileUp")
If upl.Form("add") <> nil Then    
   	messageID=trim(upl.Form("messageID")) 
   	
  
	    
	
	message_content = trim(upl.Form("message_content"))	
	
    messages_type_id=trim(upl.Form("messages_type_id"))
    'הוספת משימות לטפסים מרובים
    
    task_addressees=trim(Upl.Form("task_addressees"))    
   
	If trim(upl.Form("add")) = "1" Then
if  Upl.Form("is_replay")="on" then
Replay_Flag="1"
else
Replay_Flag="0"
end if
	
	If trim(messageID) = "" Then	 
			sqlstr = "Select EMAIL From USERS Where USER_ID = " & UserID
		set rswrk = con.getRecordSet(sqlstr)
		If not rswrk.eof Then
		  fromEmail = trim(rswrk("EMAIL"))
		End If
		set rswrk = Nothing		
		If Len(trim(task_addressees)) > 0 Then 'לשלוח במייל משימה למותבים
		arrUsers = Split(task_addressees, ",")
			If IsArray(arrUsers) Then
				For count=0 To Ubound(arrUsers)
				   If IsNumeric(arrUsers(count)) Then
					sqlstr="SET DATEFORMAT DMY; Insert Into messages (User_ID,ORGANIZATION_ID,Message_date,message_content,message_types,message_status,reciver_id,Replay_Flag) "&_
		" values (" & UserID & "," & OrgID & ",'"  &_
		FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4) & "','"  & sFix(message_content) & "','" & trim(messages_type_id) & "','1',"& arrUsers(count) &",'"& Replay_Flag &"')"		
		'Response.Write sqlstr &"<BR>"
		'Response.End
		con.executeQuery(sqlstr)
		'	messageID = rs_tmp.Fields("NewID").value
		'set rs_tmp = Nothing
					   End If				
				Next
				
			End If
		End If		
	
		

		
	
	'		response.write "bbb111"
	' response.end

	
	
					
	
		  	
		    
			'sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
			'set rs_task = con.getRecordSet(sqlstr)
			'if not rs_task.eof then
			'	message_content = trim(rs_task("message_content"))
			'	task_date = trim(rs_task("task_date"))
			'	task_status = trim(rs_task("task_status"))	
			'	activityId = trim(rs_task("parent_id"))
			'	task_sender_name = trim(rs_task("sender_name"))
			'	task_reciver_name = trim(rs_task("reciver_name"))
			'		task_replay = trim(rs_task("task_replay"))
			'	task_sender_id = trim(rs_task("User_ID"))
			'	messages_type_id = trim(rs_task("reciver_id"))
			'	private_flag = trim(rs_task("private_flag"))
			'	
			'	If IsDate(task_date) Then
			'		task_date = Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)
			'	End If	
				
													
					
			'	If false then' messages_type_id <> "" Then
			'		sqlstr = "Select EMAIL From USERS Where USER_ID = " & messages_type_id
			'		set rswrk = con.getRecordSet(sqlstr)
			'		If not rswrk.eof Then
			'		toMail = trim(rswrk("EMAIL"))
			'		End If
			'		set rswrk = Nothing					
			'	End If			
			'	
			'	mail_recivers = ""
			'	sqlstr = "Select FirstName + Char(32) + LastName From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
			'	"Where Task_ID = " & taskID
			'	set rs_names = con.getRecordSet(sqlstr)
			'	if not rs_names.eof then
			'		mail_recivers = rs_names.getString(,,",",",")
			'	end if
			'	set rs_names = nothing
			'	If Len(mail_recivers) > 1 Then
			'		mail_recivers = Left(mail_recivers,Len(mail_recivers)-1)
			'	End If					
		'
		'			
		'		sqlstr = "Exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
		'		set rs_task_types = con.getRecordSet(sqlstr)
		'		If not rs_task_types.eof Then
		'			task_types_names = rs_task_types.getString(,,",",",")
		'		Else
		'			task_types_names = ""
		'		End If		
		'	
		'		If Len(task_types_names) > 0 Then
		'			task_types_names = Left(task_types_names,(Len(task_types_names)-1))
		'		End If
	'
	'		end if
	'		set rs_task = Nothing	
		   
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
			"</tr><tr><td align=" & align_var & " valign=top dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(message_content) & "</span></td>" & vbCrLf &_
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
			"</tr><tr><td align=" & align_var & " valign=top dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(message_content) & "</span></td>" & vbCrLf &_
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
		
	       
	   
	     
	  End If
	  'Response.Write strBody
	  'Response.End	
	  
	  If Len(toMail) > 0 And Len(fromEmail) > 0 Then
	    If trim(lang_id) = "1" Then
        strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
        "</td></tr><tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/tasks/default.asp?T=IN&task_status=1&UserID="&messages_type_id&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF;font-weight:600'>לחץ כאן לפרטים</a><br></td></tr>" & vbCrLf  &_
        "</table></body></html>"
        Else
	    strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
	    "</td></tr><tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/tasks/default.asp?T=IN&task_status=1&UserID="&messages_type_id&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF;font-weight:600'>Click here for details</a><br></td></tr>" & vbCrLf  &_
        "</table></body></html>" 	  
        End If
		SendTask strBodyTo,toMail,fromEmail,0 		 
	  End If

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
			'Msg.To = "faina@cyberserve.co.il"	
			Msg.HTMLBody = strBody			
			'Msg.Send()						
		Set Msg = Nothing
	 End Sub						
		
	Else
		sqlstr = "SET DATEFORMAT DMY; Update tasks set task_date = '" & task_date & "', message_content = '" & sFix(upl.Form("message_content")) &_
		"', task_types = '" & sFix(task_types) & "', attachment = '" & sFix(File_Name) & "' Where task_id = " & taskId
		'Response.Write sqlstr
		'Response.End	
		con.executeQuery(sqlstr)	
		
		'--insert into changes table
		sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		" SELECT 'משימה', 'אל: ' + U.FIRSTNAME + ' איש קשר:'  + IsNULL(CONTACT_NAME, ''), task_id, 'עדכון', getDate(), " & UserID & _
		" FROM dbo.tasks T LEFT OUTER JOIN dbo.USERS U ON U.User_Id = T.reciver_id " & _
		" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = T.Contact_Id WHERE (Task_Id = "& taskId &")"
		con.executeQuery(sqlstr)
		
		If trim(messages_type_id) <> "" And IsNumeric(messages_type_id) Then
			sqlstr = "UPDATE tasks SET reciver_id = '" & sFix(messages_type_id) & "', task_status = 1 WHERE task_id = " & taskId
			'Response.Write sqlstr
			'Response.End	
			con.executeQuery(sqlstr)			
		End If		
		
	End If 	 %>
	<SCRIPT LANGUAGE=javascript>
	<!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	//-->
	</SCRIPT>
  <%		
    End If 
    set upl = Nothing 
   Response.End
  End If	  
%>
<%
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
<%task_open_date = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)

	


	If trim(taskID) <> "" Or trim(parentID) <> "" Then
	If trim(taskID) <> "" Then
		sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
    ElseIf trim(parentID) <> "" Then
		sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & parentID & "'"
    End If
    
    set rs_task = con.getRecordSet(sqlstr)
	if not rs_task.eof then		
		message_content = trim(rs_task("message_content"))
		If trim(taskID) <> "" Then
		task_date = trim(rs_task("task_date"))
		If IsDate(task_date) Then
			task_date = Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)
		End If
		Else
		task_date = Date()
		End If
		If trim(taskID) <> "" Then
			task_status = trim(rs_task("task_status"))	
		Else
			task_status = 1 
		End If
		If trim(taskID) <> "" Then	
		task_sender_name = trim(rs_task("sender_name"))
		Else
		task_sender_name = task_sender_name
		End If		
		task_reciver_name = trim(rs_task("reciver_name"))
		If trim(taskID) <> "" Then				
		task_sender_id = trim(rs_task("User_ID"))
		Else
		task_sender_id = UserID
		End if
		messages_type_id = trim(rs_task("reciver_id"))		
		private_flag = trim(rs_task("private_flag"))
		task_replay = trim(rs_task("task_replay"))
		If trim(parentID) <> "" And trim(task_replay) <> "" Then
			message_content = "תוכן משימה : " & message_content & vbCrLf & "תוכן סגירה : " & task_replay
		End If
		'------------------------------ מכותבים -----------------------------------------------------------------
		task_addressees = ""
		If trim(taskID) <> "" Then
			sqlstr = "Select task_to_users.User_ID From task_to_users Where Task_ID = " & taskID & " ORDER BY task_to_users.User_ID"
		ElseIf trim(parentID) <> "" Then
			sqlstr = "Select task_to_users.User_ID From task_to_users Where Task_ID = " & parentID & " ORDER BY task_to_users.User_ID"
		End If
		set rs_addresses = con.getRecordSet(sqlstr)
		if not rs_addresses.eof then
			task_addressees = rs_addresses.getString(,,",",",")
		end if
		set rs_addresses = nothing
		If Len(task_addressees) > 1 Then
			task_addressees = Left(task_addressees,Len(task_addressees)-1)
		End If
		
		mail_recivers = ""
		If trim(taskID) <> "" Then
		sqlstr = "Select FirstName + Char(32) + LastName From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
		"Where Task_ID = " & taskID
		ElseIf trim(parentID) <> "" Then
		sqlstr = "Select FirstName + Char(32) + LastName From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
		"Where Task_ID = " & parentID
		End If		
		set rs_names = con.getRecordSet(sqlstr)
		if not rs_names.eof then
			mail_recivers = rs_names.getString(,,",",",")
		end if
		set rs_names = nothing
		If Len(mail_recivers) > 1 Then
			mail_recivers = Left(mail_recivers,Len(mail_recivers)-1)
		End If				
			
		'------------------------------ מכותבים -----------------------------------------------------------------
					
		If trim(task_status) = "1" And messages_type_id = trim(UserID) And taskID <> "" Then 'משימה נפתחה פעם ראשונה
			sqlstr = "Update tasks SET task_status = '2' WHERE task_id = " & taskID
			con.executeQuery(sqlstr)
			xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_in.xml"
			'------ start deleting the new message from XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false			
			if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.getElementsByTagName("TASK")
				for j=0 to objNodes.length-1
					set objTask = objNodes.item(j)
					node_task_id = objTask.attributes.getNamedItem("TASK_ID").text
					node_user_id = objTask.attributes.getNamedItem("Reciver_ID").text
					if trim(taskId) = trim(node_task_id) And trim(UserID) = trim(node_user_id) Then					
						objDOM.documentElement.removeChild(objTask)
						exit for
					else
						set objTask = nothing
					end if
				next
				Set objNodes = nothing
				set objTask = nothing
				objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		   ' ------ end  deleting the new message from XML file ------
		End If
		
		If trim(taskID) <> "" Then		
			
			mail_recivers = ""
			sqlstr = "Select UserName = CASE task_to_users.User_ID WHEN "&UserID&" THEN '"&self_name&"' ELSE FIRSTNAME + ' ' + LASTNAME END From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
			"Where Task_ID = " & taskID & " ORDER BY Case When task_to_users.User_ID = "&UserID&" Then 0 Else 1 End, FIRSTNAME + ' ' + LASTNAME"
			set rs_names = con.getRecordSet(sqlstr)
			if not rs_names.eof then
				mail_recivers = rs_names.getString(,,",",",")
			end if
			set rs_names = nothing
			If Len(mail_recivers) > 1 Then
				mail_recivers = Left(mail_recivers,Len(mail_recivers)-1)
			End If				
		End If
	
		'Response.Write "cc" &  trim(rs_task("attachment"))
		If trim(taskID) <> "" Then
		sqlstr = "Exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
		Else
		sqlstr = "Exec dbo.get_task_types '"&parentID&"','"&OrgID&"'"
		End If
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
  Else
	  task_status = 1
  End If
%>
<tr>	     
	<td class="page_title"  dir="<%=dir_obj_var%>">&nbsp;<%If trim(taskID)="" Then%><!--הוספת--><%=arrTitles(9)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%><%Else%><!--עדכון--><%=arrTitles(10)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%><%End If%>&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap dir="<%=dir_var%>">
<FORM name="formtask" ACTION="addMessage.asp" METHOD="post" onSubmit="return CheckFields();" enctype="multipart/form-data" ID="formtask">
<table border="0" cellpadding="1" cellspacing="0" width="100%" align="center" dir="<%=dir_var%>">
<tr>
   <td align="<%=align_var%>" width="100%" nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
   <input type=hidden name="taskID" id="taskID" value="<%=taskID%>">
   <input type=hidden name="add" id="add" value="1">
   <table border="0" cellpadding="0" align="center" cellspacing="5" width="100%">        
   <tr><td colspan=2>
   <table cellpadding=0 cellspacing=0 width="100%" border=0 dir="<%=dir_obj_var%>">     
   <tr>
   <td align="<%=align_var%>" colspan=4>
   <table cellpadding=1 cellspacing=1 border=0 width="100%" dir="<%=dir_var%>" >
   	 <tr>           
            <td align="<%=align_var%>" colspan=2 nowrap><span style="width:70;text-align:center" class="task_status_num<%=task_status%>">
        <%=arr_Status(task_status)%></span>
        </td>
        <td align="<%=align_var%>" width=80 nowrap><b><!--סטטוס--><%=arrTitles(2)%></b></td>
   </tr>
   <tr>
   
   <td align="right" nowrap colspan=2><%=Date()%></td>
   <td align="<%=align_var%>" width=80 nowrap><b>תאריך</b></td>   
   </tr>   
    <tr>
	<td align="<%=align_var%>" colspan=2 nowrap  dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;height:12px">&nbsp;<%=task_sender_name%>&nbsp;</span></td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--מאת--><%=arrTitles(4)%></b></td>
	</tr>
   <tr><td align="<%=align_var%>" colspan=2> 
    <select name="messages_type_id"  dir="<%=dir_obj_var%>" class="norm" style="width:250" ID="messages_type_id" onchange='javascript:ContentOnchange(this);'>
    <option value="" id=word21 name=word21>בחר נושא ההודעה </option>
    <%set messagesList=con.GetRecordSet("SELECT messages_type_id,messages_type_name,messages_type_Content from messages_types where ORGANIZATION_ID = " & OrgID & " order  by messages_type_name")
    do while not messagesList.EOF
    messagesTypeId=messagesList(0)
    messagesTypeName=messagesList(1)
   %>
    <option value="<%=messagesTypeId%>" ><%=messagesTypeName%></option>
    <%
    messagesList.MoveNext
    loop
    set messagesList=Nothing%>
     </select>
    </td>
    <td align="<%=align_var%>" width=80 nowrap><b>נושא ההודעה</b></td>
    </tr>    
   
    <tr><td align="<%=align_var%>" colspan=2> 
      <INPUT type=image src="../../images/icon_find.gif" name=users_list id="users_list" onclick='window.open("users_list.asp?taskID=<%=taskID%>&task_addressees=" + task_addressees.value,"UsersList","left=300,top=20,width=250,height=500,scrollbars=1"); return false;'>&nbsp;
     <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px" name="users_names" id="users_names"><%=mail_recivers%></span>
     <input type=hidden name=task_addressees id=task_addressees value="<%=task_addressees%>">
    </td>
    <td align="<%=align_var%>" valign=top width=80 nowrap><b>שלח אל</b></td>
    </tr>
  
    </table></td></tr> 
     
 <tr><td valign="top" align="<%=align_var%>"><b><!--תוכן--><%=arrTitles(7)%></b></td></tr>
 <tr>             
	<td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" nowrap>
	<textarea id="message_content" class="Form" name="message_content"  dir="<%=dir_obj_var%>" style="width:330" rows=5><%=message_content%></textarea>
	</td>	
</tr>

</table></td>
</tr>
<tr>
	<td align="<%=align_var%>" class="title_show_form" nowrap width="200" dir="<%=dir_obj_var%>">רשאי לכתוב הודעה כתגובה&nbsp;</td>
	<td align="right"><input type="checkbox" dir="ltr" name="is_replay" id="is_replay" checked ></td>

</tr>
          
</table></form></td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
<tr><td align="center" colspan="2">
<table cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="center" width="50%"><input type="button" value="<%=arrButtons(2)%>" class="but_menu" style="width:100" onclick="closeWin();"></td>
<td align="center" width="50%"><A class="but_menu" href="javascript:void(0)" style="width:100px" onClick="return CheckFields('submit')">&nbsp;שלח הודעה</a></td>
</tr>
</table>
</td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
</table>
</div>
</body>
</html>
<%set con=Nothing%>