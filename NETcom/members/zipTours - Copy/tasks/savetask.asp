<% @CodePage = 1255 %>
<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title></title>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<%Response.CodePage = 1255
	  Response.CharSet = "windows-1255"
	  Session.LCID = 1037
	  Session.CodePage = 1255
  
	str_mappath="../../../download/tasks_attachments"			

	Dim uploadsDirVar
	uploadsDirVar = Server.MapPath(str_mappath)
		
	dim objUpload
	'set objUpload = New clsUpload  
	'objUpload.Upload uploadsDirVar 'Load http form data		
		
	Set objUpload = Server.CreateObject("SoftArtisans.FileUp")	

	taskID = trim(objUpload.Form("taskid"))
	parentID = trim(objUpload.Form("parentid"))
	task_appeal_id = trim(objUpload.Form("appealid"))
	appealsId = trim(objUpload.Form("appealsId"))
	task_type_id = trim(objUpload.Form("task_type_id"))
	
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
	UserID=trim(Request.Cookies("bizpegasus")("UserID"))
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))

	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	arr_Status = Array("","חדש","בטיפול","סגור")	
	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If

	If lang_id = 2 Then
		dir_var = "rtl" : align_var = "left" :  dir_obj_var = "ltr"
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
	End If	
	
	task_sender_name = trim(Request.Cookies("bizpegasus")("UserName"))
  
   If objUpload.Form("add") <> nil Then    
		
   			taskID=trim(objUpload.Form("taskid")) 
		   	
   			If objUpload.Form("attachment") <> nil Then   		
   				NewFileName=trim(objUpload.Form("attachment"))    		
   			Else
   				NewFileName=""	
   			End if     	
   
   			If trim(objUpload.Form("deleteFile")) = "1" Then
   				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   				attachment_file = trim(objUpload.Form("attachment"))
		   			
   				file_path=str_mappath & attachment_file
   				'Response.Write fs.FileExists(server.mappath(file_path))
				'Response.End
				if fs.FileExists(server.mappath(file_path)) then
					set a = fs.GetFile(server.mappath(file_path))
					a.delete			
				end if			
				sqlstr = "Update tasks set attachment = NULL Where task_id = " & taskId
				'Response.Write sqlstr
				'Response.End	
				con.executeQuery(sqlstr)		
				Set fs = Nothing		
		   	
			ElseIf  trim(objUpload.Form("add")) = "1" And Len(objUpload.Form("attachment_file")) > 0 Then
			
				set objFile = objUpload.Form("attachment_file")
	  	
  				If objFile.TotalBytes > 0  Then		
			
					File_Name = objFile.UserFileName
					extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
					NewFileName =  "task_attachment_" & taskID
					new_name = NewFileName
					upload = true
					i = 0
					set fs = server.CreateObject("Scripting.FileSystemObject") 
					do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
						i =  i + 1
						new_name = NewFileName & "_" & i
					loop
					newFileName = new_name & "." & extend
					set fs = Nothing
				
					objFile.SaveAs Server.Mappath(str_mappath) & "/" & NewFileName
					'objFile.UserFileName = NewFileName
					'objFile.Save						
				End If	
		End If
		   
		If objUpload.Form("task_types") <> nil Then
				ind = InStr(objUpload.Form("task_types"),",")
				If ind = 1 Then
					task_types = Right(objUpload.Form("task_types"),Len(objUpload.Form("task_types"))-1)
				Else
					task_types = trim(objUpload.Form("task_types"))
				End if	
			Else
				task_types = ""
			End If
			    
			If trim(objUpload.Form("task_date")) <> "" Then    
				task_date = Day(trim(objUpload.Form("task_date"))) & "/" & Month(trim(objUpload.Form("task_date"))) & "/" & Year(trim(objUpload.Form("task_date")))
			End If
			
			If IsDate(trim(objUpload.Form("task_open_date"))) Then    
				task_open_date = trim(objUpload.Form("task_open_date"))	    
				task_open_date = FormatDateTime(task_open_date,2) & " " & FormatDateTime(task_open_date,4)
				'task_open_date = Day(task_open_date) & "/" & Month(task_open_date) & "/" & Year(task_open_date) & " " & Hour(task_open_date) & ":" & Minute(task_open_date)
			Else
				task_open_date = Now()	
			End If
			
			If objUpload.Form("activity_date") <> nil And IsDate(objUpload.Form("activity_date")) Then    
				activity_date = Day(trim(objUpload.Form("activity_date"))) & "/" & Month(trim(objUpload.Form("activity_date"))) & "/" & Year(trim(objUpload.Form("activity_date")))
			End If         
			
			task_content = trim(objUpload.Form("task_content"))	
			
			companyId=trim(objUpload.Form("companyid"))
			contactID=trim(objUpload.Form("contactid"))
			task_reciver_id=trim(objUpload.Form("task_reciver_id"))
			task_appeal_id=trim(objUpload.Form("task_appeal_id")) 
			parentID=trim(objUpload.Form("parentid"))    
			task_addressees=trim(objUpload.Form("task_addressees"))    
		        
			If trim(objUpload.Form("add")) = "1" Then
			If trim(companyID) = "" Or IsNull(companyID) Then
				companyID = "NULL"
			End If	
			If trim(contactID) = "" Or IsNull(contactID) Then
				contactID = "NULL"
			End If	
			If trim(task_appeal_id) = "" Or IsNull(task_appeal_id) Then
				task_appeal_id = "NULL"
			End If
			If trim(parentID) = "" Or IsNull(parentID) Then
				parentID = "NULL"
			End If			
			
			If trim(taskID) = "" Then	  
				sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into tasks (company_id,contact_id,User_ID,ORGANIZATION_ID,appeal_ID,parent_ID,task_date,task_open_date,task_content,task_types,task_status,reciver_id,attachment) "&_
				" values (" & companyID & "," & contactID & "," & UserID & "," & OrgID & "," & task_appeal_id & "," & parentID & ",'" &_
				task_date & "','" & task_open_date & "','" & sFix(task_content) & "','" & task_types & "','1'," & task_reciver_id & ",'" & sFix(NewFileName) & "'); SELECT @@IDENTITY AS NewID"		
				'Response.Write sqlstr
				'Response.End
				set rs_tmp = con.getRecordSet(sqlstr)
					taskId = rs_tmp.Fields("NewID").value
				set rs_tmp = Nothing
				
				'--insert into changes table
				sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
				" SELECT 'משימה', 'אל: ' + U.FIRSTNAME + ' איש קשר:'  + IsNULL(CONTACT_NAME, ''), task_id, 'הוספה', getDate(), " & UserID & _
				" FROM dbo.tasks T LEFT OUTER JOIN dbo.USERS U ON U.User_Id = T.reciver_id " & _
				" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = T.Contact_Id WHERE (Task_Id = "& taskId &")"
				con.executeQuery(sqlstr)								
				
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
								sqlstr="Insert Into task_to_users (task_id,user_id) Values (" & taskId & "," & arrUsers(count) & ")"	  	
								con.executeQuery(sqlstr)		
						End If				
						Next
					End If
				End If		
				
				xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_in.xml"		
				'----start adding to XML file ------
				set objDOM = Server.CreateObject("Microsoft.XMLDOM")
				objDom.async = false
				
				if objDOM.load(server.MapPath(xmlFilePath)) then
				Set objRoot = objDOM.documentElement
				
				Set objChield = objDOM.createElement("TASK")		
 				set objNewField=objRoot.appendChild(objChield) 
		 		  
 				Set objNewAttribute = objDOM.createAttribute("Reciver_ID")
				objNewAttribute.text = task_reciver_id		
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				  
				Set objNewAttribute = objDOM.createAttribute("Sender_Name")
				objNewAttribute.text = task_sender_name		
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				  
				Set objNewAttribute = objDOM.createAttribute("TASK_ID")
				objNewAttribute.text = taskId		
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				
				Set objNewAttribute = objDOM.createAttribute("TASK_TYPE")
				objNewAttribute.text = task_types_names		  		
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing
		 		  
				If IsDate(task_date) Then
				Set objNewAttribute = objDOM.createAttribute("TASK_DATE")
				objNewAttribute.text =  Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)		  		
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing
 				End If	
		 		  
				objDom.save server.MapPath(xmlFilePath)
						    
				set objNewField=nothing
				Set objChield = nothing
			end if ' objDOM.load	
			set objDOM = nothing
			'----end adding to XML file ------					
		    
			sqlstr = "EXECUTE dbo.get_task  '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
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
				task_replay = trim(rs_task("task_replay"))
				task_close_date = trim(rs_task("task_close_date"))
				task_sender_id = trim(rs_task("User_ID"))
				task_reciver_id = trim(rs_task("reciver_id"))
				task_contact_id = trim(rs_task("contact_id"))
				task_company_id = trim(rs_task("company_id"))
				task_appeal_id = trim(rs_task("appeal_id"))
				parentID = trim(rs_task("parent_ID"))
				private_flag = trim(rs_task("private_flag"))
				
				If IsDate(task_date) Then
					task_date = Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)
				End If	
				
				If IsDate(task_open_date) Then
					task_open_date = FormatDateTime(task_open_date,2) & " " & FormatDateTime(task_open_date,4)
				End If												
					
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
					
					'לאחר שליחת משימה ראשונה לתלונה סטטוס תלונה עובר לActive
					sqlstr = "Update Appeals Set appeal_status = '2' Where appeal_id = " & task_appeal_id &_
					" And appeal_status = '1'"
					con.ExecuteQuery(sqlstr)
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
		End If
		set rs_task = Nothing	
		set objUpload = Nothing 
		    
		'<!-- start sending mail -->		
	    Set xml = Server.CreateObject("Microsoft.XMLHTTP")
	    If trim(lang_id) = "1"
			FileToRead = strLocal & "netCom/members/tasks/mailtask.asp?OrgID=" & OrgID & "&TaskId=" & TaskId & "&LangId=" & lang_id & "&UserID=" & UserID
		Else
			FileToRead = strLocal & "netCom/members/tasks/mailtask_eng.asp?OrgID=" & OrgID & "&TaskId=" & TaskId & "&LangId=" & lang_id & "&UserID=" & UserID
		End If	
		'Response.Write FileToRead
		'Response.End 
		xml.Open "GET", FileToRead, False, "bgu4u", "u142u3"
		xml.SetRequestHeader "Content-type", "text/html; charset=windows-1255"
		xml.Send		
		strBody = xml.responseText       
		Set xml = Nothing

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
		SendTask strBodyTo,toMail,fromEmail,task_sender_name,0 		 
	  End If
	  
	  sqlstr = "Select DISTINCT Email From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
	  "Where Task_ID = " & taskId
	  set rs_emails = con.getRecordSet(sqlstr)
	  If not rs_emails.eof Then
	    strBodyCC = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
	    "</td></tr></table></body></html>" 
	  while not rs_emails.eof
			toMail = rs_emails(0)	
			If Len(toMail) > 0 And Len(fromEmail) > 0 Then
				SendTask strBodyCC,toMail,fromEmail,task_sender_name,1  
			End If			
			rs_emails.moveNext	
	  wend 			
	  End If
	  set rs_emails = Nothing			  	
	   
	  Sub SendTask(strBody,toMail,fromEmail,fromName,flag)
		Dim Msg
		Set Msg = Server.CreateObject("CDO.Message")
			Msg.BodyPart.Charset = "windows-1255"
			Msg.From = fromName & " <" & fromEmail & ">"
			Msg.MimeFormatted = true
			
			If trim(task_appeal_id) <> "" Then
				If flag = 0 Then 'מקבל המשימה
					strSub = task_appeal_id & " - " & productName & " - חדשה " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - מערכת פניות למדור אקדמיה"				
				Else
					strSub = task_appeal_id & " - " & productName & " - חדשה " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - מערכת פניות למדור אקדמיה - מכותב"				
				End if
			Else
				If flag = 0 Then 'מקבל המשימה
					strSub = "חדשה " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - מערכת פניות למדור אקדמיה"
				Else
					strSub = "חדשה " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - מערכת פניות למדור אקדמיה - מכותב"	
				End if
			End If
			Msg.Subject = strSub				
			Msg.To = toMail			
			Msg.HTMLBody = strBody			
			Msg.Send()						
		Set Msg = Nothing
	 End Sub						
		
	Else
		sqlstr = "SET DATEFORMAT DMY; UPDATE tasks set task_date = '" & task_date & "', task_content = '" & sFix(objUpload.Form("task_content")) &_
		"', task_types = '" & sFix(task_types) & "', attachment = '" & sFix(NewFileName) & "' Where task_id = " & taskId
		'Response.Write sqlstr
		'Response.End	
		con.executeQuery(sqlstr)			
		
		'--insert into changes table
		sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		" SELECT 'משימה', 'אל: ' + U.FIRSTNAME + ' איש קשר:'  + IsNULL(CONTACT_NAME, ''), task_id, 'עדכון', getDate(), " & UserID & _
		" FROM dbo.tasks T LEFT OUTER JOIN dbo.USERS U ON U.User_Id = T.reciver_id " & _
		" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = T.Contact_Id WHERE (Task_Id = "& taskId &")"
		con.executeQuery(sqlstr)		
		
	End If%>
	<script language="javascript" type="text/javascript">
	<!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	//-->
	</script>
  <%		
    End If    
  End If	  %>
<%set con=Nothing%>
</head>
<body></body>
</html>