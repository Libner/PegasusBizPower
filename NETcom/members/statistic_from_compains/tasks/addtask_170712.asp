<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%Response.CharSet = "windows-1255"

	task_company_id=trim(Request("companyId"))
	task_contact_id=trim(Request("contactID"))
	task_reciver_id=trim(Request("task_reciver_id"))
	task_project_id=trim(Request("project_id")) 
	taskID=trim(Request("taskID"))
	parentID=trim(Request("parentID"))
	task_appeal_id=trim(Request("appealID"))
	'הוספת משימות לטפסים מרובים
	appealsId=trim(Request("appealsId"))
  
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
    
	If trim(task_appeal_id) <> "" And  trim(taskID) = "" Then
		sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & task_appeal_id & "'"
		'Response.Write sqlstr
		'Response.End
		set rs_app = con.getRecordSet(sqlstr)
		if not rs_app.eof then
			task_appeal_id=trim(rs_app("appeal_id"))
			task_company_id=trim(rs_app("company_id"))
			task_contact_id=trim(rs_app("contact_id"))	
			task_project_id=trim(rs_app("project_id"))
			task_company_name = trim(rs_app("company_Name"))
			task_contact_name = trim(rs_app("contact_Name"))
			task_project_name = trim(rs_app("project_Name"))	
			productName = trim(rs_app("product_Name"))
			private_flag = trim(rs_app("private_flag"))			
		end if
		set rs_app = nothing
	End If   %>
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
	function openCompaniesList()
	{
		window.open('../appeals/companies_list.asp','winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function removeCompany()
	{
		window.document.getElementById("companyID").value = "";
		window.document.getElementById("CompanyName").value = "";
		
		window.document.getElementById("project_id").value = "";
		window.document.getElementById("projectName").value = "";
		
		window.document.getElementById("contactID").value = "";
		window.document.getElementById("contactName").value = "";		
		
		window.document.getElementById("contacter_body").style.display = "none";
		window.document.getElementById("project_body").style.display = "none";	 
		return false;
	}
	
	function openProjectsList()
	{
		var companyId = document.getElementById("companyID").value;
		window.open('../appeals/projects_list.asp?companyId=' + companyId,'winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function removeProject()
	{		
		window.document.getElementById("project_id").value = "";
		window.document.getElementById("projectName").value = "";
		
		return false;
	}	
	
	function openContactsList()
	{
		var companyId = document.getElementById("companyID").value;
		window.open('../appeals/contacts_list.asp?companyId=' + companyId,'winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function removeContact()
	{				
		window.document.getElementById("contactID").value = "";
		window.document.getElementById("contactName").value = "";		
		return false;
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
   if(window.document.all("task_reciver_id"))
   {
		if(window.document.all("task_reciver_id").value=='')
		{
   			<%
				If trim(lang_id) = "1" Then
					str_alert = "! נא לבחור את מקבל המשימה"
				Else
					str_alert = "Please, choose the task reciver !"
				End If	
			%>
 			window.alert("<%=str_alert%>");     
			window.document.all("task_reciver_id").focus();
			return false; 
		}
   }
   /*if(window.document.all("companyId").value=='')
   {
      window.alert("! נא לבחור לקוח");  
      window.document.all("companyId").focus();   
      return false; 
   }
   if(window.document.all("contactID").value=='')
   {
      window.alert("! נא לבחור <%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>"); 
      window.document.all("contactID").focus();    
      return false; 
   }   
   */
   if(window.document.all("task_types").value=='')
   {
      <%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור סוג " & trim(Request.Cookies("bizpegasus")("TasksOne"))
			Else
				str_alert = "Please, choose the " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " type !"
			End If	
	  %>
 	  window.alert("<%=str_alert%>");     
      return false; 
   }
   else if (document.all("task_content").value=='')
   {
      <%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור תוכן " & trim(Request.Cookies("bizpegasus")("TasksOne"))
			Else
				str_alert = "Please, choose the " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " content !"
			End If	
	  %>
 	  window.alert("<%=str_alert%>");        
      document.all("task_content").focus();
      return false;
   }
  if(document.formtask.attachment_file)
  {
   if (document.formtask.attachment_file.value !='')
	{
		var fname=new String();
		var fext=new String();
		var extfiles=new String();
		fname=document.formtask.attachment_file.value;
		indexOfDot = fname.lastIndexOf('.')
		fext=fname.slice(indexOfDot+1,-1)		
		fext=fext.toUpperCase();
		extfiles='HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT';		
		if ((extfiles.indexOf(fext)>-1) == false)
		{
		   <%
			If trim(lang_id) = "1" Then
				str_alert = ":סיומת הקובץ - אחת מרשימה \n HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT"
			Else
				str_alert = "The file ending should be one of the these \n HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT"
			End If	
	        %>	
			window.alert("<%=str_alert%>");
		    return false;
		}    
	  }   
   }
   if (act=='submit')
   { 
       document.formtask.submit();             
   }	 
   return true;

}

function DeleteAttachment(taskID)
{
    <%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק את הקובץ המצורף"
		Else
			str_confirm = "Are you sure want to delete the attached file?"
		End If	
	%>
	if (window.confirm("<%=str_confirm%>"))
	{
		window.document.all("add").value = "0";
		window.document.all("deleteFile").value = "1";
		window.formtask.submit();
	}
	return false;
}

function DoCal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}
//-->
</script>  
</head>
<%  
set upl=Server.CreateObject("SoftArtisans.FileUp")
If upl.Form("add") <> nil Then    
   	taskID=trim(upl.Form("taskID")) 
   	
   	If upl.Form("attachment") <> nil Then   		
   		File_Name=trim(upl.Form("attachment"))    		
   	Else
   		File_Name=""	
   	End if   	
   	
   
   	If trim(upl.Form("deleteFile")) = "1" Then
   		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		attachment_file = trim(upl.Form("attachment"))
   			
   		file_path="../../../download/tasks_attachments/" & attachment_file
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
   	
	ElseIf  trim(upl.UserFilename) <> "" And trim(upl.Form("add")) = "1" Then
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
   			
   		file_path="../../../download/tasks_attachments/" & File_Name
   		'Response.Write fs.FileExists(server.mappath(file_path))
        'Response.End
		if fs.FileExists(server.mappath(file_path)) then
			set a = fs.GetFile(server.mappath(file_path))
			a.delete			
		end if			
							
		upl.Form("attachment_file").SaveAs server.mappath("../../../download/tasks_attachments/") & "/" & File_Name
		if Err <> 0 Then
			 Response.Write("An error occurred when saving the file on the server.")			 
			 set upl = Nothing
			 Response.End
		end if			
	End If	
   
   If upl.Form("task_types") <> nil Then
		ind = InStr(upl.Form("task_types"),",")
		If ind = 1 Then
			task_types = Right(upl.Form("task_types"),Len(upl.Form("task_types"))-1)
		Else
			task_types = trim(upl.Form("task_types"))
		End if	
	Else
		task_types = ""
	End If
	    
	If trim(upl.Form("task_date")) <> "" Then    
		task_date = Day(trim(upl.Form("task_date"))) & "/" & Month(trim(upl.Form("task_date"))) & "/" & Year(trim(upl.Form("task_date")))
	End If
	
	If IsDate(trim(upl.Form("task_open_date"))) Then    
	    task_open_date = trim(upl.Form("task_open_date"))	    
		task_open_date = FormatDateTime(task_open_date,2) & " " & FormatDateTime(task_open_date,4)
		'task_open_date = Day(task_open_date) & "/" & Month(task_open_date) & "/" & Year(task_open_date) & " " & Hour(task_open_date) & ":" & Minute(task_open_date)
	Else
		task_open_date = Now()	
	End If
	
	If upl.Form("activity_date") <> nil And IsDate(upl.Form("activity_date")) Then    
		activity_date = Day(trim(upl.Form("activity_date"))) & "/" & Month(trim(upl.Form("activity_date"))) & "/" & Year(trim(upl.Form("activity_date")))
	End If         
	
	task_content = trim(upl.Form("task_content"))	
	
	companyId=trim(upl.Form("companyId"))
    contactID=trim(upl.Form("contactID"))
    task_reciver_id=trim(upl.Form("task_reciver_id"))
    task_project_id=trim(upl.Form("project_id"))
    task_appeal_id=trim(upl.Form("task_appeal_id")) 
    'הוספת משימות לטפסים מרובים
    appealsId = trim(upl.Form("appealsId")) 
    
    parentID=trim(upl.Form("parentID"))    
    task_addressees=trim(Upl.Form("task_addressees"))    
        
	If trim(upl.Form("add")) = "1" Then
	If trim(companyID) = "" Or IsNull(companyID) Then
		companyID = "NULL"
	End If	
	If trim(contactID) = "" Or IsNull(contactID) Then
		contactID = "NULL"
	End If	
	If trim(task_project_id) = "" Or IsNull(task_project_id) Then
		task_project_id = "NULL"
	End If	
	If trim(task_appeal_id) = "" Or IsNull(task_appeal_id) Then
		task_appeal_id = "NULL"
	End If
	If trim(parentID) = "" Or IsNull(parentID) Then
		parentID = "NULL"
	End If			
	
	If trim(taskID) = "" Then	  
		sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into tasks (company_id,contact_id,project_id,User_ID,ORGANIZATION_ID,appeal_ID,parent_ID,task_date,task_open_date,task_content,task_types,task_status,reciver_id,attachment) "&_
		" values (" & companyID & "," & contactID & "," & task_project_id & "," & UserID & "," & OrgID & "," & task_appeal_id & "," & parentID & ",'" &_
		task_date & "','" & task_open_date & "','" & sFix(task_content) & "','" & task_types & "','1'," & task_reciver_id & ",'" & sFix(File_Name) & "'); SELECT @@IDENTITY AS NewID"		
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
		    
			sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
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
	  
	  sqlstr = "Select DISTINCT Email From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
	  "Where Task_ID = " & taskId
	  set rs_emails = con.getRecordSet(sqlstr)
	  If not rs_emails.eof Then
	    strBodyCC = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
	    "</td></tr></table></body></html>" 
	  while not rs_emails.eof
			toMail = rs_emails(0)	
			If Len(toMail) > 0 And Len(fromEmail) > 0 Then
				SendTask strBodyCC,toMail,fromEmail,1  
			End If			
			rs_emails.moveNext	
	  wend 			
	  End If
	  set rs_emails = Nothing			  	
	   
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
		
	Else
		sqlstr = "SET DATEFORMAT DMY; Update tasks set task_date = '" & task_date & "', task_content = '" & sFix(upl.Form("task_content")) &_
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
		
		If trim(task_reciver_id) <> "" And IsNumeric(task_reciver_id) Then
			sqlstr = "UPDATE tasks SET reciver_id = '" & sFix(task_reciver_id) & "', task_status = 1 WHERE task_id = " & taskId
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

	If trim(task_company_id) <> "" Then
		sqlstr = "SELECT company_Name FROM companies WHERE company_Id = " & cLng(task_company_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			task_company_name = trim(rs_pr("company_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(task_contact_id) <> "" Then
		sqlstr = "SELECT contact_Name FROM contacts WHERE contact_Id = " & cLng(task_contact_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			task_contact_name = trim(rs_pr("contact_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(task_project_id) <> "" Then
		sqlstr = "SELECT project_Name FROM projects WHERE project_Id = " & cLng(task_project_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			task_project_name = trim(rs_pr("project_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(taskID) <> "" Or trim(parentID) <> "" Then
	If trim(taskID) <> "" Then
		sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
    ElseIf trim(parentID) <> "" Then
		sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & parentID & "'"
    End If
    
    set rs_task = con.getRecordSet(sqlstr)
	if not rs_task.eof then		
		task_content = trim(rs_task("task_content"))
		If trim(taskID) <> "" Then
		task_date = trim(rs_task("task_date"))
		If IsDate(task_date) Then
			task_date = Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)
		End If
		task_open_date = trim(rs_task("task_open_date"))
		If IsDate(task_open_date) Then
			task_open_date = FormatDateTime(task_open_date,2) & " " & FormatDateTime(task_open_date,4)
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
		task_contact_name = trim(rs_task("contact_name"))
		task_company_name = trim(rs_task("company_name"))
		task_project_name = trim(rs_task("project_name"))
		If trim(taskID) <> "" Then				
		task_sender_id = trim(rs_task("User_ID"))
		Else
		task_sender_id = UserID
		End if
		task_reciver_id = trim(rs_task("reciver_id"))		
		private_flag = trim(rs_task("private_flag"))
		task_replay = trim(rs_task("task_replay"))
		If trim(parentID) <> "" And trim(task_replay) <> "" Then
			task_content = "תוכן משימה : " & task_content & vbCrLf & "תוכן סגירה : " & task_replay
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
					
		If trim(task_status) = "1" And task_reciver_id = trim(UserID) And taskID <> "" Then 'משימה נפתחה פעם ראשונה
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
		task_contact_id = trim(rs_task("contact_id"))
		task_company_id = trim(rs_task("company_id"))
		task_project_id = trim(rs_task("project_id"))
		task_appeal_id = trim(rs_task("appeal_id"))
		If trim(taskID) <> "" Then
		parentID = trim(rs_task("parent_ID"))
		End If
		attachment = trim(rs_task("attachment"))
		childID = trim(rs_task("childID"))
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
		
		If trim(task_appeal_id) <> "" Then
			sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & task_appeal_id & "'"
			set rs_app = con.getRecordSet(sqlstr)
			if not rs_app.eof then			
				productName = trim(rs_app("product_Name"))	
			end if
			set rs_app = nothing
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
	  task_date = Day(date()) & "/" & Month(Date()) & "/" & Year(Date())	 
	  task_status = 1
  End If
%>
<tr>	     
	<td class="page_title"  dir="<%=dir_obj_var%>">&nbsp;<%If trim(taskID)="" Then%><!--הוספת--><%=arrTitles(9)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%><%Else%><!--עדכון--><%=arrTitles(10)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%><%End If%>&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap dir="<%=dir_var%>">
<FORM name="formtask" ACTION="addtask.asp" METHOD="post" onSubmit="return CheckFields();" enctype="multipart/form-data" ID="formtask">
<table border="0" cellpadding="1" cellspacing="0" width="100%" align="center" dir="<%=dir_var%>">
<tr>
   <td align="<%=align_var%>" width="100%" nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
   <input type=hidden name="parentID" id="parentID" value="<%=parentID%>">
   <input type=hidden name="taskID" id="taskID" value="<%=taskID%>">
   <input type=hidden name="add" id="add" value="1">
   <input type=hidden name="deleteFile" id="deleteFile" value="0">
   <input type=hidden name="task_open_date" id="task_open_date" value="<%=task_open_date%>">
   <table border="0" cellpadding="0" align="center" cellspacing="5" width="100%">        
   <tr><td colspan=2>
   <table cellpadding=0 cellspacing=0 width="100%" border=0 dir="<%=dir_obj_var%>">     
   <tr>
   <td align="<%=align_var%>" colspan=4>
   <table cellpadding=1 cellspacing=1 border=0 width="100%" dir="<%=dir_var%>" >
   <%If trim(taskID) <> "" Then%>
   <tr>
	<td align="<%=align_var%>" colspan=2 nowrap><span dir="ltr" class="Form_R" style="width:75;"><%=task_open_date%></span></td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--תאריך פתיחה--><%=arrTitles(1)%></b></td>
   </tr>
   <%end if%>
   	 <tr>           
        <td width="150" nowrap>
        <%If trim(taskID) <> "" And trim(childID) <> "" Then%>
         <A class="but_menu" href="#" style="width:150" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");'><%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<!--המשך--><%=arrTitles(26)%></a>
        <%Else%>&nbsp;<%End If%>
        </td>         
        <td align="<%=align_var%>" width=150 nowrap><span style="width:70;text-align:center" class="task_status_num<%=task_status%>">
        <%=arr_Status(task_status)%></span>
        </td>
        <td align="<%=align_var%>" width=80 nowrap><b><!--סטטוס--><%=arrTitles(2)%></b></td>
   </tr>
   <tr>
   <td width=150 nowrap>
   <%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
   <A class="but_menu" href="#" style="width:120" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");'><!--היסטורית--><%=arrTitles(27)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a>
   <%Else%>&nbsp;<%End If%>
   </td> 
   <td align="<%=align_var%>" width=150 nowrap><input type="text" style="width:80;" value="<%=task_date%>" id="task_date" class="Form"  onclick="return DoCal(this)" name="task_date"  dir="<%=dir_obj_var%>" ReadOnly></td>
   <td align="<%=align_var%>" width=80 nowrap><b><!--תאריך יעד--><%=arrTitles(3)%></b></td>   
   </tr>   
    <tr>
	<td align="<%=align_var%>" colspan=2 nowrap  dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;height:12px">&nbsp;<%=task_sender_name%>&nbsp;</span></td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--מאת--><%=arrTitles(4)%></b></td>
	</tr>
	<%If trim(taskID) = "" Or trim(task_status) = "1" Or trim(task_status) = "2" Then%>
	<tr><td align="<%=align_var%>" colspan=2> 
    <select name="task_reciver_id"  dir="<%=dir_obj_var%>" class="norm" style="width:250" ID="task_reciver_id">
    <option value="" id=word21 name=word21><%=arrTitles(21)%></option>
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
    <td align="<%=align_var%>" width=80 nowrap><b><!--אל--><%=arrTitles(5)%></b></td>
    </tr>    
    <tr><td align="<%=align_var%>" colspan=2> 
    <!--select name="task_addressees" id="task_addressees"  dir="<%=dir_obj_var%>" class="norm" style="width:250" size=4 multiple>
             <option value="" id=word21 name=word21><%=arrTitles(21)%></option>
             <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " ORDER BY FIRSTNAME + ' ' + LASTNAME")
             do while not UserList.EOF
                selUserID=UserList(0)
                selUserName=UserList(1)%>
                <option value="<%=selUserID%>" <%if trim(task_reciver_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
                <%
                UserList.MoveNext
                loop
                set UserList=Nothing%>
     </select-->
     <INPUT type=image src="../../images/icon_find.gif" name=users_list id="users_list" onclick='window.open("users_list.asp?taskID=<%=taskID%>&task_addressees=" + task_addressees.value,"UsersList","left=300,top=20,width=250,height=500,scrollbars=1"); return false;'>&nbsp;
     <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px" name="users_names" id="users_names"><%=mail_recivers%></span>
     <input type=hidden name=task_addressees id=task_addressees value="<%=task_addressees%>">
    </td>
    <td align="<%=align_var%>" valign=top width=80 nowrap><b><!--מכותבים--><%=arrTitles(28)%></b></td>
    </tr>
    <%Else%>    
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;"><%=task_reciver_name%></span></td>
	<td align="<%=align_var%>" valign=top width=80 nowrap><b><!--אל--><%=arrTitles(5)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:280;"><%=mail_recivers%></span></td>
	<td align="<%=align_var%>" valign=top width=80 nowrap><b><!--מכותבים--><%=arrTitles(28)%></b></td>
	</tr>    
    <%End If%>
    </table></td></tr> 
     
   <%
		 If trim(taskID) <> "" Then
		  	sqlstr="Select activity_type_name, activity_type_id, (Select CharIndex(Cast(activity_type_id as varchar) + ',', task_types + ',') from tasks Where task_id = "&taskID&" ) From activity_types WHERE ORGANIZATION_ID = " & trim(Request.Cookies("bizpegasus")("OrgID")) & " Order By activity_type_name"
		 ElseIf trim(parentID) <> "" Then
		    sqlstr="Select activity_type_name, activity_type_id, (Select CharIndex(Cast(activity_type_id as varchar) + ',', task_types + ',') from tasks Where task_id = "&parentID&" ) From activity_types WHERE ORGANIZATION_ID = " & trim(Request.Cookies("bizpegasus")("OrgID")) & " Order By activity_type_name"
		 Else
			sqlstr="Select activity_type_name, activity_type_id, 0 From activity_types WHERE ORGANIZATION_ID = " & trim(Request.Cookies("bizpegasus")("OrgID")) & " Order By activity_type_name"   
		 End If 	
		  	'Response.Write sqlstr
		  	'Response.End
		  	set rssub = con.getRecordSet(sqlstr)		 
		  	If not rssub.eof Then
		  	   arrSub = rssub.getRows()
		  	   recCount = rssub.RecordCount	
		  	   set rssub=Nothing
		  	   i=0			
		  	While i<=Ubound(arrSub,2)
		  		selSubgectId=trim(arrSub(1,i))
		  		selactivityName=trim(arrSub(0,i))
		  		is_selected=trim(arrSub(2,i))
		  %>
		 
		  <tr>
		  <%If i<=Ubound(arrSub,2) Then%>		 
		  <td width=15  align="center" nowrap><input type=checkbox id="task_types" name="task_types" <%If is_selected > 0 Then%> checked <%End If%> value="<%=selSubgectId%>"></td>		 
		  <td align="<%=align_var%>" width=170 nowrap>&nbsp;<%=selactivityName%>&nbsp;</td>
		  <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If	
		        i = i+1	
		        If i<=Ubound(arrSub,2) Then		  	    
		  	    selSubgectId=trim(arrSub(1,i))
		  		selactivityName=trim(arrSub(0,i))
		  		is_selected=trim(arrSub(2,i))		  		
		  %>	   
		   <td width=15  align="center" nowrap><input type=checkbox id="task_types" name="task_types" <%If is_selected > 0 Then%> checked <%End If%> value="<%=selSubgectId%>"></td>									  
		   <td align="<%=align_var%>" width=170 nowrap>&nbsp;<%=selactivityName%>&nbsp;</td>
		    <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If%>
		   </tr>	
		  <%  i = i+1		
		  	  Wend		  	
			  End If		
          %>         
      </table>        
 </td></tr>
 <tr><td valign="top" align="<%=align_var%>"><b><!--תוכן--><%=arrTitles(7)%></b></td></tr>
 <tr>             
	<td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" nowrap>
	<textarea id="task_content" class="Form" name="task_content"  dir="<%=dir_obj_var%>" style="width:330" rows=5><%=task_content%></textarea>
	</td>	
</tr>
<%If IsNull(attachment) Or trim(attachment) = "" Then%>
<tr><td  dir="<%=dir_obj_var%>"><!--מסמך--><%=arrTitles(8)%>&nbsp;&nbsp;&nbsp;&nbsp;<input type="file" name="attachment_file" ID="attachment_file" size=33 value=""></td></tr>
<%Else%>
<tr><td  dir="<%=dir_obj_var%>" valign=top>
<input type="hidden" name="attachment" ID="attachment" value="<%=vFix(attachment)%>">
<b><!--מסמך--><%=arrTitles(9)%></b>&nbsp;&nbsp;&nbsp;&nbsp;<a class="file_link" href="../../../download/tasks_attachments/<%=attachment%>" target=_blank><%=attachment%></a>
&nbsp;&nbsp;<input type="image" src="../../images/delete_icon.gif" style="position:relative;top:4px" onclick="return DeleteAttachment('<%=taskID%>')">
</td></tr>
<%End if%>
</table></td>
</tr>
<%If trim(task_appeal_id) = "" And trim(taskID) = "" Then%>
<tr>
   <td align="<%=align_var%>" colspan=4 bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none" style="padding:3px">
   <table cellpadding=2 cellspacing=1 border=0 width="100%" align="center">
   <tr  <%If trim(private_flag) <> "0" Then%> style="display: none;" <%End If%> >
	<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>">
		<input type="hidden" name="companyId" id="companyId" value="<%=task_company_id%>">
		<input type="text" id="CompanyName" name="CompanyName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(task_company_name) > 0 Then%><%=vFix(task_company_name)%><%Else%><%End If%>" >&nbsp;&nbsp;<input 
		type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openCompaniesList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="בטל בחירה" onclick="return removeCompany();" src="../../images/delete_icon.gif" align="absmiddle" >
       </td>
       <td align="<%=align_var%>" class="links_down" width=130 nowrap><span style="font-weight:600"><!--קישור ל--><%=arrTitles(14)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>
       </tr> 
      <tbody name="contacter_body" id="contacter_body" <%If trim(task_company_id) = ""  Or IsNull(task_company_id) Then%> style="display:none" <%End if%>>
      <tr><td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="contactID" id="contactID" value="<%=task_contact_id%>">
		<input type="text" id="contactName" name="contactName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(task_contact_name) > 0 Then%><%=vFix(task_contact_name)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input 
		type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openContactsList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="בטל בחירה" onclick="return removeContact();" src="../../images/delete_icon.gif" align="absmiddle" >
       </td>
       <td align="<%=align_var%>" class="links_down" width=130 nowrap><span style="font-weight:600"><!--קישור ל--><%=arrTitles(15)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>
       </tr>   
    </tbody>       
    <tbody name="project_body" id="project_body" <%If trim(task_company_id) = "" Or IsNull(task_company_id) Then%> style="display:none" <%End if%>>
    <TR>
		<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="project_id" id="project_id" value="<%=task_project_id%>">
		<input type="text" id="projectName" name="projectName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(task_project_name) > 0 Then%><%=vFix(task_project_name)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input 
		type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openProjectsList();" align="absmiddle">&nbsp;&nbsp;<input 
		type="image" alt="בטל בחירה" onclick="return removeProject();" src="../../images/delete_icon.gif" align="absmiddle">
		</TD>
		 <td align="<%=align_var%>" class="links_down" width=130 nowrap><span style="font-weight:600"><!--קישור ל--><%=arrTitles(16)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>
	</TR>
     </tbody>
    </table>
   </td>
  </tr> 
 <%Else%>
 <tr>
   <td align="<%=align_var%>" colspan="4" bgcolor="#C9C9C9"  style="border: 1px solid #808080; border-top:none;">
   <input type="hidden" name="task_appeal_id" id="task_appeal_id" value="<%=task_appeal_id%>">
   <input type="hidden" name="companyId" id="companyId" value="<%=task_company_id%>">
   <input type="hidden" name="contactId" id="contactId" value="<%=task_contact_id%>">
   <input type="hidden" name="project_id" id="project_id" value="<%=task_project_id%>">
   <table cellpadding="1" cellspacing="0" border="0" width="100%" align="center" >
   <tr><td height="5" colspan="2" nowrap></td></tr>   
   <%If trim(task_company_id) <> "" And trim(private_flag) = "0" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
		<A href="#" onclick="window.opener.location.href='../companies/company.asp?companyID=<%=task_company_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>
		<%=task_company_name%>
		<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(22)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(task_contact_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<%If trim(COMPANIES) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../companies/contact.asp?contactID=<%=task_contact_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>		
		<%=task_contact_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(23)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(task_project_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<%If trim(COMPANIES) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../projects/project.asp?companyID=<%=task_company_id%>&projectID=<%=task_project_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>
		<%=task_project_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(24)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>   
	</tr>
	<%End If%>
	 <%If trim(task_appeal_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<%If trim(SURVEYS)  = "1" And Request.QueryString("frommail") = nil Then%>
		<A href="#" onclick="window.opener.location.href='../appeals/appeal_card.asp?appid=<%=task_appeal_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End if%>
		<%=task_appeal_id & " - " & productName%>
		<%If trim(SURVEYS)  = "1" And Request.QueryString("frommail") = nil Then%>
		</a>
		<%End if%>
		</td>
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--קישור לטופס--><%=arrTitles(25)%></span></td>   
	</tr>
	<%End If%>		   	
	<tr><td height=5 colspan=2 nowrap></td></tr>
   </table>
   </td>
 </tr> 
 <%End If%>          
</table></form></td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
<tr><td align="center" colspan="2">
<table cellpadding="0" cellspacing="0" width="100%">
<tr>
<%If taskId <> "" And trim(task_reciver_id) = trim(UserID) Then%>
<td align="center" width="50%"><A class="but_menu" style="width:120px" href="addtask.asp?parentID=<%=taskID%>"><!--צור--><%=arrTitles(17)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%>&nbsp;<!--חדשה--><%=arrTitles(18)%></a></td>
<%Else%>
<td align="center" width="50%"><input type="button" value="<%=arrButtons(2)%>" class="but_menu" style="width:100" onclick="closeWin();"></td>
<%End If%>
<td align="center" width="50%"><A class="but_menu" href="javascript:void(0)" style="width:100px" onClick="return CheckFields('submit')"><%If taskId <> "" Then%><!--עדכן--><%=arrTitles(20)%><%Else%><!--הוסף--><%=arrTitles(19)%><%End If%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td>
</tr>
</table>
</td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
</table>
</div>
</body>
</html>
<%set con=Nothing%>