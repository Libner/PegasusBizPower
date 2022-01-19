<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%companyId=trim(Request("companyId"))
	  contactID=trim(Request("contactID"))
	  taskId=trim(Request("taskId"))  
	  UserID=trim(trim(Request.Cookies("bizpegasus")("UserID")))
	  OrgID=trim(trim(Request.Cookies("bizpegasus")("OrgID")))
	  SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	  COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
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
		dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr" : self_name = "Self"
	 Else
		dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl" : self_name = "עצמי"
	 End If		%>
<!-- #include file="../../title_meta_inc.asp" -->
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<script language="javascript" type="text/javascript">
function CheckFields(flag)
{  
   if(document.all("task_close_date").value=='' )
   {
	<%
		If trim(lang_id) = "1" Then
			str_alert = "! נא למלא תאריך סגירת המשימה"
		Else
			str_alert = "Please insert the closing date!"
		End If   
	%>   
		window.alert("<%=str_alert%>");
		return false;
   }   
   
   if(document.all("task_replay").value=='' )
   {
   	<%
		If trim(lang_id) = "1" Then
			str_alert = "! נא למלא תוכן סגירת המשימה"
		Else
			str_alert = "Please insert the closing content!"
		End If   
	%>   
		window.alert("<%=str_alert%>");	
		return false;
   }
    if (document.all("attachment_closing").value !='')
	{
		var fname=new String();
		var fext=new String();
		var extfiles=new String();
		fname=document.all("attachment_closing").value;
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
	if(flag == "2")
	{
		document.form_closing.action = "closetask.asp?parentID=<%=taskID%>";
	}   
    
   document.form_closing.submit();               
   return true;

}
function DeleteAttachmentClose(taskID)
{
   	<%
		If trim(lang_id) = "1" Then
			str_alert = "?האם ברצונך למחוק את הקובץ המצורף"
		Else
			str_alert = "Are you sure want to delete the document?"
		End If   
	%>   
	if (window.confirm("<%=str_alert%>"))
	{
		window.document.all("close").value = "0";
		window.document.all("deleteFile").value = "1";
		window.form_closing.submit();
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
function closeWin()
{
	opener.focus();
	opener.window.location.href = opener.window.location.href;
	self.close();
}
//-->
</script>  
</head>
<%  
  set upl=Server.CreateObject("SoftArtisans.FileUp")
  If upl.Form("close") <> nil Then 	    
		
	taskId = trim(upl.Form("taskId"))
	If trim(upl.Form("deleteFile")) = "1" Then
   		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		attachment_file = trim(upl.Form("attachment_closing"))
   			
   		file_path="../../../download/tasks_attachments/" & attachment_file
   		'Response.Write fs.FileExists(server.mappath(file_path))
        'Response.End
		if fs.FileExists(server.mappath(file_path)) then
			set a = fs.GetFile(server.mappath(file_path))
			a.delete			
		end if			
		sqlstr = "Update tasks set attachment_closing = NULL Where task_id = " & taskId
		'Response.Write sqlstr
		'Response.End	
		con.executeQuery(sqlstr)		
		Set fs = Nothing		
   	
	ElseIf trim(upl.UserFilename) <> "" Then		
		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
	   		
   		file_path="../../../download/tasks_attachments/" & File_Name
   		'Response.Write fs.FileExists(server.mappath(file_path))
		'Response.End
		if fs.FileExists(server.mappath(file_path)) then
			set a = fs.GetFile(server.mappath(file_path))
			a.delete			
		end if			
							
		upl.Form("attachment_closing").SaveAs server.mappath("../../../download/tasks_attachments/") & "/" & File_Name
		if Err.number <> 0 Then
			Response.Write("An error occurred when saving the file on the server.")			 
			set upl = Nothing
			Response.End
		end if	 
		sqlstr = "Update tasks set attachment_closing = '" & sFix(File_Name) & "' Where task_id = " & taskId
		con.executeQuery(sqlstr)

	End If			
		
		If upl.Form("task_close_date") <> nil And IsDate(upl.Form("task_close_date")) Then    
			task_close_date = Day(trim(upl.Form("task_close_date"))) & "/" & Month(trim(upl.Form("task_close_date"))) & "/" & Year(trim(upl.Form("task_close_date")))			
		End If  
		
		If taskId <> "" And trim(upl.Form("close")) = "1" Then
		    con.executeQuery("SET DATEFORMAT DMY")				 		
			sqlstr = "Update tasks set task_status = '3', " &_
			" task_replay = '" & sFix(upl.Form("task_replay")) & "'," &_
			" task_close_date = '" & task_close_date & "'" &_					
			" Where task_id = " & taskId
			'Response.Write sqlstr
			'Response.End	
			con.executeQuery(sqlstr)
			
				
		    If trim(taskId) <> "" Then
			sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
			set rs_task = con.getRecordSet(sqlstr)
			if not rs_task.eof then
				task_content = trim(rs_task("task_content"))
				task_date = trim(rs_task("task_date"))
				task_open_date = trim(rs_task("task_open_date"))
				task_types = trim(rs_task("task_types"))
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
				
				sqlstr = "Select EMAIL From USERS Where USER_ID = " & task_reciver_id
				set rswrk = con.getRecordSet(sqlstr)
				If not rswrk.eof Then
					fromEmail = trim(rswrk("EMAIL"))
				End If
				set rswrk = Nothing						
				
			end if
			set rs_task = Nothing	 	
			
		   xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_out.xml"		
		   '----start adding to XML file ------
		   set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		   objDom.async = false
		
		   if objDOM.load(server.MapPath(xmlFilePath)) then
			    Set objRoot = objDOM.documentElement		  
		        Set objChield = objDOM.createElement("TASK")		   
 		        set objNewField=objRoot.appendChild(objChield) 
 		
				Set objNewAttribute = objDOM.createAttribute("Sender_ID")
				objNewAttribute.text = task_sender_id		
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				
				Set objNewAttribute = objDOM.createAttribute("Reciver_Name")
				objNewAttribute.text = task_reciver_name		
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
 				
 				Set objNewAttribute = objDOM.createAttribute("TASK_DATE")
				objNewAttribute.text = task_date		  		
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing
		 		  
			'	objDom.save server.MapPath(xmlFilePath)
				    
		set objNewField=nothing
		Set objChield = nothing
	end if ' objDOM.load	
	set objDOM = nothing
	'----end adding to XML file ------	

     If trim(lang_id) = "1" Then
	        '<!-- start sending mail -->		
			strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
			"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=""#e6e6e6"" dir="&dir_var&">" & vbCrLf &_
			"<tr><td class=page_title style=""background-color:#00FF00"" dir=" & dir_obj_var & ">" & trim(Request.Cookies("bizpegasus")("TasksOne")) & " סגורה - BIZPOWER&nbsp;</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_
            "<td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
			"<table width=100% border=""0"" cellpadding=""0"" align=center cellspacing=""1"">" & vbCrLf &_
			"<tr><td colspan=4 height=""5"" nowrap></td></tr>" & vbCrLf &_
			"<tr><td align=""" & align_var & """><span  class=""Form_R"">" & task_open_date & "</span></td>" & vbCrLf &_
	        "<td align=""" & align_var & """ width=80 nowrap><b>תאריך פתיחה</b></td></tr>" & vbCrLf &_			
            "<tr><td align=" & align_var & "><span style=""width:75;text-align:center"" class=""task_status_num" & task_status & """>" & vbCrLf &_
            arr_Status(task_status) & "</span></td><td align=" & align_var & " width=90 style=""padding-right:10px;padding-left:10px"" nowrap><b>סטטוס</b></td>" & vbCrLf &_
            "</tr><tr><td align=" & align_var & " dir=" & dir_obj_var & "><span  class=Form_R style=""width:75;text-align:right"">" & vbCrLf &_
            task_date & "</span></td><td align=" & align_var & " width=90 style=""padding-right:10px;padding-left:10px"" nowrap><b>תאריך יעד</b></td>" & vbCrLf &_
	        "</tr><tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_sender_name & "</span></td>" & vbCrLf &_
	        "<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>מאת</b></td></tr><tr>" & vbCrLf &_
			"<td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_reciver_name & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>אל</b></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=Form_R style=""width:280;"">" & mail_recivers & "</span></td>" & vbCrLf &_
			"<td align="& align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>מכותבים</b></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & task_types_names & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>סוג " & trim(Request.Cookies("bizpegasus")("TasksOne")) & "</b></td>" & vbCrLf &_
			"</tr><tr><td align=" & align_var & " valign=top dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(task_content) & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap valign=top><b>תוכן</b></td></tr>" & vbCrLf &_
			"<tr><td height=5 colspan=2 nowrap></td></tr>"
			If not (IsNull(attachment) Or trim(attachment) = "") Then
	        strBody = strBody & "<tr><td dir=" & dir_obj_var & " valign=top>" & vbCrLf &_
	        "<a class=""file_link"" href=""" & strLocal & "download/tasks_attachments/" & attachment & """ target=_blank>" & attachment & "</a>" & vbCrLf &_
	        "</td><td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap valign=top><b>מסמך</b></td></tr>" & vbCrLf
			End if			
			strBody = strBody & "</table></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
	        "<table border=0 cellpadding=0 align=center cellspacing=1><tr><td height=5 colspan=2 nowrap></td></tr>" & vbCrLf &_
		    "<tr><td align=" & align_var & " width=100% ><span  class=Form_R style=""width:75;text-align:right"">" & task_close_date & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " width=90 style=""padding-right:10px;padding-left:10px"" nowrap><b>תאריך סגירה</b></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " bgcolor=""#e6e6e6"" valign=top nowrap><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(task_replay) & "</span></td>" & vbCrLf &_
		    "<td valign=top style=""padding-right:10px;padding-left:10px"" align=" & align_var & "><b>תוכן סגירה</b></td></tr>"
			If NOT (IsNull(attachment_closing) Or trim(attachment_closing) = "") Then
			strBody = strBody & "<tr><td valign=bottom  dir=" & dir_obj_var & ">" & vbCrLf &_
			"<a class=""file_link"" href=""" & strLocal & "download/tasks_attachments/" & attachment_closing & """ target=_blank>" & attachment_closing & "</a>" & vbCrLf &_
			"</td><td valign=bottom style=""padding-right:10px;padding-left:10px"" align=" & align_var & "><b>מסמך סגירה</b></td></tr>"
            End If
		    strBody = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf &_				
	        "<tr><td align=" & align_var & " bgcolor=""#C9C9C9""  style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
			"<table cellpadding=0 cellspacing=1 border=0 width=100% align=center><tr><td height=5 colspan=2 nowrap></td></tr>"
		    If trim(task_company_id) <> "" Then
			strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_company_name &	"</td>" & vbCrLf &_
		    "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>קישור ל" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</b></td>" & vbCrLf &_
			"</tr>"
	        End If
	        If trim(task_contact_id) <> "" Then
	        strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_contact_name & "</td>" & vbCrLf &_
		    "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>קישור ל" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</b></td>" & vbCrLf &_
	        "</tr>"
	        End If
	        If trim(task_project_id) <> "" Then
	        strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_project_name & "</td>" & vbCrLf &_
		   "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>קישור ל" &_
		    trim(Request.Cookies("bizpegasus")("Projectone")) & "</b></td></tr>"
	        End If
	        If trim(task_appeal_id) <> "" Then
	        strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_appeal_id & " - " & productName & "</td>" & vbCrLf &_
		   "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>קישור לטופס</b></td></tr>"
	        End If
        Else
		    '<!-- start sending mail -->		
			strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
			"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=""#e6e6e6"" dir="&dir_var&">" & vbCrLf &_
			"<tr><td class=page_title style=""background-color:#00FF00"" dir=" & dir_obj_var & ">Closed " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER&nbsp;</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_
            "<td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
			"<table width=100% border=""0"" cellpadding=""0"" align=center cellspacing=""1"">" & vbCrLf &_
			"<tr><td colspan=4 height=""5"" nowrap></td></tr>" & vbCrLf &_
			"<tr><td align=""" & align_var & """ colspan=2 nowrap><span  class=""Form_R"">" & task_open_date & "</span></td>" & vbCrLf &_
	        "<td align=""" & align_var & """ width=80 nowrap><b>Date</b></td></tr>"	 & vbCrLf &_
            "<tr><td align=" & align_var & "><span style=""width:75;text-align:center"" class=""task_status_num" & task_status & """>" & vbCrLf &_
            arr_Status(task_status) & "</span></td><td align=" & align_var & " width=90 style=""padding-right:10px;padding-left:10px"" nowrap><b>Status</b></td>" & vbCrLf &_
            "</tr><tr><td align=" & align_var & " dir=" & dir_obj_var & "><span  class=Form_R style=""width:75;text-align:right"">" & vbCrLf &_
            task_date & "</span></td><td align=" & align_var & " width=90 style=""padding-right:10px;padding-left:10px"" nowrap><b>Target date</b></td>" & vbCrLf &_
	        "</tr><tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_sender_name & "</span></td>" & vbCrLf &_
	        "<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>From</b></td></tr><tr>" & vbCrLf &_
			"<td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=""Form_R"" style=""width:150;"">" & task_reciver_name & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>To</b></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " nowrap dir=" & dir_obj_var & "><span  class=Form_R style=""width:280;"">" & mail_recivers & "</span></td>" & vbCrLf &_
			"<td align="& align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>Addressees</b></td></tr>" & vbCrLf &_			
			"<tr><td align=" & align_var & " dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & task_types_names & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap><b>Type</b></td>" & vbCrLf &_
			"</tr><tr><td align=" & align_var & " valign=top dir=" & dir_obj_var & "><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(task_content) & "</span></td>" & vbCrLf &_
			"<td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap valign=top><b>Content</b></td></tr>" & vbCrLf &_
			"<tr><td height=5 colspan=2 nowrap></td></tr>"
			If not (IsNull(attachment) Or trim(attachment) = "") Then
	        strBody = strBody & "<tr><td dir=" & dir_obj_var & " valign=top>" & vbCrLf &_
	        "<a class=""file_link"" href=""" & strLocal & "download/tasks_attachments/" & attachment & """ target=_blank>" & attachment & "</a>" & vbCrLf &_
	        "</td><td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap valign=top><b>Document</b></td></tr>" & vbCrLf
			End if			
			strBody = strBody & "</table></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
	        "<table border=0 cellpadding=0 align=center cellspacing=1><tr><td height=5 colspan=2 nowrap></td></tr>" & vbCrLf &_
		    "<tr><td align=" & align_var & " width=100% ><span  class=Form_R style=""width:75;text-align:right"">" & task_close_date & "</span></td>" & vbCrLf &_
			"<tr><td align=" & align_var & " width=90 style=""padding-right:10px;padding-left:10px"" nowrap><b>Closing date</b></td></tr>" & vbCrLf &_
			"<tr><td align=" & align_var & " bgcolor=""#e6e6e6"" valign=top nowrap><span class=""Form_R"" style=""width:280;line-height:120%;"">" & breaks(task_replay) & "</span></td>" & vbCrLf &_
		    "<td valign=top style=""padding-right:10px;padding-left:10px"" align=" & align_var & "><b>Closing content</b></td></tr>"
			If NOT (IsNull(attachment_closing) Or trim(attachment_closing) = "") Then
			strBody = strBody & "<tr><td valign=bottom  dir=" & dir_obj_var & ">" & vbCrLf &_
			"<a class=""file_link"" href=""" & strLocal & "download/tasks_attachments/" & attachment_closing & """ target=_blank>" & attachment_closing & "</a>" & vbCrLf &_
			"</td><td valign=bottom style=""padding-right:10px;padding-left:10px"" align=" & align_var & "><b>Closing document</b></td></tr>"
            End If
		    strBody = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf &_				
	        "<tr><td align=" & align_var & " bgcolor=""#C9C9C9""  style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
			"<table cellpadding=0 cellspacing=1 border=0 width=100% align=center><tr><td height=5 colspan=2 nowrap></td></tr>"
		    If trim(task_company_id) <> "" Then
			strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_company_name &	"</td>" & vbCrLf &_
		    "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>Link to " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</b></td>" & vbCrLf &_
			"</tr>"
	        End If
	        If trim(task_contact_id) <> "" Then
	        strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_contact_name & "</td>" & vbCrLf &_
		    "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>Link to " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</b></td>" & vbCrLf &_
	        "</tr>"
	        End If
	        If trim(task_project_id) <> "" Then
	        strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_project_name & "</td>" & vbCrLf &_
		   "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>Link to " &_
		    trim(Request.Cookies("bizpegasus")("Projectone")) & "</b></td></tr>"
	        End If
	        If trim(task_appeal_id) <> "" Then
	        strBody = strBody & "<tr><td align=" & align_var & " width=100% dir=" & dir_obj_var & ">" & task_appeal_id & " - " & productName & "</td>" & vbCrLf &_
		   "<td align=" & align_var & " width=120 nowrap style=""padding-right:10px;padding-left:10px""><b>Link to form</b></td></tr>"
	        End If
        End If  
        
	  'If Len(toMail) > 0 And Len(fromEmail) > 0 Then
	    If trim(lang_id) = "1" Then
        strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
        "</td></tr><tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/tasks/default.asp?T=IN&task_status=3&UserID="&task_reciver_id&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF;font-weight:600'>לחץ כאן לפרטים</a><br></td></tr>" & vbCrLf  &_
        "</table></body></html>"
        Else
	    strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
	    "</td></tr><tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/tasks/default.asp?T=IN&task_status=3&UserID="&task_reciver_id&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF;font-weight:600'>Click here for details</a><br></td></tr>" & vbCrLf  &_
        "</table></body></html>" 	  
        End If
		'SendTask strBodyTo,toMail,fromEmail,0		 
	  'End If
	  
	 'שליחת סגירת המשימה לשולחי משימות הקודמות 
	 Sub sendToParents(taskID, strBodyTo)
			sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_id & "','" & taskId & "'"
			set rs_task = con.getRecordSet(sqlstr)
			If not rs_task.eof Then		   					 
				taskParentID = trim(rs_task("parent_id"))	
				taskSenderID = trim(rs_task("USER_ID"))
				taskSenderMail = trim(rs_task("Sender_Email"))
			End If
			set rs_task = nothing
			
			toMail = taskSenderMail
			If Len(toMail) > 0 And Len(fromEmail) > 0 Then	     	              
				SendTask strBodyTo,toMail,fromEmail,0	
			End If			
		   
			If trim(taskParentID) <> "" And IsNull(taskParentID) = false Then			        
				sendToParents taskParentID, strBodyTo
			Else	       	        
				Exit Sub
			End If	
	  End Sub	
	
	 sendToParents taskID, strBodyTo
	 
	 'שליחת סגירת משימה למכותבים 
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
			If flag = 0 Then 'מקבל המשימה				
				If trim(lang_id) = "1" Then
					strSub = "" & trim(Request.Cookies("bizpegasus")("TasksOne")) & " סגורה - BIZPOWER"
				Else
					strSub = "Closed " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER"
				End If	
			Else
				If trim(lang_id) = "1" Then
					strSub = "" & trim(Request.Cookies("bizpegasus")("TasksOne")) & " סגורה - BIZPOWER - לידיעה"
				Else
					strSub = "Closed " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER - FYI"
				End If				
			End If	
			Msg.Subject = strSub
			Msg.To = toMail							
			
			Msg.HTMLBody = strBody			
			
			'Msg.Send()						
			Set Msg = Nothing					
		End Sub	
	End If	
End If	
    If Request.QueryString("parentID") = nil Then	
	%>
	<SCRIPT LANGUAGE=javascript>
	<!--
		if(opener)
		{
			opener.focus();
			opener.window.location.href = opener.window.location.href;
		}
		self.close();

	//-->
	</SCRIPT>
	<%Else%>
  <script language=javascript>
	<!--
		document.location.href = "addtask.asp?parentID=<%=taskID%>";
	//-->
  </script>	
  <%End If		
  End If  
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 27 Order By word_id"				
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
<body style="margin:0px;" onload="window.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor=#e6e6e6 dir="<%=dir_var%>">
<%
  If trim(taskId) <> "" Then
	Set rs_task = Server.CreateObject("ADODB.RECORDSET") 
    sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & taskId & "'"
    set rs_task = con.getRecordSet(sqlstr)
	if not rs_task.eof then
		task_content = trim(rs_task("task_content"))
		task_date = trim(rs_task("task_date"))
		task_open_date = trim(rs_task("task_open_date"))
		task_types = trim(rs_task("task_types"))
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
		childID = trim(rs_task("childID"))
		
		If IsDate(task_date) Then
			task_date = Day(task_date) & "/" & Month(task_date) & "/" & Year(task_date)
		End If	
		
		If IsDate(task_open_date) Then
			task_open_date = FormatDateTime(task_open_date,2) & " " & FormatDateTime(task_open_date,4)
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
		
		If trim(task_status) = "1" And task_reciver_id = trim(UserID) Then 'משימה נפתחה פעם ראשונה
			sqlstr = "Update tasks SET task_status = '2' WHERE task_id = " & taskId
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
					if trim(taskId) = trim(node_task_id) And trim(node_user_id) = trim(UserID)  Then					
						objDOM.documentElement.removeChild(objTask)
						exit for
					else
						set objTask = nothing
					end if
				next
				Set objNodes = nothing
				set objTask = nothing
				'objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		' ------ end  deleting the new message from XML file ------
		End If	
	
	If trim(task_status) = "3" And task_sender_id = trim(UserID) Then 'משימ
	
			xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_out.xml"
			'------ start deleting the new message from XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false			
			if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.getElementsByTagName("TASK")
				for j=0 to objNodes.length-1
					set objTask = objNodes.item(j)
					node_task_id = objTask.attributes.getNamedItem("TASK_ID").text					
					node_user_id = objTask.attributes.getNamedItem("Sender_ID").text
					if trim(taskId) = trim(node_task_id) And trim(node_user_id) = trim(UserID) Then					
						objDOM.documentElement.removeChild(objTask)
						exit for
					else
						set objTask = nothing
					end if
				next
				Set objNodes = nothing
				set objTask = nothing
			'	objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		' ------ end  deleting the new message from XML file ------
		End If
	end if
	set rs_task = Nothing	 	
  End If 
%>
<tr>	     
	<td class="page_title" dir="<%=dir_obj_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%> BIZPOWER&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap>

<table border="0" cellpadding="1" cellspacing="0" width=100% align=center>
<tr>
   <td align="<%=align_var%>" width=100% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
    <table width=100% border="0" cellpadding="0" align=center cellspacing="1"> 
	<!--tr><td colspan=3 align="<%=align_var%>" class="title">פרטי המשימה&nbsp;</td></tr-->
	<tr><td colspan=3 height="5" nowrap></td></tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap><span class="Form_R" dir="ltr" style="width:75;"><%=task_open_date%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--תאריך פתיחה--><%=arrTitles(1)%></b></td>
	</tr>
	 <tr>
		<td>
        <%If trim(taskID) <> "" And (trim(childID) <> "" Or trim(parentID) <> "") Then%>        
        <A class="but_menu" href="#" style="width:150" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");'>&nbsp;<!--היסטוריה--><%=arrTitles(21)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a>
        <%End If%>
        </td>
        <td align="<%=align_var%>"><span style="width:75;text-align:center" class="task_status_num<%=task_status%>">
         <%=arr_Status(task_status)%></span>
         </td>
         <td align="<%=align_var%>" width=90 style="padding-right:10px;padding-left:10px" nowrap><b><!--סטטוס--><%=arrTitles(2)%></b></td>
     </tr>
	<tr>	
	<td colspan=2 align="<%=align_var%>" dir="ltr"><span  class="Form_R" style="width:75;text-align:right"><%=task_date%></span></td>
	<td align="<%=align_var%>" width=90 style="padding-right:10px;padding-left:10px" nowrap><b><!--תאריך יעד--><%=arrTitles(3)%></b></td>
	</tr>	
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;"><%=task_sender_name%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--מאת--><%=arrTitles(4)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;"><%=task_reciver_name%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--אל--><%=arrTitles(5)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:280;"><%=mail_recivers%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--מכותבים--><%=arrTitles(22)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:280;line-height:120%;"><%=task_types_names%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--סוג--><%=arrTitles(6)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 valign=top dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:280;line-height:120%;"><%=breaks(task_content)%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap valign=top><b><!--תוכן--><%=arrTitles(7)%></b></td></tr>
	<tr><td height=5 colspan=2 nowrap></td></tr>
	<%If not (IsNull(attachment) Or trim(attachment) = "") Then%>	
	<tr><td dir="<%=dir_obj_var%>" colspan=2 valign=top>
	<input type="hidden" name="attachment" ID="attachment" value="<%=vFix(attachment)%>">
	<a class="file_link" href="<%=strLocal%>download/tasks_attachments/<%=attachment%>" target=_blank><%=attachment%></a>	
	</td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap valign=top><b><!--מסמך--><%=arrTitles(8)%></b></td></tr>
	<%End if%>
	</table></td></tr>
	
	<%If ((trim(task_status) = "1" Or trim(task_status) = "2") And trim(task_reciver_id) = trim(UserID)) Or trim(task_status) = "3" Then%>
	<FORM name="form_closing" ACTION="closetask.asp" METHOD="post" enctype="multipart/form-data">
	<input type="hidden" name="taskId"  value="<%=taskId%>" ID="taskId">
	<input type="hidden" name="deleteFile" value="0" ID="deleteFile">
	<input type="hidden" name="close" value="1" ID="close">
	<tr>
	<td align="<%=align_var%>" width=100% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
	<table border="0" cellpadding="0" align=center cellspacing="1">     
	<!--tr><td colspan=2 align="<%=align_var%>" class="title">פרטי סגירה</td></tr-->      
	<tr><td height=5 colspan=2 nowrap></td></tr>
	<%If IsNull(task_close_date) Then
		task_close_date = Date()
	  End If 
	%> 
	<tr>
	<td align="<%=align_var%>" width=100%><input type="text" style="width:80;" value="<%=task_close_date%>" id="task_close_date" <%If trim(task_reciver_id) = trim(UserID)  Then%> class="Form" onclick="return DoCal(this)" <%Else%> class="Form_R" readonly <%End If%> name="task_close_date" dir="rtl" ReadOnly></td>
	<td align="<%=align_var%>" width=90 style="padding-right:10px;padding-left:10px" nowrap><b><!--תאריך סגירה--><%=arrTitles(9)%></b></td>   
	</tr>   	
	<tr>             
		<td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" nowrap>
		<textarea id="task_replay" name="task_replay" dir="<%=dir_obj_var%>" style="height:80;width:280;line-height:120%;" <%If trim(task_reciver_id) = trim(UserID)  Then%> class="Form" <%Else%> class="Form_R" readonly <%End If%>><%=task_replay%></textarea>
		</td>	
		<td valign="top" style="padding-right:10px;padding-left:10px" align="<%=align_var%>"><b><!--תוכן סגירה--><%=arrTitles(10)%></b></td>
	</tr>  
	<%If (IsNull(attachment_closing) Or trim(attachment_closing) = "") And (trim(task_reciver_id) = trim(UserID)) Then%>
	<tr><td valign="top"  dir="<%=dir_obj_var%>"><input type="file" name="attachment_closing" ID="attachment_closing" size=30 value="<%=vFix(attachment_closing)%>"></td>
	<td valign="top" style="padding-right:10px;padding-left:10px" align="<%=align_var%>"><b><!--מסמך סגירה--><%=arrTitles(11)%></b></td>
	</tr>
	<%ElseIf NOT (IsNull(attachment_closing) Or trim(attachment_closing) = "") Then%>
	<tr><td valign="bottom"  dir="<%=dir_obj_var%>">
	<input type="hidden" name="attachment_closing" ID="attachment_closing" value="<%=vFix(attachment_closing)%>">
	<a class="file_link" href="../../../download/tasks_attachments/<%=attachment_closing%>" target=_blank><%=attachment_closing%></a>
	&nbsp;&nbsp;<input type=image src="../../images/delete_icon.gif" border=0 hspace=0 vspace=0 style="position:relative;top:5px" onclick="return DeleteAttachmentClose('<%=taskID%>')">
	</td>
	<td valign="bottom" style="padding-right:10px;padding-left:10px" align="<%=align_var%>"><b><!--מסמך סגירה--><%=arrTitles(12)%></b></td>
	</tr>
	<%End if%>
	<tr><td height=5 colspan=2 nowrap></td></tr>       
	</table></td></tr>
	</form>
	<% End If%>
	</table></td></tr>	
	<tr>
	<td align="<%=align_var%>" bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none">
	<table cellpadding=1 cellspacing=0 border=0 width=100% align=center>	
	<tr><td height=5 colspan=2 nowrap></td></tr>   
   <%If trim(task_company_id) <> "" Then%>
	<tr  <%If trim(private_flag) <> "0" Then%> style="display: none;" <%End If%> >
		<td align="<%=align_var%>" width=100% dir="<%=dir_obj_var%>">
		<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
		<A href="#" onclick="window.opener.location.href='../companies/company.asp?companyID=<%=task_company_id%>';" class="links_down" dir="<%=dir_obj_var%>">
		<%End If%>
		<%=task_company_name%>
		<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(13)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(task_contact_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width=100% dir="<%=dir_obj_var%>">
		<%If trim(COMPANIES) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../companies/contact.asp?contactID=<%=task_contact_id%>';" class="links_down" dir="<%=dir_obj_var%>">
		<%End If%>		
		<%=task_contact_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(14)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(task_project_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width=100% dir="<%=dir_obj_var%>">
		<%If trim(COMPANIES) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../projects/project.asp?companyID=<%=task_company_id%>&projectID=<%=task_project_id%>';" class="links_down" dir="<%=dir_obj_var%>">
		<%End If%>
		<%=task_project_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(15)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>   
	</tr>
	<%End If%>
	 <%If trim(task_appeal_id) <> "" Then%>
	<tr>
	    <td align="<%=align_var%>" width=100% dir="<%=dir_obj_var%>">
		<%If trim(SURVEYS) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../appeals/appeal_card.asp?appid=<%=task_appeal_id%>';" class="links_down" dir="<%=dir_obj_var%>">
		<%End if%>
		<%=task_appeal_id & " - " & productName%>
		<%If trim(SURVEYS) = "1" Then%>
		</a>
		<%End if%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור לטופס--><%=arrTitles(16)%></span></td>
	</tr>
	<%End If%>		   	
	<tr><td height=5 colspan=2 nowrap></td></tr>
	</table>
	</td></tr>
<%If trim(task_reciver_id) = trim(UserID) Then%>
<tr><td align=center colspan="2" height=5 nowrap></td></tr>
<tr><td align=center colspan="2">
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td align=center width=140 nowrap>
<A class="but_menu" style="width:120" onclick='document.location.href="addtask.asp?parentID=<%=taskID%>";'><span id="word17" name=word17><!--צור--><%=arrTitles(17)%></span> <%=trim(Request.Cookies("bizpegasus")("TasksOne"))%> <span id="word18" name=word18><!--חדשה--><%=arrTitles(18)%></span></a></td>
<td align=center width=100%>
<A class="but_menu" style="width:170" onClick="return CheckFields(2)"><!--סגור וצור משימה--><%=arrTitles(23)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td>
</td>
<td align=center width=120 nowrap>
<A class="but_menu" style="width:100" onClick="return CheckFields(1)"><!--סגור--><%=arrTitles(19)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a>
</td>
</tr>
<tr><td align=center colspan="2" height=5 nowrap></td></tr>
<%End If%>
</table>
</div>
</body>
</html>
<%set con=Nothing%>