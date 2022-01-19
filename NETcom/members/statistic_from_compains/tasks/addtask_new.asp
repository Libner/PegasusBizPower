<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<% Response.CodePage = 1255
	  Response.CharSet = "windows-1255"
	  Session.LCID = 1037
	  Session.CodePage = 1255 

	task_company_id=trim(Request("companyId"))
	task_contact_id=trim(Request("contactID"))
	task_reciver_id=trim(Request("task_reciver_id"))
	task_project_id=trim(Request("project_id")) 
	taskID=trim(Request("taskID"))
	parentID=trim(Request("parentID"))
	task_appeal_id=trim(Request("appealID"))
	'����� ������ ������ ������
	appealsId=trim(Request("appealsId"))
  
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
	UserID=trim(Request.Cookies("bizpegasus")("UserID"))
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))

	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	if lang_id = "1" then
		arr_Status = Array("","���","������","����")	
	else
		arr_Status = Array("","new","active","close")	
	end if
	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  align_var = "left"  :  dir_obj_var = "ltr"  :  self_name = "Self"
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"  :  self_name = "����"
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
	End If   
	
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
	  set rsbuttons=nothing	 %>	
<%task_open_date = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)

	If trim(task_company_id) <> "" Then
		sqlstr = "SELECT company_Name FROM companies WHERE company_Id = " & cInt(task_company_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			task_company_name = trim(rs_pr("company_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(task_contact_id) <> "" Then
		sqlstr = "SELECT contact_Name FROM contacts WHERE contact_Id = " & cInt(task_contact_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			task_contact_name = trim(rs_pr("contact_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(task_project_id) <> "" Then
		sqlstr = "SELECT project_Name FROM projects WHERE project_Id = " & cInt(task_project_id)
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
			task_content = "���� ����� : " & task_content & vbCrLf & "���� ����� : " & task_replay
		End If
		'------------------------------ ������� -----------------------------------------------------------------
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
			
		'------------------------------ ������� -----------------------------------------------------------------
					
		If trim(task_status) = "1" And task_reciver_id = trim(UserID) And taskID <> "" Then '����� ����� ��� ������
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
	    Set rs_task_types = con.getRecordSet(sqlstr)
	    If not rs_task_types.eof Then
		    task_types_names = rs_task_types.getString(,,",",",")
	    Else
			task_types_names = ""
	    End If		
	
		If Len(task_types_names) > 0 Then
			task_types_names = Left(task_types_names,(Len(task_types_names)-1))
		End If		
	End If
	Set rs_task = Nothing
  Else
	  task_date = Day(date()) & "/" & Month(Date()) & "/" & Year(Date())	 
	  task_status = 1
  End If %>	  
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
						str_alert = "! �� ����� �� ���� ������"
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
		window.alert("! �� ����� ����");  
		window.document.all("companyId").focus();   
		return false; 
	}
	if(window.document.all("contactID").value=='')
	{
		window.alert("! �� ����� <%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>"); 
		window.document.all("contactID").focus();    
		return false; 
	}   
	*/
	if(window.document.all("task_types").value=='')
	{
		<%
				If trim(lang_id) = "1" Then
					str_alert = "! �� ����� ��� " & trim(Request.Cookies("bizpegasus")("TasksOne"))
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
					str_alert = "! �� ����� ���� " & trim(Request.Cookies("bizpegasus")("TasksOne"))
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
					str_alert = ":����� ����� - ��� ������ \n HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT"
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
			str_confirm = "?��� ������ ����� �� ����� ������"
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
<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr>	     
	<td class="page_title"  dir="<%=dir_obj_var%>">&nbsp;<%If trim(taskID)="" Then%><!--�����--><%=arrTitles(9)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%><%Else%><!--�����--><%=arrTitles(10)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%><%End If%>&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap dir="<%=dir_var%>">
<FORM name="formtask" ACTION="savetask.asp" METHOD="post" onSubmit="return CheckFields();" enctype="multipart/form-data" ID="formtask">
<table border="0" cellpadding="1" cellspacing="0" width="100%" align=center dir="<%=dir_var%>">
<tr>
   <td align="<%=align_var%>" width="100%" nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
   <input type=hidden name=parentID id="parentID" value="<%=parentID%>">
   <input type=hidden name=taskID id=taskID value="<%=taskID%>">
   <input type=hidden name=add id=add value="1">
   <input type=hidden name=deleteFile id="deleteFile" value="0">
   <input type=hidden name=task_open_date id="task_open_date" value="<%=task_open_date%>">
   <table border="0" cellpadding="0" align=center cellspacing="5" width="100%">        
   <tr><td colspan=2>
   <table cellpadding=0 cellspacing=0 width="100%" border=0 dir="<%=dir_obj_var%>">     
   <tr>
   <td align="<%=align_var%>" colspan=4>
   <table cellpadding=1 cellspacing=1 border=0 width="100%" dir="<%=dir_var%>"
   <%If trim(taskID) <> "" Then%>
   <tr>
	<td align="<%=align_var%>" colspan=2 nowrap><span dir="ltr" class="Form_R" style="width:75;"><%=task_open_date%></span></td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--����� �����--><%=arrTitles(1)%></b></td>
   </tr>
   <%end if%>
   	 <tr>           
        <td width="150" nowrap>
        <%If trim(taskID) <> "" And trim(childID) <> "" Then%>
         <A class="but_menu" href="#" style="width:150" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=750,height=480,align=center,resizable=0");'><%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<!--����--><%=arrTitles(26)%></a>
        <%Else%>&nbsp;<%End If%>
        </td>         
        <td align="<%=align_var%>" width=150 nowrap><span style="width:70;text-align:center" class="task_status_num<%=task_status%>">
        <%=arr_Status(task_status)%></span>
        </td>
        <td align="<%=align_var%>" width=80 nowrap><b><!--�����--><%=arrTitles(2)%></b></td>
   </tr>
   <tr>
   <td width=150 nowrap>
   <%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
   <A class="but_menu" href="#" style="width:120" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=750,height=480,align=center,resizable=0");'><!--��������--><%=arrTitles(27)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a>
   <%Else%>&nbsp;<%End If%>
   </td> 
   <td align="<%=align_var%>" width=150 nowrap><input type="text" style="width:80;" value="<%=task_date%>" id="task_date" class="Form"  onclick="return DoCal(this)" name="task_date"  dir="<%=dir_obj_var%>" ReadOnly></td>
   <td align="<%=align_var%>" width=80 nowrap><b><!--����� ���--><%=arrTitles(3)%></b></td>   
   </tr>   
    <tr>
	<td align="<%=align_var%>" colspan=2 nowrap  dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;height:12px">&nbsp;<%=task_sender_name%>&nbsp;</span></td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--���--><%=arrTitles(4)%></b></td>
	</tr>
	<%If trim(taskID) = "" Then%>
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
    <td align="<%=align_var%>" width=80 nowrap><b><!--��--><%=arrTitles(5)%></b></td>
    </tr>    
    <tr><td align="<%=align_var%>" colspan="2">   
     <INPUT type=image src="../../images/icon_find.gif" name=users_list id="users_list" onclick='window.open("users_list.asp?taskID=<%=taskID%>&task_addressees=" + task_addressees.value,"UsersList","left=300,top=20,width=250,height=500,scrollbars=1"); return false;'>&nbsp;
     <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px" name="users_names" id="users_names"><%=mail_recivers%></span>
     <input type=hidden name=task_addressees id=task_addressees value="<%=task_addressees%>">
    </td>
    <td align="<%=align_var%>" valign=top width=80 nowrap><b><!--�������--><%=arrTitles(28)%></b></td>
    </tr>
    <%Else%>    
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;"><%=task_reciver_name%></span></td>
	<td align="<%=align_var%>" valign=top width=80 nowrap><b><!--��--><%=arrTitles(5)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:280;"><%=mail_recivers%></span></td>
	<td align="<%=align_var%>" valign=top width=80 nowrap><b><!--�������--><%=arrTitles(28)%></b></td>
	</tr>    
    <%End If%>
    </table></td></tr>      
<%	 If trim(taskID) <> "" Then
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
		  		is_selected=trim(arrSub(2,i)) 	  %>		 
		  <tr>
		  <%If i<=Ubound(arrSub,2) Then%>		 
		  <td width=15  align=center nowrap><input type=checkbox id="task_types" name="task_types" <%If is_selected > 0 Then%> checked <%End If%> value="<%=selSubgectId%>"></td>		 
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
		   <td width=15  align=center nowrap><input type=checkbox id="task_types" name="task_types" <%If is_selected > 0 Then%> checked <%End If%> value="<%=selSubgectId%>"></td>									  
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
 <tr><td valign="top" align="<%=align_var%>"><b><!--����--><%=arrTitles(7)%></b></td></tr>
 <tr>             
	<td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" nowrap>
	<textarea id="task_content" class="Form" name="task_content"  dir="<%=dir_obj_var%>" style="width:330" rows=5><%=task_content%></textarea>
	</td>	
</tr>
<%If IsNull(attachment) Or trim(attachment) = "" Then%>
<tr><td  dir="<%=dir_obj_var%>"><!--����--><%=arrTitles(8)%>&nbsp;&nbsp;&nbsp;&nbsp;<input type="file" name="attachment_file" ID="attachment_file" size=33 value=""></td></tr>
<%Else%>
<tr><td  dir="<%=dir_obj_var%>" valign=top>
<input type="hidden" name="attachment" ID="attachment" value="<%=vFix(attachment)%>">
<b><!--����--><%=arrTitles(9)%></b>&nbsp;&nbsp;&nbsp;&nbsp;<a class="file_link" href="../../../download/tasks_attachments/<%=attachment%>" target=_blank><%=attachment%></a>
&nbsp;&nbsp;<input type=image src="../../images/delete_icon.gif" border=0 hspace=0 vspace=0 style="position:relative;top:4px" onclick="return DeleteAttachment('<%=taskID%>')">
</td></tr>
<%End if%>
</table></td>
</tr>
<%If (trim(task_appeal_id) = "" AND trim(appealsId) = "") And trim(taskID) = "" Then%>
<tr>
   <td align="<%=align_var%>" colspan=4 bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none" style="padding:3px">
   <table cellpadding=2 cellspacing=1 border=0 width="100%" align=center>
   <tr><td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="companyId" id="companyId" value="<%=task_company_id%>">
		<input type="text" id="CompanyName" name="CompanyName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(task_company_name) > 0 Then%><%=vFix(task_company_name)%><%Else%><%End If%>" >&nbsp;&nbsp;<input 
		type="image" alt="��� ������" src="../../images/icon_find.gif" onclick="return openCompaniesList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="��� �����" onclick="return removeCompany();" src="../../images/delete_icon.gif" align="absmiddle" >
       </td>
       <td align="<%=align_var%>" class="links_down" width=130 nowrap><span style="font-weight:600"><!--����� �--><%=arrTitles(14)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>
       </tr> 
      <tbody name="contacter_body" id="contacter_body" <%If trim(task_company_id) = ""  Or IsNull(task_company_id) Then%> style="display:none" <%End if%>>
      <tr><td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="contactID" id="contactID" value="<%=task_contact_id%>">
		<input type="text" id="contactName" name="contactName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(task_contact_name) > 0 Then%><%=vFix(task_contact_name)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input 
		type="image" alt="��� ������" src="../../images/icon_find.gif" onclick="return openContactsList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="��� �����" onclick="return removeContact();" src="../../images/delete_icon.gif" align="absmiddle" >
       </td>
       <td align="<%=align_var%>" class="links_down" width=130 nowrap><span style="font-weight:600"><!--����� �--><%=arrTitles(15)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>
       </tr>   
    </tbody>       
    <tbody name="project_body" id="project_body" <%If trim(task_company_id) = "" Or IsNull(task_company_id) Then%> style="display:none" <%End if%>>
    <TR>
		<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="project_id" id="project_id" value="<%=task_project_id%>">
		<input type="text" id="projectName" name="projectName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(task_project_name) > 0 Then%><%=vFix(task_project_name)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input 
		type="image" alt="��� ������" src="../../images/icon_find.gif" onclick="return openProjectsList();" align="absmiddle">&nbsp;&nbsp;<input 
		type="image" alt="��� �����" onclick="return removeProject();" src="../../images/delete_icon.gif" align="absmiddle">
		</TD>
		 <td align="<%=align_var%>" class="links_down" width=130 nowrap><span style="font-weight:600"><!--����� �--><%=arrTitles(16)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>
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
   <%If trim(task_company_id) <> "" Then%>
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
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--����� �--><%=arrTitles(22)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>   
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
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--����� �--><%=arrTitles(23)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(task_project_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<%If trim(COMPANIES) = "1" Then%>
		<A href="javascript:void(0)" onclick="window.opener.location.href='../projects/project.asp?companyID=<%=task_company_id%>&projectID=<%=task_project_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>
		<%=task_project_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--����� �--><%=arrTitles(24)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>   
	</tr>
	<%End If%>
	 <%If trim(task_appeal_id) <> "" Or trim(appealsId) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<%If trim(task_appeal_id) <> "" Then%>
		<%If trim(SURVEYS)  = "1" Then%>
		<A href="javascript:void(0)" onclick="window.opener.location.href='../appeals/appeal_card.asp?appid=<%=task_appeal_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End if%>
		<%=task_appeal_id & " - " & productName%>
		<%If trim(SURVEYS)  = "1" Then%>
		</a>
		<%End If%>
		<%Else%><%=appealsId%> :�����<%End if%>
		<input type="hidden" id="appealsId" name="appealsId" value="<%=appealsId%>" >
		</td>
		<td align="<%=align_var%>" class="links_down" width="135" nowrap style="padding-right:5px;padding-left:5px"><span style="font-weight:600"><!--����� �����--><%=arrTitles(25)%></span></td>   
	</tr>
	<%End If%>		   	
	<tr><td height="5" colspan="2" nowrap></td></tr>
   </table>
   </td>
 </tr> 
 <%End If%>          
</table>
</form></td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
<tr><td align="center" colspan="2">
<table cellpadding="0" cellspacing="0" width="100%">
<tr>
<%If taskId <> "" And trim(task_reciver_id) = trim(UserID) Then%>
<td align="center" width="50%"><A class="but_menu" style="width:120px" href="addtask.asp?parentID=<%=taskID%>"><!--���--><%=arrTitles(17)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%>&nbsp;<!--����--><%=arrTitles(18)%></a></td>
<%Else%>
<td align="center" width="50%"><input type="button" value="<%=arrButtons(2)%>" class="but_menu" style="width:100" onclick="closeWin();"></td>
<%End If%>
<td align="center" width="50%"><A class="but_menu" href="javascript:void(0)" style="width:100px" onClick="return CheckFields('submit')"><%If taskId <> "" Then%><!--����--><%=arrTitles(20)%><%Else%><!--����--><%=arrTitles(19)%><%End If%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td>
</tr>
</table>
</td></tr>
<tr><td align="center" colspan="2" height="10" nowrap></td></tr>
</table>
</div>
</body>
</html>
<%set con=Nothing%>