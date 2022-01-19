<!--#include file="../../connect.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<script>
function ClearAllS()
{
document.getElementById("contentMess").value='';
document.getElementById("contentClose").value='';
document.getElementById("status_type").value='0';
document.getElementById("task_type").value='0';
document.getElementById("search_contact").value='';
document.getElementById("start_date").value='';
document.getElementById("end_date").value='';

document.form_searchTop.submit();
}
</script>
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
 T = trim(Request.QueryString("T"))
 reciver_id = trim(Request("reciver_id"))
 sender_id = trim(Request("sender_id"))
 task_status = trim(Request.QueryString("task_status")) 
 contentClose=trim(Request("contentClose"))
 ContentMess=trim(Request("ContentMess"))
  if IsNumeric(Request.Form("status_type")) then
status=Request.Form("status_type")
end if


'if status="" then
'status = trim(Request.QueryString("status"))
'end if

if status="" then
status=0
end if
 
if IsNumeric(Request.Form("task_type")) then
Ttype=Request.Form("task_type")
else
Ttype= trim(Request.QueryString("Ttype"))
end if
if  Ttype="" then
Ttype=0
end if
taskTypeID=Ttype

'task_status= trim(Request.QueryString("status"))
'if status>0 then
'where_status = " And (task_status in (" & sFix(status) & "))"
'end if
 'response.Write contentClose &":"& ContentMess
 If trim(T) = "OUT" Then 'לפתוח משימות יוצאות סגורות	
    If trim(sender_id) = "" then
		sender_id = UserID	
	End If
 Else ' לפתוח משימות נכנסות פתוחות
    T = "IN"	
    If trim(reciver_id) = "" Then	
		reciver_id = UserID		
	End If		
	If trim(task_status) = "" Then
		task_status = "1,2"
	End If	
 End If
 
 where_sender = " AND user_id = " & sender_id
 where_reciver = " AND reciver_id = " & reciver_id  

If trim(T) = "OUT" Then ' מחק הודעה על משימות סגורות שלי
    xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_out.xml"	
	set objDOMOut = Server.CreateObject("Microsoft.XMLDOM")
	objDOMOut.async = false			
	if objDOMOut.load(server.MapPath(xmlFilePath)) then
		'set objNodesOut = objDOMOut.documentElement.childNodes 		
		Set objNodesOut = objDOMOut.getElementsByTagName("TASK")
		for j=0 to objNodesOut.length-1		
			set objTaskOUT = objNodesOut.item(j)								
			node_user_id = objTaskOUT.attributes.getNamedItem("Sender_ID").text				
			if trim(sender_id) = trim(node_user_id) Then					
				objDOMOut.documentElement.removeChild(objTaskOUT)										
			end if	
			Set objTaskOUT = nothing					
		next
		Set objNodesOut = nothing
		set objTaskOUT = nothing
	'	objDOMOut.save server.MapPath(xmlFilePath)
	end if
	set objDOMOut = nothing	
	
 ElseIf trim(T) = "IN" Then 
 		xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_in.xml"
		'------ start deleting the new message from XML file ------
		set objDOMIn = Server.CreateObject("Microsoft.XMLDOM")
		objDOMIn.async = false			
		if objDOMIn.load(server.MapPath(xmlFilePath)) then
			Set objNodesIn = objDOMIn.getElementsByTagName("TASK")
			for j=0 to objNodesIn.length-1
				set objTaskIN = objNodesIn.item(j)						
				node_user_id = objTaskIN.attributes.getNamedItem("Reciver_ID").text				
				if trim(reciver_id) = trim(node_user_id) Then					
					objDOMIn.documentElement.removeChild(objTaskIN)											
				end if
				set objTaskIN = nothing												
			next								
			Set objNodesIn = nothing
			set objTaskIN = nothing
			'objDOMIn.save server.MapPath(xmlFilePath)				
		end if
		set objDOMIn = nothing		    
	End If 

	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 29 Order By word_id"				
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
	  set rstitle = Nothing	  	%> 
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html;">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../../tooltip.js"></script>
<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

<script language="javascript" type="text/javascript">
<!--

	//var oPopup = window.createPopup();
	function callCalendar(pf,pid)
	{
	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}
	function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}
	function StatusDropDown(obj)
	{
		oPopup.document.body.innerHTML = Status_Popup.innerHTML; 
		oPopup.document.charset = "windows-1255";
		oPopup.show(-20, 17,65, 82, obj);    
	}
  
	function closeTask(contactID,companyID,taskID)
	{
			h = parseInt(520);
			w = parseInt(460);
			window.open("closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	
	function task_typeDropDown(obj)
	{
	    oPopup.document.body.innerHTML = task_type_Popup.innerHTML;
	    oPopup.document.charset = "windows-1255"; 
	    oPopup.show(0, 17, obj.offsetWidth+50, 182, obj);    
	}

	function DoCal(elTarget)
	{
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
    
    function addtask(contactID,companyID,taskID)
	{
		h = parseInt(530);
		w = parseInt(460);
		window.open("addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	function CheckDelTask(T,taskID)
	{
     <%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק לצמיתות את המשימה"
		Else
			str_confirm = "Are you sure want to delete the task?"
		End If	
	 %>	
		if (confirm("<%=str_confirm%>")){     
	//	alert( document.getElementById("search_contact").value)
			document.location.href = "default.asp?T=" + T + "&delTaskID=" + taskID +"&search_contact=" +  document.getElementById("search_contact").value+ "&reciver_id=" + document.getElementById("reciver_id").value;
			return false;
		}else{
			return false;
		}
	}

function cball_onclick() {
	var strid = new String(document.getElementById("ids").value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.getElementById('cb'+ arrid[i]).checked = document.getElementById("cb_all").checked;	
}

	function checkMoveTasks()
	{
		var fl = 0;
		document.getElementById("trapp").value = '';
		if(document.getElementById("ids"))
		{
		var strid = new String(document.getElementById("ids").value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.getElementById('cb'+ arrid[i]).checked)
				{	document.getElementById("trapp").value = document.getElementById("trapp").value + arrid[i] + ',';
					fl = 1;
				}	
			}		  
			if (fl && confirm("? האם ברצונך להעביר את המשימות הנבחרות בצורה גורפת")){
				var txtnew = new String(document.getElementById("trapp").value);
				document.getElementById("trapp").value = txtnew.substr(0,txtnew.length - 1);
				
				h = parseInt(520);
				w = parseInt(520);
				window.open("../tasks/movetasks.asp?tasksId=" + document.getElementById("trapp").value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
				return true;
			}
			else if (fl) return false;
		}
		window.alert("! נא לסמן משימות על מנת לבצע העברה גורפת");
		return false;
	}	
	}

//-->
</script>
</head>
<%
 if lang_id = "1" then
    arr_Status = Array("","חדש","בטיפול","סגור")	
    self_name = "עצמי"
 else
    arr_Status = Array("","new","active","close")	
    self_name = "Self"
 end if
 
 'where_status = " And (task_status in (" & sFix(task_status) & "))"
 if status=0 then
 status = task_status
 end if

 If trim(T) = "OUT"  Then
	class_ = "4"
 Else
	class_ = "7"
 End if	  
 
 If Request("start_date") <> nil Then	
	start_date = Request("start_date")
	start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)   
 Else
	start_date_ = ""  
 End If

 If Request("end_date") <> nil Then
	end_date = Request("end_date")
    end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date)
 Else
    end_date_ = ""
 End If
 
 if trim(Request.QueryString("page"))<>"" then
    page=Request.QueryString("page")
 else
    Page=1
 end if  
 
 if trim(Request.QueryString("numOfRow"))<>"" then
    numOfRow=Request.QueryString("numOfRow")
 else    
    numOfRow = 1
 end if  

 sort = Request.QueryString("sort")	
 if trim(sort)="" then  sort=0 end if  

 if trim(Request.Form("search_company")) <> "" Or trim(Request.QueryString("search_company")) <> "" then
	 search_company = trim(Request.Form("search_company"))
	 if trim(Request.QueryString("search_company")) <> "" then
		search_company = trim(Request.QueryString("search_company"))
	 end if					
	 where_company = " And company_Name LIKE '%"& sFix(search_company) &"%'"
	 task_status = "all"			
 else
	 where_company = ""		
 end if


if trim(Request.Form("search_contact")) <> "" Or trim(Request.QueryString("search_contact")) <> "" then
	 search_contact = trim(Request.Form("search_contact"))
	 if trim(Request.QueryString("search_contact")) <> "" then
		search_contact = trim(Request.QueryString("search_contact"))
	 end if					
	 where_contact = " And CONTACT_NAME LIKE '%"& sFix(search_contact) &"%'"
	 task_status = "all"					
else
	 where_contact = ""		
end if

if trim(Request("search_project")) <> "" then		
	search_project = trim(Request("search_project"))	
	where_project = " And project_Name LIKE '%"& sFix(search_project) &"%'"			
	status = "all"	
else
	where_project = ""	
	search_project = ""		
end if

  delTaskID = trim(Request("delTaskID"))
  If delTaskID<>nil And delTaskID<>"" Then 		
		sqlstr = "EXECUTE get_task '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & delTaskID & "'"
		set rs_task = con.getRecordSet(sqlstr)
		if not rs_task.eof then
				task_content = trim(rs_task("task_content"))
				task_date = trim(rs_task("task_date"))
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
				
				mail_recivers = ""
				sqlstr = "Select FirstName + Char(32) + LastName From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
				"Where Task_ID = " & delTaskID
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
				sqlstr = "Exec dbo.get_task_types '"&delTaskID&"','"&OrgID&"'"
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
					toMail = trim(rswrk("EMAIL"))
				End If
				set rswrk = Nothing
						
				sqlstr = "Select EMAIL From USERS Where USER_ID = " & task_sender_id
				set rswrk = con.getRecordSet(sqlstr)
				If not rswrk.eof Then
					fromEmail = trim(rswrk("EMAIL"))
				End If
				set rswrk = Nothing		
	
			end if
			set rs_task = Nothing	 				
		   
            If trim(task_status) = "1" Or trim(task_status) = "2" Then  
	        '<!-- start sending mail -->
	        If trim(lang_id) = "1" Then		
			strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
			"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=""#e6e6e6"" dir=" & dir_var & ">" & vbCrLf &_
			"<tr><td class=page_title style=""background-color:#000000;color:#FFFFFF"" dir=" & dir_obj_var & ">&nbsp;ביטול&nbsp;" & trim(Request.Cookies("bizpegasus")("TasksOne")) & "&nbsp;-&nbsp;BIZPOWER&nbsp;</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_
            "<td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
			"<table width=100% border=""0"" cellpadding=""0"" align=center cellspacing=""1"" dir=" & dir_var & ">" & vbCrLf &_
			"<tr><td colspan=4 height=""5"" nowrap></td></tr><tr>" & vbCrLf &_
            "<td align=" & align_var & "><span style=""width:75;text-align:center"" class=""task_status_num" & task_status & """>" & vbCrLf &_
            arr_Status(task_status) & "</span></td><td align=" & align_var & " width=80 style=""padding-right:10px;padding-left:10px"" nowrap><b>סטטוס</b></td>" & vbCrLf &_
            "</tr><tr><td align=" & align_var & " dir=" & dir_obj_var & "><span  class=Form_R style=""width:75;text-align:right"">" & vbCrLf &_
            task_date & "</span></td><td align=" & align_var & " width=80 style=""padding-right:10px;padding-left:10px"" nowrap><b>תאריך יעד</b>&nbsp;</td>" & vbCrLf &_
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
	        "<a class=""file_link"" href=" & strLocal & "download/tasks_attachments/" & attachment & " target=_blank>" & attachment & "</a>" & vbCrLf &_
	        "</td><td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap valign=top><b>מסמך</b>&nbsp;</td></tr>" & vbCrLf
			End if
			strBody = strBody & "</table></td></tr>" & vbCrLf &_
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
		   	strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=""" & strLocal & "netcom/IE4.css"" rel=STYLESHEET type=""text/css""></head>" & vbCrLf  &_			
			"<body><table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=""#e6e6e6"" dir=" & dir_var & ">" & vbCrLf &_
			"<tr><td class=page_title style=""background-color:#000000;color:#FFFFFF"" dir=" & dir_obj_var & ">" & trim(Request.Cookies("bizpegasus")("TasksOne")) & "&nbsp;cancellation&nbsp;-&nbsp;BIZPOWER&nbsp;</td>" & vbCrLf &_
            "</tr><tr>" & vbCrLf &_
            "<td align=" & align_var & " width=100% nowrap bgcolor=""#e6e6e6"" valign=top style=""border: 1px solid #808080;border-top:none"">" & vbCrLf &_
			"<table width=100% border=""0"" cellpadding=""0"" align=center cellspacing=""1"" dir=" & dir_var & ">" & vbCrLf &_
			"<tr><td colspan=4 height=""5"" nowrap></td></tr><tr>" & vbCrLf &_
            "<td align=" & align_var & "><span style=""width:75;text-align:center"" class=""task_status_num" & task_status & """>" & vbCrLf &_
            arr_Status(task_status) & "</span></td><td align=" & align_var & " width=80 style=""padding-right:10px;padding-left:10px"" nowrap><b>Status</b>&nbsp;</td>" & vbCrLf &_
            "</tr><tr><td align=" & align_var & " dir=" & dir_obj_var & "><span  class=Form_R style=""width:75;text-align:right"">" & vbCrLf &_
            task_date & "</span></td><td align=" & align_var & " width=80 style=""padding-right:10px;padding-left:10px"" nowrap><b>Target date</b>&nbsp;</td>" & vbCrLf &_
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
	        "<a class=""file_link"" href=" & strLocal & "download/tasks_attachments/" & attachment & " target=_blank>" & attachment & "</a>" & vbCrLf &_
	        "</td><td align=" & align_var & " style=""padding-right:10px;padding-left:10px"" nowrap valign=top><b>Document</b>&nbsp;</td></tr>" & vbCrLf
			End if
			strBody = strBody & "</table></td></tr>" & vbCrLf &_
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
	  
	  If Len(toMail) > 0 And Len(fromEmail) > 0 Then
	   
        strBodyTo = strBody & "<tr><td height=5 colspan=2 nowrap></td></tr></table>" & vbCrLf  &_
        "</td></tr></table></body></html>"
        
		SendTask strBodyTo,toMail,fromEmail,0 		 
	  End If
	  
	  sqlstr = "Select DISTINCT Email From task_to_users Inner Join Users On Users.User_ID = task_to_users.User_ID " &_
	  "Where Task_ID = " & delTaskID
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
					strSub = "ביטול " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER"
				Else
					strSub = trim(Request.Cookies("bizpegasus")("TasksOne")) & " cancellation " & " - BIZPOWER"
				End If	
			Else
				If trim(lang_id) = "1" Then
					strSub = "ביטול " & trim(Request.Cookies("bizpegasus")("TasksOne")) & " - BIZPOWER - לידיעה"
				Else
					strSub = trim(Request.Cookies("bizpegasus")("TasksOne")) & " cancellation " & " - BIZPOWER - FYI"
				End If				
			End If				
			Msg.Subject = strSub
			Msg.To = toMail			
			Msg.HTMLBody = strBody				
			Msg.Send()						
			Set Msg = Nothing        	         
		End Sub
	End If 	
    
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
				
				if trim(delTaskID) = trim(node_task_id) Then					
					objDOM.documentElement.removeChild(objTask)							
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
				if trim(delTaskID) = trim(node_task_id) Then					
					objDOM.documentElement.removeChild(objTask)						
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
     			
	sqlstr = "Select attachment, attachment_closing FROM tasks WHERE task_id = " & delTaskID	
	set rs_file = con.getRecordSet(sqlstr)
	if not rs_file.eof then
		attachment = trim(rs_file(0))
		attachment_closing = trim(rs_file(1))
	end if
	set rs_file = nothing
	
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object! 
	sqlstr = "Select TOP 1 task_id FROM tasks WHERE attachment LIKE '" & sFix(attachment) & "' AND task_id <> " & delTaskID
	set rs_check = con.getRecordSet(sqlstr)
	if rs_check.eof then		  			
   		file_path="../../../download/tasks_attachments/" & attachment
   		'Response.Write fs.FileExists(server.mappath(file_path))
		'Response.End
		if fs.FileExists(server.mappath(file_path)) then
			set a = fs.GetFile(server.mappath(file_path))
			a.delete			
		end if	
	end if
	set rs_check = nothing		
	
	sqlstr = "Select TOP 1 task_id FROM tasks WHERE attachment_closing LIKE '" & sFix(attachment_closing) & "' AND task_id <> " & delTaskID
	set rs_check = con.getRecordSet(sqlstr)
	if rs_check.eof then
		file_path="../../../download/tasks_attachments/" & attachment_closing   
		if fs.FileExists(server.mappath(file_path)) then
			set a = fs.GetFile(server.mappath(file_path))
			a.delete			
		end if	
	end if
	set rs_check = nothing	
	set fs = nothing
	
	sqlstr = "UPDATE tasks SET PARENT_ID = NULL WHERE PARENT_ID = " & delTaskID
	con.executeQuery(sqlstr)
	
	'--insert into changes table
	sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
	" SELECT 'משימה', 'אל: ' + U.FIRSTNAME + ' איש קשר:'  + IsNULL(CONTACT_NAME, ''), task_id, 'מחיקה', getDate(), " & UserID & _
	" FROM dbo.tasks T LEFT OUTER JOIN dbo.USERS U ON U.User_Id = T.reciver_id " & _
	" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = T.Contact_Id WHERE (Task_Id = "& delTaskID &")"
	con.executeQuery(sqlstr)	
	
	con.executeQuery "DELETE From tasks WHERE task_id = " & delTaskID	
	if search_contact<>"" then
    Response.Redirect "default.asp?T=" & T & "&search_contact="&search_contact
elseif reciver_id<>"" then

Response.Redirect "default.asp?T=" & T & "&reciver_id="&reciver_id
	else
    Response.Redirect "default.asp?T=" & T 
    end if
				
End If %>
<body>
<div id="ToolTip"></div>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 5%>
  <%If trim(TASKS) = "1" And trim(T) = "IN" Then%>
  <%numOfLink = 1%>
<%topLevel2 = 39 'current bar ID in top submenu - added 03/10/2019%>
  <%ElseIf trim(TASKS) = "1" And trim(T) = "OUT" Then%>
  <%numOfLink = 2%>
<%topLevel2 = 40 'current bar ID in top submenu - added 03/10/2019%>
  <%End If%> 
<!--#include file="../../top_in.asp"-->
</td></tr>
<FORM action="default.asp?T=<%=T%>" method=POST id="form_searchTop" name=form_searchTop target="_self">   

<tr><td bgcolor=#FFD011 dir=rtl align=left>
<table border=0 cellpadding=0 cellspacing=0 align=left dir=ltr>
<tr height=30>
<td align="left" width="30"><a href="#" onclick="javascript:ClearAllS();" class="button_edit_1" style="width:50px">&nbsp;הכל&nbsp;</a></td>
<td align="left"><a class="button_edit_1" style="width:60;" href="#" onclick="document.form_searchTop.submit();">&nbsp;<!--עדכן תאריך-->חיפוש&nbsp;</a></td>
<td>&nbsp;</td>
<TD align=right><input type="text" name="contentClose" id="contentClose" value="<%=contentClose%>"  class="search"  style="width: 60px;" dir="rtl"></TD>
<td align="right" nowrap>&nbsp;תוכן סגירה</td>
<td>&nbsp;</td>

<TD align=right><input type="text" name="contentMess" id="contentMess" value="<%=contentMess%>"  class="search" style="width:60px;" dir="rtl"></TD>
<td align="right" nowrap>&nbsp;תוכן משימה</td>
<td>&nbsp;</td>

	<td align="right" >
			<a href='' onclick='callCalendar(document.form_searchTop.end_date,"Asend_date");return false;' id="Asend_date">
	<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70"  readonly></a>
	</td>	
<td align="right" nowrap>&nbsp;עד תאריך</td>
<td>&nbsp;</td>

	<td align="right">
		<a href='' onclick='callCalendar(document.form_searchTop.start_date,"Asstart_date");return false;' id='Asstart_date'>
		<input dir='ltr' type='text' class='passw' size=6 id="start_date" name='start_date' value='<%=start_date%>' maxlength=10 readonly>		</a>
	</td>
	<td align="right">&nbsp;מתאריך</td>
<td>&nbsp;</td>
<%If trim(T) = "OUT" Then%>
<td align="right"><!-- onChange="form_searchTop.submit();"-->
<select name="reciver_id" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="reciver_id">
    <option value=""><!-- כולם --><%=arrTitles(19)%></option>
    <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE (ORGANIZATION_ID = " & OrgID & ") AND (ACTIVE = 1) ORDER BY FIRSTNAME + ' ' + LASTNAME")
    do while not UserList.EOF
    selUserID=UserList(0)
    selUserName=UserList(1)%>
    <option value="<%=selUserID%>" <%if trim(reciver_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
</select>
</td>
<td align="right">&nbsp;אל</td>

<%End If%>
<%If trim(T) = "IN" Then%>
<td align="right" ><!--onChange="form_search.submit();"-->
<select name="sender_id" dir="<%=dir_obj_var%>" class="norm"  ID="sender_id">
    <option value=""><!-- כולם --><%=arrTitles(17)%></option>
    <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE (ORGANIZATION_ID = " & OrgID & ") AND (ACTIVE = 1) ORDER BY FIRSTNAME + ' ' + LASTNAME")
    do while not UserList.EOF
    selUserID=UserList(0)
    selUserName=UserList(1)%>
    <option value="<%=selUserID%>" <%if trim(sender_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
</select>
</td>
<td  align="right">&nbsp;מאת&nbsp;&nbsp;&nbsp;</td>
<%End If%>
<td>&nbsp;</td>
<td align="right"><input type="text" class="search" dir="rtl" style="width:70;" value="<%=search_contact%>" name="search_contact" ID="search_contact"></td>
<td  align="right" nowrap>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>

<td>&nbsp;</td>

	<td>
<select name="task_type" dir="<%=dir_obj_var%>" class="norm"  ID="task_type" >
    <option value="0">כל הסוגים</option>
    <%sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = " & OrgID & " Order By activity_type_id"
	'response.Write sqlstr
	'response.end
	set rstask_type = con.getRecordSet(sqlstr)
    do while not rstask_type.EOF
    seltypeID=rstask_type(0)
    seltypeName=rstask_type(1)%>
    <option value="<%=seltypeID%>" <%if cint(TType)=cInt(seltypeID) then%> selected <%end if%>><%=seltypeName%></option>
    <%
    rstask_type.MoveNext
    loop
    set rstask_type=Nothing%>
</select>
</td>	
<td>&nbsp;&nbsp;</td>
	<td>
<select name="status_type" dir="<%=dir_obj_var%>" class="norm"  ID="status_type" >
    <option value="0" <%if status="0" then%>selected<%end if%>>כל סטטוסים</option>
    <option value="1" <%if status="1" then%>selected<%end if%> >חדש</option>
    <option value="2" <%if status="2" then%>selected<%end if%>>בטיפול</option>
    <option value="3" <%if status="3" then%>selected<%end if%>>סגור</option>

</select>
</td>

<td width=2%>&nbsp;&nbsp;</td>
<td  width=10%>&nbsp;&nbsp;</td>



</tr>
</table>

</td></tr>		
</form>   
<tr><td width=100%>
<%  
'If trim(Request("taskTypeID")) <> nil Then
	'taskTypeID = trim(Request("taskTypeID"))
'Else 
'	taskTypeID = ""	
'End If	  
		
dim sortby(12)	
If trim(T) = "OUT" Then 
sortby(0) = "task_status, task_date, task_id DESC"
Else
sortby(0) = "task_status, task_date, task_id DESC"
End If
sortby(1) = "company_name, task_date DESC, task_id DESC"
sortby(2) = "company_name DESC, task_date DESC, task_id DESC"
sortby(3) = "task_date, task_id DESC"
sortby(4) = "task_date DESC, task_id DESC"
sortby(5) = "contact_name, task_date DESC, task_id DESC"
sortby(6) = "contact_name DESC, task_date DESC, task_id DESC"
sortby(7) = "U.FIRSTNAME, U.LASTNAME, task_date DESC,task_id DESC"
sortby(8) = "U.FIRSTNAME DESC, U.LASTNAME DESC,task_date DESC,task_id DESC"
sortby(9) = "U1.FIRSTNAME,  U1.LASTNAME,task_date DESC,task_id DESC"
sortby(10) = "U1.FIRSTNAME DESC,  U1.LASTNAME DESC,task_date DESC,task_id DESC"
sortby(11) = "project_name,task_date DESC,task_id DESC"
sortby(12) = "project_name DESC,task_date DESC,task_id DESC"

urlSort="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;task_status="&task_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;tasktypeID=" & tasktypeID & "&amp;T=" & T
UrlStatus="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;tasktypeID=" & tasktypeID & "&amp;T=" & T
urlType="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact) &"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date&"&amp;task_status="&task_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id  & "&amp;T=" & T
%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellpadding=0 cellspacing=0 dir="<%=dir_var%>">
   <tr>    
    <td width="100%" valign="top" align="center">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
    <tr>
    <td bgcolor=#FFFFFF align="left" width="100%" valign=top>
    <table width="100%" cellspacing="1" cellpadding="0" border=0>      
     <tr style="line-height:18px"> 	  
          <%If trim(sender_id) <> "" Then%>   
          <td align="center" class="title_sort"><!--מחק--><%=arrTitles(3)%></td> 
          <%End If%>
          <td align="center" class="title_sort" width=29>&nbsp;</td>   
  	      <td width=30% align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן סגירה&nbsp;</td>	                         
  	      <td width=30% align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן משימה&nbsp;</td>	          
	      <td id="td_type" width="150" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--סוג--><%=arrTitles(5)%>&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(24)%>" align=absmiddle onclick='window.open("Types_list.asp","TypeList","left=300,top=20,width=250,height=182,scrollbars=1"); return false;'>&nbsp;</td>
	      <!--td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="11" OR trim(sort)="12" then%>_act<%end if%>"><%if trim(sort)="11" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="12" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=11"  title="<%=arrTitles(26)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>bot<%elseif trim(sort)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td-->
	      <td  align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=5"  title="<%=arrTitles(26)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>	      
	      <%If trim(is_companies) = "1" Then%>
          <td width="90" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort"  href="<%=urlSort%>&amp;sort=1" title="<%=arrTitles(26)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>	                  
          <%End If%>
          <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="9" OR trim(sort)="10" then%>_act<%end if%>"><%if trim(sort)="9" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="10" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=9"  title="<%=arrTitles(26)%>"><%end if%>&nbsp;<!--אל--><%=arrTitles(6)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="9" then%>bot<%elseif trim(sort)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
          <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=7"  title="<%=arrTitles(26)%>"><%end if%>&nbsp;<!--מאת--><%=arrTitles(7)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="65" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=3"  title="<%=arrTitles(26)%>"><%end if%><!--תאריך יעד--><%=arrTitles(8)%><img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="40" id="td_status" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--'סט--><%=arrTitles(9)%>&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(24)%>" align=absmiddle onclick='window.open("Status_list.asp","StatusList","left=300,top=20,width=250,height=80,scrollbars=1"); return false;'></td>
	      <td class="title_sort">&nbsp;</td>
	      <%If trim(T) = "OUT" Then%>
	      <td width="20" nowrap dir="<%=dir_obj_var%>" class="title_sort"><INPUT type="checkbox" LANGUAGE="javascript" onclick="return cball_onclick()" title="<%=arrTitles(20)%>" id="cb_all" name="cb_all"></td>
	      <%End If%>
    </tr>   
<%  
   PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
   If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
     	PageSize = 10
   End If	
   
   Set tasksList = Server.CreateObject("ADODB.RECORDSET")
   sqlstr = "exec dbo.get_tasks_paging " & Page & "," & PageSize & ",'" & sFix(search_company) & "','" & sFix(search_contact) & "','" & sFix(search_project) & "','" & status & "','" & UserID & "','" & OrgID & "','" & lang_id & "','" & taskTypeID & "','" & reciver_id & "','" & sender_id & "','" & sortby(sort) & "','" & start_date_ & "','" & end_date_ & "',' ',' ',' ', ' ','"  & ContentMess & "','" & ContentClose & "'"
   
   'Response.Write sqlstr &"<BR>" &status
   
  ' Response.End
   set  tasksList = con.getRecordSet(sqlstr)
   current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
   dim	IS_DESTINATION 
   ids = ""
   If not tasksList.EOF then		
        recCount = tasksList("CountRecords")		
		       
   while not tasksList.EOF       
       taskId = trim(tasksList(1))   :  ids = ids & taskId 		
       companyID =  trim(tasksList(2))
       contactID = trim(tasksList(3))
       companyName = trim(tasksList(4))      
       contactName = trim(tasksList(5))
       task_date = trim(tasksList(6))
       projectName = trim(tasksList(7))  
       status = trim(tasksList(8))       
	   sender_name = trim(tasksList(9))
	   reciver_name = trim(tasksList(10))      
       task_content = trim(tasksList(11))          
       parentID = trim(tasksList(12))  
       childID = trim(tasksList(15))  
       task_replay = trim(tasksList(16))  
       
       If Len(task_content) > 25 Then
			task_content_short = Left(task_content,22) & "..."
	   Else
		   task_content_short = task_content 		
       End If
       task_content_short = task_content 
       If Len(task_replay) > 25 Then
			task_replay_short = Left(task_replay,22) & "..."
	   Else
		   task_replay_short = task_replay 		
       End If       
        task_replay_short = task_replay 
     
       If IsDate(trim(task_date)) Then
		  task_date = Day(trim(task_date)) & "/" & Month(trim(task_date)) & "/" & Right(Year(trim(task_date)),2)
		  if DateDiff("d",task_date,current_date) >= 0 then
			IS_DESTINATION = true
		  else
			IS_DESTINATION = false
		  end if
	   Else
		   task_date=""
		   IS_DESTINATION = false
	   End If     				
	   task_types_n = ""
	   sqlstr = "exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
	   set rs_task_types = con.getRecordSet(sqlstr)
	   If not rs_task_types.eof Then
		    task_types_n = rs_task_types.getString(,,"<br>","<br>")
	   Else
			task_types_n = ""
	   End If
       
       If (trim(status) = "1" OR trim(status) = "2") And trim(UserID) = trim(sender_id) Then
			href = "href=""javascript:addtask('" & contactID & "','" & companyID & "','" & taskID & "')"""   
       Else
			href = "href=""javascript:closeTask('" & contactID & "','" & companyID & "','" & taskID & "')"""     
       End If      %>
  	<tr height=24 bgcolor="#ffd7d7" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#ffd7d7';">

        <%If trim(sender_id) <> "" Then%>  
          <td align="center"  valign="top"><input type="image" src="../../images/delete_icon.gif" onclick="return CheckDelTask('<%=T%>','<%=taskID%>')"></td>         
        <%End If%>  
           <td align="center"  valign="top">
           <%If trim(taskID) <> "" And trim(childID) <> "" Then%>
           <input type=image src="../../images/hets4.gif" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' align="absmiddle">
           <%End If%>
           <%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
           <input type=image src="../../images/hets4a.gif" onclick='window.open("task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' align="absmiddle">
           <%End If%>
           </td>  
	      <td  valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(4)%>','<%=Escape(breaks(task_replay))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(task_replay_short))%>&nbsp;</td>	      	      	      
	      <td  valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(4)%>','<%=Escape(breaks(task_content))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(task_content_short))%>&nbsp;</td>	      	      	      
          <td  valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>><%=task_types_n%></a></td>	
		  <!--td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=projectName%>&nbsp;</a></td-->
          <td  valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=contactName%>&nbsp;</a></td>
          <%If trim(is_companies) = "1" Then%>
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=companyName%>&nbsp;</a></td>
          <%End If%>
          <td valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=reciver_name%>&nbsp;</a></td>
          <td  valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=sender_name%>&nbsp;</a></td>
          <td  valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%> style="font-size: 8pt;" <%if IS_DESTINATION and status <> 3 then%> name=word27 title="<%=arrTitles(27)%>"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%><%If isDate(task_date) Then%>&nbsp;<%=task_date%>&nbsp;<%End If%></a></td>         
	      <td  valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num<%=status%>" <%=href%>><%=arr_Status(status)%></a></td>	  
		  <td  align="center" valign="top"><%If trim(contactID) <> "" And trim(contactID) <> "0" Then%><img src="../../images/forms_icon.gif" border="0" 
			style="cursor: pointer;"	alt="היסטוריית טפסים של הלקוח" 
			onclick="window.open('../appeals/contact_appeals.asp?contactID=<%=contactID%>','winCA','top=20, left=10, width=950, height=500, scrollbars=1');"  ><%End If%></td>
	      <%If trim(T) = "OUT" Then%>
	      <td class="card<%=class_%>" align="center" valign="top"><INPUT type="checkbox" id="cb<%=taskId%>" name="cb<%=taskId%>"></td>
	      <%End If%>			
      </tr> 
	  <%	tasksList.MoveNext
		if not tasksList.eof then
		ids = ids & ","
		end if	  
	  Wend
	  
	  NumberOfPages = Fix((recCount / PageSize)+0.9)
	  if NumberOfPages > 1 then
	  urlSort = urlSort & "&amp;sort=" & sort 	  %>
	  <tr>
		<td width="100%" align="center" colspan="12" nowrap class="card<%=class_%>">
			<table border="0" cellspacing="0" cellpadding="2" dir="ltr" >
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" align="<%=align_var%>"><A class="pageCounter" name=word28 title="<%=arrTitles(28)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="<%=align_var%>"><A class="pageCounter" href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(29)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<!--tr><td  class="title_form" align="center" colspan=9><%=recCount%> :ואצמנ תומושר כ"הס</td></tr-->										 
	<%End If%> 
	<tr>
	   <td colspan="12" height=18 class="card<%=class_%>" align="center" dir="ltr" style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(10)%>&nbsp;<%=recCount%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> &nbsp;</td>
	</tr>
	<%Else%>
	<tr><td colspan="12" class="card<%=class_%>" align="center">&nbsp;&nbsp;</td></tr>
<% End If
   set tasksList = Nothing%>
</table><input type="hidden" id="ids" value="<%=ids%>">
<input type="hidden" name="trapp" value="" ID="trapp">
</td>
<td width=80 nowrap valign="top" class="td_menu" style="border: 1px solid #808080; border-top: 0px">
<FORM action="default.asp?sort=<%=sort%>&amp;T=<%=T%>" method=POST id=form_search name=form_search target="_self">   
<table cellpadding="1" cellspacing="0" width="100%" >
<tr><td align="<%=align_var%>" colspan="2" height=20 class="title_search"><!--חיפוש--><%=arrTitles(11)%></td></tr>
<%If trim(is_companies) = "1" Then%>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b></td></tr>
<tr>
<td align="<%=align_var%>"><input type="image" onclick="form_search.submit();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:70;" value="<%=search_company%>" name="search_company" ID="search_company"></td></tr>
<%End If%>
<tr><td colspan=2 height=7 nowrap></td></tr>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:80px;" href="#" onclick="addtask('','','')"><!--הוסף--><%=arrTitles(18)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td></tr>
<%If trim(T) = "OUT" Then%>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:78px;" href='javascript:void(0)' onclick="return checkMoveTasks();">העבר משימות</a></td></tr>
<%End If%>
<tr><td colspan=2 height=10 nowrap></td></tr>
</table>
</form>
</td></tr></table>
</td></tr></table>
</td></tr></table>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:65; height:82; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
<%For i=1 To uBound(arr_Status)	%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand; border-bottom:1px solid black"
	ONCLICK="parent.location.href='<%=UrlStatus%>&amp;task_status=<%=i%>'">
    &nbsp;<%=arr_Status(i)%>&nbsp;
    </DIV>
<%Next%>    
    <DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=UrlStatus%>&amp;task_status=1,2,3'">&nbsp;<!--כל הרשימה--><%=arrTitles(20)%>&nbsp;</DIV>
</div>
</DIV>
<DIV ID="task_type_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="overflow: scroll; position:absolute; top:0; left:0; width:100%; height:182; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = " & OrgID & " Order By activity_type_id"
	set rstask_type = con.getRecordSet(sqlstr)
	while not rstask_type.eof %>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>&amp;tasktypeID=<%=rstask_type(0)%>&amp;task_status=1,2,3'">&nbsp;<%=rstask_type(1)%>&nbsp;</DIV>
	<%
		rstask_type.moveNext
		Wend
		set rstask_type=Nothing
	%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>'">
    &nbsp;<!--כל הרשימה--><%=arrTitles(21)%>&nbsp;</DIV>
</div>
</DIV>
<script language="javascript" type="text/javascript">
<!--
document.body.onmousemove = overhere;
//-->
</script>
 	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
	
</body>
</html>
<%set con=Nothing%>