<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%
	Response.CharSet = "windows-1255"  
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	UserID=trim(trim(Request.Cookies("bizpegasus")("UserID")))
	CurrUserName=trim(trim(Request.Cookies("bizpegasus")("UserName")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))  
	if lang_id = "1" then
		arr_Status = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status = Array("","Future","Done","Summary added","Postponed")	
	end if	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  align_var = "left"  :  dir_obj_var = "ltr"
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
	End If
	meeting_company_id=trim(Request("companyId"))
	meeting_contact_id=trim(Request("contactID"))
	meeting_project_id=trim(Request("project_id"))
	meetingID = trim(Request("meetingID"))
	meeting_date = trim(Request("meeting_date"))
	participant_id = trim(Request("participant_id"))
	If trim(participant_id) = "" Then
		participant_id = UserID
	End If
	If trim(meeting_date) = "" Then
		meeting_date = Date()
	End If
 
	If trim(meeting_company_id) <> "" Then
		sqlstr = "SELECT company_Name FROM companies WHERE company_Id = " & cLng(meeting_company_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			meeting_company_name = trim(rs_pr("company_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(meeting_contact_id) <> "" Then
		sqlstr = "SELECT contact_Name FROM contacts WHERE contact_Id = " & cLng(meeting_contact_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			meeting_contact_name = trim(rs_pr("contact_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(meeting_project_id) <> "" Then
		sqlstr = "SELECT project_Name FROM projects WHERE project_Id = " & cLng(meeting_project_id)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			meeting_project_name = trim(rs_pr("project_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(participant_id) <> "" Then
		sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & participant_id & " And ORGANIZATION_ID = " & OrgID
		set rs_user = con.getRecordSet(sqlstr)
		if not rs_user.eof then
			userName = trim(rs_user(0)) & " " & trim(rs_user(1))
			found_user = true
		else
			found_user = false	
		end if
		set rs_user = nothing
	End If 
 
	If Request.Form("add") <> nil Then	
		meeting_date = trim(Request.Form("meeting_date")) 
		meeting_status = trim(Request.Form("meeting_status"))   
		meeting_participants = trim(Request.Form("meeting_participants"))   
		start_time = trim(Request.Form("start_time"))   
		end_time = trim(Request.Form("end_time"))   
		meeting_content = trim(Request.Form("meeting_content")) 
		meeting_content2 = trim(Request.Form("meeting_content2")) 
		meeting_closing = trim(Request.Form("meeting_closing"))
		companyId=trim(Request.Form("companyId"))
        contactID=trim(Request.Form("contactID"))
		projectID=trim(Request.Form("project_id"))
		
		If trim(companyID) = "" Or IsNull(companyID) Then
			companyID = "NULL"
		End If	
		If trim(contactID) = "" Or IsNull(contactID) Then
			contactID = "NULL"
		End If	
		If trim(projectID) = "" Or IsNull(projectID) Then
			projectID = "NULL"
		End If		
			
		If trim(meetingID) = "" Then
			sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into meetings (company_id,contact_id,project_id,ORGANIZATION_ID,User_id,meeting_date,start_time,end_time,meeting_content,meeting_content2,meeting_status) "&_
			" values (" & companyID & "," & contactID & "," & projectID & "," & OrgID & "," & UserID & ",'" &_
			meeting_date & "','" & start_time & "','" & end_time & "','" & sFix(meeting_content)& "','" &sFix(meeting_content2) & "','1'); SELECT @@IDENTITY AS NewID"		
			'Response.Write sqlstr
			'Response.End
			set rs_tmp = con.getRecordSet(sqlstr)
				meetingID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing						
		
		Else	
			
			sqlstr = "SET DATEFORMAT DMY; Update meetings set meeting_status = '" & meeting_status &_
			"', meeting_date = '" & meeting_date & "', meeting_content = '" & sFix(meeting_content) &_
			"', meeting_content2 = '" & sFix(meeting_content2) &_
			"', start_time = '" & start_time & "', end_time = '" & end_time & "' Where meeting_id = " & meetingId
			con.executeQuery(sqlstr)	
								   
		End If
		
		If trim(meeting_closing) <> "" Then	
			sqlstr = "Update meetings set meeting_closing = '" & sFix(meeting_closing) & "', meeting_status = '3' Where meeting_id = " & meetingId
			con.executeQuery(sqlstr)	
		End If			 
		
		If trim(meetingId) <> "" Then	
		    sqlstr = "Delete From meeting_to_users Where meeting_id = " & meetingId
		    con.executeQuery(sqlstr)	
			If Len(trim(meeting_participants)) > 0 Then 'לשלוח במייל משימה למותבים
				arrUsers = Split(meeting_participants, ",")
					If IsArray(arrUsers) Then
						For count=0 To Ubound(arrUsers)
							If IsNumeric(arrUsers(count)) Then														    
								sqlstr="Insert Into meeting_to_users (meeting_id,User_id,organization_id) Values (" &_
								meetingID & "," & arrUsers(count) & "," & OrgID & ")"	  	
								con.executeQuery(sqlstr)		
							End If				
						Next
					End If
			End If
		End If						
	 %>
	<SCRIPT LANGUAGE=javascript>
	<!--	
		opener.focus();
		var loc_str = new String(opener.window.location.href);
		if(loc_str.search("default.asp")>0)
		{
			opener.window.location.href = "default.asp?date_=<%=meeting_date%>&participant_id=<%=participant_id%>";
		}
		else
		{
			opener.window.location.reload(true);
		}	
		self.close();
	//-->
	</SCRIPT>
  <%	 End If%>
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
	self.close();
}	
function checkHours() 
{     
     var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");		
     
     meeting_date = window.document.formmeeting.meeting_date.value;
     start_time = window.document.formmeeting.start_time.value;
     end_time = window.document.formmeeting.end_time.value;
     meeting_participants = window.document.formmeeting.meeting_participants.value; 
     meetingID = "<%=meetingID%>";
     OrgID = "<%=OrgID%>";    

	 xmlHTTP.open("POST", 'getHours.asp', false);
	 xml_str = '<?xml version="1.0" encoding="UTF-8"?><request><OrgID>' + OrgID + '</OrgID>';
	 xml_str = xml_str + '<meetingID>' + meetingID + '</meetingID>';
	 xml_str = xml_str + '<meeting_date>'+ meeting_date + '</meeting_date>';
	 xml_str = xml_str + '<start_time>'+ start_time + '</start_time><end_time>'+ end_time + '</end_time>';
	 xml_str = xml_str + '<meeting_participants>'+ meeting_participants + '</meeting_participants></request>';
	 xmlHTTP.send(xml_str);	 
		
	 result = new String(xmlHTTP.ResponseText);						
	 //window.alert(result)
	 
	 if(result.length == 0)	 
		return true;
	 else
	 {<%
			If trim(lang_id) = "1" Then
				str_alert = "! שים לב, קיימות פגישות בשעות שבחרת"
			Else
				str_alert = "Pay attention, another meeting is scheduled in this time!"
			End If	
		%>
		window.alert("<%=str_alert%>");
		return false;	
	 } 	 
}

function CheckFields(meeting_id, meeting_status)
{
  	
  	if(window.document.all("meeting_participants").value=='')
	{
   		<%
			If trim(lang_id) = "1" Then
				str_alert = "! חובה לבחור לפחות משתתף אחד"
			Else
				str_alert = "Please, choose meetings participants !"
			End If	
		%>
 		window.alert("<%=str_alert%>");     
		return false; 
	}
  	if(window.document.all("start_time").value=='')
	{
   		<%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור שעת התחלה של הפגישה"
			Else
				str_alert = "Please, choose the meeting start time !"
			End If	
		%>
 		window.alert("<%=str_alert%>");     
		window.document.all("start_time").focus();
		return false; 
	}
	if(window.document.all("end_time").value=='')
	{
   		<%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור שעת סיום של הפגישה"
			Else
				str_alert = "Please, choose the meeting end time !"
			End If	
		%>
 		window.alert("<%=str_alert%>");     
		window.document.all("end_time").focus();
		return false; 
	}    
	else if (checkHours() == false)
	{
		return false;
	} 
  
   document.formmeeting.submit();               	 
   return true;
}

function DoCal(elTarget){
	if (showModalDialog){
		var sRtn;
		sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=140pt; status=0; help=0;");
		if (sRtn!="")
			elTarget.value = sRtn;
		}else
		alert("Internet Explorer 4.0 or later is required.");
	return false;
	window.document.focus;   
}

function checkTime(field)
{	
		// Checks if time is in HH:MM:SS AM/PM format.
		// The seconds and AM/PM are optional.
        timeStr = new String(field.value);
		var timePat = /^(\d{1,2}):(\d{2})(:(\d{2}))?(\s?(AM|am|PM|pm))?$/;

		var matchArray = timeStr.match(timePat);
		if (matchArray == null) {
		alert("שים לב, השעה שהזנת אינה בפורמט נכון \n\n\ <%=Space(15)%> HH:MM פורמט שעה צריך להיות \n\n <%=Space(43)%> לדוגמה 12:00");
		field.value = "";
		field.focus();
		return false;
		}
		hour = matchArray[1];
		minute = matchArray[2];		

		if (hour < 0  || hour > 23) {
		alert("שעה צריכה להיות בין 0 - 24");
		field.value = "";
		field.focus();
		return false;
		}			
		if (minute<0 || minute > 59) {
		alert ("דקות צריכות להיות בין 0 - 60");
		field.value = "";
		field.focus();
		return false;
		}
		
		//אם הוכנסה שעת סיום הפגישה בודקים ששעת ההתחלה מוקדת משעת הסיום
		if(window.document.all("start_time").value!='' && window.document.all("end_time").value!='')	
		{
			timeStrStart = new String(window.document.all("start_time").value);
			timeStrEnd = new String(window.document.all("end_time").value);
			var timePat = /^(\d{1,2}):(\d{2})(:(\d{2}))?(\s?(AM|am|PM|pm))?$/;

			var startArray = timeStrStart.match(timePat);
			hourStart = startArray[1];
			minuteStart = startArray[2];	
			
			var endArray = timeStrEnd.match(timePat);
			hourEnd = endArray[1];
			minuteEnd = endArray[2];	
			//window.alert(eval(hourStart) + " - " + eval(hourEnd))			   
			if((eval(hourStart) > eval(hourEnd)) || ((eval(hourStart) == eval(hourEnd)) && (eval(minuteStart) > eval(minuteEnd))))
			{
				window.alert("שעת התחלה של הפגישה צריכה להיות\n\n<%=Space(08)%>מוקדמת משעת הסיום של הפגישה");
				window.document.all("start_time").value = '';
				window.document.all("end_time").value = '';				
				return false;
			}	
			if((eval(hourStart) == eval(hourEnd)) && (eval(minuteStart) == eval(minuteEnd)))
			{
				window.alert("שים לב, בחרת שעת התחלת הפגישה\n\n<%=Space(20)%>! ושעת סיום הפגישה זהים");
				window.document.all("start_time").value = '';
				window.document.all("end_time").value = '';
				return false;
			}				
			   
		}	
		return true;			   
}

//פתח רשימת כל הפגישות של המשתתפים בפגישה
function openList()
{
	h = parseInt(500);
	w = parseInt(300);
	meeting_date = window.formmeeting.meeting_date.value;
	meeting_users = window.formmeeting.meeting_participants.value;
	window.open("meetings_list.asp?meeting_date="+meeting_date+"&meeting_users="+meeting_users, "OpenList" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
}
//-->
</script>  
</head>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 74 Order By word_id"				
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
<%  
	If trim(meetingID) <> "" Then
		sqlstr = "Select meeting_status,meeting_content,meeting_content2,meeting_closing,meeting_date,start_time,end_time, "&_
		" contact_name,company_name, project_name, contact_id, company_id, project_id, private_flag, User_ID "&_
		" From meetings_view WHERE meeting_id = " & meetingId
        'response.Write sqlstr
        'response.End
    set rs_meeting = con.getRecordSet(sqlstr)
	if not rs_meeting.eof then		
		meeting_content = trim(rs_meeting("meeting_content"))
		meeting_content2 = trim(rs_meeting("meeting_content2"))				
		meeting_date = trim(rs_meeting("meeting_date"))
		If IsDate(meeting_date) Then
			meeting_date = Day(meeting_date) & "/" & Month(meeting_date) & "/" & Year(meeting_date)
		End If				
		meeting_status = trim(rs_meeting("meeting_status"))			
		start_time = trim(rs_meeting("start_time"))		
		end_time = trim(rs_meeting("end_time"))
		meeting_contact_name = trim(rs_meeting("contact_name"))
		meeting_company_name = trim(rs_meeting("company_name"))	
		meeting_contact_id = trim(rs_meeting("contact_id"))
		meeting_company_id = trim(rs_meeting("company_id"))		
		meeting_project_id = trim(rs_meeting("project_id"))		
		meeting_project_name = trim(rs_meeting("project_name"))
		meeting_closing = trim(rs_meeting("meeting_closing"))	
		private_flag = trim(rs_meeting("private_flag"))	
		meeting_user_id = trim(rs_meeting("User_ID"))
		
		If trim(meeting_user_id) <> "" Then
			sqlstr = "Select FIRSTNAME + ' ' + LASTNAME From Users Where User_ID = " & meeting_user_id &_
			" And Organization_ID = " & OrgID
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				meeting_user_name = trim(rs_name(0))	
			Else
				meeting_user_name = ""
			End If
			set rs_name = Nothing	
		End If
		
		'------------------------------ משתתפים ------------------------------------------------------------
		users_names= ""
		sqlstr = "Select FIRSTNAME + ' ' + LASTNAME From meeting_to_users Inner Join Users On Users.User_ID = meeting_to_users.User_ID " &_
		" Where meeting_ID = " & meetingID & " ORDER BY FIRSTNAME + ' ' + LASTNAME"
		set rs_participants = con.getRecordSet(sqlstr)
		if not rs_participants.eof then
			users_names= rs_participants.getString(,,"<br>","<br>")
		end if
		set rs_participants = nothing
		If Len(users_names) > 1 Then
			users_names= Left(users_names,Len(users_names)-4)
		End If
		meeting_participants = ""
		sqlstr = "Select User_ID From meeting_to_users Where meeting_ID = " & meetingID & " ORDER BY User_ID"
		set rs_participants = con.getRecordSet(sqlstr)
		if not rs_participants.eof then
			meeting_participants= rs_participants.getString(,,",",",")
		end if
		set rs_participants = nothing	
		If Len(meeting_participants) > 1 Then
			meeting_participants= Left(meeting_participants,Len(meeting_participants)-1)
		End If	
	end if
	set rs_meeting = Nothing
  Else
	  meeting_status = 1 : users_names = userName : meeting_participants = participant_id : meeting_user_name = CurrUserName
  End If
  
  'משתמש מורשה להוסיף פגישות לאחרים
	sqlstr = "Select IsNULL(Add_Meetings,0) From Users Where User_ID = " & UserID
	set rs_check = con.getRecordSet(sqlstr)
	if not rs_check.eof Then
		AddMeetings = trim(rs_check.Fields(0))
	else
		AddMeetings = 0
	end if			  
  
  If (trim(meeting_user_id) = trim(UserID) Or InStr(meeting_participants,UserID) > 0) Or trim(meetingID) = "" Or trim(AddMeetings) = "1" Then
	  update_flag = true
  Else
	  update_flag = false
  End If	 
%>
<tr>	     
	<td class="page_title"  dir="<%=dir_obj_var%>">&nbsp;<%If trim(meetingID)="" Then%><!--הוספת--><%=arrTitles(1)%><%Else%><!--עדכון--><%=arrTitles(2)%><%End If%>&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap dir="<%=dir_var%>">
<table border="0" cellpadding="1" cellspacing="0" width="100%" align=center dir="<%=dir_var%>">
<FORM name="formmeeting" ACTION="addmeeting.asp" METHOD="post">
<tr>
   <td align="<%=align_var%>" width=100% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
   <input type=hidden name=meetingID id=meetingID value="<%=meetingID%>">
   <input type=hidden name=meeting_status id="meeting_status" value="<%=meeting_status%>">
   <input type=hidden name=participant_id id="participant_id" value="<%=participant_id%>">
   <input type=hidden name=add id=add value="1">     
   <table border="0" cellpadding="0" align=center cellspacing="2" width=100%>          
   <tr>
   <td align="<%=align_var%>" colspan=4 style="padding:10px">
   <table cellpadding=1 cellspacing=1 border=0 width=100% dir="<%=dir_var%>"    
   <tr>
	<td align="<%=align_var%>" width=100%>
	<input type="text" style="width:75;" value="<%=meeting_date%>" id="meeting_date" name="meeting_date" ReadOnly<%If update_flag Then%> class="Form" onclick="return DoCal(this)" <%Else%> class="Form_R" <%End If%>  dir="<%=dir_obj_var%>">
	</td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--תאריך--><%=arrTitles(3)%></b></td>
   </tr>
   <tr>
   <tr>
	<td align="<%=align_var%>" width=100%>
	<span style="width:110;" class="Form_R" dir="<%=dir_obj_var%>"><%=meeting_user_name%></span>
	</td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--הוזן ע''י--><%=arrTitles(14)%></b></td>
   </tr>
   <tr>   
    <tr><td align="<%=align_var%>" valign=top>   
    <%If update_flag = true Then%>
    <input type=button style="vertical-align:top" value="<%=arrTitles(12)%>" onclick="openList()" class="but_menu" style="width:130" ID="Button1" NAME="Button1"><!--בדוק זמינות-->
 	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
     <INPUT type=image style="vertical-align:top" src="../../images/icon_find.gif" onclick='window.open("users_list.asp?meetingID=<%=meetingID%>&participant_id=<%=participant_id%>&meeting_participants=" + meeting_participants.value,"UsersList","left=300,top=20,width=250,height=500,scrollbars=1"); return false;'>
    <%End If%> 
     <span class="Form_R" dir="<%=dir_obj_var%>" style="width:110;line-height:14px" name="users_names" id="users_names"><%=users_names%></span>
     <input type=hidden name=meeting_participants id="meeting_participants" value="<%=meeting_participants%>">
    </td>
    <td align="<%=align_var%>" width=80 valign=top><b><!--משתתפים--><%=arrTitles(5)%></b></td>   
   </tr>   
   <tr>
	<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>">
	<input type=text <%If update_flag Then%> class="Form" <%Else%> class="Form_R" ReadOnly <%End If%> style="width:50;" dir=ltr value="<%=start_time%>" name="start_time" id="start_time" onchange="checkTime(this)">
	</td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--משעה--><%=arrTitles(6)%></b></td>
   </tr>
   <tr>
	<td align="<%=align_var%>" nowrap  dir="<%=dir_obj_var%>"><input type=text <%If update_flag Then%> class="Form" <%Else%> class="Form_R" ReadOnly <%End If%> style="width:50;" dir=ltr value="<%=end_time%>" name="end_time" id="end_time" onchange="checkTime(this)"></td>
	<td align="<%=align_var%>" width=80 nowrap><b><!--עד שעה--><%=arrTitles(7)%></b></td>
   </tr>	 
   <tr>
   <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">   
	<textarea id="meeting_content"  <%If update_flag Then%> class="Form" <%Else%> class="Form_R" ReadOnly style="height:80" <%End If%> name="meeting_content"  dir="<%=dir_obj_var%>" style="width:290" rows=5><%=meeting_content%></textarea>
   </td>	
   <td valign="top" align="<%=align_var%>"><b><!--תאור--><%=arrTitles(8)%></b></td>
   </tr>  
   <tr>             
	<td align="<%=align_var%>" nowrap  dir="<%=dir_obj_var%>">
	<textarea id="meeting_closing" name="meeting_closing" dir="<%=dir_obj_var%>" style="height:80;width:290;line-height:120%;" <%If update_flag Then%> class="Form" <%Else%> class="Form_R" readonly <%End If%>><%=meeting_closing%></textarea>
	</td>	
	<td align="<%=align_var%>" width=80 nowrap valign=top><b><!--סיכום פגישה--><%=arrTitles(13)%></b></td>
	</tr>  
	  <tr>             
	<td align="<%=align_var%>" nowrap  dir="<%=dir_obj_var%>">
	<textarea id="meeting_content2" name="meeting_content2" dir="<%=dir_obj_var%>" style="height:40;width:290;line-height:120%;" class="Form" ><%=meeting_content2%></textarea>
	</td>	
	<td align="<%=align_var%>" width=80 nowrap valign=top><b><!--הערכת סיכונים--><%=arrTitles(15)%></b></td>
	</tr> 
</table></td></tr>
</table></td></tr>
<%If trim(meetingID) = "" Then%>
<tr>
   <td align="<%=align_var%>" colspan=4 bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none">
   <table cellpadding=2 cellspacing=1 border=0 width=100% align=center>
   <tr <%If trim(private_flag) <> "0" Then%> style="display: none;" <%End If%> >
		<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>">
		<input type="hidden" name="companyId" id="companyId" value="<%=meeting_company_id%>">
		<input type="text" id="CompanyName" name="CompanyName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(meeting_company_name) > 0 Then%><%=vFix(meeting_company_name)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input 
		type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openCompaniesList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="בטל בחירה" onclick="return removeCompany();" src="../../images/delete_icon.gif" align="absmiddle" >
       </td>
       <td align="<%=align_var%>" class="links_down" width=135 nowrap><span style="font-weight:600"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>
       </tr> 
      <tbody name="contacter_body" id="contacter_body" <%If trim(meeting_company_id) = ""  Or IsNull(meeting_company_id) Then%> style="display:none" <%End if%>>
      <tr><td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="contactID" id="contactID" value="<%=meeting_contact_id%>">
		<input type="text" id="contactName" name="contactName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(meeting_contact_name) > 0 Then%><%=vFix(meeting_contact_name)%><%Else%><%End If%>" >&nbsp;&nbsp;<input 
		type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openContactsList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="בטל בחירה" onclick="return removeContact();" src="../../images/delete_icon.gif" align="absmiddle" >
       </td>
       <td align="<%=align_var%>" class="links_down" width=135 nowrap><span style="font-weight:600"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>
       </tr>   
    </tbody>       
    <tbody name="project_body" id="project_body" <%If trim(meeting_company_id) = "" Or IsNull(meeting_company_id) Then%> style="display:none" <%End if%>>
    <TR>
		<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="project_id" id="project_id" value="<%=meeting_project_id%>">
		<input type="text" id="projectName" name="projectName" dir="<%=dir_obj_var%>" style="width: 250px" class="Form_R" readonly 
		value="<%If Len(meeting_project_name) > 0 Then%><%=vFix(meeting_project_name)%><%Else%><%End If%>"  >&nbsp;&nbsp;<input 
		type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openProjectsList();" align="absmiddle" >&nbsp;&nbsp;<input 
		type="image" alt="בטל בחירה" onclick="return removeProject();" src="../../images/delete_icon.gif" align="absmiddle" >
		</TD>
		 <td align="<%=align_var%>" class="links_down" width=135 nowrap><span style="font-weight:600"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>
	</TR>
     </tbody>
    </table>
   </td>
  </tr> 
 <%Else%>
 <tr>
   <td align="<%=align_var%>" colspan=4 bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none" style="padding:3px">
   <input type=hidden name="companyId" id="companyId" value="<%=meeting_company_id%>">
   <input type=hidden name="contactId" id="contactId" value="<%=meeting_contact_id%>">
   <input type=hidden name="project_id" id="project_id" value="<%=meeting_project_id%>">
   <table cellpadding=0 cellspacing=0 border=0 width=100% align=center>
   <tr><td height=5 colspan=2 nowrap></td></tr>   
   <%If trim(meeting_company_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width=100%>
		<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
		<A href="#" onclick="window.opener.location.href='../companies/company.asp?companyID=<%=meeting_company_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>
		<%=meeting_company_name%>
		<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(meeting_contact_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width=100%>
		<%If trim(COMPANIES) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../companies/contact.asp?contactID=<%=meeting_contact_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>		
		<%=meeting_contact_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></span></td>   
	</tr>
	<%End If%>
	<%If trim(meeting_project_id) <> "" Then%>
	<tr>
		<td align="<%=align_var%>" width=100%>
		<%If trim(COMPANIES) = "1" Then%>
		<A href="#" onclick="window.opener.location.href='../projects/project.asp?companyID=<%=meeting_company_id%>&projectID=<%=meeting_project_id%>';" class="links_down"  dir="<%=dir_obj_var%>">
		<%End If%>
		<%=meeting_project_name%>
		<%If trim(COMPANIES) = "1" Then%>
		</a>
		<%End If%>
		</td>
		<td align="<%=align_var%>" class="links_down" width=145 nowrap style="padding-right:10px;padding-left:10px"><span style="font-weight:600"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></span></td>   
	</tr>
	<%End If%>	
	<tr><td height=5 colspan=2 nowrap></td></tr>
   </table>
   </td>
 </tr> 
 <%End If%>          
</form>
</table></td></tr>
<tr><td align=center colspan="2" height=10 nowrap></td></tr>
<%If update_flag = true Then%>
<tr><td align=center colspan="2">
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td align=center width=50%>
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:120" onclick="closeWin();">
</td>
<td align=center width=50%>
<A class="but_menu" href="#" style="width:120;line-height:120%;padding:4px" onClick="return CheckFields('<%=meetingID%>','<%=meeting_status%>')"><%If meetingId <> "" Then%><!--עדכן--><%=arrTitles(11)%><%Else%><!--הוסף--><%=arrTitles(10)%><%End If%></a>
</td>
</tr>
</table>
</td></tr>
<%End If%>
<tr><td align=center colspan="2" height=10 nowrap></td></tr>
</table>
</div>
</body>
<%set con=Nothing%>
</html>