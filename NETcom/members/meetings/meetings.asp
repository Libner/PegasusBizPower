<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%	If trim(lang_id) = "1" Then
		Session.LCID = 1037				
	Else
		Session.LCID = 2057
	End If

   if Request("Page")<>"" then
		Page=Request("Page")
	else
		Page=1
	end if	
	
	If Request.QueryString("numOfRow") <> nil Then
	        numOfRow = Request.QueryString("numOfRow")
	Else numOfRow = 1
	End If	
	If trim(Request("participant_id")) <> "" Then
		participant_id = trim(Request("participant_id"))
	Else
		participant_id = ""'UserID	
	End If	

	if lang_id = "1" then
		arr_Status = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status = Array("","Future","Done","Summary added","Postponed")	
	end if	
	 
	'משתמש מורשה להוסיף פגישות לאחרים
	sqlstr = "Select IsNULL(Add_Meetings,0) From Users Where User_ID = " & UserID
	set rs_check = con.getRecordSet(sqlstr)
	if not rs_check.eof Then
		AddMeetings =  rs_check.Fields(0)
	else
		AddMeetings = 0
	end if	
	
  	'עדכן סטטוס פגישות שעברו
	sqlstr = " execute UpdateMeetingStatus '" & OrgID & "'"
    'Response.Write sqlstr
    con.ExecuteQuery(sqlstr) 
    
	sort = Request.QueryString("sort")	
	if trim(sort)="" then  sort=0 end if     
    
	dim sortby(12)	
	sortby(0) = "meeting_date ,start_time, end_time"
	sortby(1) = "rtrim(ltrim(company_name)),task_date DESC,task_id DESC"
	sortby(2) = "rtrim(ltrim(company_name)) DESC,task_date DESC,task_id DESC"
	sortby(3) = "meeting_date ,start_time, end_time"
	sortby(4) = "meeting_date DESC,start_time, end_time"
	sortby(5) = "contact_name,task_date DESC,task_id DESC"
	sortby(6) = "contact_name DESC,task_date DESC,task_id DESC"
	sortby(7) = "sender_name,task_date DESC,task_id DESC"
	sortby(8) = "sender_name DESC,task_date DESC,task_id DESC"
	sortby(9) = "reciver_name,task_date DESC,task_id DESC"
	sortby(10) = "reciver_name DESC,task_date DESC,task_id DESC"
	sortby(11) = "project_name,task_date DESC,task_id DESC"
	sortby(12) = "project_name DESC,task_date DESC,task_id DESC"
	
	If trim(Request("start_date")) <> "" Then	
		start_date = trim(Request("start_date"))		
	Else
		start_date = DateAdd("d", -14, Now())
	End If
	start_date = Day(start_date) & "/" & Month(start_date) & "/" & Year(start_date) 
	
	If IsDate(start_date) Then
		start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date) 
	Else
		start_date_ = ""
	End If

	If trim(Request("end_date")) <> "" Then
		end_date = trim(Request("end_date"))		
	Else
		end_date = DateAdd("d", 14, Now())
	End If	
	end_date = Day(end_date) & "/" & Month(end_date) & "/" & Year(end_date) 
	
	If IsDate(end_date) Then
		end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date)
	Else
		end_date_ = ""
	End If	
	
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
	
	search_status = trim(Request("search_status"))
	
 	If Request("delete") <> nil And Request.QueryString("meeting_id") <> nil Then	    
		con.executeQuery("DELETE FROM meetings WHERE meeting_id = " & Request.QueryString("meeting_id"))	
		Response.Redirect "meetings.asp?search_company="& Server.URLEncode(search_company) &"&search_contact="& Server.URLEncode(search_contact)&"&search_project="&Server.URLEncode(search_project)&"&start_date="&start_date&"&end_date="&end_date & "&participant_id=" & participant_id
		Response.End
	End If
	
	If trim(Request.QueryString("postpone")) = "1" And Request.QueryString("meeting_id") <> nil Then 'דחה פגישה
		sqlstr = "Update meetings Set meeting_status = 4 Where meeting_id = " &  Request.QueryString("meeting_id")
		con.ExecuteQuery(sqlstr)
		Response.Redirect "meetings.asp?search_company="& Server.URLEncode(search_company) &"&search_contact="& Server.URLEncode(search_contact)&"&search_project="&Server.URLEncode(search_project)&"&start_date="&start_date&"&end_date="&end_date & "&participant_id=" & participant_id
		Response.End
	End If	
	
	If trim(Request.QueryString("make")) = "1" And Request.QueryString("meeting_id") <> nil Then 'דחה פגישה
		sqlstr = "Update meetings Set meeting_status = 1 Where meeting_id = " &  Request.QueryString("meeting_id")
		con.ExecuteQuery(sqlstr)
		Response.Redirect "meetings.asp?search_company="& Server.URLEncode(search_company) &"&search_contact="& Server.URLEncode(search_contact)&"&search_project="&Server.URLEncode(search_project)&"&start_date="&start_date&"&end_date="&end_date & "&participant_id=" & participant_id
		Response.End
	End If		

	urlSort="meetings.asp?search_company="& Server.URLEncode(search_company) &"&search_contact="& Server.URLEncode(search_contact)&"&search_project="&Server.URLEncode(search_project)&"&start_date="&start_date&"&end_date="&end_date & "&participant_id=" & participant_id & "&search_status=" & search_status
	
    sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 73 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
			End If
		Next
	End If
	set rstitle = Nothing		
%> 
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT language="JavaScript" src='../../tooltip.js'></SCRIPT>
<script LANGUAGE="JavaScript">
<!--
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
	
	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 	
		
	function CheckDelMeeting(meeting_id)
	{
	 <%
		If trim(lang_id) = "1" Then
			str_confirm = "? האם ברצונך למחוק את הפגישה"
		Else
			str_confirm = "Are you sure want to delete the meeting ?"
		End If   
     %>		
		if(window.confirm("<%=str_confirm%>") == true)
		{
		    document.location.href = "<%=urlSort%>&delete=1&meeting_id=" + meeting_id; 
		}
		return true;
	}
	function CheckPostpone(meeting_id)
	{
	  <%If trim(lang_id) = "1" Then	%>str_confirm = "? האם ברצונך לדחות את הפגישה";<%Else%>str_confirm = "Are you sure want to postpone the meeting?";<%End If%>		
	  var answer = window.confirm(str_confirm);
	  if(answer == true)
		document.location.href = "<%=urlSort%>&postpone=1&meeting_id=" + meeting_id;
	  return false;
	}
	
	function CheckMake(meeting_id)
	{
	  <%If trim(lang_id) = "1" Then	%>str_confirm = "? האם ברצונך לקבוע מחדש את הפגישה";<%Else%>str_confirm = "Are you sure want to reschedule the meeting?";<%End If%>		
	  var answer = window.confirm(str_confirm);
	  if(answer == true)
		document.location.href = "<%=urlSort%>&make=1&meeting_id=" + meeting_id;
	  return false;
	}		
    function addmeeting(meetingID)
	{
		h = parseInt(500);
		w = parseInt(465);
		window.open("addmeeting.asp?meetingID=" + meetingID + "&meeting_date=<%=dtCurrentDate%>&participant_id=<%=participant_id%>", "AddMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	function closemeeting(meetingID)
	{
		h = parseInt(490);
		w = parseInt(465);
		window.open("addmeeting.asp?meetingID=" + meetingID + "&meeting_date=<%=dtCurrentDate%>&participant_id=<%=participant_id%>", "AddMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}		
	
	function OpenExcel()
	{
		window.document.form_search.action = "meetings_report.asp";
		window.document.form_search.target = "_blank";
		window.document.form_search.submit();
		return false;
	}
	
    function OpenSearch()
	{
		window.document.form_search.action = "meetings.asp";
		window.document.form_search.target = "_self";
		window.document.form_search.submit();
		return false;
	}

//-->
</script>
</head>
<body>
<div id="ToolTip"></div>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
<tr><td width=100% align="<%=align_var%>">
<%numOftab = 0%>
<%numOfLink = 5%>
<%topLevel2 = 46 'current bar ID in top submenu - added 03/10/2019%>

<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>"> 
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>		   
<tr><td width=100%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellpadding=0 cellspacing=0 dir="<%=dir_var%>">
   <tr>    
    <td width="100%" valign="top" align="center">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
    <tr>
    <td bgcolor=#FFFFFF align="left" width="100%" valign=top>
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>">
		<tr>	
		<td align=center class="title_sort" nowrap><!--מחק--><%=arrTitles(12)%></td>
		<td align=center class="title_sort" nowrap><!--דחה--><%=arrTitles(30)%></td>
		<td width="85" nowrap align="center" class="title_sort"><!--הוזן ע''י--><%=arrTitles(26)%></td>		
		<td width="100%" nowrap align="center" class="title_sort"><!--תאור--><%=arrTitles(11)%></td>
		<td width="90" nowrap align="center" class="title_sort"><!--משתתפים--><%=arrTitles(24)%></td>
		<td width="100" nowrap align="center" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
		<td width="150" nowrap align="center" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
		<!--td width="70" nowrap align="center" class="title_sort"><--שעת סיום><arrTitles(10)%></td-->
		<td width="50" nowrap align="center" class="title_sort"><!--שעת התחלה--><%=arrTitles(9)%></td>
		<td width="55" nowrap align="center" class="title_sort">&nbsp;</td>
		<td width="58" nowrap align="center" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="<%=arrTitles(22)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="<%=arrTitles(23)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" title="<%=arrTitles(23)%>"><%end if%><!--תאריך--><%=arrTitles(2)%><img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="65" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--'סט--><%=arrTitles(13)%>&nbsp;</td>
		</tr>    
<%
If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
	PageSize = RowsInList
Else	
    PageSize = 25
End If	 
'sqlStr = "SET DATEFORMAT DMY; Select DISTINCT meeting_date,meeting_id,meeting_content,start_time,end_time,meeting_status, "&_
'" company_name, contact_name from meetings_view where ORGANIZATION_ID = "& OrgID & " Order By " & sortby(sort)
sqlstr = "exec dbo.get_meetings_paging " & Page & "," & PageSize & ",'" & sFix(search_company) & "','" & sFix(search_contact) & "','" & sFix(search_project) & "','" & search_status & "','" & participant_id & "','" & OrgID & "','" & sortby(sort) & "','" & start_date_ & "','" & end_date_ & "'"
'Response.Write sqlStr
'Response.End
set rs_meetings = con.GetRecordSet(sqlStr)
if not rs_meetings.EOF then	
	  
	recCount = rs_meetings("CountRecords")
	do while not rs_meetings.eof
	meetingID = trim(rs_meetings(1))
	company_name = trim(rs_meetings(4))
	contact_name = trim(rs_meetings(5))
	meetingDate = trim(rs_meetings(6))
	If IsDate(meetingDate) Then
		nameOfDay = WeekDayName(WeekDay(meetingDate)) 
	Else
		nameOfDay = ""
	End If
	status = trim(rs_meetings(7))
	startTime = trim(rs_meetings(8))
	endTime = trim(rs_meetings(9))
	meetingContent = trim(rs_meetings(10))
	UserName = trim(rs_meetings(12))
	
    If Len(meetingContent) > 30 Then
		meetingContentS = Left(meetingContent,29) & "..."
	Else
		meetingContentS = meetingContent 		
    End If	
	
	meeting_participants = ""
	sqlstr = "exec dbo.get_meeting_participants '"&meetingID&"','"&OrgID&"'"
	set rs_meeting_participants = con.getRecordSet(sqlstr)
	If not rs_meeting_participants.eof Then
	    meeting_participants = rs_meeting_participants.getString(,,"<br>","<br>")
	Else
		meeting_participants = ""
	End If		
	
    If trim(status) = "1" Then
       href = "href=""javascript:addmeeting('" & meetingID & "')"""   
    Else
       href = "href=""javascript:closemeeting('" & meetingID & "')"""     
    End If	
%>
	<tr name="title<%=rownum%>" id="title<%=rownum%>"  class="card">
	    <td align=center valign=top><input type=image src="../../images/delete_icon.gif" border=0 hspace=0 vspace=0 onclick="return CheckDelMeeting('<%=meetingID%>')"></td>
	    <td align=center valign=top><%If trim(status) <> "4" Then%><input type=image src="../../images/notok_icon.gif" border=0 hspace=0 vspace=0 title="<%=vFix(arrTitles(30))%>" onclick="return CheckPostpone('<%=meetingID%>')"><%Else%><input type=image src="../../images/ok_icon.gif" border=0 hspace=0 vspace=0 title="<%=vFix(arrTitles(31))%>" onclick="return CheckMake('<%=meetingID%>')"><%End If%></td> 	            
	    <td align="center" valign="top"><a class="link_categ" <%=href%>><%=UserName%></a></td>
	    <td align="<%=align_var%>" valign="top" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(11)%>','<%=Escape(breaks(meetingContent))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=meetingContentS%>&nbsp;</td>	    
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>"><%=meeting_participants%></a></td>	    
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%>><%=contact_name%></a></td>
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%>><%=company_name%></a></td>
		<!--td align="center" valign="top"><a class="link_categ" <%=href%>><%=endTime%></a></td-->
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=startTime%></a></td>
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%>><%=nameOfDay%></a></td>
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%>><%=FormatMediumDateShort(meetingDate)%></a></td>
		<td align="center" valign="top" dir="<%=dir_obj_var%>"><a class="task_status_num<%=status%>" <%=href%>><%=arr_Status(status)%></a></td>	  
	</tr>
	
<% 
 	  rs_meetings.movenext		
	  loop
	  set rs_meetings = nothing	
		  
	  NumberOfPages = Fix((recCount / PageSize)+0.9)
	  if NumberOfPages > 1 then
	  urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan=11 nowrap class="card">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(20)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(19)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<!--tr><td  class="title_form" align=center colspan=9><%=recCount%> :ואצמנ תומושר כ"הס</td></tr-->										 
	<%End If%> 
	<tr>
	   <td colspan=11 height=18 class="card<%=class_%>" align=center dir="ltr" style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(21)%>&nbsp;<%=recCount%>&nbsp;<%=arrTitles(1)%> &nbsp;</td>
	</tr>
	<%Else%>
	<tr><td colspan=11 class="card<%=class_%>" align=center>&nbsp;&nbsp;</td></tr>
<% End If %>
</table></td>
<td width=80 nowrap valign=top class="td_menu" style="border: 1px solid #808080; border-top: 0px">
<table cellpadding=1 cellspacing=0 width=100%>
<FORM action="meetings.asp?sort=<%=sort%>" method=POST id=form_search name=form_search target="_self">   
<tr><td align="<%=align_var%>" colspan=2><b><!--מתאריך--><%=arrTitles(17)%></b>&nbsp;</td></tr>
<tr>
	<td align=center colspan=2 nowrap>
	<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70" onclick="return DoCal(this);" readonly></td>
</tr>
<tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b><!--עד תאריך--><%=arrTitles(18)%></b>&nbsp;</td></tr>
<TR>
	<td align="center" colspan=2 nowrap>
	<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return DoCal(this);" readonly></td>		
</TR>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:80;" href="#" onclick="return OpenSearch();">&nbsp;<!--עדכן תאריך--><%=arrTitles(8)%>&nbsp;</a></td></tr>
<tr><td nowrap colspan=2 height=10 nowrap></td></tr>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b></td></tr>
<tr>
<td align="<%=align_var%>"><input type="image" onclick="return OpenSearch();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:70;" value="<%=vFix(search_company)%>" name="search_company" ID="search_company"></td></tr>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></b></td></tr>
<tr><td align="<%=align_var%>"><input type="image" onclick="return OpenSearch();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:70;" value="<%=vFix(search_contact)%>" name="search_contact" ID="search_contact"></td></tr>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></b></td></tr>
<tr><td align="<%=align_var%>"><input type="image" onclick="return OpenSearch();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:70;" value="<%=vFix(search_project)%>" name="search_project" ID="search_project"></td>
</tr>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><!--עובד--><%=arrTitles(15)%></b></td></tr>
<tr>
<td align="<%=align_var%>" colspan=2>
<select name="participant_id" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="participant_id" onChange="return OpenSearch();">
<option value=""><!-- כולם --><%=arrTitles(25)%></option>
<%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " And Active = 1 ORDER BY FIRSTNAME + ' ' + LASTNAME")
    do while not UserList.EOF
    selUserID=UserList(0)
    selUserName=UserList(1)%>
    <option value="<%=selUserID%>" <%if trim(participant_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
</select>
</td>
</tr>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><!--סטטוס--><%=arrTitles(13)%></b></td></tr>
<tr>
<td align="<%=align_var%>" colspan=2>
<select name="search_status" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="search_status" onChange="return OpenSearch();">
<option value=""><!-- כולם --><%=arrTitles(25)%></option>
<%For i=1 To Ubound(arr_Status)%>
    <option value="<%=i%>" <%if trim(search_status)=trim(i) then%> selected <%end if%>><%=arr_Status(i)%></option>
<%Next%>
</select>
</td>
</tr>
<tr><td nowrap colspan=2 height=10 nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:80;" href="#" onclick="return OpenExcel()">&nbsp;<!--הצג דוח--><%=arrTitles(29)%>&nbsp;</a></td></tr>
<tr><td nowrap colspan=2 height=10 nowrap></td></tr>
</FORM>
<tr><td colspan=2 height=5 nowrap></td></tr>
</table></td></tr></table>
</td></tr></table>
</td></tr></table>
<script language="javascript">
<!--
document.body.onmousemove = overhere;
//-->
</script>
</body>
<%set con=Nothing%>
</html>

