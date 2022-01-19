<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 
	If trim(lang_id) = "1" Then
		Session.LCID = 1037 
	Else
		Session.LCID = 2057
	End If
	
	'משתמש מורשה להוסיף פגישות לאחרים
	sqlstr = "Select IsNULL(Add_Meetings,0) From Users Where User_ID = " & UserID
	set rs_check = con.getRecordSet(sqlstr)
	if not rs_check.eof Then
		AddMeetings = trim(rs_check.Fields(0))
	else
		AddMeetings = 0
	end if			

	If AddMeetings = "1" Then
		participant_id = ""
	Else
		participant_id = UserID	
	End If	
	
	if lang_id = "1" then
		arr_Status = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status = Array("","Future","Done","Summary added","Postponed")	
	end if
	
  	'עדכן סטטוס פגישות שעברו
	sqlstr = "exec dbo.UpdateMeetingStatus '" & OrgID & "'"
    'Response.Write sqlstr
    con.ExecuteQuery(sqlstr) 

	found_user = false
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
		userName = trim(userName) & " - "
	End If

' ---------- Page Variables ----------
    Const intCharToShow = 19		' The number of characters shown in each day
    Const bolEditable   = True		' If the calendar is editable or not (Can be tied into password verification)

    Dim dtToday 			' Today's meeting_date
    Dim dtCurrentDate			' The current meeting_date
    Dim aCalendarDays(42)		' Array of possible calendar dates
    Dim iFirstDayOfMonth		' The first day of the month
    Dim iDaysInMonth	 		' The number of days in the month
    Dim iColumns, iRows	, iDay, iWeek	' The numer of columns and rows in the table, and counters to print them
    Dim objConn, strConn, strSQL, objRS ' Database Variables
    Dim counter 			' Loop counter
    Dim strNextMonth, strPrevMonth	' The next and previous month dates
    Dim dailyMsg			' The message for the day
    Dim current_date				' The current day being displayed by the loops
    Dim strPage				' The link that each day takes you too

' ---------- Variable Definitions ----------
    dtToday = Date()
    date_ = Date()
    If Request("date_") <> nil Then
		date_ = Request("date_")
	End If
	
	If Request("currentMonth") = "" And Request("currentYear") = ""  Then	 
		currentMonth = Month(date_)
		currentYear = Year(date_)
	Else
		currentMonth = Request("currentMonth")
		currentYear = Request("currentYear")	
	End If

    If currentMonth <> "" And currentYear <> ""  Then    
		If Request("date_") = nil Then
			If trim(currentMonth) <> cStr(Month(Date())) Then
				dtCurrentDate = "1" & "/" & currentMonth & "/" &  currentYear
			Else
				dtCurrentDate = Day(Date()) & "/" & currentMonth & "/" &  currentYear
			End If	
		Else
			dtCurrentDate = date_
		End If
    Else
         dtCurrentDate = dtToday
    End If    

    iFirstDayOfMonth = DatePart("w", DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), 1))
    iDaysInMonth = DatePart("d", DateSerial(Year(dtCurrentDate), Month(dtCurrentDate)+1, 1-1))

    For counter = 1 to iDaysInMonth
      aCalendarDays(counter + iFirstDayOfMonth - 1) = counter
    Next

    iColumns = 7
    iRows    = 6 - Int((43 - (iFirstDayOfMonth + iDaysInMonth)) / 7)    
    
    'Response.Write (43 - (iFirstDayOfMonth + iDaysInMonth)) / 7

    strPrevMonth = Server.URLEncode(DateAdd("m", -1, dtCurrentDate))
    strNextMonth = Server.URLEncode(DateAdd("m",  1, dtCurrentDate))
    
    If trim(lang_id) = "1" Then
    daysname = array(" 'א"," 'ב"," 'ג"," 'ד"," 'ה"," 'ו","שבת")
    Else
    daysname = array("S","M","T","W","T","F","S")
    End If    
    
	If Request("delete") <> nil And Request.QueryString("meeting_id") <> nil Then
	    delete_meeting = Request.Form("delete_meeting")
	    If trim(delete_meeting) = "0" Then
			con.executeQuery("DELETE FROM meetings WHERE meeting_id = " & Request.QueryString("meeting_id"))
		Else
			sqlstr = "Delete From meeting_to_users Where meeting_id = " & Request.QueryString("meeting_id") &_
			" And User_ID = " & delete_meeting
			con.executeQuery(sqlstr)
			
			sqlstr = "Select Top 1 User_ID From meeting_to_users WHERE meeting_id = " & Request.QueryString("meeting_id")
			set rs_check = con.getRecordSet(sqlstr)
			If rs_check.eof Then 'לא נמצאו משתתפים בפגישה
				con.executeQuery("DELETE FROM meetings WHERE meeting_id = " & Request.QueryString("meeting_id"))
			End If		    	
		End If
		Response.Redirect "default.asp?date_=" & dtCurrentDate & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&participant_id=" & participant_id
	End If
	
	If trim(Request.QueryString("postpone")) = "1" And Request.QueryString("meeting_id") <> nil Then 'דחה פגישה
		sqlstr = "Update meetings Set meeting_status = 4 Where meeting_id = " &  Request.QueryString("meeting_id")
		con.ExecuteQuery(sqlstr)
		Response.Redirect "default.asp?date_=" & dtCurrentDate & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&participant_id=" & participant_id
	End If	
	
	If trim(Request.QueryString("make")) = "1" And Request.QueryString("meeting_id") <> nil Then 'דחה פגישה
		sqlstr = "Update meetings Set meeting_status = 1 Where meeting_id = " &  Request.QueryString("meeting_id")
		con.ExecuteQuery(sqlstr)
		Response.Redirect "default.asp?date_=" & dtCurrentDate & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&participant_id=" & participant_id
	End If	
	
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
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
   function popupcal(elTarget){
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
	function CheckPostpone(meeting_id)
	{
	  <%If trim(lang_id) = "1" Then	%>str_confirm = "? האם ברצונך לדחות את הפגישה";<%Else%>str_confirm = "Are you sure want to postpone the meeting?";<%End If%>		
	  var answer = window.confirm(str_confirm);
	  if(answer == true)
		document.location.href = "default.asp?postpone=1&date_=<%=dtCurrentDate%>&meeting_id=" + meeting_id + "&currentMonth=<%=currentMonth%>&currentYear=<%=currentYear%>&participant_id=<%=participant_id%>"; 
	  return false;
	}
	
	function CheckMake(meeting_id)
	{
	  <%If trim(lang_id) = "1" Then	%>str_confirm = "? האם ברצונך לקבוע מחדש את הפגישה";<%Else%>str_confirm = "Are you sure want to reschedule the meeting?";<%End If%>		
	  var answer = window.confirm(str_confirm);
	  if(answer == true)
		document.location.href = "default.asp?make=1&date_=<%=dtCurrentDate%>&meeting_id=" + meeting_id + "&currentMonth=<%=currentMonth%>&currentYear=<%=currentYear%>&participant_id=<%=participant_id%>"; 
	  return false;
	}	
	function getRealTop(imgElem)
	{
		yPos = eval(imgElem).offsetTop;
		tempEl = eval(imgElem).offsetParent;
		while (tempEl != null)
		{
			yPos += tempEl.offsetTop;
			tempEl = tempEl.offsetParent;
		}
		//alert(yPos);
		return yPos;
	}	
		
	function CheckDelMeeting(meeting_id,participant_id,participant_name,imgObj)
	{
	  participant_name = unescape(participant_name);
	  <%If trim(lang_id) = "1" Then	%>	
		str_all = "מחק את הפגישה כולה";
		str_me = "מחק את " + participant_name + " מהפגישה";
	  <%Else%>
		str_all = "Delete the meeting from all participants";
		str_me = "Delete " + participant_name + " from the meeting participants";
	  <%End If%>	
	
	    window.titleALL.innerHTML = str_all;
	    window.titleMe.innerHTML = str_me;	
	    window.delete_form.user_delete_id.value = participant_id;
	    window.delete_form.action="default.asp?delete=1&date_=<%=dtCurrentDate%>&meeting_id=" + meeting_id + '&currentMonth=<%=currentMonth%>&currentYear=<%=currentYear%>&participant_id=<%=participant_id%>'; 
	    
  		document.all["div_alert"].style.zIndex=11; 
  		document.all["div_alert"].style.top = getRealTop(imgObj)-45;
  		document.all["div_alert"].style.left = 10;
		document.all["div_alert"].style.display='inline';
		document.all["div_alert"].style.visibility="visible";	
       
		//if(window.confirm("<%=str_confirm%>") == true)
		//{
		 //   document.location.href = "default.asp?delete=1&date_=<%=dtCurrentDate%>&meeting_id=" + meeting_id + '&currentMonth=<%=currentMonth%>&currentYear=<%=currentYear%>&participant_id=<%=participant_id%>'; 

		//}
		return false;
	}
	function hideDive() 
	{ 
		document.all["div_alert"].style.display="none"; 
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
		h = parseInt(550);
		w = parseInt(465);
		window.open("addmeeting.asp?meetingID=" + meetingID + "&meeting_date=<%=dtCurrentDate%>&participant_id=<%=participant_id%>", "AddMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}		

//-->
</script>
<STYLE TYPE="text/css">
    <!--   
    .blackBacking   {background-color: #000000;}
    .names 	    {background-color: #CCCCCC; font-size: 13px; color: #000000; text-decoration: none; text-align:  center;  font-weight: bold; border-top : 1px solid black; border-right : 1px solid black; line-height: 150%}
    .calendarBody   {background-color: #F0F0F0; font-size: 12px; color: #000000; text-decoration: none; text-align:  center; border : 1px solid black; border-right : none}
    .calCurrentDay  {background-color: #F0F0F0; font-size: 12px; color: #000000; text-align:  center; font-weight:bold; border: solid #FF5959 2px}
    .calOtherDay    {background-color: #F0F0F0; font-size: 12px; color: #000000; text-align:center; line-height : 140%; font-weight:bold; border-top : 1px solid black; border-right : 1px solid black}
    .calNotDay	    {background-color: #F0F0F0; font-size: 12px; color: #000000; line-height : 140%; text-align:  center; border-top : 1px solid black; border-right : 1px solid black}
    .calFormMenu    {background-color: #4C5D87; font-size: 13px; color: #FFFFFF; text-decoration: none; text-align:  center;  font-weight: bold; border : 1px solid black}
    -->
</STYLE>
</HEAD>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 0%>
<%numOfLink = 4%>
<%topLevel2 = 46 'current bar ID in top submenu - added 03/10/2019%>

<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td class="page_title" colspan=2 dir="<%=dir_obj_var%>"><font color="#6F6DA6"><%=userName%></font>&nbsp;<%=arrTitles(1)%>&nbsp;<%=FormatDateTime(date_,1)%></td></tr>
<tr><td width=100% valign=top>
 <table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" dir="<%=dir_var%>"> 
 <%
	 sqlstr = "SET DATEFORMAT DMY; Select DISTINCT participant_id, ParticipantName From meetings_view Where ORGANIZATION_ID = "& OrgID &_
	 " And meeting_date = '" & dtCurrentDate & "'"
	 If trim(participant_id) <> "" Then
		sqlstr = sqlstr & " And participant_id = " & participant_id
	 End If
	 sqlstr = sqlstr & " Order BY ParticipantName"
	 'Response.Write sqlstr
	 'Response.End
	 set rs_users = con.getRecordSet(sqlstr)
	 If not rs_users.eof Then
		users_arr = rs_users.getRows()
	 End If
	 set rs_users = Nothing
	 
	 If IsArray(users_arr) Then
	 For i=0 To Ubound(users_arr,2)	
	 If trim(participant_id) = "" Then
		participant_name = users_arr(1,i)
 %>
   <tr><td class="title" align="<%=align_var%>" style="padding: 10 10 5 5"><%=participant_name%></td></tr>
   <%End If%>
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>">
	<tr>	
	<td align=center class="title_sort" nowrap><!--מחק--><%=arrTitles(12)%></td>
	<td align=center class="title_sort" nowrap><!--דחה--><%=arrTitles(30)%></td>
	<td width="150" nowrap align="center" class="title_sort"><!--משתתפים--><%=arrTitles(24)%></td>
	<td width="100%" nowrap align="center" class="title_sort"><!--תאור--><%=arrTitles(11)%></td>
	<td width="150" nowrap align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
	<td width="70" nowrap align="center" class="title_sort"><!--שעת התחלה--><%=arrTitles(9)%></td>
	<td width="70" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--'סט--><%=arrTitles(13)%>&nbsp;</td>
	</tr>
<%
'sqlStr = "SET DATEFORMAT DMY; Select meeting_date,meeting_id,meeting_content,start_time,end_time,meeting_status, "&_
'" company_name, contact_name from meetings_view where ORGANIZATION_ID = "& OrgID &_
'" AND participant_id = " & participant_id & " And meeting_date = '" & dtCurrentDate & "'" &_
'" Order BY start_time, end_time"
sqlstr = "exec dbo.get_meetings '" & users_arr(0,i) & "','" & OrgID & "','" & dtCurrentDate & "'"
'Response.Write sqlStr
'Response.End
set rs_meetings = con.GetRecordSet(sqlStr)
if not rs_meetings.EOF then
 while not rs_meetings.EOF
	meetingID = rs_meetings(0)
	meetingDate = rs_meetings(1) 
	startTime = rs_meetings(2)
	endTime = rs_meetings(3)			
	company_name = rs_meetings(4)
	contact_name = rs_meetings(5)
	status = rs_meetings(6)
	meeting_desc = trim(rs_meetings(11))
	If Len(meeting_desc) > 50 Then
		meeting_desc_short = left(meeting_desc,50) & "..."
	Else
		meeting_desc_short = meeting_desc
	End If	
	
	users_names= ""
	sqlstr = "Select FIRSTNAME + ' ' + LASTNAME From meeting_to_users Inner Join Users On Users.User_ID = meeting_to_users.User_ID " &_
	" Where meeting_ID = " & meetingID & " ORDER BY FIRSTNAME + ' ' + LASTNAME"
	set rs_participants = con.getRecordSet(sqlstr)
	if not rs_participants.eof then
		users_names= rs_participants.getString(,,",",",")
	end if
	set rs_participants = nothing
	If Len(users_names) > 1 Then
		users_names= Left(users_names,Len(users_names)-1)
	End If	
	
    If trim(status) = "1" Then
       href = "href=""javascript:addmeeting('" & meetingID & "')"""   
    Else
       href = "href=""javascript:closemeeting('" & meetingID & "')"""     
    End If	
%>
	<tr class="card">
	    <td align=center valign=top>
	    <%If trim(status) = "1" Or trim(AddMeetings) = "1" Then%>
	    <input type=image src="../../images/delete_icon.gif" border=0 hspace=0 vspace=0 onclick="return CheckDelMeeting('<%=meetingID%>','<%=users_arr(0,i)%>','<%=escape(participant_name)%>',this)">
	    <%End If%>
	    </td> 
	    <td align=center valign=top><%If trim(status) <> "4" Then%>
	    <input type=image src="../../images/notok_icon.gif" border=0 hspace=0 vspace=0 title="<%=vFix(arrTitles(30))%>" onclick="return CheckPostpone('<%=meetingID%>')">
	    <%Else%><input type=image src="../../images/ok_icon.gif" border=0 hspace=0 vspace=0 title="<%=vFix(arrTitles(31))%>" onclick="return CheckMake('<%=meetingID%>')"><%End If%>
	    </td> 	            	    
		<td align="<%=align_var%>" valign="top" ><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>"><%=users_names%></a></td>	    
		<td align="<%=align_var%>" dir=<%=dir_obj_var%> valign="top"><a class="link_categ" <%=href%> title="<%=vFix(meeting_desc)%>"><%=meeting_desc_short%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=company_name%></a></td>	
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=startTime%></a></td>
		<td align="center" valign="top" nowrap dir="<%=dir_obj_var%>"><a class="task_status_num<%=status%>" <%=href%>><%=arr_Status(status)%></a></td>	  
	</tr>	
<%  rs_meetings.moveNext
   Wend  
 Else	
%>   <tr class="card">		
		<td align="center" valign="top" colspan=8><!--לא נמצאו--><%=arrTitles(3)%>&nbsp;</td>	    
	</tr>
<% End If
   set rs_meetings = Nothing    
%>
</table></td></tr>
<%Next
  End If
%>
</table></td>
<td width=170 nowrap valign=top class="td_menu">
<table cellpadding=2 cellspacing=0 width=100%  dir="<%=dir_var%>">
 <tr>    
    <td width="100%" valign="top" align="center">
    <table cellpadding=1 cellspacing=2 width=100% align=center border=0  dir="<%=dir_obj_var%>">
    <FORM NAME="form_search" ACTION="default.asp" METHOD="POST" ID="form_search">
    <tr>		
		<td align=center nowrap>
		  <SELECT NAME="currentMonth" CLASS="norm" ID="currentMonth" dir="<%=dir_var%>" style="width:95" onchange="form_search.submit();">
	      <% For counter = 1 to 12 %>
	        <OPTION VALUE="<%=counter%>" <% If (DatePart("m", dtCurrentDate) = counter) Then Response.Write "SELECTED"%>><%=MonthName(counter)%></OPTION>
	      <% Next %>
	      </SELECT>
	    </td>					
   	
		<td align="center" nowrap>
		<SELECT NAME="currentYear" CLASS="norm" ID="currentYear" dir="<%=dir_var%>" style="width:60" onchange="form_search.submit();">
	      <% For counter = -2 to 2 %>
	        <OPTION VALUE="<%=Year(dtCurrentDate)+counter%>" <% If (DatePart("yyyy", dtCurrentDate) = Year(dtCurrentDate)+counter) Then Response.Write "SELECTED"%>><%=Year(dtCurrentDate)+counter%></OPTION>
	      <% Next %>
	      </SELECT>
		</td>		
	</TR>		
    </table>
    </td>
   </tr>
    <tr><td align="center" colspan=2>
     <TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=160 dir="<%=dir_obj_var%>" class="calendarBody" ID="Table1">
  	 <tr><th colspan=7 align=center style="border-right : 1px solid black"><%=MonthName(Month(dtCurrentDate))%>&nbsp;<%=Year(dtCurrentDate)%></th></tr>
	  <!-- Writring the days of the week for headers -->
        <TR VALIGN="TOP" ALIGN="CENTER" dir=ltr>
        <% For iDay = vbSunday To vbSaturday%>
        <TD WIDTH="14%" class="names"><%=daysname(iDay-1) %></TD>
        <% Next %>
        </TR>
<%   
    For iWeek = 1 To iRows
      Response.Write "<TR VALIGN=TOP>"
      For iDay = 1 To iColumns
	' Checks to see if there is a day this month on the meeting_date being written
	If aCalendarDays((iWeek-1)*7 + iDay) > 0 then
		current_date = DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), aCalendarDays((iWeek-1)*7 + iDay))	  
		strPage = "default.asp?date_=" & current_date & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&participant_id=" & participant_id	  
	 
		meeting_count = 0
		style = "style=""background-color:#F0F0F0; """ 		
		meetings_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>0"
		meetings_str = meetings_str & "<span style='font-size:11px'> "&arrTitles(1)&"</span></A>"		
	  
		sqlStr = "SET DATEFORMAT DMY; Select meeting_date,Count(meeting_id) from meetings_view where ORGANIZATION_ID = "& OrgID
		If trim(participant_id) <> "" Then
			sqlStr = sqlStr & " AND participant_id = " & participant_id
		End If
		sqlStr = sqlStr & " And meeting_date = '" & current_date & "' Group BY meeting_date"
		'Response.Write sqlStr
		'Response.End
		set rs_meetings = con.GetRecordSet(sqlStr)
		if not rs_meetings.EOF then
			meeting_count = rs_meetings(1)
			date_pr = rs_meetings(0) 			
		
			meetings_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>" & meeting_count & "<span style='font-size:11px'> "&arrTitles(1)&"</span></span>"
			if meeting_count < 1 then
				style = "style=""background-color:#F0F0F0; """ 			
			else
				style = "style=""background-color:#53FE01; """ 
			end if
	   end if
	   set rs_meetings = nothing	
	   
	  ' Checks to see if the day being printed is today
	  If DateDiff("d",current_date,dtCurrentDate) = 0 Then
	    Response.Write "<TD CLASS='calCurrentDay'" 
 	  Else
   	    Response.Write "<TD CLASS='calOtherDay'" 
 	  End If	  
 	 
 	  Response.Write " style=""cursor:hand"" onClick=""document.location.href='" & strPage & "'""" 	  
 	  Response.Write style & ">" 	  
 	  Response.Write ("&nbsp;" & aCalendarDays((iWeek-1)*7 + iDay))
 	   	  
	Else 
	  Response.Write ("<TD CLASS='calNotDay'>&nbsp;")
	End IF

	Response.Write "</TD>"
      Next
      Response.Write "</TR>"
    Next  
%>
  </TABLE> </td></tr>
  <%If trim(AddMeetings) = "1" Then%>
  
  <%Else%>
  <input type=hidden name="participant_id" id="participant_id" value="<%=participant_id%>">	
  <%End If%>   
   </form>
   <tr><td width="100%" height=5 nowrap></td></tr>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="#" onclick="addmeeting('')">&nbsp;<!--הוסף פגישה --><%=arrTitles(4)%>&nbsp;</a></td></tr>  
   <tr><td colspan=2 height=5 nowrap></td></tr>
   <%If trim(AddMeetings) = "1" Then%>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="meetings.asp">&nbsp;<!--רשימת הפגישות --><%=arrTitles(16)%>&nbsp;</a></td></tr>  
   <tr><td colspan=2 height=5 nowrap></td></tr>
   <%End If%>
</table>
</td></tr></table>
<div id=div_alert name=div_alert style="display:none;visibility:hidden;position:absolute;left:100;top:300;width:250;height:90;z-index:11;">
<table cellpadding=0 cellspacing=1 border=1 bgcolor="#ffffff" width=100% dir="<%=dir_var%>">
 <tr>
        <td width="100%" class="card align="<%=align_var%>">
        <form id=delete_form name=delete_form action="" method=post>
            <table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">                
            <tr><td colspan=2 height=10 align=center></td></tr>
			<tr>
            <td width="100%" id="titleALL" align="<%=align_var%>" class=card style="padding:5px" dir="<%=dir_obj_var%>">
            </td>
			<td width=20 nowrap align=center class=card> 
			<input type=radio name="delete_meeting" value="0" style="vertical-align:middle">   
            </td>            
            </tr>	
			<tr>					
            <td width="100%" id="titleMe" align="<%=align_var%>" class=card style="padding:5px" dir="<%=dir_obj_var%>">
            </td>
			<td width=20 nowrap align=center class=card> 
			<input type=radio name="delete_meeting" value="" id="user_delete_id" style="vertical-align:middle" checked>   
            </td>	            
            </tr>
            <tr><td colspan=2 height=10 align=center></td></tr>
            <tr>
				<td colspan=2 width="100%" align=center>
				<input type=button value="<%=arrButtons(2)%>" onclick="hideDive()" class="but_menu" style="width:100">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type=submit value="אישור מחיקה" name="delete_button" id="delete_button" class="but_menu" style="width:100">
				</td>
			</tr>
            </table>
         </form>   
        </td>  
  </tr>                     
</table>
</div>
</BODY>
</HTML>
