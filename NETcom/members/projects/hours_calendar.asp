<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 
If trim(lang_id) = "1" Then
	Session.LCID = 1037 
Else
	Session.LCID = 2057
End If	

If trim(Request("USER_ID")) <> "" Then
	USER_ID = trim(Request("USER_ID"))
Else
	USER_ID = UserID	
End If	

found_user = false
If trim(USER_ID) <> "" Then
	sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & USER_ID & " And ORGANIZATION_ID = " & OrgID
	set rs_user = con.getRecordSet(sqlstr)
	if not rs_user.eof then
		userName = trim(rs_user(0)) & " " & trim(rs_user(1))
		found_user = true
	else
		found_user = false	
	end if
	set rs_user = nothing
End If

' ---------- Page Variables ----------
    Const intCharToShow = 19		' The number of characters shown in each day
    Const bolEditable   = True		' If the calendar is editable or not (Can be tied into password verification)

    Dim dtToday 			' Today's Date
    Dim dtCurrentDate			' The current date
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
	
	If Request("currentMonth") = "" And Request("currentYear") = ""  Then
		currentMonth = Month(dtToday)
		currentYear = Year(dtToday)
	Else
		currentMonth = Request("currentMonth")
		currentYear = Request("currentYear")	
	End If

    If currentMonth <> "" And currentYear <> "" And Request("dtCurrentDate") = ""  Then
		dtCurrentDate = "1" & "/" & currentMonth & "/" &  currentYear
    ElseIf Request("dtCurrentDate") <> "" Then
		dtCurrentDate = Request("dtCurrentDate")
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
    
	If Request("update")<>nil Then		
		sqlstr = "SET DATEFORMAT DMY; Select hour_id FROM hours WHERE ORGANIZATION_ID = "& OrgID &_
		" AND USER_ID = " & USER_ID & " AND date = '" & dtCurrentDate & "'"
		'Response.Write sqlstr
		'Response.End
		set rs_hours = con.getRecordSet(sqlstr)
		While not rs_hours.eof
			hour_id = trim(rs_hours(0))
			hours = trim(Request.Form("hours"&hour_id))   
			If trim(hours) <> "" Then
				arrhours = split(hours,":")
				if arrhours(0) <> "" then
					minuts = cdbl(arrhours(0)) * 60
				else
					minuts = 0
				end if	
				if ubound(arrhours) > 0 then
					if arrhours(1) <> "" then
						minuts = minuts + cdbl(arrhours(1))
					end if	
				end if	
				con.executeQuery("Update hours Set minuts = " & minuts & " WHERE hour_id = " & hour_id)			
			End If
	 		rs_hours.moveNext
		Wend 
		set rs_hours = nothing	
		Response.Redirect "hours_calendar.asp?dtCurrentDate=" & dtCurrentDate & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&User_ID=" & USER_ID
	end if

	If Request("delete") <> nil And Request.QueryString("hour_id") <> nil Then
		con.executeQuery("DELETE FROM hours WHERE hour_id = " & Request.QueryString("hour_id"))
		Response.Redirect "hours_calendar.asp?dtCurrentDate=" & dtCurrentDate & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&User_ID=" & USER_ID
	End If

%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 52 Order By word_id"				
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
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 58 ;
	} 	
		
	function delete_hour(hour_id,dtCurrentDate)
	{
	 <%
		If trim(lang_id) = "1" Then
			str_confirm = "? האם ברצונך למחוק את שעות עבודה על ה" & trim(Request.Cookies("bizpegasus")("Projectone"))
		Else
			str_confirm = "Are you sure want to delete the labor hours on the " & trim(Request.Cookies("bizpegasus")("Projectone")) & " ?"
		End If   
     %>		
		if(window.confirm("<%=str_confirm%>") == true)
		{
		    document.location.href = "hours_calendar.asp?delete=1&dtCurrentDate=" + dtCurrentDate + "&hour_id=" + hour_id + '&currentMonth=<%=currentMonth%>&currentYear=<%=currentYear%>&User_ID=<%=USER_ID%>'; 

		}
		return true;
	}

//-->
</script>
<STYLE TYPE="text/css">
    <!--   
    .blackBacking   {background-color: #000000;}
    .names 	    {background-color: #CCCCCC; font-size: 13px; color: #000000; text-decoration: none; text-align:  center;  font-weight: bold; border-top : 1px solid black; border-right : 1px solid black; line-height: 150%}
    .calendarBody   {background-color: #F0F0F0; font-size: 12px; color: #000000; text-decoration: none; text-align:  center; border : 1px solid black; border-right : none}
    .calCurrentDay  {background-color: #F0F0F0; font-size: 12px; color: #000000; line-height : 150%; 
    text-align:  center; font-weight:bold; border: ridge #FF5959 5px}
    .calOtherDay    {background-color: #F0F0F0; font-size: 12px; color: #000000; text-align:center; line-height : 150%; font-weight:bold; border-top : 1px solid black; border-right : 1px solid black}
    .calNotDay	    {background-color: #F0F0F0; font-size: 12px; color: #000000; line-height : 150%; text-align:  center; border-top : 1px solid black; border-right : 1px solid black}
    .calFormMenu    {background-color: #4C5D87; font-size: 13px; color: #FFFFFF; text-decoration: none; text-align:  center;  font-weight: bold; border : 1px solid black}
    -->
  </STYLE>
</HEAD>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 0%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td class="page_title" dir="<%=dir_obj_var%>"><font color="#6F6DA6"><%=userName%></font>&nbsp;</td></tr>
<%If found_user Then%>		   		       	
<tr><td width=100%>
 <table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>"> 
 <tr>
 <td align=center width="100%" nowrap> 
 <table width=100% border="0" width="100%" align=center dir="<%=dir_var%>">
  <tr><td>
  <TABLE CELLPADDING=0 CELLSPACING=0 WIDTH=600 BORDER=0 align=center>
    <TR>
      <TD>
	<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=600 dir="<%=dir_obj_var%>" class="calendarBody">
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
	' Checks to see if there is a day this month on the date being written
	If aCalendarDays((iWeek-1)*7 + iDay) > 0 then
	  current_date = DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), aCalendarDays((iWeek-1)*7 + iDay))
	  
	  strPage = "hours_calendar.asp?dtCurrentDate=" & current_date & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&User_ID=" & USER_ID
	  
	  If DateDiff("d",current_date,date()) >= 0 Then
		hours = 0
		style = "style=""background-color:#FF5959; """ 		
		hours_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>0"
		If trim(lang_id) = "1" Then
			hours_str = hours_str & "<span style='font-size:11px'> שעות</span></A>"
		Else
			hours_str = hours_str & "<span style='font-size:11px'> Hours</span></A>"
		End If
	  Else
		hours = ""
		style = "" 	  
		hours_str = "&nbsp;"
	  End If  
	  
	  is_projects = false
	  
	  sqlStr = "SET DATEFORMAT DMY; Select date,Sum(minuts) from hours_view where ORGANIZATION_ID = "& OrgID &_
	  " AND USER_ID = " & USER_ID & " And date = '" & current_date & "' Group BY date"
	  'Response.Write sqlStr
	  'Response.End
	  set rs_PROJECTS = con.GetRecordSet(sqlStr)
	  if not rs_PROJECTS.EOF then
	    is_projects = true		
		minuts = rs_PROJECTS(1)
		date_pr = rs_PROJECTS(0) 			
		hours = GetTime(minuts)
		hours_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>" & hours
		If trim(lang_id) = "1" Then
			hours_str = hours_str & "<span style='font-size:11px'> שעות</span></A>"
		Else
			hours_str = hours_str & "<span style='font-size:11px'> Hours</span></A>"
		End If
		if minuts < 120 then
			style = "style=""background-color:#FF5959; """ 
		elseif (minuts < 360 and minuts >= 120) then
			style = "style=""background-color:yellow; """ 
		elseif (minuts >= 360) then
			style = "style=""background-color:#ACE920; """ 
		end if
	end if
	set rs_PROJECTS = nothing	
 ' Checks to see if the day being printed is today
	  If DateDiff("d",current_date,dtCurrentDate) = 0 Then
	    Response.Write "<TD CLASS='calCurrentDay'" 
 	  Else
   	    Response.Write "<TD CLASS='calOtherDay'" 
 	  End If
 	  
 	  If DateDiff("d",date(),current_date) <=0 Then
 		Response.Write " style=""cursor:hand"" onClick=""document.location.href='" & strPage & "'"""
 	  End If 
 	  
 	  Response.Write style & ">" 
	  
 	  Response.Write ("&nbsp;" & aCalendarDays((iWeek-1)*7 + iDay) & "<BR>" & hours_str)
 	  
	Else 
	  Response.Write ("<TD CLASS='calNotDay'>&nbsp;")
	End IF

	Response.Write "</TD>"
      Next
      Response.Write "</TR>"
    Next
 
%>

	</TABLE>
      </TD>
    </TR>
   </table></td></tr> 
   <tr><td height=10 nowrap></td></tr> 
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>">
	<tr>	
	<td width="100%" align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> / <%=trim(Request.Cookies("bizpegasus")("Projectone"))%> / <%=arrTitles(14)%></td>
	<td width="80" nowrap align="center" class="title_sort"><!--שעות עבודה--><%=arrTitles(3)%></td>
	<td width="100" nowrap align="center" valign="top" class="title_sort" colspan=2><!--תאריך--><%=arrTitles(4)%></td>
	</tr>
<%

sqlStr = "SET DATEFORMAT DMY; Select date,Sum(minuts) from hours_view where ORGANIZATION_ID = "& OrgID &_
" AND USER_ID = " & USER_ID & " And date = '" & dtCurrentDate & "' Group BY date"
'Response.Write sqlStr
'Response.End
set rs_PROJECTS = con.GetRecordSet(sqlStr)
if not rs_PROJECTS.EOF then
	is_projects = true	
	minuts = rs_PROJECTS(1)
	date_pr = rs_PROJECTS(0) 
	day_num = DatePart("w",dtCurrentDate)
	day_name = daysname(day_num-1)

	projects_names = ""
	If IsDate(date_pr) Then
		sqlstr = "Select company_name,project_name,mechanism_name From hours_view WHERE ORGANIZATION_ID = "& OrgID &_
		" AND USER_ID = " & USER_ID & " AND date = '" & date_pr & "'"
		set rs_pr = con.getRecordSet(sqlstr)
		projects_names = ""
		while not rs_pr.eof
			projects_names = projects_names & "<span dir=rtl><font color=""#DF0000"">" & rs_pr(2) & "</font></span><span dir=ltr>&nbsp;/&nbsp;</span><span dir=rtl><font color=""#060165""> " & rs_pr(1) & "</font></span><span dir=ltr>&nbsp;/&nbsp;</span><span dir=rtl>" &  rs_pr(0) & "</span>" & ", "
			rs_pr.moveNext
		wend
		set rs_pr = nothing
		projects_names = Left(projects_names, Len(projects_names)-2)
	End If
	
	hours = GetTime(minuts)	
	if minuts < 120 then
		style = "style=""background-color:#FF5959""" 
	elseif (minuts < 366 and minuts >= 120) then
		style = "style=""background-color:yellow""" 
	elseif (minuts >= 360) then
		style = "style=""background-color:#ACE920""" 
	end if
%>
	<tr name="title<%=rownum%>" id="title<%=rownum%>"  class="card3">		
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><%=projects_names%></td>	    
		<td align="center" valign="top"><%=hours%></td>
		<td align="<%=align_var%>" width=65 nowrap valign="top" <%=style%>><%=FormatMediumDateShort(date_pr)%>&nbsp;</td>
		<td align="<%=align_var%>" width=35 nowrap valign="top" <%=style%>><%=day_name%>&nbsp;</td>
	</tr>
	
<% Else
	style =  "style=""background-color:#FF5959""" 
	day_num = DatePart("w",dtCurrentDate)
	day_name = daysname(day_num-1)
%>
    <tr class="card3">		
		<td align="<%=align_var%>" valign="top" ><span id=word5 name=word5><!--לא נמצאו--><%=arrTitles(5)%></span> <%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))	%>&nbsp;</td>	    
		<td align="center" valign="top" >0</td>
		<td align="center" width=65 nowrap valign="top" <%=style%> ><%=FormatMediumDateShort(dtCurrentDate)%>&nbsp;</td>
		<td align="center" width=35 nowrap valign="top" <%=style%> ><%=day_name%>&nbsp;</td>
	</tr>
<% End If
   set rs_PROJECTS = Nothing 
   is_projects = false  
%>
<form name="form<%=dtCurrentDate%>" id="form<%=dtCurrentDate%>" method=post action="hours_calendar.asp?update=1&User_ID=<%=USER_ID%>">	
<%       
   sqlstr = "Select project_name,company_name,mechanism_name,minuts,hour_id from hours_view Where ORGANIZATION_ID = "& OrgID &_
   " AND USER_ID = " & USER_ID & " AND date = '" & dtCurrentDate & "'"
   set rs_project = con.getRecordset(sqlstr)
   if not rs_project.eof then
	is_projects = true
%>
  
  <tr>
  <input type=hidden name="dtCurrentDate" id="dtCurrentDate" value="<%=dtCurrentDate%>">	
  <input type=hidden name="currentMonth"  value="<%=currentMonth%>">	
  <input type=hidden name="currentYear" value="<%=currentYear%>">	  
  <td class="title_sort1" align=center width=100% nowrap><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> / <%=trim(Request.Cookies("bizpegasus")("Projectone"))%> / <%=arrTitles(14)%></td>
  <td class="title_sort1" align=center width="78" nowrap><span id=word6 name=word6><!--שעות--><%=arrTitles(6)%></span></td>
  <td class="title_sort1" align=center width="98" nowrap colspan=2>&nbsp;</td>	
  </tr>
<%	
  while not rs_project.eof
	company_name = rs_project(1)	
	project_name = rs_project(0)
	mechanism_name = rs_project(2)
	minuts_pr = rs_project(3)
	hour_id = rs_project(4)
	hours_pr = GetTime(minuts_pr)
%>
	<tr>				
	<td class="card1" align="<%=align_var%>" dir="<%=dir_obj_var%>"><span dir=rtl>&nbsp;<%=company_name%></span>
	<span dir=ltr>&nbsp;/&nbsp;</span><span dir=rtl><font color="#060165"><%=project_name%></font></span><span dir=ltr>&nbsp;/&nbsp;</span>
	<span dir=rtl><font color="#DF0000"><%=mechanism_name%></font>&nbsp;</span></td>	
	<td class="card1" align=center><input type=text class="texts" name="hours<%=hour_id%>" id="hours<%=hour_id%>" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=5></td>	
	<td class="card1" align=center colspan=2><A class="button_delete hand" href="#" onclick="return delete_hour('<%=hour_id%>','<%=dtCurrentDate%>')">&nbsp;<span id=word7 name=word7><!--מחק--><%=arrTitles(7)%></span>&nbsp;</a></td>
	</tr>	
	<%
		rs_project.moveNext
		wend
		end if
		set rs_project = nothing
	%>
	<tr>		
	<td class="card1" align=center>
	<A class="but_menu1" nohref style="width:145" onclick="window.open('addhour.asp?date=<%=dtCurrentDate%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=USER_ID%>','','left=200,top=200,width=400,height=250')"><span id="word9" name=word9><!--הוספת שעות עבודה--><%=arrTitles(9)%></span></a></td>
	<td class="card1" align=center><%If is_projects Then%><A class="but_menu1" style="width:50" onclick='document.all("form<%=dtCurrentDate%>").submit()'><span id=word8 name=word8><!--עדכן--><%=arrTitles(8)%></span></a><%End if%>
	<td class="card1" colspan=2>&nbsp;</td>
	</td>
	</tr>
	</form>	

</table>
  </tr>  
 </table></td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=2 cellspacing=0 width=100%  dir="<%=dir_var%>">
 <tr>    
    <td width="100%" valign="top" align="center">
    <table cellpadding=1 cellspacing=2 width=100% align=center border=0  dir="<%=dir_var%>">
    <FORM NAME="form_search" ACTION="hours_calendar.asp?USER_ID=<%=USER_ID%>" METHOD="POST" ID="form_search">
    <tr><td align="<%=align_var%>" colspan=2><b><span id="word10" name=word10><!--חודש--><%=arrTitles(10)%></span></b>&nbsp;</td></tr>
    <tr>		
		<td align=center nowrap>
		  <SELECT NAME="currentMonth" CLASS="norm" ID="currentMonth" dir="<%=dir_obj_var%>" style="width:90">
	      <% For counter = 1 to 12 %>
	        <OPTION VALUE="<%=counter%>" <% If (DatePart("m", dtCurrentDate) = counter) Then Response.Write "SELECTED"%>><%=MonthName(counter)%></OPTION>
	      <% Next %>
	      </SELECT>
	    </td>					
	</tr>
	<tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b><span id="word11" name=word11><!--שנה--><%=arrTitles(11)%></span></b>&nbsp;</td></tr>
     <TR>    	
		<td align="center" nowrap>
		<SELECT NAME="currentYear" CLASS="norm" ID="currentYear" dir="<%=dir_obj_var%>" style="width:90">
	      <% For counter = -2 to 2 %>
	        <OPTION VALUE="<%=Year(dtCurrentDate)+counter%>" <% If (DatePart("yyyy", dtCurrentDate) = Year(dtCurrentDate)+counter) Then Response.Write "SELECTED"%>><%=Year(dtCurrentDate)+counter%></OPTION>
	      <% Next %>
	      </SELECT>
		</td>		
	</TR>		
    </table>
    </td>
   </tr>  
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="#" onclick="document.form_search.submit();">&nbsp;<span id="word12" name=word12><!--עדכן תאריך--><%=arrTitles(12)%></span>&nbsp;</a></td></tr>  
   </form>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="hours.asp?User_ID=<%=USER_ID%>">&nbsp;<span id="word13" name=word13><!--תצוגת רשימה--><%=arrTitles(13)%></span>&nbsp;</a></td></tr>
   <tr><td width="100%" height=10 nowrap></td></tr>
</table>
</td>
<%End If%>
</tr></table>
</BODY>
</HTML>
