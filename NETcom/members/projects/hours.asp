<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  Session.LCID = 1037 

	date_ = trim(Request("date_"))
	end_date = date()
	start_date = DateAdd("d",-13,date())
	If Request("end_date") <> nil And Request("start_date") <> nil Then
		end_date = Request("end_date")
		start_date = Request("start_date")
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

if Request("update")<>nil then		
    sqlstr = "SET DATEFORMAT DMY; Select hour_id FROM hours WHERE ORGANIZATION_ID = "& OrgID &_
	" AND USER_ID = " & USER_ID & " AND date = '" & date_ & "'"
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
	Response.Redirect "hours.asp?date_=" & date_ & "&start_date=" & start_date & "&end_date=" & end_date & "&User_ID=" & USER_ID
end if

If Request("delete") <> nil And Request.QueryString("hour_id") <> nil Then
    con.executeQuery("DELETE FROM hours WHERE hour_id = " & Request.QueryString("hour_id"))
    Response.Redirect "hours.asp?date_=" & date_ & "&start_date=" & start_date & "&end_date=" & end_date & "&User_ID=" & USER_ID
End If

If trim(date_) = "" Then
	date_ = end_date
End If	

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
	
	function openDetails(date_tbody,tr_obj)
	{
		if(window.document.all(""+date_tbody+"").style.display == "none")
		{
			window.document.all(""+date_tbody+"").style.display = "inline"; 
			tr_obj.className = "card3 hand";
		}	
		else
		{
			window.document.all(""+date_tbody+"").style.display = "none"; 
			tr_obj.className = "card hand";	
		}	
	}
	
	function delete_hour(hour_id,date_)
	{
		if(window.confirm("? האם ברצונך למחוק את שעות עבודה על הפרויקט") == true)
		{
		    document.location.href = "hours.asp?delete=1&date_=" + date_ + "&hour_id=" + hour_id + "&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=USER_ID%>"; 

		}
		return true;
	}
<!--End-->
</script> 
</head> 
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 0%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title" dir=rtl>&nbsp;שעות העבודה שלי&nbsp;-&nbsp;<font color="#6F6DA6"><%=userName%></font>&nbsp;</td></tr>		   
   </table></td></tr>    
<%If found_user Then%>        
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center>
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
	<tr>	
	<td width="100%" align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> / <%=trim(Request.Cookies("bizpegasus")("Projectone"))%> / מנגנון</td>
	<td width="80" nowrap align="center" class="title_sort">שעות עבודה</td>
	<td width="100" nowrap align="center" valign="top" class="title_sort" colspan=2>תאריך</td>
	</tr>
<%
current_date = end_date
daysname = array(" 'א"," 'ב"," 'ג"," 'ד"," 'ה"," 'ו","שבת")
rownum = 0
While DateDiff("d",current_date,start_date) <= 0
is_projects = false
rownum = rownum + 1
sqlStr = "SET DATEFORMAT DMY; Select date,Sum(minuts) from hours_view where ORGANIZATION_ID = "& OrgID &_
" AND USER_ID = " & USER_ID & " And date = '" & current_date & "' Group BY date"
'Response.Write sqlStr
'Response.End
set rs_PROJECTS = con.GetRecordSet(sqlStr)
if not rs_PROJECTS.EOF then
	is_projects = true
	minuts = rs_PROJECTS(1)
	date_pr = rs_PROJECTS(0) 
	day_num = DatePart("w",date_pr)
	day_name = daysname(day_num-1)

	projects_names = ""
	If IsDate(date_pr) Then
		sqlstr = "Select company_name,project_name,mechanism_name From hours_view WHERE ORGANIZATION_ID = "& OrgID &_
		" AND USER_ID = " & USER_ID & " AND date = '" & date_pr & "'"
		set rs_pr = con.getRecordSet(sqlstr)
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
	<tr name="title<%=rownum%>" id="title<%=rownum%>" <%If DateDiff("d", date_ ,date_pr) = 0 Then%> class="card3 hand" <%Else%> class="card hand" <%End If%>>		
		<td align="right" valign="top"  onClick="openDetails('<%=date_pr%>',title<%=rownum%>)" dir="<%=dir_obj_var%>"><%=projects_names%>&nbsp;</td>	    
		<td align="center" valign="top" onClick="openDetails('<%=date_pr%>',title<%=rownum%>)"><%=hours%></td>
		<td align="right" width=65 nowrap valign="top" <%=style%> onClick="openDetails('<%=date_pr%>',title<%=rownum%>)"><%=FormatMediumDateShort(date_pr)%>&nbsp;</td>
		<td align="right" width=35 nowrap valign="top" <%=style%> onClick="openDetails('<%=date_pr%>',title<%=rownum%>)"><%=day_name%>&nbsp;</td>
		
	</tr>
	
<% Else
	style =  "style=""background-color:#FF5959""" 
	day_num = DatePart("w",current_date)
	day_name = daysname(day_num-1)
%>
    <tr name="title<%=rownum%>" id="title<%=rownum%>" <%If DateDiff("d", date_ ,current_date) = 0 Then%> class="card3 hand" <%Else%> class="card hand" <%End If%>>		
		<td align="right" valign="top" onClick="openDetails('<%=current_date%>',title<%=rownum%>)">לא נמצאו <%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))	%>&nbsp;</td>	    
		<td align="center" valign="top" onClick="openDetails('<%=current_date%>',title<%=rownum%>)">0</td>
		<td align="right" width=65 nowrap valign="top" <%=style%> onClick="openDetails('<%=current_date%>',title<%=rownum%>)"><%=FormatMediumDateShort(current_date)%>&nbsp;</td>
		<td align="right" width=35 nowrap valign="top" <%=style%> onClick="openDetails('<%=current_date%>',title<%=rownum%>)"><%=day_name%>&nbsp;</td>
	</tr>
<% End If
   set rs_PROJECTS = Nothing   
%>
<tbody name="<%=current_date%>" id="<%=current_date%>" <%If DateDiff("d", date_ ,current_date) <> 0 Then%> style="display:none" <%End If%>>
<form name="form<%=current_date%>" id="form<%=current_date%>" method=post action="hours.asp?update=1&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=USER_ID%>">	
<%       
   sqlstr = "Select project_name,company_name,mechanism_name,minuts,hour_id from hours_view Where ORGANIZATION_ID = "& OrgID &_
   " AND USER_ID = " & USER_ID & " AND date = '" & current_date & "'"
   set rs_project = con.getRecordset(sqlstr)
   if not rs_project.eof then
%>  
  <tr>
  <input type=hidden name="date_" id="date_" value="<%=current_date%>">	    
  <td class="title_sort1" align=center width=100% nowrap><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> / <%=trim(Request.Cookies("bizpegasus")("Projectone"))%> / מנגנון</td>
  <td class="title_sort1" align=center width="78" nowrap>שעות</td>
  <td class="title_sort1" align=center width="98" nowrap colspan=2>&nbsp;</td>	
  </tr>
<%	
  while not rs_project.eof
	company_name = trim(rs_project(1))
	project_name = trim(rs_project(0))
	mechanism_name = trim(rs_project(2))
	minuts_pr = trim(rs_project(3))
	hour_id = trim(rs_project(4))
	hours_pr = GetTime(minuts_pr)
%>
	<tr>				
	<td class="card1" align="<%=align_var%>" dir="<%=dir_obj_var%>"><span dir=rtl>&nbsp;<%=company_name%></span>
	<span dir=ltr>&nbsp;/&nbsp;</span><span dir=rtl><font color="#060165"><%=project_name%></font></span><span dir=ltr>&nbsp;/&nbsp;</span>
	<span dir=rtl><font color="#DF0000"><%=mechanism_name%></font>&nbsp;</span></td>	
	<td class="card1" align=center><input type=text class="texts" name="hours<%=hour_id%>" id="hours<%=hour_id%>" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=5></td>	
	<td class="card1" align=center colspan=2><input type=button class="button_delete hand" onclick="return delete_hour('<%=hour_id%>','<%=current_date%>')" value="מחק"></td>
	</tr>	
	<%
		rs_project.moveNext
		wend
		end if
		set rs_project = nothing
	%>
	<tr>		
	<td class="card1" align=center>
	<input type=button class="but_menu1" style="width:145" value="הוספת שעות עבודה" onclick="window.open('addhour.asp?date=<%=current_date%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=USER_ID%>','','left=200,top=200,width=400,height=250')"></td>
	<td class="card1" align=center><%If is_projects Then%><input type=submit class="but_menu1" value="עדכן" style="width:50" ID="Submit1" NAME="Submit1"><%End if%>
	<td class="card1" colspan=2>&nbsp;</td>
	</td>
	</tr>
	</form>	
</tbody>
<% 
  current_date = DateAdd("d",-1,current_date)
  Wend
%>
</table>
</td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=2 cellspacing=1 width=100% ID="Table1">
<FORM action="hours.asp?search=1&User_ID=<%=USER_ID%>" method=POST id=form_search name=form_search target="_self">
 <tr>    
    <td width="100%" valign="top" align="center">
    <table cellpadding=1 cellspacing=2 width=100% align=center border=0>
    <tr><td align="right" colspan=2><b>מתאריך</b>&nbsp;</td></tr>
    <tr>
		<!--td width=20 nowrap>
		<input type=image src="../../images/delete_icon.gif" border=0 onClick="start_date.value='';return false;" title="מחק תאריך" hspace=0 vspace=0 id=image2 name=image2></td-->
		<td align=center nowrap>
		<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70"onclick="return popupcal(this);" readonly></td>					
	</tr>
	<tr><td width=100% align=right colspan=2>&nbsp;<b>עד תאריך</b>&nbsp;</td></tr>
     <TR>
    	<!--td align="right"><input type=image hspace=0 vspace=0 src="../../images/delete_icon.gif" border=0 onClick="end_date.value='';return false;" title="מחק תאריך" id=image1 name=image1></td-->
		<td align="center" nowrap>
		<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return popupcal(this);" readonly></td>		
	</TR>		
    </table>
    </td>
   </tr> 
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="#" onclick="document.form_search.submit();">&nbsp;עדכן תאריך&nbsp;</a></td></tr>
   </form>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="hours_calendar.asp?User_ID=<%=USER_ID%>">&nbsp;תצוגה קלנדרית&nbsp;</a></td></tr>
   <tr><td width="100%" height=10 nowrap></td></tr>
</table>
</td>
</tr></table></td>
<%End If%>
</tr></table>
</body>
</html>
<%set con = nothing%>