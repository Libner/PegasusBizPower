<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
date_ = trim(Request("date_"))
current_date = date_

If trim(Request("USER_ID")) <> "" Then
	UserID = trim(Request("USER_ID"))
End If	

If trim(UserID) <> "" Then
	sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & userID
	set rs_user = con.getRecordSet(sqlstr)
	if not rs_user.eof then
		userName = trim(rs_user(0)) & " " & trim(rs_user(1))
	end if
	set rs_user = nothing
End If

if Request("update")<>nil then		
    con.executeQuery("SET DATEFORMAT dmy")
    sqlstr = "Select hour_id FROM hours WHERE ORGANIZATION_ID = "& OrgID &_
	" AND USER_ID = " & UserId & " AND date = '" & date_ & "'"
	'Response.Write sqlstr
	'Response.End
	set rs_hours = con.getRecordSet(sqlstr)
	While not rs_hours.eof
		hour_id = trim(rs_hours(0))
		hours = trim(Request.Form("hours"&hour_id))   
		If trim(hours) <> "" Then
			hours = cDbl(hours)
			minuts = hours * 60		
			con.executeQuery("Update hours Set minuts = " & minuts & " WHERE hour_id = " & hour_id)			
		End If
	 	rs_hours.moveNext
	Wend 
	set rs_hours = nothing	
	Response.Redirect "hours_day.asp?date_=" & date_ & "&start_date=" & start_date & "&end_date=" & end_date & "&User_ID=" & UserID
end if

If Request("delete") <> nil And Request.QueryString("hour_id") <> nil Then
    con.executeQuery("DELETE FROM hours WHERE hour_id = " & Request.QueryString("hour_id"))
    Response.Redirect "hours_day.asp?date_=" & date_ & "&start_date=" & start_date & "&end_date=" & end_date & "&User_ID=" & UserID
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
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 	
		
	function delete_hour(hour_id,date_)
	{
		if(window.confirm("? האם ברצונך למחוק את שעות עבודה על הפרויקט הזאת") == true)
		{
		    document.location.href = "hours.asp?delete=1&date_=" + date_ + "&hour_id=" + hour_id + "&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>"; 

		}
		return true;
	}
<!--End-->
</script> 
</head> 
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table7">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title" dir=rtl>&nbsp;שעות העבודה שלי&nbsp;-&nbsp;<font color="#6F6DA6"><%=userName%></font>&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>         
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center>
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
	<tr>	
	<td width="100%" align="center" valign="top" class="title_sort">פרויקט / ארגון</td>
	<td width="80" nowrap align="center" class="title_sort">שעות עבודה</td>
	<td width="100" nowrap align="center" valign="top" class="title_sort" colspan=2>תאריך</td>
	</tr>
<%
Function FormatMediumDateShort(DateValue)
	Dim strYYYY
	Dim strMM
	Dim strDD

		strYYYY = CStr(DatePart("yyyy", DateValue))
		If Len(strYYYY) > 1 Then
			strYY = Right(strYYYY,2)
		Else
			strYY = strYYYY
		End If

		strMM = CStr(DatePart("m", DateValue))
		If Len(strMM) = 1 Then strMM = "0" & strMM

		strDD = CStr(DatePart("d", DateValue))
		If Len(strDD) = 1 Then strDD = "0" & strDD

		FormatMediumDateShort = strDD & "/" & strMM & "/" & strYY

End Function 
daysname = array(" 'א"," 'ב"," 'ג"," 'ד"," 'ה"," 'ו","שבת")

con.executeQuery("SET DATEFORMAT dmy")
sqlStr = "Select date,Sum(minuts) from hours_view where ORGANIZATION_ID = "& OrgID &_
" AND USER_ID = " & UserId & " And date = '" & current_date & "' Group BY date"
'Response.Write sqlStr
'Response.End
set rs_PROJECTS = con.GetRecordSet(sqlStr)
if not rs_PROJECTS.EOF then
	is_projects = true	
	minuts = rs_PROJECTS(1)
	date_pr = rs_PROJECTS(0) 
	day_num = DatePart("w",current_date)
	day_name = daysname(day_num-1)

	projects_names = ""
	If IsDate(date_pr) Then
		sqlstr = "Select project_name, company_name From hours_view WHERE ORGANIZATION_ID = "& OrgID &_
		" AND USER_ID = " & UserId & " AND date = '" & date_pr & "'"
		set rs_pr = con.getRecordSet(sqlstr)
		if not rs_pr.eof then
			projects_names = rs_pr.getString(,," / "," , ")
		end if
		set rs_pr = nothing
		projects_names = Left(projects_names, Len(projects_names)-2)
	End If
	hours = Round((minuts / 60),1)
	if hours < 2 then
		style = "style=""background-color:red""" 
	elseif (hours < 6 and hours >= 2) then
		style = "style=""background-color:yellow""" 
	elseif (hours >= 6) then
		style = "style=""background-color:#ACE920""" 
	end if
%>
	<tr name="title<%=rownum%>" id="title<%=rownum%>"  class="card3">		
		<td align="right" valign="top" ><%=projects_names%>&nbsp;</td>	    
		<td align="center" valign="top"><%=hours%></td>
		<td align="right" width=65 nowrap valign="top" <%=style%>><%=FormatMediumDateShort(date_pr)%>&nbsp;</td>
		<td align="right" width=35 nowrap valign="top" <%=style%>><%=day_name%>&nbsp;</td>
	</tr>
	
<% Else
	style =  "style=""background-color:red""" 
	day_num = DatePart("w",current_date)
	day_name = daysname(day_num-1)
%>
    <tr class="card3">		
		<td align="right" valign="top" >לא נמצאו פרויקטים&nbsp;</td>	    
		<td align="center" valign="top" >0</td>
		<td align="right" width=65 nowrap valign="top" <%=style%> ><%=FormatMediumDateShort(current_date)%>&nbsp;</td>
		<td align="right" width=35 nowrap valign="top" <%=style%> ><%=day_name%>&nbsp;</td>
	</tr>
<% End If
   set rs_PROJECTS = Nothing   
%>
<form name="form<%=current_date%>" id="form<%=current_date%>" method=post action="hours_day.asp?update=1&User_ID=<%=UserID%>">	
<%       
   sqlstr = "Select project_name,company_name, minuts,hour_id from hours_view Where ORGANIZATION_ID = "& OrgID &_
   " AND USER_ID = " & UserId & " AND date = '" & current_date & "'"
   set rs_project = con.getRecordset(sqlstr)
   if not rs_project.eof then
	 'Response.Write date_
%>
  
  <tr>
  <input type=hidden name="date_" id="date_" value="<%=current_date%>">	    
  <td class="title_sort1" align=center width=100% nowrap>פרויקט / ארגון</td>
  <td class="title_sort1" align=center width="78" nowrap>שעות</td>
  <td class="title_sort1" align=center width="98" nowrap colspan=2>&nbsp;</td>	
  </tr>
<%	
  while not rs_project.eof
	company_name = rs_project(1)	
	project_name = rs_project(0)
	minuts_pr = rs_project(2)
	hour_id = rs_project(3)
	hours_pr = Round((minuts_pr / 60),1)
%>
	<tr>				
	<td class="card1" align=right>&nbsp;<%=project_name%>&nbsp;/&nbsp;<%=company_name%>&nbsp;</td>	
	<td class="card1" align=center><input type=text class="texts" name="hours<%=hour_id%>" id="hours<%=hour_id%>" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=4></td>	
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
	<input type=button class="but_menu1" style="width:145" value="הוספת שעות עבודה" onclick="window.open('addhour.asp?date=<%=current_date%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>','','left=200,top=200,width=400,height=200')"></td>
	<td class="card1" align=center><%If is_projects Then%><input type=submit class="but_menu1" value="עדכן" style="width:50" ID="Submit1" NAME="Submit1"><%End if%>
	<td class="card1" colspan=2>&nbsp;</td>
	</td>
	</tr>
	</form>	

</table>
</td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=2 cellspacing=1 width=100% ID="Table1">
<tr><td align="right" colspan=2 class="title_search">תפריט</td></tr>
 <tr><td width="100%" height=10 nowrap></td></tr>
 <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="hours.asp?User_ID=<%=UserID%>&date_=<%=current_date%>&start_date=<%=DateAdd("d",-14,current_date)%>&end_date=<%=current_date%>">&nbsp;תצוגת רשימה&nbsp;</a></td></tr>
 <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="hours_calendar.asp?User_ID=<%=UserID%>&currentMonth=<%=Month(current_date)%>&currentYear=<%=Year(current_date)%>">&nbsp;תצוגה קלנדרית&nbsp;</a></td></tr>
 <tr><td width="100%" height=10 nowrap></td></tr>
</table>
</td>
</tr></table>
</td></tr></table>
</body>
</html>
<%
set con = nothing
%>