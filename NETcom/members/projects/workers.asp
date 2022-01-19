<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	if trim(Request.Form("search_worker")) <> "" Or trim(Request.QueryString("search_worker")) <> "" then
		search_worker = trim(Request.Form("search_worker"))
		if trim(Request.QueryString("search_worker")) <> "" then
			search_worker = trim(Request.QueryString("search_worker"))
		end if					
		where_worker = " And FIRSTNAME + ' ' + LASTNAME LIKE '%"& sFix(search_worker) &"%'"
		search = true			
	else
		where_worker = ""			
	end if
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 51 Order By word_id"				
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
function CheckDel(str) {
 	 <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את העובד"
     Else
		str_confirm = "Are you sure want to delete the employee ?"
     End If   
     %>		
	 return window.confirm("<%=str_confirm%>");
}
//-->
</script> 
</head> 
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width=100% dir="<%=dir_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 3%>
<%numOfLink = 0%>
<%If trim(wizard_id) = "" Then%>
<tr><td width=100% dir="<%=dir_var%>">
<!--#include file="../../top_in.asp"-->
</td></tr>
<%End If%>
<tr><td width=100% dir="<%=dir_var%>">
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width="100%" class="page_title" dir="<%=dir_var%>" colspan=2>&nbsp;<%If IsNumeric(wizard_id) Then%> <%=page_title_%> <%End If%>&nbsp;</td></tr>         
<tr><td width=100% valign=top>
	<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" dir="<%=dir_var%>">
	<tr>	
		<td width="15%" align="center" valign="top" class="title_sort"><span id="word8" name=word8><!--שעות עבודה--><%=arrTitles(8)%></span></td>	
		<td width="15%" align="center" class="title_sort"><span id=word11 name=word11><!--&#8362;--><%=arrTitles(11)%></span>&nbsp;<span id="word7" name=word7><!--עלות שעה--><%=arrTitles(7)%></span></td>
		<td width="30%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--סוג עובד--><%=arrTitles(5)%></span>&nbsp;</td>
		<td width="40%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word9" name=word9><!--שם עובד--><%=arrTitles(9)%></span>&nbsp;</td>
	</tr>	
<%
sqlStr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME, job_name, ACTIVE, hour_pay from USERS_VIEW where ORGANIZATION_ID = "& OrgID & where_worker & " order by FIRSTNAME + ' ' + LASTNAME" 
'Response.Write sqlStr
set rs_USERS = con.GetRecordSet(sqlStr)
if not rs_USERS.EOF then
recCount=rs_USERS.RecordCount 
do while not rs_USERS.EOF
	USER_ID = rs_USERS("USER_ID")
	If recCount = 1 Then ' נמצא רק פרויקט אחד
		Response.Redirect "../projects/hours_calendar.asp?USER_ID=" & USER_ID
    End If
	user_name = rs_USERS(1)
	job_name = rs_USERS("job_name")
	active = rs_USERS("ACTIVE")
	hour_pay = rs_USERS("hour_pay")
%>
	<tr>
		<td align="center" valign="top" class="card"><a href="../projects/hours.asp?USER_ID=<%=USER_ID%>"><img src="../../images/report_icon.gif" border="0"></a></td>
	    <td align="center" class="card"><a href="../projects/hours_calendar.asp?USER_ID=<%=USER_ID%>" class="link_categ" target=_self><%=hour_pay%>&nbsp;</a></td>
	    <td align="<%=align_var%>" class="card"><a href="../projects/hours_calendar.asp?USER_ID=<%=USER_ID%>" class="link_categ" target=_self><%=job_name%>&nbsp;</a></td>
		<td align="<%=align_var%>" class="card"><a href="../projects/hours_calendar.asp?USER_ID=<%=USER_ID%>" class="link_categ" target=_self><strong><%=user_name%></strong>&nbsp;</a></td>
	</tr>
<%
rs_USERS.movenext
loop
set rs_USERS = nothing
else
%>
<tr><td colspan=4 class="title_sort1" align=center><span id="word10" name=word10><!--לא נמצאו עובדים--><%=arrTitles(10)%></span></td></tr>
<%
end if
%>
</table>
</td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td align="<%=align_var%>" colspan=2 class="title_search"><span id=word3 name=word3><!--חיפוש--><%=arrTitles(3)%></span></td></tr>
<FORM action="workers.asp" method=POST id=form_search name=form_search target="_self">   
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><span id="word4" name=word4><!--שם עובד--><%=arrTitles(4)%></span></td></tr>
<tr>
<td align=right><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ID="Image1" NAME="Image1"></td>
<td align=right><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(search_worker)%>" name="search_worker" ID="Text1"></td></tr>
</FORM>
<tr><td colspan=2 height=10 nowrap></td></tr>
</table>
</td>
</tr></table></td></tr></table>
</body>
</html>
<%
set con = nothing
%>