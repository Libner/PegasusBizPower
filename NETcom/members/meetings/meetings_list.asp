<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 
	meeting_date = trim(Request("meeting_date"))
	meeting_users = trim(Request("meeting_users"))
	arr_users = Split(meeting_users, ",")
	
	if lang_id = "1" then
		arr_Status = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status = Array("","Future","Done","Summary added","Postponed")	
	end if	
	
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
<script language=javascript>
<!--
function closeWin()
{
	opener.focus();	
	self.close();
}	
//-->
</script>
</HEAD>
<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td class="page_title" align=center dir="<%=dir_var%>"><font color="#6F6DA6"><%=meeting_date%></font>&nbsp;<%=arrTitles(1)%></td></tr>
<tr><td height=10 nowrap></td></tr>
<%If trim(meeting_users) <> "" Then%>		   		       	
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table align="center" border="0" cellpadding="0" cellspacing="0" dir="<%=dir_var%>">
	<tr>	
    <%
	For i=0 To Ubound(arr_users)
	If trim(arr_users(i)) <> "" Then
		sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & arr_users(i) & " And ORGANIZATION_ID = " & OrgID
		set rs_user = con.getRecordSet(sqlstr)
		if not rs_user.eof then
			userName = trim(rs_user(0)) & " " & trim(rs_user(1))			
		end if
		set rs_user = nothing
	End If	
	%>
	<td valign=top width=250 nowrap align=center><table cellpadding=0 cellspacing=0 width=250 border=0>
	<tr><td colspan=6 class="title_table" align=center><%=userName%></td></tr>	
	<%
	'sqlStr = "SET DATEFORMAT DMY; Select meeting_date,meeting_id,meeting_content,start_time,end_time,meeting_status, "&_
	'" company_name, contact_name from meetings_view where ORGANIZATION_ID = "& OrgID &_
	'" AND participant_id = " & arr_users(i) & " And meeting_date = '" & meeting_date & "'" &_
	'" Order BY start_time, end_time"
	sqlstr = " Execute dbo.get_meetings '" & arr_users(i) & "','" & OrgID & "','" & meeting_date & "'"
	'Response.Write sqlStr
	'Response.End
	set rs_meetings = con.GetRecordSet(sqlStr)	
	%>
	<tr><td width=250 nowrap align=center>
	<table align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>" bgcolor="white" width=250 ID="Table1">
	<tr>	
		<!--td nowrap width=85 align="center" valign="top" class="title_sort3"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td-->
		<td nowrap width=135 align="center" class="title_sort3"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
		<td nowrap width=55 align="center" class="title_sort3"><!--שעת סיום--><%=arrTitles(10)%></td>
		<td nowrap width=55 align="center" class="title_sort3"><!--שעת התחלה--><%=arrTitles(9)%></td>
	</tr>
<%if not rs_meetings.EOF then
	while not rs_meetings.EOF
		meetingID = rs_meetings(0)
		meetingDate = rs_meetings(1) 
		startTime = rs_meetings(2)
		endTime = rs_meetings(3)			
		company_name = rs_meetings(4)
		contact_name = rs_meetings(5)
		status = rs_meetings(6)
		If Len(company_name) > 20 Then
			company_name_short = Left(company_name, 20) & ".."
		Else
			company_name_short = company_name
		End If		
%>
   
	<tr name="title<%=rownum%>" id="title<%=rownum%>">
		<!--td align="center" valign="top" class="title_sort1"><%=contact_name%></td-->
		<td align="center" valign="top" class="title_sort1" title="<%=vFix(company_name)%>"><%=company_name_short%></td>
		<td align="center" valign="top" class="title_sort1"><%=endTime%></td>
		<td align="center" valign="top" class="title_sort1"><%=startTime%></td>
	</tr>
	
<% 
   rs_meetings.moveNext
   Wend  
 Else
%>
    <tr class="card">		
		<td align="center" valign="top" colspan=6 nowrap class="title_sort1"><!--לא נמצאו--><%=arrTitles(3)%>&nbsp;</td>	    
	</tr>
<% End If
   set rs_meetings = Nothing 
 %>
 </table></td></tr></table></td></tr>
 <tr><td height=10 nowrap></td></tr>
 <%  
  Next
%>
</table></td></tr>
<%End If%>
<tr><td height=10 nowrap></td></tr>
<tr><td align=center width=100%>
<input type=button value="<%=arrButtons(8)%>" class="but_menu" style="width:110" onclick="closeWin();">
</td></tr>
</table>
</BODY>
</HTML>
