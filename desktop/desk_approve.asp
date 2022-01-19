<!--#include file="../netcom/connect.asp"-->
<!--#include file="../netcom/reverse.asp"-->
<%
Function GetTime(inMin)
dim str_hour,str_min
    If cdbl(inMin) >= 60 Then
        str_hour = Right("0" & CStr(cdbl(Split(CStr(inMin / 60), ".")(0)) Mod 60),2) & ":"
    Else
        str_hour = "00:"
    End If
    str_min = Right("0" & CStr(cdbl(inMin) Mod 60),2)
    GetTime = str_hour & str_min
End Function
if Request.QueryString("User_Id") <> nil then
	UserId = Request("User_Id")
else
	UserId = ""
end if
if Request.QueryString("Org_Id") <> nil then
	OrgID = Request("Org_Id")
else
	OrgID = ""
end if	
if Request.QueryString("date_") <> nil then
	date_ = Request("date_")
else
	date_ = ""
end if
current_date = date()
if Request.QueryString("hour_id") <> nil then
	hour_id = trim(Request.QueryString("hour_id"))
else
	hour_id = ""
end if	

if Request.QueryString("clId") <> nil then
	clId = trim(Request.QueryString("clId"))
else
	clId = ""
end if

if Request.QueryString("prId") <> nil then
	prId = trim(Request.QueryString("prId"))
else
	prId = ""
end if
if Request.QueryString("mcId") <> nil then
	mcId = trim(Request.QueryString("mcId"))
else
	mcId = ""
end if

If Request("delete") <> nil And hour_id <> "" Then
    con.executeQuery("DELETE FROM hours WHERE hour_id = " & hour_id)
End If
if Request("update")<>nil and hour_id <> "" then		
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
		con.executeQuery("SET DATEFORMAT dmy")
		con.executeQuery("Update hours Set minuts = " & minuts & " WHERE hour_id = " & hour_id)			
	End If
end if
if Request.Form("Submit1") <> nil and date_ <> "" and UserId <> "" and OrgID <> "" then
	con.executeQuery("SET DATEFORMAT dmy")
	con.executeQuery("Update hours Set approved=1 WHERE ORGANIZATION_ID="& OrgID & " AND USER_ID=" & UserId & " AND date='" & date_ & "'")	
end if
%>
<html>
<head>
<title>BizPower</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../netcom/IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
    
	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 58) ;
	} 
	
	function delete_hour(hour_id)
	{
		if(window.confirm("? האם ברצונך למחוק את שעות עבודה על הפרויקט הזאת") == true)
		{
		    document.location.href = "desk_approve.asp?delete=1&hour_id=" + hour_id + "&User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>"; 

		}
		return false;
	}
	function update_hour(hour_id)
	{
		document.form1.action = "desk_approve.asp?update=1&hour_id=" + hour_id + "&User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>"; 
		document.form1.submit();
		return false;
	}
<!--End-->
</script> 
</head> 
<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0" bgcolor="#EFF1EF"  scroll=yes style="MARGIN: 0pt 0px;">
<%If trim(UserID) <> "" Then%>
<table border="0" cellpadding="0" cellspacing="1" width="100%">
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center>
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
	
<%
daysname = array("ראשון","שני","שלישי","רביעי","חמישי","שישי","שבת")
   con.executeQuery("SET DATEFORMAT dmy")   
   sqlstr = "Select * from hours_view Where ORGANIZATION_ID = "& OrgID &_
   " AND USER_ID = " & UserId & " AND date <> '" & current_date & "' AND approved is null Order by [date]"
   set rs_project = con.getRecordset(sqlstr)
	all_hours = 0
    if not rs_project.eof then
		is_projects = true
		rs_date = rs_project("date")
		day_num = DatePart("w",rs_date)
		day_name = daysname(day_num-1)
%>
	<tr>	
		<td width="100%" align="center" valign="top" colspan=4><span dir=rtl style="font-size:10pt"><b>&nbsp; להלן פירוט שעות העבודה שלך ביום &nbsp;<%=day_name%>&nbsp;<%=rs_date%>&nbsp;<br>אנא בצע שינוים ואשר את דיווח השעות</b></span></td>
	</tr>
	<tr>	
	<td width="100%" align="center" valign="top" class="title_sort">ארגון/פרויקט/מנגנון
	</td>
	<td width="80" nowrap align="center" class="title_sort">שעות עבודה</td>
	<td width="50" nowrap align="center" valign="top" class="title_sort">עדכון</td>
	<td width="50" nowrap align="center" valign="top" class="title_sort">מחיקה</td>
	</tr>
<form name="form1" id="form1" method=post action="desk_approve.asp?date_=<%=rs_date%>&User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>">	
<%    
		do while not rs_project.eof 
			if rs_date <> rs_project("date") then
				exit do
			end if
			hour_id = rs_project("hour_id")
			company_name = rs_project("company_name")	
			project_name = rs_project("project_name")
			mechanism_name = rs_project("mechanism_name")
			minuts_pr = rs_project("minuts")
			all_hours = all_hours + minuts_pr
			hours_pr = GetTime(minuts_pr) 
%>
	<tr>				
	<td class="card1" align=right dir=rtl> <%=company_name%> / <%=project_name%> / <%=mechanism_name%> </td>	
	<td class="card1" align=center>
		<input type=text class="texts" name="hours<%=hour_id%>" id="hours<%=hour_id%>" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=5 style="text-align:center">
	</td>	
	<td class="card1" align=center><img src="../icons/check.gif" border="0" align="absmiddle" onclick="return update_hour('<%=hour_id%>')" style="cursor:hand"></td>
	<td class="card1" align=center><img src="../icons/delete_icon.gif" border="0" align="absmiddle" onclick="return delete_hour('<%=hour_id%>')" style="cursor:hand"></td>
	</tr>	
	<%
		rs_project.moveNext
		loop
		
		else 'not rs_project.eof %>
			<SCRIPT LANGUAGE=javascript>
			<!--
			document.location.href = "desk_hours.asp?User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>"; 
			//-->
			</SCRIPT>
<%		end if
		set rs_project = nothing
		%>
	<tr>		
	<td class="card1" align=center>
	<input type=button class="but_menu1" style="width:145" value="הוספת שעות עבודה" onclick="window.open('desk_addhour.asp?date=<%=rs_date%>&User_ID=<%=UserID%>&Org_Id=<%=OrgId%>','','left=200,top=200,width=400,height=200')"></td>
	<td class="card1" align=center><b><%=GetTime(all_hours)%></b></td>
	<td class="card1" colspan=2>&nbsp;סה''כ שעות</td>
	</td>
	</tr>
	<tr>		
		<td class="card1" align=center colspan=4 height=50 valign=middle>
	<%If is_projects Then%><input type=submit class="but_menu1" value="אישור" style="width:100;height:30;font-size:12pt" ID="Submit1" NAME="Submit1"><%End if%>
		</td>
	</tr>
	</form>	
</table>
</td>
</tr></table>
</td></tr></table>
<%else 'If trim(UserID) <> ""%>
<br><br>
<center><font color=red face=arial size=3><b><span dir=rtl>אנא להגדיר שם משתמש וסיסמא במערכת BizPower בחלון הגדרות</span></b></font></p></center>
<%end if 'If trim(UserID) <> ""%>
</body>
</html>
<%
set con = nothing
%>