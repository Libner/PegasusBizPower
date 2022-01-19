<!--#include file="../netcom/connect.asp"-->
<!--#include file="../netcom/reverse.asp"-->
<%
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
response.buffer=true
Response.Expires = 0 

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
'Response.Write (clId &":"& prId&":"&mcid)

If Request("delete") <> nil And hour_id <> "" Then
    con.executeQuery("DELETE FROM hours WHERE hour_id = " & hour_id)
End If

if Request("update")<>nil and hour_id <> "" then		
    hours = trim(Request.QueryString("hours"))   
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

%>
<html>
<head>
<title>BizPower</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=strLocal%>/netcom/IE4.css" rel="STYLESHEET" type="text/css">
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
		    document.location.href = "desk_hours.asp?delete=1&hour_id=" + hour_id + "&User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>"; 

		}
		return true;
	}
	function update_hour(hour_id)
	{
		document.location.href = "desk_hours.asp?update=1&hour_id=" + hour_id + "&hours=" + document.form1.all("hours" + hour_id).value + "&User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>"; 
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
	' check not approved days
   con.executeQuery("SET DATEFORMAT dmy")   
   sqlstr = "Select hour_id,company_id,company_name,project_id,project_name,mechanism_name,mechanism_id,minuts from hours_view Where ORGANIZATION_ID = "& OrgID &_
   " AND USER_ID = " & UserId & " AND date <> '" & current_date & "' AND approved is null"
   set rs_tmp = con.getRecordset(sqlstr)
	if not rs_tmp.eof then %>
		<SCRIPT LANGUAGE=javascript>
		<!--
		document.location.href = "desk_approve.asp?User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>"; 
		//-->
		</SCRIPT>
<%	else 'not rs_tmp.eof
	
sqlstr = "Select hour_id,company_id,company_name,project_id,project_name,mechanism_name,mechanism_id,minuts from hours_view Where ORGANIZATION_ID = "& OrgID &_
" AND USER_ID = " & UserId & " AND date = '" & current_date & "' AND approved is null"
   set rs_project = con.getRecordset(sqlstr)
	all_hours = 0
    rs_date = current_date
	day_num = DatePart("w",rs_date)
	day_name = daysname(day_num-1)
%>
	<tr>	
		<td width="100%" align="center" valign="top" colspan=4><span dir=rtl style="font-size:10pt"><b>&nbsp; להלן שעות העבודה של היום &nbsp;(&nbsp;<%=day_name%>&nbsp;<%=rs_date%>&nbsp;)&nbsp;</b></span></td>
	</tr>
	<tr>	
	<td width="100%" align="center" valign="top" class="title_sort">חברה/פרויקט/מנגנון</td>
	<td width="70" nowrap align="center" class="title_sort">שעות עבודה</td>
	<td width="30" nowrap align="center" valign="top" class="title_sort">עדכון</td>
	<td width="30" nowrap align="center" valign="top" class="title_sort">מחיקה</td>
	</tr>
<form name="form1" id="form1" method=post action="desk_hours.asp?date_=<%=rs_date%>&User_ID=<%=UserID%>&Org_Id=<%=OrgID%>&clId=<%=clId%>&prId=<%=prId%>&mcId=<%=mcId%>">	
<%    
		do while not rs_project.eof 
			is_projects = true
			hour_id = rs_project("hour_id")
			company_name = rs_project("company_name")	
			project_name = rs_project("project_name")
			mechanism_name = rs_project("mechanism_name")
			mechanism_Id = rs_project("mechanism_Id")
			company_id = rs_project("company_id")	
			project_id = rs_project("project_id")
			minuts_pr = rs_project("minuts")
			all_hours = all_hours + minuts_pr
			hours_pr = GetTime(minuts_pr) 
%>
<tr>				
	<td class="card1" align=right >
	<table cellspacing=0 border=0 cellpadding=0>
	<tr><td align=right><b><%=company_name%></b></tD></tr>
	<tr><td align=right> <%=project_name%></td></tr>
	<tR><tD align=right><%=mechanism_name%></td></tr>
	
	</table>
	 </td>	
	<td class="card1" align=center>
		<input type=text class="texts" name="hours<%=hour_id%>" id="hours<%=hour_id%>" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=5 style="text-align:center">
	</td>
	<%if (cstr(clId) = cstr(company_id)) and (cstr(prId) = cstr(project_id))  and (cstr(mcId) = cstr(mechanism_id)) then%>
	<td class="card1" align=center colspan=2 style="background-color:#ACE920">בעבודה</td>
	<%else%>	
	<td class="card1" align=center><img src="../icons/check.gif" border="0" align="absmiddle" onclick="return update_hour('<%=hour_id%>')" style="cursor:hand"></td>
	<td class="card1" align=center><img src="../icons/delete_icon.gif" border="0" align="absmiddle" onclick="return delete_hour('<%=hour_id%>')" style="cursor:hand"></td>
	<%end if%>
	</tr>	
	<%
		rs_project.moveNext
		loop
	end if 'not rs_tmp.eof
	set rs_tmp = nothing
	set rs_project = nothing
		%>
	<tr>		
	<td class="card1" align=center>
	<input type=button class="but_menu1" style="width:145" value="הוספת שעות עבודה" onclick="window.open('desk_addhour.asp?date=<%=rs_date%>&User_ID=<%=UserID%>&Org_Id=<%=OrgId%>','','left=200,top=200,width=400,height=250')" id=button1 name=button1></td>
	<td class="card1" align=center><b><%=GetTime(all_hours)%></b></td>
	<td class="card1" colspan=2>&nbsp;סה''כ שעות</td>
	</td>
	</tr>
	<tr>		
		<td class="card1" align=center colspan=4 height=20 valign=middle>
	<%If 0 Then%><input type=submit class="but_menu1" value="אישור" style="width:100;height:30;font-size:12pt" ID="Submit1" NAME="Submit1"><%End if%>
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