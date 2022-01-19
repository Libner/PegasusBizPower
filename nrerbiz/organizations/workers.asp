<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckDel(str) {
  return (confirm("? האם ברצונך למחוק את העובד"))    
}
//-->
</script>  
<%
If Request.QueryString("delUSER_ID")<>nil Then
    con.ExecuteQuery "DELETE from Responsibles_to_Groups where Responsible_ID=" & Request.QueryString("delUSER_ID")
	con.ExecuteQuery "delete from Users_to_Groups where USER_ID=" & Request.QueryString("delUSER_ID")
	con.ExecuteQuery "delete from Users_To_Products where USER_ID=" & Request.QueryString("delUSER_ID")	
	con.ExecuteQuery "delete from bar_users where USER_ID=" & Request.QueryString("delUSER_ID")	
	con.ExecuteQuery "delete from hours where USER_ID=" & Request.QueryString("delUSER_ID")	
	con.ExecuteQuery "delete from USERS where USER_ID=" & Request.QueryString("delUSER_ID")
End If

ORGANIZATION_ID = Request.QueryString("ORGANIZATION_ID")
if trim(ORGANIZATION_ID) <> "" then
	sqlStr = "select ORGANIZATION_NAME from  ORGANIZATIONS where  ORGANIZATION_ID = " & ORGANIZATION_ID
	set rs_ORGANIZATIONS = con.GetRecordSet(sqlStr)
	if not rs_ORGANIZATIONS.eof then
		if trim(rs_ORGANIZATIONS("ORGANIZATION_NAME")) <> "" then
			ORGANIZATION_NAME = rs_ORGANIZATIONS("ORGANIZATION_NAME")
		end if
	end if
	set rs_ORGANIZATIONS = nothing
	
	sqlStr = "Select Top 1 Bar_Id from bar_organizations WHERE Bar_ID = 43 And IsNull(is_visible,0) = 1 and ORGANIZATION_ID= "& ORGANIZATION_ID
	'Response.Write sqlStr
	set rs_groups = con.GetRecordSet(sqlStr)
	If not rs_groups.eof Then
		is_groups = "1"
	Else
		is_groups = "0"	
	End If
	set rs_groups = Nothing		
end if	
%>
<body bgcolor="#FFFFFF">
<div align="right">

<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5" class="page_title">דף ניהול עובדים - <%=ORGANIZATION_NAME%></td>
  </tr>
  <tr>
     <td width="5%"  align="center" ><a class="button_admin_1" href="addWorker.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>">הוספת עובד חדש</a></td>    
     <td width="5%"  align="center" ><a class="button_admin_1" href="default.asp">חזרה לדף ארגונים</a></td>         
     <td width="5%"  align="center" ><a class="button_admin_1" href="../../nrerbiz/choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="5%"  align="center" ></td> 
     <td width="*%"  align="center" ></td>      
  </tr>
</table>

<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1">
<tr bgcolor="#B4B4B4">
   <td width="7%" align="center" valign="top"></td>
   <td width="7%" align="center" valign="top"></td>
   <td width="7%" align="center" valign="top" class="11normalB">פעיל</td>
   <%If trim(is_groups) = "1" Then%>
   <td width="20%" align="right" class="11normalB">&nbsp;אחראי קבוצות&nbsp;</td>
   <td width="20%" align="right" class="11normalB">&nbsp;שייך לקבוצות&nbsp;</td>
   <%End If%>
   <td width="20%" align="right" class="11normalB">&nbsp;תפקיד&nbsp;</td>
   <td width="20%" align="right" class="11normalB">&nbsp;שם עובד&nbsp;</td>
</tr>
<%
sqlStr = "select USER_ID, FIRSTNAME + ' ' + LASTNAME, ACTIVE,job_name from USERS_VIEW() where ORGANIZATION_ID = "& ORGANIZATION_ID &" order by FIRSTNAME + ' ' + LASTNAME" 
''Response.Write sqlStr
set rs_USERS = con.GetRecordSet(sqlStr)

do while not rs_USERS.EOF
	USER_ID = rs_USERS(0)
	User_Name = trim(rs_USERS(1))
	perSite = trim(rs_USERS(2))
	job_name = trim(rs_USERS(3))
	Groups=""		
	sqlstr="Select Group_Name From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_id"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		Groups = rssub.getString(,,",",",")		
	Else	
		Groups = ""
	End If		    
	set rssub=Nothing
	If Len(Groups) > 0 Then
		Groups = Left(Groups,(Len(Groups)-1))
	End If	
	
	ResInGroups=""		
	sqlstr="Select Group_Name From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_id"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		ResInGroups = rssub.getString(,,",",",")	
	Else
		ResInGroups=""
	End If		    
	set rssub=Nothing
	If Len(ResInGroups) > 0 Then
		ResInGroups = Left(ResInGroups,(Len(ResInGroups)-1))
	End If		

%>
	<tr bgcolor="#DDDDDD">
		<td align="center" class="10normal"><a href="workers.asp?delUSER_ID=<%=USER_ID%>&ORGANIZATION_ID=<%=ORGANIZATION_ID%>" ONCLICK="return CheckDel()"><img src="../images/delete_icon.gif" border="0" alt="מחיקה"></a></td>
		<td align="center" class="10normal"><a href="addWorker.asp?USER_ID=<%=USER_ID%>&ORGANIZATION_ID=<%=ORGANIZATION_ID%>"><img src="../images/edit_icon.gif" border="0" alt="עדכון"></a></td>
	    <td align="center" class="10normal"><a href="vsbPress_USER.asp?idsite=<%=USER_ID%>&ORGANIZATION_ID=<%=ORGANIZATION_ID%>"><%if perSite = "0" then%><img src="../images/lamp_off.gif" alt="לא מופיע באתר" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../images/lamp_on.gif" alt="מופיע באתר" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>
		<%If trim(is_groups) = "1" Then%>
		<td align="right" class="10normal">&nbsp;<%=ResInGroups%>&nbsp;</td>	   	    
	    <td align="right" class="10normal">&nbsp;<%=Groups%>&nbsp;</td>	   
	    <%End If%>
	    <td align="right" class="10normal">&nbsp;<%=job_name%>&nbsp;</td>
		<td align="right" class="11normalB">&nbsp;<%=User_Name%>&nbsp;</td>
	</tr>
<%
rs_USERS.movenext
loop
set rs_USERS = nothing
%>
</table>
</div>
</body>
</html>
<%
set con = nothing
%>