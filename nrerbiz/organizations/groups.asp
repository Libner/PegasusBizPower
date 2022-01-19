<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<%
	OrgId = trim(Request("ORGANIZATION_ID"))
	If trim(OrgId) <> "" Then
	sqlStr = "select ORGANIZATION_NAME from ORGANIZATIONS where ORGANIZATION_ID=" & OrgId  
	''Response.Write sqlStr
	set rs_org = con.GetRecordSet(sqlStr)
	if not rs_org.eof then
		orgName = trim(rs_org(0))
	end if
	set rs_org = nothing
	End If  
	
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then
		    con.executeQuery("Delete From Responsibles_to_Groups Where group_id = " & delId)
		    con.executeQuery("Delete From Users_Groups Where group_id = " & delId)		    
			con.executeQuery("Delete From Users_to_Groups Where group_id = " & delId)
		End If
		Response.Redirect "groups.asp?ORGANIZATION_ID=" & OrgId
	End If
%>
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function checkDelete()
	{
		return window.confirm("?האם ברצונך למחוק את הקבוצה");	
	}
	
	function openUsers_Groups(groupID,OrgID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addgroup.asp?group_id=" + groupID + "&ORGANIZATION_ID="+OrgID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	
//-->
</SCRIPT>

</HEAD>

<body bgcolor="#FFFFFF">
<div align="right">

<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="4" class="page_title" width=100% dir=rtl><%=orgName%> - רשימת קבוצות</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1"  href="" onClick="return openUsers_Groups('','<%=OrgID%>')">הוספת קבוצה</a></td>    
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדף ארגונים</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>       
     <td width="*%" align="center"></td>      
  </tr>
</table>
<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr><td height=20 nowrap></td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table1">
<tr>
<td width=100% valign=top>
<table width="650" cellspacing="1" cellpadding="2" align=center border="0" bgcolor="#ffffff">
<tr bgcolor="#B4B4B4" height="22">
	<th width="80" nowrap align="center" class="11normalB">&nbsp;מחיקה&nbsp;</th>	
	<th width="80" nowrap align="center" class="11normalB">&nbsp;עריכה&nbsp;</th>	
	<th width="50%" align="right" class="11normalB">&nbsp;אחראי&nbsp;</th>
	<th width="50%" align="right" class="11normalB">&nbsp;קבוצה&nbsp;</th>	
</tr>
<%	set rs_Users_Groups = con.GetRecordSet("Select Users_Groups.group_id, Users_Groups.group_name from Users_Groups WHERE Users_Groups.Organization_ID = " & trim(OrgID) & " order by Users_Groups.group_id")
    if not rs_Users_Groups.eof then 
		do while not rs_Users_Groups.eof
    	group_Id = trim(rs_Users_Groups(0))
    	group_name = trim(rs_Users_Groups(1))
    	If IsNumeric(trim(group_Id)) Then
    		sqlstr = "Select FIRSTNAME + Char(32) + LASTNAME From Responsibles_To_Groups_View WHERE Group_ID = " &_
    		group_Id & " And Organization_ID = " & trim(OrgID)
    		set rs_resp = con.getRecordSet(sqlstr)
    		If not rs_resp.eof Then
    			reponsibles_list = rs_resp.getString(,,",",",")
    			If Len(reponsibles_list) > 1 Then
    				reponsibles_list = Left(reponsibles_list, Len(reponsibles_list)-1)
    			End If
    		End If
    		set rs_resp = nothing
    	End If
%>
<tr>
	<td align="center" class="normalB" bgcolor="#e6e6e6" nowrap><a href="groups.asp?deleteId=<%=group_id%>&ORGANIZATION_ID=<%=OrgId%>" ONCLICK="return checkDelete()"><IMG SRC="../images/delete_icon.gif" BORDER=0 alt="מחיקת קבוצת עובדים"></a></td>
	<td align="center" class="normalB" bgcolor="#e6e6e6" nowrap><a href="" onClick="return openUsers_Groups('<%=group_Id%>','<%=OrgID%>')" target="_blank"><IMG SRC="../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align=right class="normalB" bgcolor="#e6e6e6"><%=reponsibles_list%>&nbsp;</td>	
	<td align=right class="normalB" bgcolor="#e6e6e6"><b><%=group_name%></b>&nbsp;</td>	
</tr>
<%		rs_Users_Groups.MoveNext
		loop
	end if
	set rs_Users_Groups=nothing%>
</table>
</td>
</tr>
<tr><td height=10 nowrap></td></tr>
</table>
</BODY>
</HTML>
