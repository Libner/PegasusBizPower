<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  UserId = cInt(trim(Request.Cookies("bizpegasus")("UserID")))
    OrgId = cInt(trim(Request.Cookies("bizpegasus")("OrgId")))
    grp_type = trim(Request.QueryString("grp_type"))	
    If Not isNumeric(grp_type) Then
		grp_type = "1"
    End If
	
	groups_id = trim(Request.QueryString("groups_id"))		
	If trim(Request.QueryString("add")) <> nil Then	
	'For each item In Request.Form
		'Response.Write item & "-" & Request.Form(item) & "<br>"
	'Next
	'Response.End
		If Request.Form("group") <> nil Or Request.Form("all_groups_id") <> nil Then			
			groups_id = trim(Request.Form("group"))			
			sqlstr = "Select GroupName From Groups WHERE ORGANIZATION_ID = " & OrgId &_			
			" And GroupID IN (" & groups_id & ")"			
			set rs_n = con.getRecordSet(sqlstr)
			If not rs_n.eof Then
				groups_name = trim(rs_n.getString(,," ; "," ; "))
				groups_name = Left(groups_name, Len(groups_name)-1)
			Else
				groups_name = ""	
			End If
			set rs_n = nothing
			sqlstr="Select COUNT(PEOPLE_ID) FROM PEOPLES WHERE GROUPID IN (" & groups_id & ")"
			set rs_count = con.getRecordSet(sqlstr)
			If not rs_count.eof Then
				count_all = trim(rs_count(0))
			End If		
			set rs_count = Nothing				
		%>
		<script language=javascript>
		<!--
			window.opener.document.all("groups_id").value = "<%=groups_id%>";
			window.opener.document.all("groups_name").innerHTML = "&nbsp;<span dir='<%=dir_obj_var%>'><%=vFix(groups_name)%></span>&nbsp;<font dir='<%=dir_obj_var%>' color=red>(<%=count_all%>&nbsp;נמענים)</font>";
			window.close();
		//-->
		</script>
		<%
	End If
   End If%>
<%	If trim(Request.QueryString("sort"))<>"" Then
		sort = cInt(Request.QueryString("sort"))
	Else
		sort = 0
	End If 
	
	dim sortby(2) 
	sortby(0) = "GROUPDATE DESC"
	sortby(1) = "groupName"
	sortby(2) = "groupName DESC"

	urlSort = "groups.asp?1=1"
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 44 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
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
	  set rsbuttons=nothing%>   
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
function updateGroups()
{   
	return true;
}
function check_all_groups(objChk)
{
	input_arr = document.getElementsByName("group");	
	for(i=0;i<input_arr.length;i++)
	{
		//input_arr(i).disabled = objChk.checked;
		input_arr(i).checked = objChk.checked;
		if(objChk.checked == true)
			window.document.all("tr"+input_arr(i).value).style.background = "#B1B0CF";
		else
			window.document.all("tr"+input_arr(i).value).style.background = "#E6E6E6";
	}
	return true;
}
function selectGroup(objChk,groupID)
{
	if(objChk.checked == true)
		window.document.all("tr"+groupID).style.background = "#B1B0CF";
	else
		window.document.all("tr"+groupID).style.background = "#E6E6E6";
}
function openPreview(groupId)
{
	page = window.open("../groups_clients/group_clients.asp?groupId="+groupId,"Group","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=680, height=320, left=50, top=50");	
	return false;
}
</SCRIPT>
</head>
<body style="margin:0px" onload="window.focus()">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr><td width="100%" class="page_title" dir="<%=dir_obj_var%>">&nbsp;<!--רשימת קבוצות דיוור--><%=arrTitles(17)%>&nbsp;</td></tr>         
<tr><td width="100%">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" valign="top" align="<%=align_var%>">    
<%'//start of groups%>
<form name="form1" id="form1" action="groups.asp?add=1" method=post>
	<table border=0 width=100% cellpadding=0 cellspacing=0>							
		<tr><td align="<%=align_var%>" valign=top>
		<input type=hidden name="groups_id" id="groups_id" value="<%=groups_id%>">		
			<table border=0 width=100% cellpadding=0 cellspacing=1 bgcolor="#FFFFFF">
			<%
				sqlStr = "select groupId, groupName,GROUPDATE from groups WHERE ORGANIZATION_ID=" & OrgId &_
				"And GROUPTYPE = " & grp_type & " order by " & sortby(sort)
				set rs_groups = con.GetRecordSet(sqlStr)
				If not rs_groups.eof Then												
			%>
				<tr>					
					<td width=70 class="title_sort" align=center nowrap>&nbsp;<!--נמענים--><%=arrTitles(16)%>&nbsp;</td>
					<td width=80 class="title_sort" align=center nowrap><!--תאריך יצירה--><%=arrTitles(5)%></td>
					<td width=50 class="title_sort" align=center nowrap><!--כמות--><%=arrTitles(6)%></td>												
					<td width=100% align="<%=align_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self"><!--שם קבוצה--><%=arrTitles(7)%><img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
					<td width=25 class="title_sort" align=center nowrap><INPUT type=checkbox name="all_groups_id" id="all_groups_id" onclick="return check_all_groups(this)"></td>
				</tr>					
			<%
				do while not rs_groups.eof
				groupId = trim(rs_groups("groupId"))				
				groupName = trim(rs_groups("groupName"))			
				GROUPDATE = trim(rs_groups("GROUPDATE"))
				If IsDate(GROUPDATE) And trim(GROUPDATE) <> "" Then
					GROUPDATE = DateValue(Day(GROUPDATE) & "/" & Month(GROUPDATE) & "/" & Year(GROUPDATE))
				End if
				sqlstr = "Select Count(PEOPLE_ID) From PEOPLES WHERE GROUPID = " & groupId
				set rs_count = con.getRecordSet(sqlstr)
				If not 	rs_count.eof Then
					count_people = trim(rs_count(0))
				Else
					count_people = 0
				End If
				Set	rs_count = Nothing
				If trim(groups_id) <> "" Then
				sqlstr = "Select * From Groups WHERE GROUPID = " & groupId &_
				" AND GROUPID IN (" & groups_id & ")"
				set rs_chk = con.getRecordSet(sqlstr)
				If not rs_chk.eof Then
					checked = "checked"
				Else
					checked = ""
				End If	
				End If	
			%>		
			<tr id="tr<%=groupId%>" name="tr<%=groupId%>" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align=center>&nbsp;<INPUT type=image OnClick="return openPreview('<%=groupId%>')" SRC="../../images/preview_icon.gif" BORDER=0 ID="Image1" NAME="Image1">&nbsp;</td>
				<td align=center><%=GROUPDATE%></td>
				<td align=center><%=count_people%></td>
				<td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=groupName%>&nbsp;</td>
				<td align=center><INPUT type=checkbox name="group" <%=checked%> id="<%=groupId%>" value="<%=groupId%>" onclick=" selectGroup(this,<%=groupId%>)"></td>
			</tr>					
			<%
		rs_groups.movenext
		loop
		set rs_groups = nothing
		If Len(ids) > 1 Then
			ids = Left(ids, Len(ids)-1)
		End If
		If Len(names) > 1 Then
			names = Left(names, Len(names)-1)
		End If
		End If
		%>					
		</table>						
	  </td></tr>
	<tr><td colspan=2 height="15" nowrap>
	<input type=hidden name="all_ids" id="all_ids" value="<%=ids%>">
	<input type=hidden name="all_names" id="all_names" value="<%=vFix(names)%>">
	</td></tr>
	<tr>
		<td colspan=2 align="center" nowrap>
			<input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button1" NAME="Button1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type=submit class="but_menu" onclick="return updateGroups()" value="<%=arrButtons(1)%>" style="width:80" ID="Button2" NAME="Button2">
		</td>		 
	</tr>
	</table>	
	</td></tr>		
	</table>
	</form>
	<%'//end of groups%>			
</td>	
</tr></table>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
</html>
<%set con=nothing%>