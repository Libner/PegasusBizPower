<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%  OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	USER_ID = trim(Request("USER_ID"))
    res_groups_id = trim(Request("res_groups_id"))
    lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))	 
	If lang_id = "2" Then
			dir_var = "rtl"
			align_var = "left"
			dir_obj_var = "ltr"
	Else
			dir_var = "ltr"
			align_var = "right"
			dir_obj_var = "rtl"
	End If           
   	
	If trim(Request.QueryString("add")) <> nil Then		
		If Request.Form("group") <> nil Or Request.Form("all_group_id") <> nil Then			
			groups_id = trim(Request.Form("group"))			
			sqlstr = "Select group_name From Users_Groups WHERE ORGANIZATION_ID = " & OrgId &_			
			" And GROUP_ID IN (" & groups_id & ") Order BY group_name"
			'Response.Write sqlstr
			'Response.End						
			set rs_n = con.getRecordSet(sqlstr)
			If not rs_n.eof Then
				groups_name = trim(rs_n.getString(,,",",","))
				groups_name = Left(groups_name, Len(groups_name)-1)
			Else
				groups_name = ""	
			End If
			set rs_n = nothing							
	End If
	%>
		<script language=javascript>
		<!--
		    var groups_id = new String("<%=groups_id%>")
			window.opener.document.all("R_GROUP_ID").value = groups_id;
			window.opener.document.all("ResInGroups").innerHTML = "<%=Replace(groups_name,Chr(34),Chr(39)&Chr(39))%>";			
			// אחראי מורשה צפיה או גם עדכון
			if(groups_id.length > 0)
			{
				window.opener.document.all("edit_appeal_tbody").style.display='inline';
				window.opener.document.all("EDIT_APPEAL")[0].checked = true;				
				window.opener.document.all("EDIT_APPEAL")[1].checked = false;
			}	
			else
			{
				window.opener.document.all("edit_appeal_tbody").style.display='none';
				window.opener.document.all("EDIT_APPEAL")[0].checked = false;				
				window.opener.document.all("EDIT_APPEAL")[1].checked = false;				
			}	
			window.close();
		//-->
		</script>

	<%
   End If
   
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
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" Users_Groups="text/css">
<SCRIPT LANGUAGE=javascript>
function check_all_Groups(objChk)
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
function selecttype(objChk,groupID)
{
	if(objChk.checked == true)
		window.document.all("tr"+groupID).style.background = "#B1B0CF";
	else
		window.document.all("tr"+groupID).style.background = "#E6E6E6";
}
</SCRIPT>
</head>
<body style="margin:0px;background-color:#e6e6e6" onload="window.focus()">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF ID="Table1">
<tr>
<td width="100%" class="page_title" dir=rtl>&nbsp;
<%If trim(lang_id) = "1" Then%>
רשימת קבוצות
<%Else%>
Groups list
<%End If%>
&nbsp;</td></tr>         
<tr><td width="100%">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" valign="top" align="right">    
<%'//start of Users_Groups%>
	<table border=0 width=100% cellpadding=0 cellspacing=0 bgcolor="#e6e6e6">								
		<tr><td align=right valign=top>
		<form name="form1" id="form1" action="res_groups_list.asp?add=1" method=post>		
			<table border=0 width=100% cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">
			<%
				sqlStr = "select GROUP_ID, group_name from Users_Groups WHERE ORGANIZATION_ID=" & OrgId & " order by group_name"
				set rs_groups = con.GetRecordSet(sqlStr)
				If not rs_groups.eof Then												
			%>
				<tr>								
					<td width=100% style="line-height:18px" align="<%=align_var%>" class="title_sort">
					<%If trim(lang_id) = "1" Then%>
					כל הקבוצות
					<%Else%>
					All groups
					<%End if%>
					&nbsp;</td>
					<td width=25 class="title_sort" align=center nowrap><INPUT type=checkbox name="all_group_id" id="all_group_id" onclick="return check_all_Groups(this)"></td>
				</tr>					
			<%
				do while not rs_groups.eof
				groupID = trim(rs_groups(0))				
				groupName = trim(rs_groups(1))						
				If Len(trim(res_groups_id)) > 0 Then
					sqlstr = "Select Top 1 GROUP_ID From Users_Groups WHERE GROUP_ID = " & groupID &_
					" AND GROUP_ID IN (" & res_groups_id & ")"
					'Response.Write sqlstr
					'Response.End
					set rs_chk = con.getRecordSet(sqlstr)
					If not rs_chk.eof Then
						checked = "checked"
					Else
						checked = ""
					End If	
				Else
					checked = ""
				End If
			%>		
			<tr id="tr<%=groupID%>" name="tr<%=groupID%>" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align="<%=align_var%>" dir=rtl>&nbsp;<%=groupName%>&nbsp;</td>
				<td align=center><INPUT type=checkbox name="group" <%=checked%> id="<%=groupID%>" value="<%=groupID%>" onclick=" selecttype(this,<%=groupID%>)"></td>
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
		<td colspan=2 align="center" nowrap dir="<%=dir_var%>">
			<input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button1" NAME="Button1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type=submit class="but_menu" value="<%=arrButtons(1)%>" style="width:80" ID="Button2" NAME="Button2">
		</td>		 
	</tr>
	<tr><td height=10 nowrap colspan=2></td></tr>
	</table></td></tr>
	</form>
	</table></td></tr></table>
</body>
<%set con=nothing%>
</html>
