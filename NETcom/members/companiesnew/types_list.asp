<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%  
    If Request.QueryString("OrgId") = nil Then
		OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	Else
		OrgId = trim(Request.QueryString("OrgId"))
	End If	
    contactID = trim(Request("contactID"))
    contact_types = trim(Request("contact_types"))
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
		If Request.Form("type") <> nil Or Request.Form("all_contact_type_id") <> nil Then			
			types_id = trim(Request.Form("type"))			
			sqlstr = "Select type_name From contact_type WHERE ORGANIZATION_ID = " & OrgId &_			
			" And type_id IN (" & types_id & ") Order BY type_name"			
			set rs_n = con.getRecordSet(sqlstr)
			If not rs_n.eof Then
				types_name = trim(rs_n.getString(,,",",","))
				types_name = Left(types_name, Len(types_name)-1)
			Else
				types_name = ""	
			End If
			set rs_n = nothing							
	End If
	%>
		<script language=javascript>
		<!--
			window.opener.document.all("contact_types").value = "<%=types_id%>";
			window.opener.document.all("types_names").innerHTML = "<%=Replace(types_name,Chr(34),Chr(39)&Chr(39))%>";			
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
<link href="../../IE4.css" rel="STYLESHEET" contact_type="text/css">
<SCRIPT LANGUAGE=javascript>
function check_all_contact_type(objChk)
{
	input_arr = document.getElementsByName("type");	
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
function selecttype(objChk,typeID)
{
	if(objChk.checked == true)
		window.document.all("tr"+typeID).style.background = "#B1B0CF";
	else
		window.document.all("tr"+typeID).style.background = "#E6E6E6";
}
</SCRIPT>
</head>
<body style="margin:0px;background-color:#e6e6e6" onload="window.focus()">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF ID="Table1">
<tr>
<td width="100%" class="page_title" dir=rtl>&nbsp;
<%If trim(lang_id) = "1" Then%>
רשימת סיווגים
<%Else%>
Groups list
<%End If%>
&nbsp;</td></tr>         
<tr><td width="100%">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" valign="top" align="right">    
<%'//start of contact_type%>
	<table border=0 width=100% cellpadding=0 cellspacing=0 bgcolor="#e6e6e6">								
		<tr><td align=right valign=top>
		<form name="form1" id="form1" action="types_list.asp?add=1" method=post>		
			<table border=0 width=100% cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">
			<%
				sqlStr = "select type_Id, type_Name from contact_type WHERE ORGANIZATION_ID=" & OrgId & " order by type_Name"
				set rs_contact_type = con.GetRecordSet(sqlStr)
				If not rs_contact_type.eof Then												
			%>
				<tr>								
					<td width=100% style="line-height:18px" align="<%=align_var%>" class="title_sort">
					<%If trim(lang_id) = "1" Then%>
					כל הסיווגים
					<%Else%>
					All groups
					<%End if%>
					&nbsp;</td>
					<td width=25 class="title_sort" align=center nowrap><INPUT type=checkbox name="all_contact_type_id" id="all_contact_type_id" onclick="return check_all_contact_type(this)"></td>
				</tr>					
			<%
				do while not rs_contact_type.eof
				typeId = trim(rs_contact_type(0))				
				type_Name = trim(rs_contact_type(1))						
				If trim(contactID) <> "" And trim(typeId) <> "" Then					
				sqlstr = "Select * From contact_to_types WHERE type_ID = " & typeId &_				
				" AND contact_id = " & contactID				
				set rs_chk = con.getRecordSet(sqlstr)
				If not rs_chk.eof Then
					checked = "checked"
				Else
					checked = ""					
				End If	
				End If
				If trim(contact_types) <> "" Then
				sqlstr = "Select * From contact_type WHERE type_id = " & typeId &_
				" AND type_id IN (" & contact_types & ")"
				'Response.Write sqlstr
				'Response.End
				set rs_chk = con.getRecordSet(sqlstr)
				If not rs_chk.eof Then
					checked = "checked"
				Else
					checked = ""
				End If	
				End If					
			%>		
			<tr id="tr<%=typeId%>" name="tr<%=typeId%>" <%If checked <> "" Then%> style="background:#B1B0CF" <%Else%> style="background:#E6E6E6" <%End If%>>		    
				<td align="<%=align_var%>" dir=rtl>&nbsp;<%=type_Name%>&nbsp;</td>
				<td align=center><INPUT type=checkbox name="type" <%=checked%> id="<%=typeId%>" value="<%=typeId%>" onclick=" selecttype(this,<%=typeId%>)"></td>
			</tr>					
			<%
		rs_contact_type.movenext
		loop
		set rs_contact_type = nothing
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
