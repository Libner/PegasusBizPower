<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%		
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
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 64 Order By word_id"				
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

%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function checkDelete()
	{
		<%If trim(lang_id) = "1" Then%>
			  str_confirm = "?האם ברצונך למחוק את הקבוצה"
		<%Else%>		
			  str_confirm = "Are you sure want to delete the group?"		
		<%End If%>		
		return window.confirm(str_confirm);		
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
</head>
<body>
<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>" ID="Table2">
<tr><td width=100%>
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 4%>
<%numOfLink = 5%>
<tr><td width=100%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td width=100% valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff" >
<tr>
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name=word4><!--עדכון--><%=arrTitles(4)%></span>&nbsp;</td>	
	<td width="50%" align="<%=align_var%>" class="title_sort">&nbsp;<span id=word7 name=word7><!--אחראים--><%=arrTitles(7)%></span>&nbsp;</td>
	<td width="50%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--קבוצה--><%=arrTitles(5)%></span>&nbsp;</td>	
</tr>
<%	set rs_Users_Groups = con.GetRecordSet("Select Users_Groups.group_id, Users_Groups.group_name from Users_Groups WHERE Users_Groups.Organization_ID = " & trim(OrgID) & " order by Users_Groups.group_id")
    if not rs_Users_Groups.eof then 
		do while not rs_Users_Groups.eof
    	group_Id = trim(rs_Users_Groups(0))
    	group_name = trim(rs_Users_Groups(1))
    	reponsibles_list = ""
    	If IsNumeric(trim(group_Id)) Then
    		sqlstr = "Select FIRSTNAME + Char(32) + LASTNAME From Responsibles_To_Groups_View WHERE Group_ID = " &_
    		group_Id & " And Organization_ID = " & trim(OrgID)
    		set rs_resp = con.getRecordSet(sqlstr)
    		If not rs_resp.eof Then
    			reponsibles_list = rs_resp.getString(,,",",",")
    			If Len(reponsibles_list) > 1 Then
    				reponsibles_list = Left(reponsibles_list, Len(reponsibles_list)-1)
    			End If
    		Else
    			reponsibles_list = ""	
    		End If
    		set rs_resp = nothing
    	End If
    		
%>
<tr>
	<td align="center" class="card"><a href="groups.asp?deleteId=<%=group_id%>&ORGANIZATION_ID=<%=OrgId%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת קבוצת עובדים שיחה"></a></td>
	<td align="center" class="card"><a href="" onClick="return openUsers_Groups('<%=group_Id%>','<%=OrgID%>')" target="_blank"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align="<%=align_var%>" class="card"><%=reponsibles_list%>&nbsp;</td>	
	<td align="<%=align_var%>" class="card"><b><%=group_name%></b>&nbsp;</td>	
</tr>
<%		rs_Users_Groups.MoveNext
		loop
	end if
	set rs_Users_Groups=nothing%>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% ID="Table1">
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openUsers_Groups('','<%=OrgID%>')"><span id=word6 name=word6><!--הוספת קבוצה--><%=arrTitles(6)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href="excel_groups.asp" target=_blank><!--הצג דוח--><%=arrTitles(8)%></a></td></tr>

</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</BODY>
</HTML>
