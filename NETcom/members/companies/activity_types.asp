<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	If Request.QueryString("deleteId") <> nil Then
		con.executeQuery "Delete From activity_types Where activity_type_id = " & Request.QueryString("deleteId")
		Response.Redirect "activity_types.asp"
	End If
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 34 Order By word_id"				
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
	  set rsbuttons=nothing	 	  
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
	   var str_confirm = "?האם ברצונך למחוק את סוג המשימה"
	   <%Else%>
	   var str_confirm = "Are you sure want to delete the task type?"
	   <%End If%>
		return window.confirm(str_confirm);		
	}
	
	function openactivity_type(typeID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addactivityType.asp?activity_type_id=" + typeID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	
//-->
</SCRIPT>

</HEAD>

<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 4%>
  <%numOfLink = 3%>
<%topLevel2 = 34 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align="<%=align_var%>" valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td width=100% valign=top>
<table width="100%" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff">
<tr>
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name=word4><!--עדכון--><%=arrTitles(4)%></span>&nbsp;</td>	
	<td width="100%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--סוג--><%=arrTitles(5)%></span>&nbsp;</td>	
</tr>
<%	set rs_activity_type = con.GetRecordSet("Select activity_type_id, activity_type_name from activity_types WHERE Organization_ID = " & trim(OrgID) & " order by activity_type_name")
    if not rs_activity_type.eof then 
		do while not rs_activity_type.eof
    	activity_type_id = trim(rs_activity_type("activity_type_id"))
    	activity_type_name = trim(rs_activity_type("activity_type_name"))  
    	sqlstr = "Select Count(task_id) From tasks WHERE CHARINDEX('" & activity_type_id & "',task_types) > 0" 	
    	'Response.Write sqlstr
    	'Response.End
    	set rs_count = con.getRecordSet(sqlstr)
    	If not rs_count.eof Then
    		tasks_count = cLng(trim(rs_count(0)))
    	Else
    		tasks_count = 0
    	End If
    	set rs_count = Nothing    	  	
 %>
<tr>
	<td align="center" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" nowrap>
	<%If tasks_count = 0 Then%>
	<a href="activity_types.asp?deleteId=<%=activity_type_id%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת סוג <%=trim(Request.Cookies("bizpowerh")("Tasks_one"))%>"></a>
	<%Else%>
	<%If trim(lang_id) = "1" Then
	    str_alert = "  קיימות משימות עם סוג זה במערכת\n\n" & Space(20) & "על מנת למחוק את הסוג\n\nאליך למחוק קודם את המשימות הנ\'\'ל"
	Else
	    str_alert = "There are tasks with this type in the system,\n\n therefore you can\'t delete this task type"
	End If%>	
	<input type=image SRC="../../images/delete_icon.gif" border=0 Onclick="window.alert('<%=str_alert%>');return false;">
	<%End If%>	
	</td>
	<td align="center" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" nowrap><a href="" onClick="return openactivity_type(<%=activity_type_id%>)" target="_blank"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6">&nbsp;<b><%=activity_type_name%></b>&nbsp;</td>	
</tr>
<%		rs_activity_type.MoveNext
		loop
	end if
	set rs_activity_type=nothing%>
</table>
</td>
<td width=115 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openactivity_type('')"><span id=word6 name=word6><!--הוספת סוג--><%=arrTitles(6)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href="excel_task_types.asp" target=_blank><!--הצג דוח--><%=arrTitles(7)%></a></td></tr>

</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
</table>
</BODY>
</HTML>
