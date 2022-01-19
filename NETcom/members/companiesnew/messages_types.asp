<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	If Request.QueryString("deleteId") <> nil Then
		con.executeQuery "Delete From messages_types Where messages_type_id = " & Request.QueryString("deleteId")
		Response.Redirect "messages_types.asp"
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
	   var str_confirm = "?האם ברצונך למחוק את התבנית"
	   <%Else%>
	   var str_confirm = "Are you sure want to delete the message template?"
	   <%End If%>
		return window.confirm(str_confirm);		
	}
	
	function openmessages_type(typeID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addmessagesType.asp?messages_type_id=" + typeID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	
//-->
</SCRIPT>

</HEAD>

<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>" ID="Table1">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 4%>
  <%numOfLink = 8%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align="<%=align_var%>" valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table2">
<tr>
<td width=100% valign=top>
<table width="100%" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff" ID="Table3">
<tr>
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name=word4><!--עדכון--><%=arrTitles(4)%></span>&nbsp;</td>	
	<td width="100%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--סוג-->תבנית<%'=arrTitles(5)%></span>&nbsp;</td>	
</tr>
<%	set rs_messages_type = con.GetRecordSet("Select messages_type_id, messages_type_name from messages_types WHERE Organization_ID = " & trim(OrgID) & " order by messages_type_name")
    if not rs_messages_type.eof then 
		do while not rs_messages_type.eof
    	messages_type_id = trim(rs_messages_type("messages_type_id"))
    	messages_type_name = trim(rs_messages_type("messages_type_name"))  
    	sqlstr = "Select Count(Message_types) From messages WHERE CHARINDEX('" & messages_type_id & "',Message_types) > 0" 	
    	'Response.Write sqlstr
    	'Response.End
    	set rs_count = con.getRecordSet(sqlstr)
    	If not rs_count.eof Then
    		messages_count = cLng(trim(rs_count(0)))
    	Else
    		messages_count = 0
    	End If
    	set rs_count = Nothing    	  	
 %>
<tr>
	<td align="center" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" nowrap>
	<%If messages_count = 0 Then%>
	<a href="messages_types.asp?deleteId=<%=messages_type_id%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת תבנית הודעה"></a>
	<%Else%>
	<%If trim(lang_id) = "1" Then
	    str_alert = "  קיימות הודעות עם סוג זה במערכת\n\n" & Space(20) & "על מנת למחוק את הסוג\n\nאליך למחוק קודם את הודעות הנ\'\'ל"
	Else
	    str_alert = "There are messages with this type in the system,\n\n therefore you can\'t delete this message type"
	End If%>	
	<input type=image SRC="../../images/delete_icon.gif" border=0 Onclick="window.alert('<%=str_alert%>');return false;" ID="Image1" NAME="Image1">
	<%End If%>	
	</td>
	<td align="center" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" nowrap><a href="" onClick="return openmessages_type(<%=messages_type_id%>)" target="_blank"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6">&nbsp;<b><%=messages_type_name%></b>&nbsp;</td>	
</tr>
<%		rs_messages_type.MoveNext
		loop
	end if
	set rs_messages_type=nothing%>
</table>
</td>
<td width=115 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% ID="Table4">
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openmessages_type('')"><span id=word6 name=word6><!--הוספת סוג--><%'=arrTitles(6)%>הוספת תבנית</span></a></td></tr>

</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
</table>
</BODY>
</HTML>
