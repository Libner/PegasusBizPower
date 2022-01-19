<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then		    
		    con.ExecuteQuery "delete from movement_lines where movement_ID IN(Select movement_ID FROM movements WHERE movement_type_id = " & delId & ")"
			con.ExecuteQuery "delete from movements where movement_type_id=" & delId
			con.executeQuery("Delete From movements_types Where movement_type_id = " & delId)
		End If
		Response.Redirect "movement_types.asp"
	End If
	
	If trim(lang_id) = "1" Then
	     type_name = Array("","הוצאה" ,"הכנסה")    	
	Else
		 type_name = Array("","Expense" ,"Income")
	End If	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 42 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
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
		<%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק את סוג התנועה עם כל התנועות בחשבון"
		Else
			str_confirm = "Are you sure want to delete the type of movement?"
		End If   
		%>		
		return window.confirm("<%=str_confirm%>");		
	}
	
	function openmovement_type(typeID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addmovementType.asp?movement_type_id=" + typeID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
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
  <%numOfLink = 8%>
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
	<td width="100%" align="<%=align_var%>" class="title_sort">&nbsp;<span id=word7 name=word7><!--שם תנועה--><%=arrTitles(7)%></span>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word5" name=word5><!--סוג תנועה--><%=arrTitles(5)%></span>&nbsp;</td>
</tr>
<%	set rs_movement_type = con.GetRecordSet("Select movement_type_id, movement_type_name,type_id from movements_types WHERE Organization_ID = " & trim(OrgID) & " order by type_id,movement_type_name")
    if not rs_movement_type.eof then 
		do while not rs_movement_type.eof
    	movement_type_id = trim(rs_movement_type("movement_type_id"))
    	movement_type_name = trim(rs_movement_type("movement_type_name")) 
    	type_id  = trim(rs_movement_type("type_id"))    	
 %>
<tr>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="movement_types.asp?deleteId=<%=movement_type_id%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" BORDER=0></a></td>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="" onClick="return openmovement_type(<%=movement_type_id%>)" target="_blank"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align="<%=align_var%>" bgcolor="#e6e6e6"><b><%=movement_type_name%></b>&nbsp;</td>	
	<td align="<%=align_var%>" bgcolor="#e6e6e6"><b><%=type_name(type_id)%></b>&nbsp;</td>	
</tr>
<%		rs_movement_type.MoveNext
		loop
	end if
	set rs_movement_type=nothing%>
</table>
</td>
<td width=115 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:110;line-height:110%;padding:3px"  href='#' OnClick="return openmovement_type('')"><span id="word6" name=word6><!--הוספת סוג תנועה--><%=arrTitles(6)%></span></a></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
</table>
</BODY>
</HTML>
