<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then		   
			con.executeQuery("Delete From MESSANGERS Where item_id = " & delId)
		End If
		Response.Redirect "messangers.asp"
	End If
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
			str_confirm = "?��� ������ ����� �� ������"
		Else
			str_confirm = "Are you sure want to delete the position ?"
		End If   
		%>			
		return window.confirm("<%=str_confirm%>");		
	}
	
	function openmessanger(itemID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addmessanger.asp?item_id=" + itemID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	
//-->
</SCRIPT>
</HEAD>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 31 Order By word_id"				
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
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 4%>
  <%numOfLink = 1%>
<%topLevel2 = 32 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align="<%=align_var%>" valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td width=100% valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff">
<tr>
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--�����--><%=arrTitles(3)%></span>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name=word4><!--�����--><%=arrTitles(4)%></span>&nbsp;</td>	
	<td width="100%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--�����--><%=arrTitles(5)%></span>&nbsp;</td>	
</tr>
<%	set rs_contact_type = con.GetRecordSet("Select item_id, item_name from messangers WHERE Organization_ID = " & trim(OrgID) & " order by item_name")
    if not rs_contact_type.eof then 
		do while not rs_contact_type.eof
    	item_id = trim(rs_contact_type("item_id"))
    	item_name = trim(rs_contact_type("item_name"))    	
%>
<tr>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="messangers.asp?deleteId=<%=item_id%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="����� �����"></a></td>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="" onClick="return openmessanger(<%=item_id%>)" target="_blank"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align="<%=align_var%>" bgcolor="#e6e6e6" dir="<%=dir_obj_var%>">&nbsp;<b><%=item_name%></b>&nbsp;</td>	
</tr>
<%		rs_contact_type.MoveNext
		loop
	end if
	set rs_contact_type=nothing%>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openmessanger('')"><span id=word6 name=word6><!--����� �����--><%=arrTitles(6)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href="excel_messangers.asp" target=_blank><!--��� ���--><%=arrTitles(7)%></a></td></tr>

</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
</table>
</BODY>
</HTML>
