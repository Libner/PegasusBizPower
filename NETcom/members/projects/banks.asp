<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then
		    con.executeQuery("Delete From movement_lines Where movement_id IN (Select movement_id From movements WHERE bank_id = " & delId & ")")
		    con.executeQuery("Delete From movements Where bank_id = " & delId)
			con.executeQuery("Delete From banks Where bank_id = " & delId)
		End If
		Response.Redirect "banks.asp"
	End If
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 40 Order By word_id"				
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
<meta http-equiv=Content-bank content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" bank="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function checkDelete(countComp)
	{
		if(countComp > 0)
		<%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק את חשבון הבנק עם כל התנועות"
		Else
			str_confirm = "Are you sure want to delete the bank account with all movements?"
		End If   
		%>
		<%
		If trim(lang_id) = "1" Then
			str_confirm1 = "?האם ברצונך למחוק את חשבון הבנק"
		Else
			str_confirm1 = "Are you sure want to delete the bank account ?"
		End If   
		%>		
			return window.confirm("<%=str_confirm%>");
		else return window.confirm("<%=str_confirm1%>");	
	}
	
	function openbanks(bankID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("addbank.asp?bank_id=" + bankID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
	function openUpload(bankID)
	{
		h = 200;
		w = 500;
		S_Wind = window.open("uploadExcel.asp?bank_id=" + bankID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
//-->
</SCRIPT>

</HEAD>

<body>
<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr><td width=100%>
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 4%>
<%numOfLink = 7%>
<tr><td width=100%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td width=100% valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff">
<tr>
	<td width="70" nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>	
	<td width="70" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name=word4><!--עדכון--><%=arrTitles(4)%></span>&nbsp;</td>	
	<td width="120" nowrap align="<%=align_var%>" class="title_sort"><span id=word9 name=word9><!--&#8362;--><%=arrTitles(9)%></span>&nbsp;<span id=word7 name=word7><!--מסגרת אשראי--><%=arrTitles(7)%></span>&nbsp;</td>	
	<td width="100%" align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--חשבון--><%=arrTitles(5)%></span>&nbsp;</td>	
</tr>
<%	set rs_banks = con.GetRecordSet("Select bank_id, bank_name, credit from banks WHERE Organization_ID = " & trim(OrgID) & " order by bank_id")
    if not rs_banks.eof then 
	   do while not rs_banks.eof
    	bank_Id = trim(rs_banks("bank_id"))
    	bank_name = trim(rs_banks("bank_name"))
    	credit = trim(rs_banks("credit"))
    	If IsNumeric(credit) Then
    		credit = Round(credit)
    		credit = FormatNumber(credit,0,-1,0,-1)
    	End If    	    	
%>
<tr>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="banks.asp?deleteId=<%=bank_id%>" ONCLICK="return checkDelete(<%=countComp %>)"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת בנק שיחה"></a></td>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="addBank.asp?bank_Id=<%=bank_Id%>" target="_self"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap><%=credit%>&nbsp;</td>	
	<td align="<%=align_var%>" bgcolor="#e6e6e6"><a class="link_categ" href="addBank.asp?bank_Id=<%=bank_Id%>" target="_self"><b><%=bank_name%></b>&nbsp;</a></td>	
</tr>
<%		rs_banks.MoveNext
		loop
	else	
%>
<tr><td align="center" colspan=5 class="title_sort1"><span id=word8 name=word8><!--לא נמצאו חשבונות בנק--><%=arrTitles(8)%></span></td></tr>
<%		
	end if
	set rs_banks=nothing%>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% ID="Table2">
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href="addBank.asp" target=_self><span id="word6" name=word6><!--הוספת חשבון--><%=arrTitles(6)%></span></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr></table>
</BODY>
</HTML>
