<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-type_id content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type_id="text/css">
<script LANGUAGE="JavaScript">
<!--
    function CheckDel()
    {
	<%
		If trim(lang_id) = "1" Then
			str_confirm = "? האם ברצונך למחוק את התנועה"
		Else
			str_confirm = "Are you sure want to delete the movement?"
		End If   
	%>			
    return (window.confirm("<%=str_confirm%>"))    
    }
	
	var oPopup_type = window.createPopup();
	function typeDropDown(obj)
	{
	    oPopup_type.document.body.innerHTML = type_Popup.innerHTML;
	    var popupBodyObjtype = oPopup_type.document.body;	      	 
	    var dHeight = 67;	      			
	    oPopup_type.show(0, obj.scrollHeight + 3, obj.scrollWidth, dHeight, obj); 
	    popupBodyObjtype = null;
	    return true;   
	}
	function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

	function activateMovement(movementID,movement_rank)
	{		
		h = parseInt(300);
		w = parseInt(470);
		window.open("activate_movement.asp?movement_ID=" + movementID + "&movement_rank=" + movement_rank, "Activate" ,"scrollbars=1,toolbar=0,top=150,left=100,width="+w+",height="+h+",align=center,resizable=0");
		return false;			
	}	
	function updateActivation(lineID,movementID)
	{
		h = parseInt(300);
		w = parseInt(470);
		window.open("activate_movement.asp?line_ID=" + lineID + "&movement_ID=" + movementID, "Update" ,"scrollbars=1,toolbar=0,top=150,left=100,width="+w+",height="+h+",align=center,resizable=0");
		return false;		
	}
	function updateMovement(movementID)
	{		
		h = parseInt(300);
		w = parseInt(470);
		window.open("editMovement.asp?movement_ID=" + movementID, "Edit" ,"scrollbars=1,toolbar=0,top=150,left=100,width="+w+",height="+h+",align=center,resizable=0");
		return false;		
	}
//-->	
</script> 
</head> 
<%

start_date = "1/" & Month(Date()) & "/" & Year(date())
end_date = DateAdd("d",30,start_date)

sqlstr = "Select TOP 1 bank_id, bank_name FROM banks WHERE ORGANIZATION_ID = " & OrgID & " Order BY bank_name"
set rs_bank = con.getRecordSet(sqlstr)
if not rs_bank.eof then
	bank_id = rs_bank(0)
end if
set rs_bank = Nothing

If Request("end_date") <> nil And Request("start_date") <> nil Then
	end_date = Request("end_date")
	start_date = Request("start_date")
	con.executeQuery("SET DATEFORMAT dmy")
	'Response.Write "EXECUTE insert_movement_payments '" & DateValue(start_date) & "','" & DateValue(end_date) & "','" & OrgID & "'"
	'Response.End
	con.executeQuery("EXECUTE insert_movement_payments '" & DateValue(start_date) & "','" & DateValue(end_date) & "','" & OrgID & "'")
End If

If Request("bank_id") <> nil Then
	bank_id = trim(Request("bank_id"))
	where_bank = " AND bank_id = " & bank_id	
End If

where_dates = " AND movements_view.payment_date BETWEEN '" & start_date & "' AND '" & end_date & "'"
urlSort  = "movements_activate.asp?start_date=" & start_date & "&end_date=" & end_date & "&bank_id=" & bank_id

If trim(Request("delete")) <> nil Then
	line_id = trim(Request("line_id"))
	con.executeQuery("DELETE From movement_lines WHERE line_id = " & line_id)
	Response.Redirect "movements_activate.asp?start_date=" & start_date & "&end_date=" & end_date & "&bank_id=" & bank_id
End If

dim sortby(12)	
sortby(0) = "movements_view.payment_date"
sortby(1) = "movements_view.payment_date DESC"

sort = Request.QueryString("sort")	
if trim(sort)="" then  sort=0 end if 

%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 60 Order By word_id"				
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
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 5%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title" colspan=2>&nbsp;</td></tr>		   
	<tr>    
    <td width="100%" valign="top" align="center">
    <%If trim(bank_id) <> "" Then%>
	<table width="100%" align="center" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" dir="<%=dir_obj_var%>">	
<%
sqlStr = "SELECT DISTINCT bank_id, bank_name, balance, credit from banks where ORGANIZATION_ID = "& OrgID & " AND bank_id = " & bank_id & " Order BY bank_id"
set rs_banks = con.getRecordSet(sqlStr)
if not rs_banks.eof then
%>
<tr style="line-height:22px">
<td colspan=5 align=center class="title_sort"><span id=word3 name=word3><!--תכנון--><%=arrTitles(3)%></span></td>
<td colspan=4 align=center class="title_sort"><span id="word4" name=word4><!--ביצוע--><%=arrTitles(4)%></span></td>
</tr>
<tr style="line-height:22px">
<td width="53" align="<%=align_var%>" nowrap class="title_sort2"><%if trim(sort)="0" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="<%=arrTitles(18)%>"><%elseif trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="<%=arrTitles(19)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=0" title="<%=arrTitles(19)%>"><%end if%><img src="../../images/arrow_<%if trim(sort)="0" then%>down<%elseif trim(sort)="1" then%>up<%else%>nosort<%end if%>.gif" width="9" height="6" border="0" hspace="2"><span id="word5" name=word5><!--תאריך--><%=arrTitles(5)%></span></a></td>
<td width=20%  align="<%=align_var%>" class="title_sort2">&nbsp;<span id="word6" name=word6><!--תנועה--><%=arrTitles(6)%></span>&nbsp;</td>
<td width=30%  align="<%=align_var%>" class="title_sort2">&nbsp;<span id="word7" name=word7><!--תאור--><%=arrTitles(7)%></span>&nbsp;</td>
<td width="65" nowrap align="<%=align_var%>" class="title_sort2">&nbsp;<span id="word8" name=word8><!--סכום&nbsp;&#8362;--><%=arrTitles(8)%></span></td>	
<td width="20" nowrap align="<%=align_var%>" class="title_sort2"></td>
<td width="53" align="center" nowrap class="title_sort2">&nbsp;<span id="word9" name=word9><!--תאריך--><%=arrTitles(9)%></span>&nbsp;</td>
<td width=50%  align="<%=align_var%>" class="title_sort2">&nbsp;<span id="word10" name=word10><!--פרטי ביצוע--><%=arrTitles(10)%></span>&nbsp;</td>
<td width="65" nowrap align="<%=align_var%>" class="title_sort2">&nbsp;<span id="word11" name=word11><!--סכום&nbsp;&#8362;--><%=arrTitles(11)%></span>&nbsp;</td>	
<td width="22" nowrap align="<%=align_var%>" class="title_sort2"></td>
</tr>
<%
con.executeQuery("SET DATEFORMAT dmy")
sqlStr = "select movements_view.payment_date,movements_view.type_id,movements_view.payment,movements_view.rank_id," &_
" movements_view.movement_type_name,movements_view.movement_id,movements_view.payment_frequency, "&_
" movements_view.parent_id,movements_view.movement_details from movements_view " &_
" where movements_view.ORGANIZATION_ID = "& OrgID & where_dates & where_bank & " AND rank_id IN (0,1) order by " & sortby(sort) 
'Response.Write sqlStr
set rs_payments = con.getRecordSet(sqlStr)
while not rs_payments.eof 
	payment_date = Day(rs_payments(0)) & "/" & Month(rs_payments(0)) & "/" & Right(Year(rs_payments(0)),2)
	type_id = trim(rs_payments(1))
	payment = trim(rs_payments(2)) 
	rank_id = trim(rs_payments(3))	
	movement_type_name = trim(rs_payments(4)) 
	movement_id = trim(rs_payments(5))
	payment_frequency = trim(rs_payments(6))
	parent_id = trim(rs_payments(7))
	movement_details = trim(rs_payments(8))
	
	If trim(movement_id) <> "" Then
		sqlstr = "Select TOP 1 line_id from movement_lines WHERE movement_id = " & movement_id
		set rs_check = con.getRecordSet(sqlstr)
		if not rs_check.eof then
			active_flag = true
		else
			active_flag = false	
		end if
		set rs_check = Nothing
	End If		
%>
<tr>
<td align="center" class="title_sort1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" style="padding-top:3px" valign=top><%=payment_date%></td>
<%
	If trim(type_id) = "2" Then
		style_balance = "style=color:green"
		payment = FormatNumber(payment,0,-1,0,-1)		
	Else						
		style_balance = "style=color:red"
		payment = "&nbsp;-" & FormatNumber(payment,0,-1,0,-1)		
	End If							
%>
<td align="<%=align_var%>" class="card1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" dir="<%=dir_obj_var%>" nowrap style="padding-top:3px" valign=top>&nbsp;<%=movement_type_name%></td>			
<td align="<%=align_var%>" class="card1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" dir="<%=dir_obj_var%>" valign=top style="padding-top:3px">&nbsp;<%=movement_details%></td>		
<td align="<%=align_var%>" class="card1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" style="padding-top:3px" <%=style_balance%> dir=ltr nowrap valign=top><%=payment%>&nbsp;</td>	
<td align="center" class="card1" style="padding-top:3px" valign=top>
<input type=image src="../../images/add_icon.gif" title="הוספת תנועה בפועל" onclick="return activateMovement('<%=movement_id%>','<%=rank_id%>')" vspace=2>
</td>	
<%
  If IsNumeric(movement_id) Then
  sqlStr = "select line_id,payment,movement_type_name, type_id, line_details,payment_date from movement_lines_view where "&_
  " ORGANIZATION_ID = "& OrgID & " AND movement_id = " & movement_id &_
  " order by payment_date" 
  'Response.Write sqlstr  
  set rs_line = con.getRecordSet(sqlStr)
  if not rs_line.eof then
%>
<td colspan=4 bgcolor=#FBF4D8 valign=top>
<table cellpadding=0 cellspacing=0 width=100% style="table-layout:fixed">
<COLGROUP SPAN="4"> 				
<COL width=54 nowrap>
<COL width=*% nowrap>
<COL width=66 nowrap>
<COL width=22 nowrap>			    
</COLGROUP>		
<%
  while not rs_line.eof
	line_id = trim(rs_line(0))
	line_type_name = trim(rs_line(2))
	line_payment = trim(rs_line(1))
	line_type_id = trim(rs_line(3))
	line_details = trim(rs_line(4))	
	line_payment_date = Day(rs_line(5)) & "/" & Month(rs_line(5)) & "/" & Right(Year(rs_line(5)),2)
	If trim(line_type_id) = "2" Then
		style_line = "style=color:green"
		line_payment = FormatNumber(line_payment,0,-1,0,-1)		
	Else						
		style_line = "style=color:red"
		line_payment = "&nbsp;-" & FormatNumber(line_payment,0,-1,0,-1)		
	End If		
%>
<tr height=22>
<td align="center" class="card6 hand" onclick="return updateActivation('<%=trim(line_id)%>','<%=trim(movement_id)%>')"><%=line_payment_date%></td>
<td align="<%=align_var%>" class="card4 hand" style="border-left: 1px solid #FFFFFF;border-right: 1px solid #FFFFFF" dir=ltr nowrap onclick="return updateActivation('<%=trim(line_id)%>','<%=trim(movement_id)%>')"><%=line_details%>&nbsp;</td>			
<td align="<%=align_var%>" class="card4 hand" style="border-left: 1px solid #FFFFFF" <%=style_line%> dir=ltr nowrap onclick="return updateActivation('<%=trim(line_id)%>','<%=trim(movement_id)%>')"><%=sign%>&nbsp;<%=line_payment%>&nbsp;</td>
<td align="center" class="card4"><a href="movements_activate.asp?delete=1&start_date=<%=start_date%>&end_date=<%=end_date%>&bank_id=<%=bank_id%>&line_id=<%=line_id%>" target=_self><IMG src="../../images/delete_icon.gif" title="מחק ביצוע תנועה" onclick="return CheckDel()" border=0></a></td>	
</tr>
<%
  rs_line.moveNext
  Wend 
%>
</table>
</td>  
<% 
  Else
%>
<td width="54" height=22 nowrap align="center" class="card6">&nbsp;</td>
<td width="*%" height=22 align="<%=align_var%>" class="card4" dir=ltr nowrap>&nbsp;</td>			
<td width="65" height=22 nowrap align="<%=align_var%>" class="card4" dir=ltr nowrap>&nbsp;</td>
<td width="22" height=22 nowrap align="center" class="card4">&nbsp;</td>
<%
  End If
  set rs_line = Nothing
  End if
%> 
</tr>
<% rs_payments.moveNext%>
<% Wend	%>
<% set rs_payments = Nothing%>
<% End If	%>
</table>
<%End If%>
</td>
<td width=90 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<FORM action="movements_activate.asp?search=1" method=GET id=form_search name=form_search target="_self">
 <tr>    
    <td width="100%" valign="top" align="center">
    <table cellpadding=1 cellspacing=2 width=100% align=center border=0>
    <tr><td align="<%=align_var%>" colspan=2><b><span id=word12 name=word12><!--מתאריך--><%=arrTitles(12)%></span></b>&nbsp;</td></tr>
    <tr>		
		<td align=center nowrap>
		<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70" onclick="return popupcal(this);" readonly></td>					
	</tr>
	<tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b><span id="word13" name=word13><!--עד תאריך--><%=arrTitles(13)%></span></b>&nbsp;</td></tr>
     <TR>    	
		<td align="center" nowrap>
		<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return popupcal(this);" readonly></td>		
	</TR>		    
   <tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b><span id="word14" name=word14><!--חשבון בנק--><%=arrTitles(14)%></span></b>&nbsp;</td></tr>
   <tr><td>
   <select dir="<%=dir_obj_var%>" name="bank_id" style="width:90px;" class="norm">		
	<%  sqlstr = "Select bank_id, bank_name FROM banks WHERE ORGANIZATION_ID = " & OrgID & " Order BY bank_name"
		set rs_bank = con.getRecordSet(sqlstr)
		While not rs_bank.eof
	%>
	<option value="<%=rs_bank(0)%>" <%If trim(bank_id) = trim(rs_bank(0)) Then%> selected <%End If%>><%=rs_bank(1)%></option>
	<%
		rs_bank.moveNext
		Wend
		set rs_bank = Nothing
	%>
	</select>
</table></td></tr>
   <tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:86;" href="#" onclick="document.form_search.submit();">&nbsp;<span id=word15 name=word15><!--הצג--><%=arrTitles(15)%></span>&nbsp;</a></td></tr>
   <tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>
   </form>
   <tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>  
   <tr><td colspan=2 align="center"><a class="button_edit_1" style="width:86;line-height:120%" href='addMovement.asp?type_id=1' target=_self><span id="word16" name=word16><!--הוספת הוצאה<br>מתוכננת--><%=arrTitles(16)%></span></a></td></tr>
   <tr><td colspan=2 align="center"><a class="button_edit_1" style="width:86;line-height:120%" href='addMovement.asp?type_id=2' target=_self><span id="word17" name=word17><!--הוספת הכנסה<br>מתוכננת--><%=arrTitles(17)%></span></a></td></tr>
   <tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>  
</table></td></tr></table>
</td></tr></table>
</body>
</html>
<%
set con = nothing
%>