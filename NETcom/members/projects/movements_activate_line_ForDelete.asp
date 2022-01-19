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
function CheckDel(str) {
  return (confirm("? האם ברצונך למחוק את התנועה"))    
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
			if (window.confirm('?האם ברצונך לאשר ביצוע של התנועה'))
			{
				h = parseInt(300);
				w = parseInt(450);
				window.open("activate_movement.asp?movement_ID=" + movementID + "&movement_rank=" + movement_rank, "Activate" ,"scrollbars=1,toolbar=0,top=150,left=100,width="+w+",height="+h+",align=center,resizable=0");
				return false;
			}
			return false;	
	}	
	function updateActivation(lineID,movementID)
	{
		if (window.confirm('?האם ברצונך לעדכן את ביצוע התנועה'))
		{
				h = parseInt(300);
				w = parseInt(450);
				window.open("activate_movement.asp?line_ID=" + lineID + "&movement_ID=" + movementID, "Update" ,"scrollbars=1,toolbar=0,top=150,left=100,width="+w+",height="+h+",align=center,resizable=0");
				return false;
		}
		return false;	
	}
	function updateMovement(movementID)
	{
		if (window.confirm('?האם ברצונך לעדכן את תכנון התנועה'))
		{
				h = parseInt(300);
				w = parseInt(450);
				window.open("editMovement.asp?movement_ID=" + movementID, "Edit" ,"scrollbars=1,toolbar=0,top=150,left=100,width="+w+",height="+h+",align=center,resizable=0");
				return false;
		}
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
	where_bank = " AND movements_view.bank_id = " & bank_id	
End If

where_dates = " AND movements_view.payment_date BETWEEN '" & start_date & "' AND '" & end_date & "'"
urlSort  = "movements_activate.asp?start_date=" & start_date & "&end_date=" & end_date & "&bank_id=" & bank_id
urlSortLine  = "movements_activate_line.asp?start_date=" & start_date & "&end_date=" & end_date & "&bank_id=" & bank_id

If trim(Request("delete")) <> nil Then
	line_id = trim(Request("line_id"))
	con.executeQuery("DELETE From movement_lines WHERE line_id = " & line_id)
	Response.Redirect "movements_activate_line.asp?start_date=" & start_date & "&end_date=" & end_date & "&bank_id=" & bank_id
End If

dim sortby(12)	
sortby(2) = "movement_lines.payment_date"
sortby(3) = "movement_lines.payment_date DESC"

sort = Request.QueryString("sort")	
if trim(sort)="" then  sort=0 end if 

%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 5%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title" colspan=2>תכנון מול ביצוע&nbsp;</td></tr>		   
	<tr>    
    <td width="100%" valign="top" align="center">
    <%If trim(bank_id) <> "" Then%>
	<table width="100%" align="center" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">	
<%
sqlStr = "select DISTINCT bank_id, bank_name, balance, credit from banks where ORGANIZATION_ID = "& OrgID & " AND bank_id = " & bank_id & " Order BY bank_id"
set rs_banks = con.getRecordSet(sqlStr)
if not rs_banks.eof then
%>
<tr style="line-height:22px">
<td colspan=4 align=center class="title_sort">ביצוע</td>
<td colspan=5 align=center class="title_sort">תכנון</td>
</tr>
<tr style="line-height:22px">
<td width="22" nowrap align="right" class="title_sort2" dir=rtl></td>
<td width="70" nowrap align="right" class="title_sort2" dir=rtl>&nbsp;סכום&nbsp;&#8362;</td>	
<td width=50%  align="right" class="title_sort2" dir=rtl>&nbsp;פרטי ביצוע&nbsp;</td>
<td width="53" align="right" nowrap class="title_sort2" dir=rtl><%if trim(sort)="2" then%><a class="title_sort" href="<%=urlSortLine%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="3" then%><a class="title_sort" href="<%=urlSortLine%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSortLine%>&sort=2" title="למיון בסדר עולה"><%end if%><img src="../../images/arrow_<%if trim(sort)="2" then%>down<%elseif trim(sort)="3" then%>up<%else%>nosort<%end if%>.gif" width="9" height="6" border="0" hspace="2">תאריך</a></td>
<td width="20" nowrap align="right" class="title_sort2" dir=rtl></td>
<td width="70" nowrap align="right" class="title_sort2" dir=rtl>&nbsp;סכום&nbsp;&#8362;</td>	
<td width=30%  align="right" class="title_sort2" dir=rtl>&nbsp;תאור&nbsp;</td>
<td width=20%  align="right" class="title_sort2" dir=rtl>&nbsp;תנועה&nbsp;</td>
<td width="53" align="right" nowrap class="title_sort2" dir=rtl><%if trim(sort)="0" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=0" title="למיון בסדר עולה"><%end if%><img src="../../images/arrow_<%if trim(sort)="0" then%>down<%elseif trim(sort)="1" then%>up<%else%>nosort<%end if%>.gif" width="9" height="6" border="0" hspace="2">תאריך</a></td>
</tr>
<%con.executeQuery("SET DATEFORMAT dmy")
 sqlStr = "SELECT movement_lines.line_id, movement_lines.payment, movements_view.movement_type_name, " &_
" movements_view.type_id, movement_lines.line_details, movement_lines.payment_date, movements_view.movement_id "&_
" FROM movement_lines RIGHT OUTER JOIN movements_view ON movement_lines.movement_id = movements_view.movement_id " &_
" where movements_view.ORGANIZATION_ID = "& OrgID & where_dates & where_bank & " AND rank_id IN (0,1) order by " & sortby(sort) 
  'Response.Write sqlstr  
  'Response.End
  set rs_line = con.getRecordSet(sqlStr)
  while not rs_line.eof
	line_id = trim(rs_line(0))
	line_type_name = trim(rs_line(2))
	line_payment = trim(rs_line(1))
	line_type_id = trim(rs_line(3))
	line_details = trim(rs_line(4))	
	If IsDate(rs_line(5)) Then
		line_payment_date = Day(rs_line(5)) & "/" & Month(rs_line(5)) & "/" & Right(Year(rs_line(5)),2)
	Else
		line_payment_date = ""
	End If
	movement_id = trim(rs_line(6))	
	If trim(line_type_id) = "2" And trim(line_payment) <> "" Then
		style_line = "style=color:green"
		sign = ""					
	ElseIf trim(line_type_id) = "1" And trim(line_payment) <> "" Then						
		style_line = "style=color:red"		
		sign = "-"	
	End If		
%>
<tr>
<%If trim(line_id) <> "" Then%>
<td align="center" class="card4"><a href="movements_activate.asp?delete=1&start_date=<%=start_date%>&end_date=<%=end_date%>&bank_id=<%=bank_id%>&line_id=<%=line_id%>" target=_self><IMG src="../../images/delete_icon.gif" title="מחק ביצוע תנועה" onclick="return window.confirm('?האם ברצונך למחוק את ביצוע התנועה')" border=0></a></td>	
<td align="right" class="card4 hand" <%=style_line%> nowrap onclick="return updateActivation('<%=trim(line_id)%>','<%=trim(movement_id)%>')"><%=sign%>&nbsp;<%=line_payment%>&nbsp;</td>
<td class="card4 hand" dir=rtl onclick="return updateActivation('<%=trim(line_id)%>','<%=trim(movement_id)%>')"><%=line_details%>&nbsp;</td>			
<td align="center" class="card6 hand" onclick="return updateActivation('<%=trim(line_id)%>','<%=trim(movement_id)%>')"><%=line_payment_date%></td>
<%Else%>
<td align="center" class="card4">&nbsp;</td>	
<td align="right" class="card4">&nbsp;</td>
<td class="card4">&nbsp;</td>			
<td align="center" class="card6">&nbsp;</td>

<%End if%>
<%
con.executeQuery("SET DATEFORMAT dmy")
sqlStr = "select payment_date,type_id,payment,rank_id,movement_type_name,movement_id,payment_frequency,parent_id,movement_details from movements_view where "&_
" ORGANIZATION_ID = "& OrgID & " AND rank_id IN (0,1) AND movement_id = " & movement_id
'Response.Write sqlStr
set rs_payments = con.getRecordSet(sqlStr)
if not rs_payments.eof then
	payment_date = Day(rs_payments(0)) & "/" & Month(rs_payments(0)) & "/" & Right(Year(rs_payments(0)),2)
	type_id = trim(rs_payments(1))
	payment = trim(rs_payments(2)) 
	rank_id = trim(rs_payments(3))	
	movement_type_name = trim(rs_payments(4)) 
	movement_id = trim(rs_payments(5))
	payment_frequency = trim(rs_payments(6))
	parent_id = trim(rs_payments(7))
	movement_details	= trim(rs_payments(8))
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
<%
	If trim(type_id) = "2" Then
		style_balance = "style=color:green"
		payment = FormatNumber(payment,0,-1,0,-1)		
	Else						
		style_balance = "style=color:red"
		payment = "&nbsp;-" & FormatNumber(payment,0,-1,0,-1)		
	End If							
%>
<td align="center" class="card1" style="padding-top:3px" valign=top>
<input type=image src="../../images/add_icon.gif" title="אשר ביצוע תנועה" onclick="return activateMovement('<%=movement_id%>','<%=rank_id%>')" vspace=2>
</td>
<td align="right" class="card1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" style="padding-top:3px" <%=style_balance%> dir=ltr nowrap valign=top><%=payment%>&nbsp;</td>	
<td align="right" class="card1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" dir=rtl valign=top style="padding-top:3px">&nbsp;<%=movement_details%></td>		
<td align="right" class="card1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" dir=rtl nowrap style="padding-top:3px" valign=top>&nbsp;<%=movement_type_name%></td>			
<td align="center" class="title_sort1 hand" onclick="return updateMovement('<%=trim(movement_id)%>')" style="padding-top:3px" valign=top><%=payment_date%></td>
	
<% set rs_payments = Nothing
   End If
%>
</tr>
<%
   rs_line.moveNext
   Wend 
   set rs_line = Nothing
   End If
  End If 
%>
</table></td>
<td width=90 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<FORM action="movements_activate_line.asp?search=1" method=GET id=form_search name=form_search target="_self">
 <tr>    
    <td width="100%" valign="top" align="center">
    <table cellpadding=1 cellspacing=2 width=100% align=center border=0>
    <tr><td align="right" colspan=2><b>מתאריך</b>&nbsp;</td></tr>
    <tr>		
		<td align=center nowrap>
		<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70" onclick="return popupcal(this);" readonly></td>					
	</tr>
	<tr><td width=100% align=right colspan=2>&nbsp;<b>עד תאריך</b>&nbsp;</td></tr>
     <TR>    	
		<td align="center" nowrap>
		<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return popupcal(this);" readonly></td>		
	</TR>		    
   <tr><td width=100% align=right colspan=2>&nbsp;<b>חשבון בנק</b>&nbsp;</td></tr>
   <tr><td>
   <select dir="rtl" name="bank_id" style="width:90px;" class="norm">		
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
   <tr><td align="right" colspan=2 height="5" nowrap></td></tr>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:86;" href="#" onclick="document.form_search.submit();">&nbsp;הצג&nbsp;</a></td></tr>
   <tr><td align="right" colspan=2 height="5" nowrap></td></tr>
   </form>
   <tr><td align="right" colspan=2 height="5" nowrap></td></tr>  
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:86;" href='addMovement.asp?type_id=1' target=_self>הוספת הוצאה</a></td></tr>
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:86;" href='addMovement.asp?type_id=2' target=_self>הוספת הכנסה</a></td></tr>
   <tr><td align="right" colspan=2 height="5" nowrap></td></tr>  
</table></td></tr></table>
</td></tr></table>
</body>
</html>
<%
set con = nothing
%>