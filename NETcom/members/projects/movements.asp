<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-movement_type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" movement_type="text/css">
<script LANGUAGE="JavaScript">
<!--
function CheckDel(str) {
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
	    var dHeight = 111;	      			
	    oPopup_type.show(obj.scrollWidth-170, obj.scrollHeight + 3, 170, dHeight, obj); 
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

//-->	
</script> 
</head> 
<%
	type_id = trim(Request.QueryString("type_id"))
	If trim(type_id) <> "" Then
		where_type = " AND type_id = " & type_id
	Else
		where_type = ""
	End If
	
	If trim(type_id) = "1" Then ' הוצאות
		where_type = " AND type_id  = 1"
		else_type_id = "2"
	ElseIf trim(type_id) = "2" Then ' הכנסות
		where_type = " AND type_id = 2"
		else_type_id = "1"
	End If

	If trim(lang_id) = "1" Then
	     type_name = Array("","הוצאה" ,"הכנסה")
	     type_name_r  = Array("","הוצאות" ,"הכנסות") 
	     frequency_arr = Array("חד-פעמי","יום","שבוע","חודש","חודשיים","רבעון")	 
	Else
		 type_name = Array("","Expense" ,"Income")
		 type_name_r = Array("","Expenses" ,"Incomes")
		 frequency_arr = Array("One-time","Day","Week","Month","2 Month","Quarter")	 
	End If	

movement_type = trim(Request("movement_type_id"))
If trim(movement_type) <> "" Then
	where_movement_type = " AND movement_type_id = " & movement_type
Else
	where_movement_type = ""
End If

if Request.QueryString("delmovement_ID")<>nil then			
    'con.ExecuteQuery "delete from movement_lines where movement_ID=" & Request.QueryString("delmovement_ID")
    con.ExecuteQuery "delete from movements where parent_ID=" & Request.QueryString("delmovement_ID") & " And movement_ID NOT IN (Select DISTINCT movement_id FROM movement_lines WHERE ORGANIZATION_ID = " & OrgID & ")"
	con.ExecuteQuery "delete from movements where movement_ID=" & Request.QueryString("delmovement_ID")
end if

start_date = "1/" & Month(Date()) & "/" & Year(date())
end_date = DateAdd("d",30,start_date)

If Request("end_date") <> nil And Request("start_date") <> nil Then
	end_date = Request("end_date")
	start_date = Request("start_date")
End If

where_dates = " AND payment_date BETWEEN '" & start_date & "' AND '" & end_date & "'"

urlSort = "movements.asp?type_id=" & type_id & "&movement_type_id=" & movement_type_id & "&start_date=" & start_date & "&end_date=" & end_date
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 54 Order By word_id"				
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
<%If trim(type_id) = "1" Then%>
<%numOfLink = 3%>
<%Else%>
<%numOfLink = 4%>
<%End If%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;</td></tr>	  		       	
   </table></td></tr>         
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
   <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" align="center" border="0" cellpadding="0" bgcolor="#FFFFFF" cellspacing="1" dir="<%=dir_var%>">
	<tr style="line-height:17px">
	<td width="45" nowrap align="center" class="title_sort"><span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span></td>	
	<td width="70" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort"><span id="word5" name=word5><!--&#8362;--><%=arrTitles(5)%></span>&nbsp;<span id="word4" name=word4><!--סכום--><%=arrTitles(4)%></span>&nbsp;</td>
	<td width="135" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" name="type_td" id="type_td">&nbsp;&nbsp;<%=type_name(type_id)%>&nbsp;&nbsp;<IMG id="find_stat" style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 ALT="בחר מרשימה" align=absmiddle onmousedown="typeDropDown(window.document.all('type_td'))"></td>	
	<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;&nbsp;<span id="word6" name=word6><!--חשבון בנק--><%=arrTitles(6)%></span></td>
	<td width="50%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;&nbsp;<span id="word7" name=word7><!--פרטים--><%=arrTitles(7)%></span></td>
	<td width="80" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort"><span id="word8" name=word8><!--תדירות--><%=arrTitles(8)%></span></td>
	<td width="55" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort"><span id="word9" name=word9><!--תאריך--><%=arrTitles(9)%></span></td>
	</tr>
<%
 if trim(Request.QueryString("page"))<>"" then
    page=Request.QueryString("page")
 else
    Page=1
 end if  
 
 if trim(Request.QueryString("numOfRow"))<>"" then
    numOfRow=Request.QueryString("numOfRow")
 else
    numOfRow = 1
 end if  
 
 PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
 If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
    PageSize = 10
 End If	
 
 con.executeQuery("SET DATEFORMAT dmy")
 sqlStr = "select movement_ID, movement_type_name,movement_type_id,rank_id,bank_id,bank_name,payment,payment_date,movement_details,payment_frequency from movements_view where ORGANIZATION_ID = "& OrgID 
 sqlStr = sqlStr & where_type & where_dates & where_movement_type & " AND rank_id IN (1,2) order by payment_date" 
 'Response.Write sqlStr
 'Response.End
 set rs_movementS = con.GetRecordSet(sqlStr)
 if not rs_movementS.eof then
	rs_movementS.PageSize=PageSize
	rs_movementS.AbsolutePage=Page
	recCount=rs_movementS.RecordCount 
	NumberOfPages=rs_movementS.PageCount
	i=1		    
do while (not rs_movementS.EOF and i<=rs_movementS.PageSize)
	movement_ID = rs_movementS("movement_ID")
	bank_id = rs_movementS("bank_id")	
	bank_name = rs_movementS("bank_name")
	rank_id	 = rs_movementS("rank_id")	
	movement_type_name = rs_movementS("movement_type_name")		
	payment_date = trim(rs_movementS("payment_date"))
	payment_date = Day(payment_date) & "/" & Month(payment_date) & "/" & Right(Year(payment_date),2)
	payment = trim(rs_movementS("payment"))
	payment_frequency = trim(rs_movementS("payment_frequency"))
	movement_type_id = rs_movementS("movement_type_id")			
	movement_details = trim(rs_movementS("movement_details"))
	If cInt(strScreenWidth) > 800 Then
			numOfLetters = 150
	Else
			numOfLetters = 30
	End If	
	If IsNumeric(payment) Then
    	payment = Round(payment)
    	payment = FormatNumber(payment,0,-1,0,-1)
    End If   
	If Len(movement_details) > numOfLetters Then
		movement_details_short = Left(movement_details , numOfLetters-2) & ".."
	Else movement_details_short = movement_details	
	End If   
%>
	<tr>
		<td align="center" valign="top" dir="<%=dir_obj_var%>" class="card"><A href="movements.asp?delmovement_ID=<%=movement_ID%>&type_id=<%=type_id%>" ONCLICK="return CheckDel()"><img src="../../images/delete_icon.gif" border="0" alt="מחיקה"></A></td>		
	    <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="card"><A class="link_categ" href="addmovement.asp?movement_ID=<%=movement_ID%>&type_id=<%=type_id%>&rank_id=<%=rank_id%>" target=_self>&nbsp;<%=payment%>&nbsp;</A></td>
	    <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="card"><A class="link_categ" href="addmovement.asp?movement_ID=<%=movement_ID%>&type_id=<%=type_id%>&rank_id=<%=rank_id%>" target=_self>&nbsp;<%=movement_type_name%>&nbsp;</A></td>		
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="card"><A class="link_categ" href="addmovement.asp?movement_ID=<%=movement_ID%>&type_id=<%=type_id%>&rank_id=<%=rank_id%>" target=_self>&nbsp;<%=bank_name%>&nbsp;</A></td>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="card"><A class="link_categ" href="addmovement.asp?movement_ID=<%=movement_ID%>&type_id=<%=type_id%>&rank_id=<%=rank_id%>" target=_self title="<%=vFix(movement_details)%>">&nbsp;<%=movement_details_short%>&nbsp;</A></td>
		<td align="center" valign="top" dir="<%=dir_obj_var%>" class="card"><A class="link_categ" href="addmovement.asp?movement_ID=<%=movement_ID%>&type_id=<%=type_id%>&rank_id=<%=rank_id%>" target=_self>&nbsp;<%=frequency_arr(payment_frequency)%>&nbsp;</A></td>
		<td align="center" valign="top" dir="<%=dir_obj_var%>" class="card"><A class="link_categ" href="addmovement.asp?movement_ID=<%=movement_ID%>&type_id=<%=type_id%>&rank_id=<%=rank_id%>" target=_self>&nbsp;<%=payment_date%>&nbsp;</A></td>
	</tr>
<%
rs_movementS.movenext
i=i+1
loop
set rs_movementS = nothing
if NumberOfPages > 1 then
urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" ID="Table4">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(17)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(18)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>			
	<%
	End if
	%>
	<tr>
	   <td colspan="10" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6F6DA6;font-weight:600"><span id="word10" name=word10><!--נמצאו--><%=arrTitles(10)%></span>&nbsp;<%=recCount%>&nbsp;<%=type_name_r(type_id)%></td>
	</tr>
	<% 
else
%>
<tr><td align="center" class="title_sort1" colspan=7><span id="word11" name=word11><!--לא נמצאו--><%=arrTitles(11)%></span> <%=type_name_r(type_id)%></td></tr>
<%end if%>
</table>
</td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<FORM action="movements.asp?type_id=<%=type_id%>" method=POST id=form_search name=form_search target="_self">
 <tr>    
    <td width="100%" valign="top" align="center">
    <table cellpadding=1 cellspacing=2 width=100% align=center border=0 ID="Table2">
    <tr><td align="<%=align_var%>" colspan=2><b><span id="word12" name=word12><!--מתאריך--><%=arrTitles(12)%></span></b>&nbsp;</td></tr>
    <tr>
		<!--td width=20 nowrap>
		<input type=image src="../../images/delete_icon.gif" border=0 onClick="start_date.value='';return false;" title="מחק תאריך" hspace=0 vspace=0 id=image2 name=image2></td-->
		<td align=center nowrap>
		<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70"onclick="return popupcal(this);" readonly></td>					
	</tr>
	<tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b><span id="word13" name=word13><!--עד תאריך--><%=arrTitles(13)%></span></b>&nbsp;</td></tr>
     <TR>
    	<!--td align="<%=align_var%>"><input type=image hspace=0 vspace=0 src="../../images/delete_icon.gif" border=0 onClick="end_date.value='';return false;" title="מחק תאריך" id=image1 name=image1></td-->
		<td align="center" nowrap>
		<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return popupcal(this);" readonly></td>		
	</TR>		
    </table>
    </td>
   </tr> 
   <tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:86;" href="#" onclick="document.form_search.submit();">&nbsp;<span id="word14" name=word14><!--עדכן תאריך--><%=arrTitles(14)%></span>&nbsp;</a></td></tr>
 </form>
<tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>  
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:86;" href='addMovement.asp?type_id=<%=type_id%>' target=_self><span id="word15" name=word15><!--הוספת--><%=arrTitles(15)%></span> <%=type_name(type_id)%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table>
</td></tr></table>
</td></tr></table>
<DIV ID="type_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="overflow:auto; position:absolute; width:170; top:0; left:0; height:111; border:1px solid black;" >	
<%
		sqlstr = "Select movement_type_id, movement_type_name from movements_types "
		sqlstr = sqlstr & " WHERE ORGANIZATION_ID = " & OrgID & " AND type_id = " & type_id
		sqlstr = sqlstr & " Order BY movement_type_name"
		set rs_mov = con.getRecordSet(sqlstr)
		while not rs_mov.eof
%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border-bottom:1px solid black; padding:3px; cursor:hand"
	ONCLICK="parent.location.href='movements.asp?movement_type_id=<%=rs_mov(0)%>&type_id=<%=type_id%>'">
    <%=rs_mov(1)%>
    </DIV>
<%
		rs_mov.moveNext
		Wend
		set rs_mov = Nothing
%>          
    <DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; padding:3px; cursor:hand"
	ONCLICK="parent.location.href='movements.asp?type_id=<%=type_id%>'">
    <span id=word16 name=word16><!--הכל--><%=arrTitles(16)%></span>
    </DIV>
</div>
</DIV>
</body>
</html>
<%
set con = nothing
%>