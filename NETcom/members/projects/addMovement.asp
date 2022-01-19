<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%If trim(lang_id) = "1" Then
	Session.LCID = 1037 
Else
	Session.LCID = 2057
End If	
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckFields()
{	
	if ( document.frmMain.movement_type_id.value=='' ||
	     document.frmMain.bank_id.value=='' ||
	     document.frmMain.payment.value==''	     
	     )		   
		{
				<%
				If trim(lang_id) = "1" Then
					str_alert = "'*' אנא מלאו כל השדות הצוינות בסימן"
				Else
					str_alert = "Please insert all requested fields !!"
				End If   
				%>			
				window.alert("<%=str_alert%>");		
				return false;				
		}
	else
		{
			document.frmMain.submit();
			return true;
		}		
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


function GetNumbers()
{
	var ch=event.keyCode;
	event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
} 
function populate(objForm,selectIndex) 
{
	timeA = new Date(objForm.paymentYear.options[objForm.paymentYear.selectedIndex].text, objForm.paymentMonth.options[objForm.paymentMonth.selectedIndex].value,1);
	timeDifference = timeA - 86400000;
	timeB = new Date(timeDifference);
	var daysInMonth = timeB.getDate();
	for (var i = 0; i < objForm.paymentDay.length; i++) {
	objForm.paymentDay.options[0] = null;
	}
	for (var i = 0; i < daysInMonth; i++) {
	objForm.paymentDay.options[i] = new Option(i+1);
	}
	document.frmMain.paymentDay.options[0].selected = true;
}

function getYears()
{
	// You can easily customize what years can be used
	for (var i = 0; i < document.frmMain.paymentYear.length; i++) {
	document.frmMain.paymentYear.options[0] = null;
	}
	timeC = new Date();
	currYear = timeC.getFullYear();
	for (var i = 0; i < years.length; i++) {
	document.frmMain.paymentYear.options[i] = new Option(years[i]);
	}
	document.frmMain.paymentYear.options[2].selected=true;
}
window.onLoad = getYears;
//-->
</script>
<%
movement_ID = Request("movement_ID")
type_id = trim(Request("type_id"))
movement_rank = trim(Request("movement_rank"))
If trim(movement_rank) = "" Then
	movement_rank = "1" ' קבועות
End If

If trim(type_id) = "1" Then ' הוצאות
	where_type = " AND type_id  = 1"
	title_ = "הוצאה"
	page_title_ = title_ & " מתוכננת"
ElseIf trim(type_id) = "2" Then ' הכנסות
	where_type = " AND type_id = 2"
	title_ = "הכנסה"
	page_title_ = title_ & " מתוכננת"
End If

If trim(lang_id) = "1" Then
	     type_name = Array("","הוצאה" ,"הכנסה")
	     type_name_r  = Array("","הוצאות" ,"הכנסות") 
	     frequency_arr = Array("","יום","שבוע","חודש","חודשיים","רבעון")	  
Else
		 type_name = Array("","Expense" ,"Income")
		 type_name_r = Array("","Expenses" ,"Incomes")
		 frequency_arr = Array("","Day","Week","Month","2 Month","Quarter")
End If	

if request.form("movement_type_id")<>nil then 'after form filling
	 	 
	 If Request.Form("movement_type_id") <> nil Then
		movement_type_id = trim(Request.Form("movement_type_id"))	
	 Else
		movement_type_id = "NULL"
	 End If	
	 movement_details  = sFix(trim(Request.Form("movement_details")))
	 If Request.Form("type_id") <> nil Then
		type_id = trim(Request.Form("type_id"))
	 Else
		type_id = "NULL"
	 End If
	 If Request.Form("movement_rank") <> nil Then
		rank_id = trim(Request.Form("movement_rank"))
	 Else
		rank_id = "NULL"
	 End If
	 If Request.Form("bank_id") <> nil Then		
		bank_id = trim(Request.Form("bank_id"))
	 Else
		bank_id = "NULL"
	 End If	
	 If Request.Form("payment_frequency") <> nil Then	
		payment_frequency = trim(Request.Form("payment_frequency"))
	 Else
		payment_frequency = "NULL"
	 End If	
	 payment = trim(Request.Form("payment"))
	 
	 If Request.Form("paymentDay") <> nil Then
		 payment_date = "'" & Request.Form("paymentDay") & "/" & Request.Form("paymentMonth") & "/" & Request.Form("paymentYear") & "'"
	 Else
		 payment_date = "NULL"
	 End If
	 
	 If Request.Form("payments_number") <> nil Then
		 payments_number = Request.Form("payments_number")
	 Else
		 payments_number = "NULL"
	 End If
	 
	 If Request.Form("payment_percent") <> nil Then
		 payment_percent = Request.Form("payment_percent")
	 Else
		 payment_percent = "NULL"
	 End If	
	 
	 if movement_ID=nil or movement_ID="" then 'new record in DataBase
		
			con.executeQuery("SET DATEFORMAT dmy")
			sqlStr = "insert into movements (movement_type_id,type_id,payment_frequency,movement_details,rank_id,payment_date,payment,organization_id,bank_id,payments_number,payment_percent) "
			sqlStr=sqlStr& " values (" & movement_type_id & "," & type_id &","& payment_frequency &",'"& movement_details &"'," & rank_id & "," & payment_date & "," & payment & "," & OrgID & "," & bank_id & "," & payments_number & "," & payment_percent & ")"
			'Response.Write sqlStr
			'Response.End
			con.GetRecordSet(sqlStr)		
			
			If isNumeric(wizard_id) Then
				Response.Redirect ("../wizard/wizard_"&wizard_id&"_4.asp?wmovements=true") 				
			Else
				Response.Redirect ("movements.asp?type_id=" & type_id)          
			End If		
		    
	 else			
			con.executeQuery("SET DATEFORMAT dmy")
			sqlStr = "Update movements set payment_frequency = " & payment_frequency &_
			", movement_details = '" & movement_details &_
			"', bank_id = " & bank_id & ", type_id = " & type_id &_
			", movement_type_id = " & movement_type_id &_
			", payment_date = " & payment_date &_
			", rank_id = " & rank_id &_
			", payments_number = " & payments_number & ", payment_percent = " & payment_percent &_
			", payment ='" & payment & "' where movement_ID=" & movement_ID
			'Response.Write sqlstr
			'Response.End
			con.GetRecordSet (sqlStr)			
		
			If isNumeric(wizard_id) Then
				Response.Redirect ("../wizard/wizard_"&wizard_id&"_4.asp?wmovements=true") 				
			Else
				Response.Redirect ("movements.asp?type_id=" & type_id)          
			End If			
    end if        
end if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 55 Order By word_id"				
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
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title"><%If trim(wizard_id) = "8" Then%> <%=page_title_%> <%Else%>&nbsp;&nbsp;<%End If%></td></tr>		   
	  		       	
   </table></td></tr>  
<% wizard_page_id = 3 %>
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<%If trim(wizard_id) = "8" Then%>
<tr>
<td width=100% align="<%=align_var%>" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width=100% ID="Table1">
<tr><td class="explain_title">צור טופס רישום חדש!</td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td class=explain>
בשדה <b>שפת הטופס</b> , בחר את השפה בה ייכתב טופס ההרשמה.
</td></tr>
<tr><td class=explain>
בשדה <b>שם הטופס</b> הזן שם לטופס הרישום. שם זה לא יופיע בדיוור אלא ישמש לזיהוי הטופס במערכת ה-Bizpower.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט מקדים</b> , הקלד את הטקסט שיופיע בתחילת הטופס לפני שדות הרישום.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט תודה לאחר הרישום</b> , הקלד את הטקסט שברצונך שהנרשם יראה לאחר שליחת הטופס. <br>לדוגמה: "תודה! נציגינו יחזרו אליך ב-24 השעות הקרובות"
</td></tr>
</table>
</td>
</tr>
<tr>
    <td width="100%" height="2"></td>
</tr>

<%End If%>        
<tr><td width=100% bgcolor="#E6E6E6" dir="<%=dir_var%>"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="500">    
<%
if movement_ID<>nil and movement_ID<>"" then
  sqlStr = "select movement_type_id, type_id, payment_frequency, bank_id, movement_details, start_date, end_date, payment_date, payment, payments_number, payment_percent,rank_id from movements where movement_ID=" & movement_ID  
  ''Response.Write sqlStr
  set rs_movements = con.GetRecordSet(sqlStr)
	if not rs_movements.eof then
		type_id = rs_movements("type_id")
		movement_type_id = rs_movements("movement_type_id")
		payment_frequency = rs_movements("payment_frequency")
		bank_id = rs_movements("bank_id")
		movement_details = rs_movements("movement_details")
		paymentDate = rs_movements("payment_date")
		payment = rs_movements("payment")
		dateStart = rs_movements("start_date")
		dateEnd = rs_movements("end_date")	
		payments_number = rs_movements("payments_number")
		payment_percent = rs_movements("payment_percent")
		movement_rank = rs_movements("rank_id")	
 	end if
	set rs_movements = nothing
	
	sqlstr = "Select TOP 1 movement_id from movements WHERE parent_id = " & movement_ID	
	set rs_check = con.getRecordSet(sqlstr)
	if not rs_check.eof then
		is_parent = true
	else	
	    is_parent = false
	end if
	
else
	dateStart = Day(paymentDate) & "/" & Month(paymentDate) & "/" & Year(paymentDate)
	paymentDate = Date()
	payments_number = "1"
end if
'Response.Write type_id
%>
<FORM name="frmMain" ACTION="addMovement.asp" METHOD="post" onSubmit="return CheckFields()">
<tr>
	<td align="<%=align_var%>">
	<input type=hidden name="type_id" value="<%=type_id%>">	
	<input type=hidden name="movement_ID" value="<%=movement_ID%>" ID="movement_ID">	
	</td>	
</tr>
<tr>
	<td align="<%=align_var%>">
	<select dir="<%=dir_obj_var%>" name="movement_type_id" style="width:300px;" class="norm" ID="movement_type_id">
	<option value="" id=word3 name=word3><%=arrTitles(3)%></option>
	<%
		sqlstr = "Select movement_type_id, movement_type_name from movements_types "
		sqlstr = sqlstr & " WHERE ORGANIZATION_ID = " & OrgID & " AND type_id = " & type_id & where_rank
		sqlstr = sqlstr & " Order BY movement_type_name"
		set rs_mov = con.getRecordSet(sqlstr)
		while not rs_mov.eof
	%>
	<option value="<%=rs_mov(0)%>" <%If trim(movement_type_id) = trim(rs_mov(0)) Then%> selected <%End If%>><%=rs_mov(1)%></option>
	<%
		rs_mov.moveNext
		Wend
		set rs_mov = Nothing
	%>
	</select>
	</td>
	<td align="<%=align_var%>" class="textp" nowrap width="150"><span id=word4 name=word4><!--סוג--><%=arrTitles(4)%></span>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>">
	<select dir="<%=dir_obj_var%>" name="bank_id" style="width:300px;" class="norm">	
	<option value="" id=word5 name=word5><%=arrTitles(5)%></option>
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
	</select></td>
	<td align="<%=align_var%>" class="textp" nowrap width="150"><span id=word6 name=word6><!--חשבון בנק--><%=arrTitles(6)%></span>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" dir="<%=dir_var%>"><%=type_name(type_id)%> <span id="word7" name=word7><!--חד פעמית--><%=arrTitles(7)%></span>&nbsp;
	<input type=radio dir="<%=dir_obj_var%>" name="movement_rank" <%If trim(movement_rank) = "" Or trim(movement_rank) = "1" Then%> checked <%End If%>  onclick="movement_cyclic.style.display='none';payment_frequency.disabled=true;" value=1 ID="movement_rank"></td>
	<td align=left nowrap width="150"></td>
</tr>
<tr>
	<td align="<%=align_var%>" dir="<%=dir_var%>">
	<table cellpadding=0 cellspacing=0 align="<%=align_var%>">	
	<tr>
		<td align="<%=align_var%>" name="movement_cyclic" id="movement_cyclic" <%If trim(movement_rank) = "2" Then%> style="display:inline" <%Else%> style="display:none" <%end if%>>
		<select dir="<%=dir_obj_var%>" name="payment_frequency" style="width:140px;" class="norm" ID="Select1">	
		<option value=0 <%If trim(payment_frequency) = "" Then%> selected <%End If%> id=word12 name=word12><%=arrTitles(12)%></option>
		<%For i=1 To Ubound(frequency_arr)%>
		<option value="<%=i%>" <%If trim(payment_frequency) = trim(i) Then%> selected <%End If%>><%=frequency_arr(i)%></option>
		<%Next%>
		</select>&nbsp;&nbsp;&nbsp;	
		</td>
		<td nowrap>&nbsp;<%=type_name(type_id)%> <span id="word8" name=word8><!--מחזורית--><%=arrTitles(8)%></span>&nbsp;<input type=radio dir="<%=dir_obj_var%>" name="movement_rank" <%If trim(movement_rank) <> "" And trim(movement_rank) = "2" Then%> checked <%End If%> onclick="movement_cyclic.style.display='inline';payment_frequency.disabled=false;" value=2></td>				
	</tr>
	
	</table>	
	</td>
	<td align=left nowrap width="150"></td>		
</tr>
<tr>
	<td align="<%=align_var%>">	
	<table cellpadding=0 cellspacing=0 width=100%>
	<tr>
		<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>">
		  <span id=word13 name=word13><!--יום--><%=arrTitles(13)%></span>&nbsp;&nbsp;
		  <SELECT NAME="paymentDay" CLASS="norm" ID="paymentDay" dir="<%=dir_obj_var%>" style="width:60">
	      <% For counter = 1 to 31 %>
	        <OPTION VALUE="<%=counter%>" <% If (DatePart("d", paymentDate) = counter) Then Response.Write "SELECTED"%>><%=counter%></OPTION>
	      <% Next %>
	      </SELECT>		     
		   &nbsp;&nbsp;<span id="word14" name=word14><!--חודש--><%=arrTitles(14)%></span>&nbsp;&nbsp;
		  <SELECT NAME="paymentMonth" CLASS="norm" ID="paymentMonth" dir="<%=dir_obj_var%>" style="width:90" onChange="populate(this.form,this.selectedIndex);">
	      <% For counter = 1 to 12 %>
	        <OPTION VALUE="<%=counter%>" <% If (DatePart("m", paymentDate) = counter) Then Response.Write "SELECTED"%>><%=MonthName(counter)%></OPTION>
	      <% Next %>
	      </SELECT>		    
		  &nbsp;&nbsp;<span id="word15" name=word15><!--שנה--><%=arrTitles(15)%></span>&nbsp;&nbsp;
		  <SELECT NAME="paymentYear" CLASS="norm" ID="paymentYear" dir="<%=dir_obj_var%>" style="width:60" onChange="populate(this.form,this.form.paymentMonth.selectedIndex);">
	      <% For counter = -2 to 2 %>
	        <OPTION VALUE="<%=Year(paymentDate)+counter%>" <% If (DatePart("yyyy", paymentDate) = Year(paymentDate)+counter) Then Response.Write "SELECTED"%>><%=Year(paymentDate)+counter%></OPTION>
	      <% Next %>
	      </SELECT>  	         
		</td>		
	</TR>
	</table>		
	</td>
	<td align="<%=align_var%>" class="textp" nowrap width="150">&nbsp;<span id="word9" name=word9><!--תאריך--><%=arrTitles(9)%></span>&nbsp;</td>
</tr>
<tr>	
	<td align="<%=align_var%>"><span id=word16 name=word16><!--&#8362;--><%=arrTitles(16)%></span>&nbsp;<input type=text dir="<%=dir_obj_var%>" name="payment" value="<%=payment%>" maxlength=8 size=9 onkeypress="GetNumbers()" ID="payment"></td>
	<td align="<%=align_var%>" class="textp" nowrap width="150"><span id="word10" name=word10><!--סכום--><%=arrTitles(10)%></span>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>"><textarea dir="<%=dir_obj_var%>" name="movement_details" style="width:300px;font-family:arial" ID="movement_details"><%=vfix(movement_details)%></textarea></td>
	<td align="<%=align_var%>" class="textp" nowrap width="150" valign=top><span id="word11" name=word11><!--תאור--><%=arrTitles(11)%></span>&nbsp;</td>
</tr>
<tr><td colspan="2" height=10></td></tr>

<%If trim(wizard_id) <> "" Then%>	
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>	
	
		<td width=50% align="<%=align_var%>" nowrap>
		<input style="width:90px" class="but_menu" type=button value="<< המשך" onclick="javascript:CheckFields(); return false;">
		</td>
		<td width=50 align="<%=align_var%>" nowrap></td>
		<td width=50% align="left" nowrap>
		<input type=button class="but_menu" value="חזור >>" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id%>.asp'" style="width:90px">
		</td>		
	</tr>
</table></td></tr>		
<%Else%>	
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>
<td width=50% align=center><input type=button class="but_menu" style="width:90px" onclick="document.location.href='movements.asp?type_id=<%=type_id%>';" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50 nowrap></td>
<td width=50% align=center><input type=button class="but_menu" style="width:90px" onclick="return CheckFields();" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
</table></td></tr>
<%End If%>
</form>
<tr><td colspan="2" height=10></td></tr>
</table>
</td></tr></table>
</body>
<%
set con = nothing
%>
</html>
