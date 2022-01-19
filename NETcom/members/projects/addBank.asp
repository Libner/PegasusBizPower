<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
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
	if ( document.frmMain.bank_name.value=='' || document.frmMain.balance.value == ''
	    || document.frmMain.credit.value == ''	)		   
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

function GetNumbers ()
{
	var ch=event.keyCode;
	event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
} 	
//-->
</script>  

<%

bank_ID=request.querystring("bank_ID")
errName = false

if request.form("bank_name")<>nil then 'after form filling
	 	 
	 bank_name = sFix(request.form("bank_name"))
	 bank_details  = sFix(trim(Request.Form("bank_details")))	
	 If Request.Form("credit") <> nil Then
		 credit = "'" & Request.Form("credit") & "'"
	 Else
		 credit = "NULL"
	 End If		 
	 If Request.Form("balance") <> nil Then
		 balance = "'" & Request.Form("balance") & "'"
	 Else
		 balance = "NULL"
	 End If	
	 
	 if bank_ID=nil or bank_ID="" then 'new record in DataBase
		    sqlStr = "select bank_name from banks where bank_name = '" & sFix(bank_name) & "' AND ORGANIZATION_ID = " & OrgID
		    set rs_name = con.GetRecordSet(sqlStr)
		    if  not rs_name.EOF then
		        errName=true
		    end if
		    set rs_name = nothing
		    
		    'Response.Write ("<BR>errName="& errName)		    
		    'Response.End
		    if errName=false then
				con.executeQuery("SET DATEFORMAT dmy")
				sqlStr = "insert into banks (bank_name,bank_details,credit,balance,organization_id) "
				sqlStr=sqlStr& " values ('"& bank_name &"','"& bank_details &"'," & credit & "," & balance & "," & OrgID & ")"
				'Response.Write sqlStr
				con.GetRecordSet(sqlStr)					
				If isNumeric(wizard_id) Then
					Response.Redirect ("../wizard/wizard_"&wizard_id&"_4.asp?wbanks=true") 				
				Else
					Response.Redirect ("banks.asp")          
				End If	          
			 elseif errName=true then
				 bank_name = "" %>
			     <SCRIPT LANGUAGE=javascript>
			     <!--
				   alert('שם בנק זה בשימוש, בחר שם בנק אחר')
				   history.back();
			      //-->
			     </SCRIPT>			 			 					 				
			  <%end if
		    
	 else			   
			sqlStr = "select bank_name from banks where bank_name = '" & sFix(bank_name) & "' and bank_id <> " & bank_id & " AND ORGANIZATION_ID = " & OrgID
		    set rs_name = con.GetRecordSet(sqlStr)
		    if not rs_name.EOF then
		        errName=true
		    end if
		    set rs_name = nothing
			
			'Response.Write status
			'Response.End
			if errName=false then
				con.executeQuery("SET DATEFORMAT dmy")
				sqlStr = "Update banks set bank_name = '" & bank_name &_
				"', bank_details = '" & bank_details &_
				"', credit = " & credit & ", balance = " & balance &_
				" where bank_ID=" & bank_ID
				con.GetRecordSet (sqlStr)
				If isNumeric(wizard_id) Then
					Response.Redirect ("../wizard/wizard_"&wizard_id&"_4.asp?wbanks=true") 				
				Else
					Response.Redirect ("banks.asp")          
				End If	    
			elseif errName=true then
			bank_name = ""%>
			<SCRIPT LANGUAGE=javascript>
				<!--
				alert('שם בנק זה בשימוש, בחר שם בנק אחר')
				history.back();
				//-->
			</SCRIPT>				 			 				
		<%end if
    end if        
end if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 41 Order By word_id"				
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
<%numOftab = 4%>
<%numOfLink = 7%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;</td></tr>	  		       	
   </table></td></tr>  
<tr><td width=100% align="<%=align_var%>">
</td></tr>
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="500">    
<%
if bank_ID<>nil and bank_ID<>"" then
  sqlStr = "select bank_name, bank_details, credit, balance from banks where bank_ID=" & bank_ID  
  ''Response.Write sqlStr
  set rs_banks = con.GetRecordSet(sqlStr)
	if not rs_banks.eof then		
		bank_name = rs_banks("bank_name")
		bank_details = rs_banks("bank_details")
		credit = rs_banks("credit")
		balance = rs_banks("balance")	
	end if
	set rs_banks = nothing
	
	sqlstr = "Select TOP 1 bank_id FROM movements WHERE bank_id = " & bank_id
	set rs_check = con.getRecordSet(sqlstr)	
	if not rs_check.eof then
		is_movements = true
	else
		is_movements = false
	end if
	set rs_check = nothing	
end if
'Response.Write company_id
%>
<FORM name="frmMain" ACTION="addbank.asp?bank_ID=<%=bank_ID%>" METHOD="post" onSubmit="return CheckFields()">
<tr>
	<td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="bank_name" value="<%=vfix(bank_name)%>" style="width:300px;font-family:arial"></td>
	<td align="<%=align_var%>" class="textp" nowrap width="100"><span id=word3 name=word3><!--שם חשבון--><%=arrTitles(3)%></span>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>"><textarea dir="<%=dir_obj_var%>" name="bank_details" style="width:300px;font-family:arial" ID="bank_details"><%=vfix(bank_details)%></textarea></td>
	<td align="<%=align_var%>" class="textp" nowrap width="100"><span id="word4" name=word4><!--תיאור חשבון--><%=arrTitles(4)%></span>&nbsp;</td>
</tr>
<tr>	
	<td align="<%=align_var%>"><span id=word7 name=word7>&#8362;</span>&nbsp;<input type=text dir="<%=dir_obj_var%>" name="balance" value="<%=balance%>" maxlength=8 size=8 <%if is_movements then%> disabled <%else%> onkeypress="GetNumbers()" <%end if%>>
	<%if is_movements then%>
	<input type=hidden dir="<%=dir_obj_var%>" name="balance" value="<%=balance%>" maxlength=8 size=8>
	<%end if%>
	</td>
	<td align="<%=align_var%>" class="textp" nowrap width="100"><span id="word5" name=word5><!--יתרה--><%=arrTitles(5)%></span>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>	
	<td align="<%=align_var%>"><span id="word8" name=word8>&#8362;</span>&nbsp;<input type=text dir="<%=dir_obj_var%>" name="credit" value="<%=credit%>" maxlength=8 size=8 onkeypress="GetNumbers()"></td>	
	<td align="<%=align_var%>" class="textp" nowrap width="100"><span id="word6" name=word6><!--מסגרת אשראי--><%=arrTitles(6)%></span>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>

<tr><td colspan="2" height=10></td></tr>
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>
<td width=50% align="<%=align_var%>"><input type=button class="but_menu" style="width:90px" onclick="document.location.href='banks.asp';" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
</table></td></tr>
</form>
<tr><td colspan="2" height=10></td></tr>
</table>
</td></tr></table>
</body>
<% set con = nothing %>
</html>